"""
Chat API Routes - FINAL PRODUCTION VERSION
Perfect session management with context continuity
Handles conversation flow and Excel automation orchestration
"""

import asyncio

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session
from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any, Set
import json
import re
from loguru import logger
from datetime import datetime

from app.core.config import settings
from app.core.database import get_db
from app.core.models import Session as ChatSession, Message, Operation
from app.services.llm_service import llm_service
from app.services.claude_mcp_service import claude_mcp_service
from app.services.mcp_service import mcp_service
from app.services.rag_service import rag_service

router = APIRouter()


WRITE_GUARDRAIL_TOOLS: Set[str] = {
    "write_range",
    "update_cell",
    "apply_formula",
    "update_by_search",
    "smart_update",
    "add_row",
    "delete_row",
    "bulk_update",
    "bulk_update_all",
    "calculate_column",
    "assign_grades",
    "fill_column",
    "find_replace",
    "bulk_find_replace",
    "remove_duplicates",
    "fill_down",
    "json_to_excel",
    "unpivot_columns",
    "create_excel_table",
    "fill_formula_down",
    "insert_rows",
    "delete_rows_by_index",
    "insert_columns",
    "delete_columns",
    "rename_columns",
    "fill_missing_values",
    "standardize_text_case",
    "trim_whitespace",
    "batch_update",
}

# ==================== REQUEST/RESPONSE MODELS ====================

class MessageRequest(BaseModel):
    """Request model for sending a message"""
    session_id: Optional[str] = Field(None, description="Session ID for conversation continuity. If not provided, a new session will be created.")
    message: str = Field(..., description="User's message in natural language", example="Update John Smith's salary to 62000")
    file_id: Optional[str] = Field(None, description="Excel file ID to operate on (UUID format)", example="3cf05b87-ce48-4dd7-b463-547d911ffb")
    sheet_name: str = Field(default="HR", description="Sheet name within the Excel file", example="HR")
    provider: Optional[str] = Field(None, description="LLM provider override: openai or claude", example="claude")
    use_rag: Optional[bool] = Field(
        default=None,
        description="RAG mode for non-operation messages: true=force RAG, false=disable, null=auto",
    )
    rag_top_k: int = Field(
        default=4,
        ge=1,
        le=20,
        description="Number of chunks to retrieve when RAG is used",
    )

class MessageResponse(BaseModel):
    """Response model for message"""
    session_id: str = Field(..., description="Session ID for this conversation")
    response: str = Field(..., description="AI assistant's response")
    operations: List[Dict] = Field(default=[], description="List of operations that were executed")
    context: Optional[Dict] = Field(None, description="Additional context information")

# ==================== MAIN CHAT ENDPOINT ====================

@router.post("/message", 
    response_model=MessageResponse,
    summary="Send Message to AI Assistant",
    description="Send a natural language message for Excel automation. The AI will understand your intent and execute appropriate operations.",
    tags=["Chat"])
async def send_message(
    request_data: MessageRequest,
    req: Request,
    db: Session = Depends(get_db)
):
    """
    Main chat endpoint for Excel automation via natural language.
    
    **Features:**
    - Session management with context continuity
    - Natural language understanding
    - Automatic tool selection and execution
    - Multi-turn conversations
    - Context-aware responses
    
    **Example Requests:**
    - "Update John Smith's salary to 62000"
    - "Give everyone in Sales a 10% raise"
    - "Show me all employees earning less than 50k"
    - "Calculate average salary by department"
    - "Add new student Alice with marks 85, 90, 78"
    
    **Flow:**
    1. Session management (create or resume)
    2. Store user message
    3. Detect if Excel operation or general chat
    4. Execute operations if needed
    5. Generate response with context
    6. Update session summary periodically
    """
    
    try:
        # ════════════════════════════════════════════════════════════════
        # STEP 1: SESSION MANAGEMENT
        # ════════════════════════════════════════════════════════════════
        
        if request_data.session_id:
            # Resume existing session
            chat_session = db.query(ChatSession).filter(
                ChatSession.id == request_data.session_id
            ).first()
            
            if not chat_session:
                raise HTTPException(status_code=404, detail="Session not found. Please start a new session.")
            
            logger.info(f"📂 Resumed session: {chat_session.id}")
        else:
            # Create new session
            chat_session = ChatSession(is_active=True)
            db.add(chat_session)
            db.commit()
            db.refresh(chat_session)
            logger.info(f"🆕 Created new session: {chat_session.id}")
        
        # Update session timestamp
        chat_session.updated_at = datetime.utcnow()
        db.commit()
        
        # ════════════════════════════════════════════════════════════════
        # STEP 2: SAVE USER MESSAGE
        # ════════════════════════════════════════════════════════════════
        
        user_message = Message(
            session_id=chat_session.id,
            role="user",
            content=request_data.message
        )
        db.add(user_message)
        db.commit()
        
        logger.info(f"💬 User message: {request_data.message[:100]}...")
        
        # ════════════════════════════════════════════════════════════════
        # STEP 3: CHECK FOR SPECIAL COMMANDS
        # ════════════════════════════════════════════════════════════════
        
        if _is_forget_command(request_data.message):
            # Reset session
            chat_session.summary = None
            db.commit()
            
            reset_message = "I've cleared our conversation history. How can I help you with Excel automation?"
            
            assistant_message = Message(
                session_id=chat_session.id,
                role="assistant",
                content=reset_message
            )
            db.add(assistant_message)
            db.commit()
            
            return MessageResponse(
                session_id=chat_session.id,
                response=reset_message,
                operations=[],
                context={"reset": True}
            )
        
        # ════════════════════════════════════════════════════════════════
        # STEP 4: GET CONVERSATION HISTORY (For Context)
        # ════════════════════════════════════════════════════════════════
        
        recent_messages = db.query(Message).filter(
            Message.session_id == chat_session.id
        ).order_by(Message.timestamp.desc()).limit(20).all()
        
        conversation_history = [
            {"role": msg.role, "content": msg.content}
            for msg in reversed(recent_messages)
        ]
        
        logger.info(f"📚 Loaded {len(conversation_history)} messages from history")

        provider = _resolve_provider(request_data.provider)
        logger.info(f"🤖 Using provider: {provider}")
        
        # ════════════════════════════════════════════════════════════════
        # STEP 5: DETECT INTENT (Excel Operation or General Chat)
        # ════════════════════════════════════════════════════════════════
        
        is_excel_operation = llm_service.detect_excel_intent(request_data.message, request_data.file_id)
        
        use_rag = _should_use_rag(
            message=request_data.message,
            file_id=request_data.file_id,
            explicit_use_rag=request_data.use_rag,
            is_excel_operation=is_excel_operation,
        )

        if is_excel_operation and not use_rag:
            logger.info("🔧 Detected Excel operation intent")
            
            # ═══════════════════════════════════════════════════════════
            # STEP 6: EXTRACT OR VALIDATE FILE_ID
            # ═══════════════════════════════════════════════════════════
            
            file_id = request_data.file_id
            
            # Check if this is a CREATE operation (doesn't need file_id)
            message_lower = request_data.message.lower()
            is_create_operation = any(kw in message_lower for kw in ['create', 'new file', 'new excel'])
            
            # # Try to extract file_id from message if not provided (skip for create operations)
            # if not file_id and not is_create_operation:
            #     match = re.search(
            #         r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}',
            #         request_data.message.lower()
            #     )
            #     if match:
            #         file_id = match.group(0)
            #         logger.info(f"📎 Extracted file_id from message: {file_id}")
            # Auto-fetch file_id from session
            if not file_id and not is_create_operation:
                  from app.core.models import ExcelFile
                  excel_file = db.query(ExcelFile).filter(
                     ExcelFile.session_id == chat_session.id
                 ).first()
    
                  if excel_file:
                      file_id = excel_file.id
                      logger.info(f"📎 Auto-fetched file_id from session: {file_id}")
            
            # If still no file_id, check if there's a default from session
            if not file_id and chat_session.summary:
                # Try to extract file_id from session summary
                match = re.search(
                    r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}',
                    chat_session.summary
                )
                if match:
                    file_id = match.group(0)
                    logger.info(f"📎 Using file_id from session context: {file_id}")
            
            # If STILL no file_id AND not a create operation, ask user
            if not file_id and not is_create_operation:
                response_text = """I'd love to help with that Excel operation! However, I need to know which file to work with.

You can either:
1. Include the file ID in your message
2. List available files by saying "show me my files"
3. Create a new file by saying "create a new Excel file"

Which would you prefer?"""
                
                assistant_message = Message(
                    session_id=chat_session.id,
                    role="assistant",
                    content=response_text
                )
                db.add(assistant_message)
                db.commit()
                
                return MessageResponse(
                    session_id=chat_session.id,
                    response=response_text,
                    operations=[],
                    context={"needs_file_id": True}
                )
            
            # ═══════════════════════════════════════════════════════════
            # STEP 7: EXECUTE EXCEL AUTOMATION
            # ═══════════════════════════════════════════════════════════
            
            result = await _handle_excel_operation(
                session_id=chat_session.id,
                user_message=request_data.message,
                file_id=file_id,
                sheet_name=request_data.sheet_name,
                provider=provider,
                conversation_history=conversation_history,
                session_summary=chat_session.summary,
                db=db,
                ws_manager=req.app.state.ws_manager
            )
            
            return result
        
        else:
            # ═══════════════════════════════════════════════════════════
            # STEP 8: HANDLE GENERAL CONVERSATION
            # ═══════════════════════════════════════════════════════════
            
            logger.info("💭 Detected general conversation")

            response_context: Dict[str, object] = {
                "conversation": True,
                "provider": provider,
                "rag_used": False,
            }

            if use_rag:
                # Only apply sheet filter if client explicitly set it. The default "HR"
                # should not silently restrict retrieval across other sheets.
                requested_sheet = (
                    request_data.sheet_name
                    if "sheet_name" in request_data.model_fields_set
                    else None
                )

                try:
                    rag_result = rag_service.answer_query(
                        question=request_data.message,
                        top_k=request_data.rag_top_k,
                        file_id=request_data.file_id,
                        sheet_name=requested_sheet,
                        provider=provider,
                    )
                    response = rag_result["answer"]
                    response_context.update(
                        {
                            "rag_used": True,
                            "retrieved_chunks": rag_result.get("retrieved_chunks", 0),
                            "rag_sources": rag_result.get("sources", []),
                        }
                    )
                except ValueError as rag_exc:
                    logger.warning(f"RAG unavailable, falling back to chat response: {rag_exc}")
                    response = llm_service.generate_response(
                        messages=conversation_history,
                        session_summary=chat_session.summary,
                        provider=provider,
                    )
                    response_context["rag_fallback_reason"] = str(rag_exc)
            else:
                response = llm_service.generate_response(
                    messages=conversation_history,
                    session_summary=chat_session.summary,
                    provider=provider
                )
            
            assistant_message = Message(
                session_id=chat_session.id,
                role="assistant",
                content=response,
                meta_data=json.dumps(response_context)
            )
            db.add(assistant_message)
            
            # ═══════════════════════════════════════════════════════════
            # STEP 9: UPDATE SESSION SUMMARY (Periodically)
            # ═══════════════════════════════════════════════════════════
            
            message_count = db.query(Message).filter(
                Message.session_id == chat_session.id
            ).count()
            
            # Update summary every 10 messages for context continuity
            if message_count % 10 == 0:
                logger.info("📝 Updating session summary...")
                summary = llm_service.generate_session_summary(
                    [{"role": m.role, "content": m.content} for m in recent_messages],
                    provider=provider
                )
                chat_session.summary = summary
                logger.info(f"✅ Session summary updated: {summary[:100]}...")
            
            db.commit()
            
            return MessageResponse(
                session_id=chat_session.id,
                response=response,
                operations=[],
                context=response_context
            )
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"💥 Error in send_message: {e}")
        raise HTTPException(status_code=500, detail=str(e))

# ==================== CHAT HISTORY ENDPOINT ====================

@router.get("/history/{session_id}",
    summary="Get Chat History",
    description="Retrieve full conversation history for a session",
    tags=["Chat"])
async def get_history(session_id: str, db: Session = Depends(get_db)):
    """
    Get complete chat history for a session.
    
    **Returns:**
    - Session summary (context)
    - All messages in chronological order
    - Message metadata (timestamps, operations)
    """
    
    try:
        chat_session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
        
        if not chat_session:
            raise HTTPException(status_code=404, detail="Session not found")
        
        messages = db.query(Message).filter(
            Message.session_id == session_id
        ).order_by(Message.timestamp.asc()).all()
        
        return {
            "session_id": session_id,
            "summary": chat_session.summary,
            "created_at": chat_session.created_at.isoformat(),
            "updated_at": chat_session.updated_at.isoformat(),
            "message_count": len(messages),
            "messages": [
                {
                    "id": msg.id,
                    "role": msg.role,
                    "content": msg.content,
                    "timestamp": msg.timestamp.isoformat(),
                    "metadata": json.loads(msg.meta_data) if msg.meta_data else None
                }
                for msg in messages
            ]
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Get history error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

# ==================== EXCEL OPERATION HANDLER ====================

async def _handle_excel_operation(
    session_id: str,
    user_message: str,
    file_id: str,
    sheet_name: str,
    provider: str,
    conversation_history: List[Dict],
    session_summary: Optional[str],
    db: Session,
    ws_manager
):
    """
    Handle Excel automation workflow with proper orchestration.
    
    Flow:
    1. Get available files and recent operations (context)
    2. Plan operations using LLM
    3. Execute each step in the plan
    4. Log all operations
    5. Generate natural language response
    6. Update session summary
    """
    
    try:
        tool_service = claude_mcp_service if provider == "claude" else mcp_service

        # ═══════════════════════════════════════════════════════════════
        # STEP 1: GATHER CONTEXT
        # ═══════════════════════════════════════════════════════════════
        
        # Get available files
        files = await tool_service.list_files()
        available_files = [f"{f['filename']} ({f['file_id']})" for f in files]
        
        # Get recent operations for context
        recent_ops = db.query(Operation).filter(
            Operation.session_id == session_id
        ).order_by(Operation.timestamp.desc()).limit(5).all()
        
        recent_operations = [
            {
                "tool": op.tool_name,
                "status": op.status,
                "timestamp": op.timestamp.isoformat()
            }
            for op in recent_ops
        ]
        
        logger.info(f"📋 Context: {len(available_files)} files, {len(recent_operations)} recent operations")
        
        # Get column headers for context (if file_id provided)
        column_headers = []
        if file_id:
            try:
                column_headers = await tool_service.get_column_headers(file_id, sheet_name)
                logger.info(f"📊 Column headers: {column_headers}")
            except Exception as e:
                logger.warning(f"Could not get column headers: {e}")
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 2: BROADCAST PLANNING STATUS
        # ═══════════════════════════════════════════════════════════════
        
        await ws_manager.broadcast({
            "type": "planning",
            "session_id": session_id,
            "message": "Analyzing your request and planning operations..."
        })

        # ═══════════════════════════════════════════════════════════════
        # STEP 2.5: OPTIONAL RAG EVIDENCE FOR WRITE/UPDATE REQUESTS
        # ═══════════════════════════════════════════════════════════════

        write_request_detected = _is_write_or_update_request(user_message)
        write_confirmation_requested = _contains_write_confirmation(user_message)
        pending_write_plan: Optional[Dict[str, Any]] = None

        if write_confirmation_requested:
            pending_write_plan = _get_pending_write_confirmation(db=db, session_id=session_id)

            if pending_write_plan:
                file_id = pending_write_plan.get("file_id") or file_id
                sheet_name = pending_write_plan.get("sheet_name") or sheet_name

                confirmed_provider = pending_write_plan.get("provider")
                if isinstance(confirmed_provider, str) and confirmed_provider in {"openai", "claude"}:
                    provider = confirmed_provider

                logger.info(
                    "✅ Reusing previously confirmed write plan "
                    f"with {len(pending_write_plan.get('planned_steps', []))} step(s)"
                )
            else:
                response_text = (
                    "I couldn't find a pending write plan to confirm. "
                    "Please resend your full update request so I can generate a fresh plan."
                )

                assistant_message = Message(
                    session_id=session_id,
                    role="assistant",
                    content=response_text,
                    meta_data=json.dumps(
                        {
                            "confirmation_missing": True,
                            "file_id": file_id,
                            "sheet_name": sheet_name,
                            "provider": provider,
                        }
                    ),
                )
                db.add(assistant_message)
                db.commit()

                await ws_manager.broadcast(
                    {
                        "type": "operations_complete",
                        "session_id": session_id,
                        "total_steps": 0,
                        "successful_steps": 0,
                    }
                )

                return MessageResponse(
                    session_id=session_id,
                    response=response_text,
                    operations=[],
                    context={
                        "file_id": file_id,
                        "sheet_name": sheet_name,
                        "provider": provider,
                        "confirmation_missing": True,
                    },
                )

        write_rag_evidence: Optional[Dict[str, Any]] = None
        if pending_write_plan and isinstance(pending_write_plan.get("write_rag_evidence"), dict):
            write_rag_evidence = pending_write_plan.get("write_rag_evidence")
        elif file_id and write_request_detected:
            if settings.RAG_ENABLED and settings.RAG_WRITE_ENABLED:
                try:
                    write_rag_evidence = rag_service.retrieve_for_write(
                        question=user_message,
                        file_id=file_id,
                        sheet_name=sheet_name,
                        top_k=settings.RAG_WRITE_TOP_K,
                    )

                    logger.info(
                        "🧭 Write RAG evidence: "
                        f"top_score={write_rag_evidence.get('top_score')} "
                        f"confident_hits={write_rag_evidence.get('confident_hits')} "
                        f"estimated_rows={write_rag_evidence.get('estimated_rows')}"
                    )
                except Exception as exc:
                    logger.warning(f"Write RAG retrieval failed, enabling conservative gate: {exc}")
                    write_rag_evidence = {
                        "enabled": False,
                        "reason": str(exc),
                        "sources": [],
                        "summary": "",
                        "warnings": ["Write evidence could not be retrieved."],
                        "top_score": None,
                        "confident_hits": 0,
                        "estimated_rows": 0,
                        "built_at": None,
                        "threshold": settings.RAG_WRITE_SCORE_THRESHOLD,
                    }
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 3: CREATE EXECUTION PLAN WITH LLM
        # ═══════════════════════════════════════════════════════════════
        
        if pending_write_plan:
            plan = {
                "steps": pending_write_plan.get("planned_steps", []),
            }
            logger.info(f"📝 Using confirmed pending plan with {len(plan['steps'])} step(s)")
        else:
            logger.info("🧠 Creating execution plan with LLM...")

            plan = llm_service.plan_excel_operations(
                user_intent=user_message,
                context={
                    "file_id": file_id,
                    "sheet_name": sheet_name,
                    "session_summary": session_summary,
                    "available_files": available_files,
                    "recent_operations": recent_operations,
                    "column_headers": column_headers,
                    "write_rag_evidence": write_rag_evidence,
                },
                provider=provider
            )

            logger.info(f"📝 Plan created with {len(plan['steps'])} step(s)")

        mutating_steps = _get_mutating_steps(plan.get("steps", []))
        write_risk = {"requires_confirmation": False, "reasons": []}
        if write_request_detected and not pending_write_plan:
            write_risk = _evaluate_write_risk(
                write_rag_evidence=write_rag_evidence,
                mutating_steps=mutating_steps,
            )

        if (
            write_request_detected
            and
            mutating_steps
            and write_risk.get("requires_confirmation")
            and not _contains_write_confirmation(user_message)
        ):
            logger.warning(
                f"🛡️ Write plan requires confirmation. Reasons: {write_risk.get('reasons', [])}"
            )

            response_text = _build_write_confirmation_response(
                plan_steps=mutating_steps,
                reasons=write_risk.get("reasons", []),
                evidence=write_rag_evidence,
            )

            assistant_message = Message(
                session_id=session_id,
                role="assistant",
                content=response_text,
                meta_data=json.dumps(
                    {
                        "confirmation_required": True,
                        "risk_reasons": write_risk.get("reasons", []),
                        "planned_steps": mutating_steps,
                        "write_rag_evidence": write_rag_evidence,
                        "file_id": file_id,
                        "sheet_name": sheet_name,
                        "provider": provider,
                    }
                ),
            )
            db.add(assistant_message)
            db.commit()

            await ws_manager.broadcast(
                {
                    "type": "operations_complete",
                    "session_id": session_id,
                    "total_steps": len(plan.get("steps", [])),
                    "successful_steps": 0,
                }
            )

            return MessageResponse(
                session_id=session_id,
                response=response_text,
                operations=[],
                context={
                    "file_id": file_id,
                    "sheet_name": sheet_name,
                    "provider": provider,
                    "confirmation_required": True,
                    "risk_reasons": write_risk.get("reasons", []),
                },
            )
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 4: EXECUTE EACH STEP IN THE PLAN
        # ═══════════════════════════════════════════════════════════════
        
        operation_results = []
        
        for step in plan["steps"]:
            logger.info(f"⚙️ Executing step {step['step']}: {step['tool']}")
            
            # Broadcast step start
            await ws_manager.broadcast({
                "type": "step_start",
                "session_id": session_id,
                "step": step["step"],
                "description": step["description"]
            })
            
            # Create operation log (in progress)
            operation = Operation(
                session_id=session_id,
                tool_name=step["tool"],
                input_params=json.dumps(step["parameters"]),
                status="in_progress"
            )
            db.add(operation)
            db.commit()
            db.refresh(operation)
            
            try:
                # Execute the MCP tool (pass session_id and db for tools that need database access)
                result = await tool_service.execute_tool(
                    tool_name=step["tool"],
                    parameters=step["parameters"],
                    session_id=session_id,
                    db=db
                )
                
                # Update operation log
                operation.status = "completed" if result["success"] else "failed"
                operation.output_result = json.dumps(result)
                operation.duration = result.get("duration")
                operation.error_message = result.get("error")
                db.commit()
                
                # Store result
                operation_results.append({
                    "step": step["step"],
                    "description": step["description"],
                    "tool": step["tool"],
                    "status": "completed" if result["success"] else "failed",
                    "result": result.get("data"),
                    "error": result.get("error")
                })

                if result["success"] and step.get("tool") == "copy_sheet":
                    step_params = step.get("parameters", {})
                    result_data = result.get("data") if isinstance(result.get("data"), dict) else {}

                    requested_sheet = step_params.get("target_name") if isinstance(step_params, dict) else None
                    actual_sheet = result_data.get("new_sheet")

                    if requested_sheet and actual_sheet and requested_sheet != actual_sheet:
                        for planned_step in plan.get("steps", []):
                            if planned_step.get("step", 0) <= step.get("step", 0):
                                continue

                            planned_params = planned_step.get("parameters")
                            if not isinstance(planned_params, dict):
                                continue

                            if planned_params.get("sheet_name") == requested_sheet:
                                planned_params["sheet_name"] = actual_sheet

                            if planned_params.get("source_sheet") == requested_sheet:
                                planned_params["source_sheet"] = actual_sheet

                        logger.info(
                            f"Updated downstream steps to use copied sheet '{actual_sheet}' "
                            f"instead of requested '{requested_sheet}'"
                        )

                if result["success"] and step.get("tool") == "pivot_table":
                    step_params = step.get("parameters", {})
                    result_data = result.get("data") if isinstance(result.get("data"), dict) else {}

                    source_sheet = step_params.get("sheet_name") if isinstance(step_params, dict) else None
                    actual_pivot_sheet = result_data.get("pivot_sheet") or result_data.get("sheet_name")

                    if isinstance(actual_pivot_sheet, str) and actual_pivot_sheet.strip():
                        pivot_aliases = {
                            "PivotTable",
                            "Pivot_Table",
                            "Pivot Sheet",
                            "pivottable",
                            "pivot_table",
                            "pivot sheet",
                            actual_pivot_sheet,
                        }

                        if isinstance(source_sheet, str) and source_sheet.strip():
                            pivot_aliases.add(f"Pivot_{source_sheet}"[:31])

                        normalized_aliases = {
                            alias.strip().lower()
                            for alias in pivot_aliases
                            if isinstance(alias, str) and alias.strip()
                        }

                        updated_refs = 0

                        for planned_step in plan.get("steps", []):
                            if planned_step.get("step", 0) <= step.get("step", 0):
                                continue

                            planned_params = planned_step.get("parameters")
                            if not isinstance(planned_params, dict):
                                continue

                            for key in ("sheet_name", "source_sheet", "target_sheet"):
                                current_value = planned_params.get(key)
                                if not isinstance(current_value, str):
                                    continue

                                if current_value.strip().lower() in normalized_aliases and current_value != actual_pivot_sheet:
                                    planned_params[key] = actual_pivot_sheet
                                    updated_refs += 1

                        if updated_refs > 0:
                            logger.info(
                                f"Updated {updated_refs} downstream sheet reference(s) to actual pivot sheet "
                                f"'{actual_pivot_sheet}'"
                            )
                
                # Broadcast completion
                await ws_manager.broadcast({
                    "type": "step_complete",
                    "session_id": session_id,
                    "step": step["step"],
                    "status": "completed" if result["success"] else "failed",
                    "result": result.get("data")
                })
                
                logger.info(f"✅ Step {step['step']} completed successfully")
                
                # Stop on failure
                if not result["success"]:
                    logger.warning(f"⚠️ Step {step['step']} failed, stopping execution")
                    break
            
            except Exception as e:
                logger.error(f"💥 Step {step['step']} execution error: {e}")
                
                # Update operation as failed
                operation.status = "failed"
                operation.error_message = str(e)
                db.commit()
                
                operation_results.append({
                    "step": step["step"],
                    "description": step["description"],
                    "tool": step["tool"],
                    "status": "failed",
                    "error": str(e)
                })
                
                await ws_manager.broadcast({
                    "type": "step_failed",
                    "session_id": session_id,
                    "step": step["step"],
                    "error": str(e)
                })
                
                break
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 5: GENERATE NATURAL LANGUAGE RESPONSE
        # ═══════════════════════════════════════════════════════════════
        
        logger.info("💬 Formatting response...")

        successful_mutating_ops = [
            op
            for op in operation_results
            if op.get("status") == "completed"
            and str(op.get("tool", "")).strip().lower() in WRITE_GUARDRAIL_TOOLS
        ]

        rag_refresh_scheduled = False
        if successful_mutating_ops and settings.RAG_ENABLED:
            rag_refresh_scheduled = _schedule_rag_refresh_async("write-operation")

        executed_step_numbers = {op.get("step") for op in operation_results}
        executed_steps = [
            step for step in plan.get("steps", [])
            if step.get("step") in executed_step_numbers
        ]
        
        response_text = llm_service.format_operation_results(
            steps=executed_steps,
            results=operation_results,
            provider=provider
        )
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 6: SAVE ASSISTANT RESPONSE
        # ═══════════════════════════════════════════════════════════════
        
        assistant_message = Message(
            session_id=session_id,
            role="assistant",
            content=response_text,
            meta_data=json.dumps({
                "operations": operation_results,
                "file_id": file_id,
                "sheet_name": sheet_name,
                "provider": provider,
                "write_rag_evidence": write_rag_evidence,
                "rag_refresh_scheduled": rag_refresh_scheduled,
            })
        )
        db.add(assistant_message)
        db.commit()
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 7: UPDATE SESSION SUMMARY (Every 5 Excel operations)
        # ═══════════════════════════════════════════════════════════════
        
        excel_op_count = db.query(Operation).filter(
            Operation.session_id == session_id,
            Operation.status == "completed"
        ).count()
        
        if excel_op_count % 5 == 0:
            logger.info("📝 Updating session summary after 5 operations...")
            chat_session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
            if chat_session:
                summary = llm_service.generate_session_summary(conversation_history, [], provider=provider)
                chat_session.summary = summary
                db.commit()
                logger.info(f"✅ Session summary updated: {summary[:100]}...")
        
        # ═══════════════════════════════════════════════════════════════
        # STEP 8: BROADCAST COMPLETION
        # ═══════════════════════════════════════════════════════════════
        
        await ws_manager.broadcast({
            "type": "operations_complete",
            "session_id": session_id,
            "total_steps": len(plan["steps"]),
            "successful_steps": len([op for op in operation_results if op["status"] == "completed"])
        })
        
        logger.info("🎉 Excel operation workflow completed")
        
        return MessageResponse(
            session_id=session_id,
            response=response_text,
            operations=operation_results,
            context={
                "file_id": file_id,
                "sheet_name": sheet_name,
                "provider": provider,
                "total_operations": len(operation_results),
                "successful": len([op for op in operation_results if op["status"] == "completed"]),
                "rag_refresh_scheduled": rag_refresh_scheduled,
            }
        )
    
    except Exception as e:
        logger.error(f"💥 Excel operation workflow error: {e}")
        
        error_message = f"I encountered an error while processing your request: {str(e)}\n\nPlease try again or rephrase your request."
        
        assistant_message = Message(
            session_id=session_id,
            role="assistant",
            content=error_message
        )
        db.add(assistant_message)
        db.commit()
        
        return MessageResponse(
            session_id=session_id,
            response=error_message,
            operations=[],
            context={"error": True, "provider": provider}
        )

# ==================== HELPER FUNCTIONS ====================

def _is_forget_command(message: str) -> bool:
    """Check if message is a session reset command"""
    
    forget_patterns = [
        'forget everything',
        'forget all',
        'reset conversation',
        'clear history',
        'start over',
        'new conversation',
        'reset session'
    ]
    
    message_lower = message.lower().strip()
    return any(pattern in message_lower for pattern in forget_patterns)


def _resolve_provider(provider: Optional[str]) -> str:
    """Resolve provider from request override or app default."""

    resolved_provider = (provider or settings.LLM_PROVIDER or "openai").strip().lower()
    if resolved_provider not in {"openai", "claude"}:
        raise HTTPException(status_code=400, detail="Invalid provider. Use 'openai' or 'claude'.")
    return resolved_provider


def _should_use_rag(
    message: str,
    file_id: Optional[str],
    explicit_use_rag: Optional[bool],
    is_excel_operation: bool,
) -> bool:
    """Determine if non-operation message should use RAG retrieval."""

    if not settings.RAG_ENABLED:
        return False

    # Excel operation requests are handled by the tool planner/executor path.
    if is_excel_operation:
        return False

    if explicit_use_rag is not None:
        return explicit_use_rag

    message_lower = message.lower()
    small_talk_markers = ["hello", "hi", "hey", "thanks", "thank you", "how are you"]
    if any(marker in message_lower for marker in small_talk_markers):
        return False

    rag_markers = [
        "sheet",
        "row",
        "column",
        "table",
        "excel",
        "file",
        "in the data",
        "from the data",
        "what does",
        "how many",
        "show me",
    ]

    return bool(file_id) or message_lower.endswith("?") or any(
        marker in message_lower for marker in rag_markers
    )


def _is_write_or_update_request(message: str) -> bool:
    """Detect whether request intent likely mutates workbook data."""

    message_lower = message.lower()
    write_markers = [
        "update",
        "modify",
        "change",
        "set",
        "write",
        "replace",
        "delete",
        "remove",
        "insert",
        "add row",
        "add column",
        "rename",
        "fill",
        "apply formula",
        "create chart",
    ]

    return any(marker in message_lower for marker in write_markers)


def _contains_write_confirmation(message: str) -> bool:
    """Allow users to explicitly proceed when a write plan is risky."""

    message_lower = message.lower()
    confirmation_markers = [
        "confirm execute write",
        "confirm write",
        "proceed anyway",
        "run anyway",
        "force update",
        "yes proceed",
    ]

    return any(marker in message_lower for marker in confirmation_markers)


def _get_pending_write_confirmation(db: Session, session_id: str) -> Optional[Dict[str, Any]]:
    """Get the latest pending confirmation payload, if present."""

    last_assistant_message = db.query(Message).filter(
        Message.session_id == session_id,
        Message.role == "assistant",
    ).order_by(Message.timestamp.desc()).first()

    if not last_assistant_message or not last_assistant_message.meta_data:
        return None

    try:
        meta_data = json.loads(last_assistant_message.meta_data)
    except Exception:
        return None

    if not isinstance(meta_data, dict) or not meta_data.get("confirmation_required"):
        return None

    planned_steps_raw = meta_data.get("planned_steps")
    if not isinstance(planned_steps_raw, list) or not planned_steps_raw:
        return None

    planned_steps: List[Dict[str, Any]] = []
    for index, step in enumerate(planned_steps_raw, start=1):
        if not isinstance(step, dict):
            continue

        normalized_step = dict(step)
        if "step" not in normalized_step:
            normalized_step["step"] = index
        planned_steps.append(normalized_step)

    if not planned_steps:
        return None

    return {
        "planned_steps": planned_steps,
        "file_id": meta_data.get("file_id"),
        "sheet_name": meta_data.get("sheet_name"),
        "provider": meta_data.get("provider"),
        "write_rag_evidence": meta_data.get("write_rag_evidence"),
    }


def _get_mutating_steps(plan_steps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Extract mutating steps from planned operations."""

    mutating: List[Dict[str, Any]] = []
    for step in plan_steps:
        tool_name = str(step.get("tool", "")).strip().lower()
        if tool_name in WRITE_GUARDRAIL_TOOLS:
            mutating.append(step)
    return mutating


def _evaluate_write_risk(
    write_rag_evidence: Optional[Dict[str, Any]],
    mutating_steps: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """Determine whether write plan should require explicit confirmation."""

    reasons: List[str] = []

    if not mutating_steps:
        return {"requires_confirmation": False, "reasons": reasons}

    if not settings.RAG_WRITE_ENABLED:
        return {"requires_confirmation": False, "reasons": reasons}

    if not isinstance(write_rag_evidence, dict):
        reasons.append("Write evidence is unavailable for this request.")
        return {"requires_confirmation": True, "reasons": reasons}

    if not write_rag_evidence.get("enabled", False):
        reasons.append("Write evidence retrieval is disabled or unavailable.")

    for warning in write_rag_evidence.get("warnings", []) or []:
        if isinstance(warning, str) and warning.strip():
            reasons.append(warning.strip())

    top_score_raw = write_rag_evidence.get("top_score")
    threshold = float(write_rag_evidence.get("threshold") or settings.RAG_WRITE_SCORE_THRESHOLD)

    if top_score_raw is None:
        reasons.append("No retrieval confidence score is available.")
    else:
        try:
            top_score = float(top_score_raw)
            if top_score < threshold:
                reasons.append(
                    f"Top retrieval confidence ({top_score:.3f}) is below threshold ({threshold:.3f})."
                )
        except (TypeError, ValueError):
            reasons.append("Retrieval confidence score is invalid.")

    confident_hits = int(write_rag_evidence.get("confident_hits") or 0)
    if confident_hits < settings.RAG_WRITE_MIN_CONFIDENT_HITS:
        reasons.append(
            f"Only {confident_hits} confident hit(s); minimum is {settings.RAG_WRITE_MIN_CONFIDENT_HITS}."
        )

    estimated_rows = int(write_rag_evidence.get("estimated_rows") or 0)
    high_impact_tools = {"bulk_update_all", "delete_rows_by_index", "delete_columns", "write_range"}
    contains_high_impact = any(
        str(step.get("tool", "")).strip().lower() in high_impact_tools
        for step in mutating_steps
    )
    if contains_high_impact and estimated_rows > settings.RAG_WRITE_MAX_AFFECTED_ROWS_WARNING:
        reasons.append(
            f"Estimated affected rows ({estimated_rows}) exceed warning threshold "
            f"({settings.RAG_WRITE_MAX_AFFECTED_ROWS_WARNING})."
        )

    return {
        "requires_confirmation": bool(reasons),
        "reasons": reasons,
    }


def _build_write_confirmation_response(
    plan_steps: List[Dict[str, Any]],
    reasons: List[str],
    evidence: Optional[Dict[str, Any]],
) -> str:
    """Create a confirmation prompt for potentially risky write plans."""

    contains_global_step = any(_is_global_write_step(step) for step in plan_steps)

    lines: List[str] = [
        "I prepared a write/update plan. Confirmation is required before execution because it can modify workbook data.",
        "",
        "Planned write steps:",
    ]

    for step in plan_steps:
        impact_hint = _describe_step_impact(step)
        impact_suffix = f" ({impact_hint})" if impact_hint else ""
        lines.append(
            f"- Step {step.get('step')}: {step.get('tool')} - {step.get('description', 'No description')}{impact_suffix}"
        )

    impact_lines = _build_write_impact_summary(plan_steps=plan_steps, evidence=evidence)
    if impact_lines:
        lines.extend(["", "Execution impact summary:"])
        lines.extend([f"- {item}" for item in impact_lines])

    if reasons:
        lines.extend(["", "Risk signals:"])
        for reason in reasons:
            lines.append(f"- {reason}")

    if contains_global_step:
        lines.extend(
            [
                "",
                "Why this warning can appear for broad updates:",
                "- Retrieval confidence is semantic and row-snippet based.",
                "- Broad requests like 'update all dates' may score low even when the instruction is precise.",
                "- This is a safety confirmation, not a detected execution failure.",
            ]
        )

    if isinstance(evidence, dict):
        summary = str(evidence.get("summary") or "").strip()
        if summary:
            lines.extend(["", "Retrieved evidence preview:", summary])

    lines.extend(
        [
            "",
            "Reply with 'confirm execute write' to run this plan as-is, or refine your request with exact column/filter details.",
        ]
    )

    return "\n".join(lines)


def _is_global_write_step(step: Dict[str, Any]) -> bool:
    """Detect whether a planned write step likely affects many/all rows."""

    tool_name = str(step.get("tool", "")).strip().lower()
    if tool_name in {"bulk_update_all", "fill_column", "write_range", "batch_update"}:
        return True

    if tool_name == "bulk_update":
        params = step.get("parameters", {}) if isinstance(step.get("parameters"), dict) else {}
        filter_column = params.get("filter_column")
        filter_value = params.get("filter_value")
        if not filter_column or filter_value in (None, ""):
            return True

    return False


def _describe_step_impact(step: Dict[str, Any]) -> str:
    """Return a concise, human-readable impact hint for a planned write step."""

    params = step.get("parameters", {}) if isinstance(step.get("parameters"), dict) else {}
    tool_name = str(step.get("tool", "")).strip().lower()

    if tool_name == "bulk_update_all":
        column = params.get("column")
        value = params.get("value")
        if column and value is not None:
            return f"all rows: set {column} = {value}"
        if column:
            return f"all rows in column {column}"
        return "all rows in the target sheet"

    if tool_name == "bulk_update":
        column = params.get("column")
        value = params.get("value")
        filter_column = params.get("filter_column")
        filter_value = params.get("filter_value")
        if filter_column and filter_value not in (None, ""):
            return f"rows where {filter_column} = {filter_value}: set {column} = {value}"
        return "multi-row update"

    if tool_name == "update_by_search":
        search_column = params.get("search_column")
        search_value = params.get("search_value")
        update_column = params.get("update_column")
        return f"find {search_column} = {search_value}, update {update_column}"

    if tool_name == "fill_column":
        column_name = params.get("column_name") or params.get("column")
        fill_type = params.get("fill_type")
        if column_name and fill_type:
            return f"fill column {column_name} ({fill_type})"
        if column_name:
            return f"fill column {column_name}"

    return ""


def _build_write_impact_summary(
    plan_steps: List[Dict[str, Any]],
    evidence: Optional[Dict[str, Any]],
) -> List[str]:
    """Build user-facing impact lines for write confirmation responses."""

    summary_lines: List[str] = []

    if any(_is_global_write_step(step) for step in plan_steps):
        summary_lines.append("Scope appears broad (likely many or all rows affected).")
    else:
        summary_lines.append("Scope appears targeted (specific rows/conditions).")

    if isinstance(evidence, dict):
        estimated_rows = evidence.get("estimated_rows")
        if estimated_rows is not None:
            try:
                summary_lines.append(f"Estimated affected rows from retrieval evidence: {int(estimated_rows)}.")
            except (TypeError, ValueError):
                pass

        top_score = evidence.get("top_score")
        threshold = evidence.get("threshold")
        try:
            if top_score is not None and threshold is not None:
                summary_lines.append(
                    f"Top retrieval confidence: {float(top_score):.3f} (threshold: {float(threshold):.3f})."
                )
        except (TypeError, ValueError):
            pass

    return summary_lines


def _schedule_rag_refresh_async(reason: str) -> bool:
    """Schedule non-blocking RAG refresh after successful write operations."""

    if not settings.RAG_ENABLED:
        return False

    async def _runner() -> None:
        try:
            result = await asyncio.to_thread(rag_service.refresh_index)
            logger.info(
                f"RAG index synced after {reason} "
                f"({result.get('indexed_chunks', 0)} chunks across {result.get('indexed_files', 0)} files)"
            )
        except Exception as exc:
            logger.error(f"Failed to sync RAG index after {reason}: {exc}")

    try:
        asyncio.create_task(_runner())
        return True
    except Exception as exc:
        logger.error(f"Failed to schedule RAG refresh task after {reason}: {exc}")
        return False
