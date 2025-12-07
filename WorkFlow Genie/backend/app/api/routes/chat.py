"""
Chat API Routes - FINAL PRODUCTION VERSION
Perfect session management with context continuity
Handles conversation flow and Excel automation orchestration
"""

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session
from pydantic import BaseModel, Field
from typing import Optional, List, Dict
import json
import re
from loguru import logger
from datetime import datetime

from app.core.database import get_db
from app.core.models import Session as ChatSession, Message, Operation
from app.services.llm_service import llm_service
from app.services.mcp_service import mcp_service

router = APIRouter()

# ==================== REQUEST/RESPONSE MODELS ====================

class MessageRequest(BaseModel):
    """Request model for sending a message"""
    session_id: Optional[str] = Field(None, description="Session ID for conversation continuity. If not provided, a new session will be created.")
    message: str = Field(..., description="User's message in natural language", example="Update John Smith's salary to 62000")
    file_id: Optional[str] = Field(None, description="Excel file ID to operate on (UUID format)", example="3cf05b87-ce48-4dd7-b463-547d911ffb")
    sheet_name: str = Field(default="HR", description="Sheet name within the Excel file", example="HR")

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
        
        # ════════════════════════════════════════════════════════════════
        # STEP 5: DETECT INTENT (Excel Operation or General Chat)
        # ════════════════════════════════════════════════════════════════
        
        is_excel_operation = llm_service.detect_excel_intent(request_data.message)
        
        if is_excel_operation:
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
            
            response = llm_service.generate_response(
                messages=conversation_history,
                session_summary=chat_session.summary
            )
            
            assistant_message = Message(
                session_id=chat_session.id,
                role="assistant",
                content=response
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
                    [{"role": m.role, "content": m.content} for m in recent_messages]
                )
                chat_session.summary = summary
                logger.info(f"✅ Session summary updated: {summary[:100]}...")
            
            db.commit()
            
            return MessageResponse(
                session_id=chat_session.id,
                response=response,
                operations=[],
                context={"conversation": True}
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
        # ═══════════════════════════════════════════════════════════════
        # STEP 1: GATHER CONTEXT
        # ═══════════════════════════════════════════════════════════════
        
        # Get available files
        files = await mcp_service.list_files()
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
                column_headers = await mcp_service.get_column_headers(file_id, sheet_name)
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
        # STEP 3: CREATE EXECUTION PLAN WITH LLM
        # ═══════════════════════════════════════════════════════════════
        
        logger.info("🧠 Creating execution plan with LLM...")
        
        plan = llm_service.plan_excel_operations(
            user_intent=user_message,
            context={
                "file_id": file_id,
                "sheet_name": sheet_name,
                "session_summary": session_summary,
                "available_files": available_files,
                "recent_operations": recent_operations,
                "column_headers": column_headers
            }
        )
        
        logger.info(f"📝 Plan created with {len(plan['steps'])} step(s)")
        
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
                # Execute the MCP tool
                result = await mcp_service.execute_tool(
                    tool_name=step["tool"],
                    parameters=step["parameters"]
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
        
        response_text = llm_service.format_operation_results(
            steps=plan.get('steps', []),
            results=operation_results
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
                "sheet_name": sheet_name
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
                summary = llm_service.generate_session_summary(conversation_history, [])
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
                "total_operations": len(operation_results),
                "successful": len([op for op in operation_results if op["status"] == "completed"])
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
            context={"error": True}
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