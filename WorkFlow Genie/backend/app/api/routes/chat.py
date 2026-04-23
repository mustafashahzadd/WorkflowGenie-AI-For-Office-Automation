"""
Chat API Routes - Clean Architecture
Single-responsibility helpers. No scattered if/else. No jugaad.

Context Resolution Flow:
  1. sanitize -> strip junk values from request params
  2. resolve_file_id -> request -> session DB -> session summary (one path)
  3. resolve_sheet_name -> request -> auto-detect from file (one path)
  4. inject into every step -> by TOOL NAME, not message keywords
"""

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session
from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Tuple
import json
import re
import openpyxl
from loguru import logger
from datetime import datetime

from app.core.config import settings
from app.core.database import get_db
from app.core.models import Session as ChatSession, Message, Operation, ExcelFile, User
from app.services.llm_service import llm_service
from app.services.claude_mcp_service import claude_mcp_service
from app.services.mcp_service import mcp_service

router = APIRouter()

# ==================== REQUEST / RESPONSE MODELS ====================


class MessageRequest(BaseModel):
    """Request model for sending a message"""
    session_id: Optional[str] = Field(
        None,
        description="Session ID for conversation continuity. If not provided, a new session will be created.",
        example=None,
    )
    message: str = Field(
        ...,
        description="User message in natural language",
        example="create a new file called StudentMarks.xlsx with headers Name, Math, Science, Total and fill 10 rows with random marks",
    )
    file_id: Optional[str] = Field(
        None,
        description="Excel file ID to operate on. Leave empty to create new files or let the system auto-detect.",
        example=None,
    )
    sheet_name: Optional[str] = Field(
        None,
        description="Sheet name within the Excel file. Leave empty to auto-detect or use default.",
        example=None,
    )
    provider: Optional[str] = Field(None, description="LLM provider: 'claude' (default) or 'openai'", example="claude")
    model: Optional[str] = Field(None, description="Claude model: 'sonnet' (default) or 'opus'", example="sonnet")


class MessageResponse(BaseModel):
    """Response model for message"""
    session_id: str = Field(..., description="Session ID for this conversation")
    response: str = Field(..., description="AI assistant response")
    operations: List[Dict] = Field(default=[], description="List of operations that were executed")
    context: Optional[Dict] = Field(None, description="Additional context information")


# ==================== CONTEXT RESOLUTION HELPERS ====================
# Each function does ONE thing. No scattered if/else.

# Values that frontends/Swagger may send as "empty"
_JUNK_VALUES = frozenset({
    "string", "null", "none", "", "undefined", "example",
    "3cf05b87-ce48-4dd7-b463-547d911ffb",
})

# Tools that CREATE a brand-new workbook - they don't need an existing file_id
_CREATE_TOOLS = frozenset({"create_workbook"})

# Tools that never need sheet_name (they operate on files, not sheets)
_NO_SHEET_TOOLS = frozenset({"create_workbook", "list_files", "get_file_info", "delete_file"})


def _sanitize(value: Optional[str]) -> Optional[str]:
    """Return None if the value is junk, otherwise the stripped value."""
    if value is None:
        return None
    cleaned = value.strip()
    if cleaned.lower() in _JUNK_VALUES:
        return None
    return cleaned


def _resolve_file_id(
    request_file_id: Optional[str],
    session_id: str,
    session_summary: Optional[str],
    db: Session,
) -> Optional[str]:
    """
    Single path to resolve file_id:
      1. From request (already sanitized)
      2. From session's most recent ExcelFile in DB
      3. From session summary (regex fallback)
    """
    # 1 - explicit from request
    if request_file_id:
        return request_file_id

    # 2 - most recent file for this session
    excel_file = (
        db.query(ExcelFile)
        .filter(ExcelFile.session_id == session_id)
        .order_by(ExcelFile.uploaded_at.desc())
        .first()
    )
    if excel_file:
        logger.info(f"Auto-fetched file_id from session DB: {excel_file.id}")
        return excel_file.id

    # 3 - regex from session summary (last resort)
    if session_summary:
        match = re.search(
            r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
            session_summary,
        )
        if match:
            logger.info(f"Extracted file_id from session summary: {match.group(0)}")
            return match.group(0)

    return None


def _resolve_sheet_name(
    request_sheet_name: Optional[str],
    file_id: Optional[str],
    tool_service,
) -> Optional[str]:
    """
    Single path to resolve sheet_name:
      1. From request (already sanitized)
      2. Auto-detect first sheet from the file on disk
    """
    if request_sheet_name:
        return request_sheet_name

    if not file_id:
        return None

    try:
        filepath = tool_service._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath, read_only=True)
        name = wb.sheetnames[0]
        wb.close()
        logger.info(f"Auto-detected sheet_name from file: '{name}'")
        return name
    except Exception as e:
        logger.warning(f"Could not auto-detect sheet_name: {e}")
        return None


def _resolve_execution_context(
    request_data: MessageRequest,
    session: ChatSession,
    db: Session,
    tool_service,
) -> Tuple[Optional[str], Optional[str]]:
    """
    Master context resolver - called ONCE before the execution pipeline.
    Returns (file_id, sheet_name) fully resolved and ready to inject.
    """
    file_id = _sanitize(request_data.file_id)
    sheet_name = _sanitize(request_data.sheet_name)

    file_id = _resolve_file_id(file_id, session.id, session.summary, db)
    sheet_name = _resolve_sheet_name(sheet_name, file_id, tool_service)

    return file_id, sheet_name


_UUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)


def _is_valid_file_id(value: Optional[str]) -> bool:
    """Check if a value looks like a real UUID file_id (not a placeholder)."""
    if not value:
        return False
    return bool(_UUID_RE.match(value.strip()))


def _inject_step_context(
    step: Dict,
    file_id: Optional[str],
    sheet_name: Optional[str],
) -> Dict:
    """
    Inject file_id and sheet_name into a single step's parameters.
    Decision is based on the TOOL NAME, not message keywords.
    For file_id: only keep existing value if it's a valid UUID.
    """
    tool_name = step.get("tool", "")
    params = step.get("parameters", {})

    # file_id - skip only for tools that create new files
    if tool_name not in _CREATE_TOOLS and file_id:
        current = str(params.get("file_id", "")).strip()
        if not _is_valid_file_id(current):
            params["file_id"] = file_id

    # sheet_name - skip for tools that don't operate on sheets
    if tool_name not in _NO_SHEET_TOOLS and sheet_name:
        current = _sanitize(str(params.get("sheet_name", "")))
        if not current:
            params["sheet_name"] = sheet_name

    step["parameters"] = params
    return step


def _needs_file_id(plan_steps: List[Dict]) -> bool:
    """
    Check whether any step in the plan requires a pre-existing file_id.
    Returns False if:
    - ALL steps are create-tools, OR
    - The plan starts with create_workbook (the file_id will be generated at runtime)
    """
    if not plan_steps:
        return False
    # If the plan begins with a creation step, file_id will be produced at runtime
    if plan_steps[0].get("tool", "") in _CREATE_TOOLS:
        return False
    return any(step.get("tool", "") not in _CREATE_TOOLS for step in plan_steps)


# ==================== PROVIDER / MODEL HELPERS ====================


def _resolve_provider(provider: Optional[str]) -> str:
    """Resolve provider from request override or app default."""
    resolved = (provider or settings.LLM_PROVIDER or "claude").strip().lower()
    if resolved not in {"openai", "claude"}:
        raise HTTPException(status_code=400, detail="Invalid provider. Use 'openai' or 'claude'.")
    return resolved


def _resolve_model(model: Optional[str]) -> str:
    """Resolve Claude model shorthand to full model ID."""
    if not model:
        return settings.CLAUDE_MODEL  # default (sonnet)

    model_lower = model.strip().lower()
    model_map = {
        "sonnet": settings.CLAUDE_MODEL_SONNET,
        "opus": settings.CLAUDE_MODEL_OPUS,
    }
    if model_lower in model_map:
        return model_map[model_lower]
    if model_lower.startswith("claude-"):
        return model_lower
    raise HTTPException(status_code=400, detail=f"Invalid model '{model}'. Use 'sonnet' or 'opus'.")


# ==================== FORGET / RESET ====================

_FORGET_PATTERNS = frozenset([
    "forget everything", "forget all", "reset conversation",
    "clear history", "start over", "new conversation", "reset session",
])


def _is_forget_command(message: str) -> bool:
    msg = message.lower().strip()
    return any(p in msg for p in _FORGET_PATTERNS)


# ==================== SESSION / HISTORY / CHAT HELPERS ====================


def _get_or_create_default_user(db: Session) -> str:
    """Get or create an anonymous user for unauthenticated sessions."""
    anon = db.query(User).filter(User.username == "anonymous").first()
    if not anon:
        import hashlib
        anon = User(
            username="anonymous",
            password_hash=hashlib.sha256(b"anonymous").hexdigest(),
            is_active=True,
        )
        db.add(anon)
        db.commit()
        db.refresh(anon)
        logger.info(f"Created anonymous user: {anon.id}")
    return anon.id


def _get_or_create_session(session_id: Optional[str], db: Session) -> ChatSession:
    """Resume or create a chat session."""
    if session_id:
        session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
        if not session:
            raise HTTPException(status_code=404, detail="Session not found. Please start a new session.")
        logger.info(f"Resumed session: {session.id}")
    else:
        user_id = _get_or_create_default_user(db)
        session = ChatSession(is_active=True, user_id=user_id)
        db.add(session)
        db.commit()
        db.refresh(session)
        logger.info(f"Created session: {session.id}")

    session.updated_at = datetime.utcnow()
    db.commit()
    return session


def _load_history(session_id: str, db: Session, limit: int = 20) -> List[Dict]:
    """Load recent messages for LLM context."""
    msgs = (
        db.query(Message)
        .filter(Message.session_id == session_id)
        .order_by(Message.timestamp.desc())
        .limit(limit)
        .all()
    )
    history = [{"role": m.role, "content": m.content} for m in reversed(msgs)]
    logger.info(f"Loaded {len(history)} messages from history")
    return history


def _handle_forget(session: ChatSession, db: Session) -> MessageResponse:
    """Reset session context."""
    session.summary = None
    db.commit()
    text = "I've cleared our conversation history. How can I help you with Excel automation?"
    db.add(Message(session_id=session.id, role="assistant", content=text))
    db.commit()
    return MessageResponse(session_id=session.id, response=text, operations=[], context={"reset": True})


def _handle_general_chat(
    session: ChatSession,
    history: List[Dict],
    provider: str,
    model: str,
    db: Session,
) -> MessageResponse:
    """Handle non-Excel conversation."""
    logger.info("General conversation")
    response = llm_service.generate_response(
        messages=history, session_summary=session.summary, provider=provider, model=model
    )
    db.add(Message(session_id=session.id, role="assistant", content=response))

    # Update summary every 10 messages
    count = db.query(Message).filter(Message.session_id == session.id).count()
    if count % 10 == 0:
        logger.info("Updating session summary...")
        summary = llm_service.generate_session_summary(history, provider=provider, model=model)
        session.summary = summary

    db.commit()
    return MessageResponse(
        session_id=session.id,
        response=response,
        operations=[],
        context={"conversation": True, "provider": provider, "model": model},
    )


# ==================== MAIN CHAT ENDPOINT ====================

@router.post(
    "/message",
    response_model=MessageResponse,
    summary="Send Message to AI Assistant",
    description="Send a natural language message for Excel automation.",
    tags=["Chat"],
)
async def send_message(
    request_data: MessageRequest,
    req: Request,
    db: Session = Depends(get_db),
):
    """
    Main chat endpoint - clean flow:
      1. Session management
      2. Save user message
      3. Special commands (forget)
      4. Detect intent (Excel vs chat)
      5. Resolve context ONCE (file_id, sheet_name)
      6. Plan + execute + respond
    """
    try:
        # 1. SESSION
        chat_session = _get_or_create_session(request_data.session_id, db)

        # 2. SAVE USER MESSAGE
        db.add(Message(session_id=chat_session.id, role="user", content=request_data.message))
        db.commit()
        logger.info(f"[{chat_session.id[:8]}] {request_data.message[:100]}...")

        # 3. FORGET COMMAND
        if _is_forget_command(request_data.message):
            return _handle_forget(chat_session, db)

        # 4. CONVERSATION HISTORY + PROVIDER
        conversation_history = _load_history(chat_session.id, db)
        provider = _resolve_provider(request_data.provider)
        model = _resolve_model(request_data.model)
        logger.info(f"Provider: {provider}, Model: {model}")

        # 5. DETECT INTENT
        is_excel = llm_service.detect_excel_intent(request_data.message, request_data.file_id)

        if not is_excel:
            return _handle_general_chat(
                chat_session, conversation_history, provider, model, db
            )

        # 6. RESOLVE CONTEXT (ONCE)
        logger.info("Excel operation detected - resolving context...")
        tool_service = claude_mcp_service if provider == "claude" else mcp_service
        file_id, sheet_name = _resolve_execution_context(request_data, chat_session, db, tool_service)
        logger.info(f"Resolved context -> file_id={file_id}, sheet_name={sheet_name}")

        # 7. EXECUTE EXCEL PIPELINE
        return await _handle_excel_operation(
            session_id=chat_session.id,
            user_message=request_data.message,
            file_id=file_id,
            sheet_name=sheet_name,
            provider=provider,
            model=model,
            conversation_history=conversation_history,
            session_summary=chat_session.summary,
            db=db,
            ws_manager=req.app.state.ws_manager,
            tool_service=tool_service,
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"send_message error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# ==================== EXCEL OPERATION PIPELINE ====================

async def _handle_excel_operation(
    session_id: str,
    user_message: str,
    file_id: Optional[str],
    sheet_name: Optional[str],
    provider: str,
    model: str,
    conversation_history: List[Dict],
    session_summary: Optional[str],
    db: Session,
    ws_manager,
    tool_service,
) -> MessageResponse:
    """
    Clean Excel pipeline:
      1. Gather context (files, recent ops, column headers)
      2. Plan with LLM
      3. Check if plan needs file_id - ask user if missing
      4. Execute steps with context injection
      5. Format response
      6. Save and broadcast
    """
    try:
        # 1. GATHER CONTEXT
        files = await tool_service.list_files()
        available_files = [f"{f['filename']} ({f['file_id']})" for f in files]

        recent_ops = (
            db.query(Operation)
            .filter(Operation.session_id == session_id)
            .order_by(Operation.timestamp.desc())
            .limit(5)
            .all()
        )
        recent_operations = [
            {"tool": op.tool_name, "status": op.status, "timestamp": op.timestamp.isoformat()}
            for op in recent_ops
        ]

        column_headers = []
        if file_id and sheet_name:
            try:
                column_headers = await tool_service.get_column_headers(file_id, sheet_name)
                logger.info(f"Column headers: {column_headers}")
            except Exception as e:
                logger.warning(f"Could not get column headers: {e}")

        logger.info(f"Context: {len(available_files)} files, {len(recent_operations)} recent ops")

        # 2. BROADCAST PLANNING
        await ws_manager.broadcast({
            "type": "planning",
            "session_id": session_id,
            "message": "Analyzing your request and planning operations...",
        })

        # 3. CREATE PLAN
        logger.info("Creating execution plan...")
        plan = llm_service.plan_excel_operations(
            user_intent=user_message,
            context={
                "file_id": file_id,
                "sheet_name": sheet_name,
                "session_summary": session_summary,
                "available_files": available_files,
                "recent_operations": recent_operations,
                "column_headers": column_headers,
            },
            provider=provider,
            model=model,
        )
        steps = plan.get("steps", [])
        logger.info(f"Plan: {len(steps)} step(s)")

        # 4. CHECK IF FILE_ID IS NEEDED BUT MISSING
        if not file_id and _needs_file_id(steps):
            return _ask_for_file_id(session_id, db)

        # 5. EXECUTE STEPS
        operation_results = []
        active_file_id = file_id
        active_sheet_name = sheet_name

        for step in steps:
            step_num = step.get("step", "?")
            tool_name = step.get("tool", "unknown")
            logger.info(f"Step {step_num}: {tool_name}")

            # Inject context - uses tool name, not message keywords
            _inject_step_context(step, active_file_id, active_sheet_name)

            # Broadcast step start
            await ws_manager.broadcast({
                "type": "step_start",
                "session_id": session_id,
                "step": step_num,
                "description": step.get("description", ""),
            })

            # Create operation log
            operation = Operation(
                session_id=session_id,
                tool_name=tool_name,
                input_params=json.dumps(step.get("parameters", {})),
                status="in_progress",
            )
            db.add(operation)
            db.commit()
            db.refresh(operation)

            try:
                result = await tool_service.execute_tool(
                    tool_name=tool_name,
                    parameters=step.get("parameters", {}),
                    session_id=session_id,
                    db=db,
                )

                # Update operation log
                operation.status = "completed" if result["success"] else "failed"
                operation.output_result = json.dumps(result)
                operation.duration = result.get("duration")
                operation.error_message = result.get("error")
                db.commit()

                # Capture file_id from create_workbook for subsequent steps
                if tool_name == "create_workbook" and result.get("success"):
                    data = result.get("data", {})
                    if isinstance(data, dict) and data.get("file_id"):
                        active_file_id = data["file_id"]
                        logger.info(f"Captured new file_id: {active_file_id}")
                        # Also auto-detect sheet_name for the new file
                        active_sheet_name = _resolve_sheet_name(None, active_file_id, tool_service)

                operation_results.append({
                    "step": step_num,
                    "description": step.get("description", ""),
                    "tool": tool_name,
                    "status": "completed" if result["success"] else "failed",
                    "result": result.get("data"),
                    "error": result.get("error"),
                })

                await ws_manager.broadcast({
                    "type": "step_complete",
                    "session_id": session_id,
                    "step": step_num,
                    "status": "completed" if result["success"] else "failed",
                    "result": result.get("data"),
                })

                logger.info(f"Step {step_num} {'completed' if result['success'] else 'FAILED'}")

                if not result["success"]:
                    logger.warning(f"Step {step_num} failed - stopping execution")
                    break

            except Exception as e:
                logger.error(f"Step {step_num} error: {e}")
                _safe_rollback(db)
                operation.status = "failed"
                operation.error_message = str(e)
                db.commit()

                operation_results.append({
                    "step": step_num,
                    "description": step.get("description", ""),
                    "tool": tool_name,
                    "status": "failed",
                    "error": str(e),
                })
                await ws_manager.broadcast({
                    "type": "step_failed",
                    "session_id": session_id,
                    "step": step_num,
                    "error": str(e),
                })
                break

        # 6. FORMAT RESPONSE
        executed_step_nums = {r["step"] for r in operation_results}
        executed_steps = [s for s in steps if s.get("step") in executed_step_nums]

        response_text = llm_service.format_operation_results(
            steps=executed_steps,
            results=operation_results,
            provider=provider,
            model=model,
        )

        # 7. SAVE ASSISTANT MESSAGE
        db.add(Message(
            session_id=session_id,
            role="assistant",
            content=response_text,
            meta_data=json.dumps({
                "operations": operation_results,
                "file_id": active_file_id,
                "sheet_name": active_sheet_name,
                "provider": provider,
                "model": model,
            }),
        ))
        db.commit()

        # 8. UPDATE SESSION SUMMARY (every 5 ops)
        completed_ops = db.query(Operation).filter(
            Operation.session_id == session_id,
            Operation.status == "completed",
        ).count()
        if completed_ops > 0 and completed_ops % 5 == 0:
            logger.info("Updating session summary...")
            session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
            if session:
                session.summary = llm_service.generate_session_summary(
                    conversation_history, [], provider=provider, model=model
                )
                db.commit()

        # 9. BROADCAST COMPLETION
        successful = len([r for r in operation_results if r["status"] == "completed"])
        await ws_manager.broadcast({
            "type": "operations_complete",
            "session_id": session_id,
            "total_steps": len(steps),
            "successful_steps": successful,
        })

        logger.info(f"Excel workflow done: {successful}/{len(steps)} steps succeeded")

        return MessageResponse(
            session_id=session_id,
            response=response_text,
            operations=operation_results,
            context={
                "file_id": active_file_id,
                "sheet_name": active_sheet_name,
                "provider": provider,
                "model": model,
                "total_operations": len(operation_results),
                "successful": successful,
            },
        )

    except Exception as e:
        logger.error(f"Excel pipeline error: {e}")
        error_msg = f"I encountered an error while processing your request: {str(e)}\n\nPlease try again or rephrase your request."
        db.add(Message(session_id=session_id, role="assistant", content=error_msg))
        db.commit()
        return MessageResponse(
            session_id=session_id,
            response=error_msg,
            operations=[],
            context={"error": True, "provider": provider},
        )


# ==================== SMALL UTILITIES ====================


def _ask_for_file_id(session_id: str, db: Session) -> MessageResponse:
    """Prompt user to provide a file when the plan needs one but none was resolved."""
    text = (
        "I'd love to help with that Excel operation! However, I need to know which file to work with.\n\n"
        "You can either:\n"
        "1. Include the file ID in your message\n"
        "2. List available files by saying 'show me my files'\n"
        "3. Create a new file by saying 'create a new Excel file'\n\n"
        "Which would you prefer?"
    )
    db.add(Message(session_id=session_id, role="assistant", content=text))
    db.commit()
    return MessageResponse(
        session_id=session_id,
        response=text,
        operations=[],
        context={"needs_file_id": True},
    )


def _safe_rollback(db: Session):
    """Rollback DB transaction without raising."""
    try:
        db.rollback()
    except Exception:
        pass


# ==================== CHAT HISTORY ENDPOINT ====================

@router.get(
    "/history/{session_id}",
    summary="Get Chat History",
    description="Retrieve full conversation history for a session",
    tags=["Chat"],
)
async def get_history(session_id: str, db: Session = Depends(get_db)):
    """Get complete chat history for a session."""
    try:
        chat_session = db.query(ChatSession).filter(ChatSession.id == session_id).first()
        if not chat_session:
            raise HTTPException(status_code=404, detail="Session not found")

        messages = (
            db.query(Message)
            .filter(Message.session_id == session_id)
            .order_by(Message.timestamp.asc())
            .all()
        )

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
                    "metadata": json.loads(msg.meta_data) if msg.meta_data else None,
                }
                for msg in messages
            ],
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Get history error: {e}")
        raise HTTPException(status_code=500, detail=str(e))
