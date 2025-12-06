"""
Session History API - For Chat UI Display
Shows session preview with first prompt/response like WhatsApp chat list
"""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from sqlalchemy import desc
from pydantic import BaseModel
from typing import List, Optional
from datetime import datetime

from app.core.database import get_db
from app.core.models import User, Session as ChatSession, Message, Task, ExcelFile
from app.api.routes.auth import get_current_user

router = APIRouter(prefix="/api/history", tags=["Session History"])

# ============================================================================
# RESPONSE MODELS
# ============================================================================

class SessionHistoryItem(BaseModel):
    """Session preview for chat UI list"""
    session_id: str
    created_at: datetime
    updated_at: datetime
    first_prompt: Optional[str]  # First user message
    last_response: Optional[str]  # Last assistant response
    summary: Optional[str]  # Session summary
    filename: Optional[str]  # Linked Excel file name
    message_count: int
    task_count: int
    is_active: bool

# ============================================================================
# API ENDPOINTS
# ============================================================================

@router.get("/sessions", response_model=List[SessionHistoryItem])
async def get_session_history(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
    limit: int = 50,
    active_only: bool = True
):
    """
    Get session history for chat UI display
    
    Returns list of sessions with preview info (like WhatsApp chat list)
    
    - **limit**: Max sessions to return (default: 50)
    - **active_only**: Show only active sessions (default: true)
    
    Requires Authorization header: Bearer <token>
    """
    
    # Get user's sessions
    query = db.query(ChatSession).filter(ChatSession.user_id == current_user.id)
    
    if active_only:
        query = query.filter(ChatSession.is_active == True)
    
    sessions = query.order_by(desc(ChatSession.updated_at)).limit(limit).all()
    
    history_items = []
    
    for session in sessions:
        # Get first user message (prompt)
        first_message = db.query(Message).filter(
            Message.session_id == session.id,
            Message.role == "user"
        ).order_by(Message.timestamp).first()
        
        # Get last assistant response
        last_response = db.query(Message).filter(
            Message.session_id == session.id,
            Message.role == "assistant"
        ).order_by(desc(Message.timestamp)).first()
        
        # Get message count
        message_count = db.query(Message).filter(
            Message.session_id == session.id
        ).count()
        
        # Get task count
        task_count = db.query(Task).filter(
            Task.session_id == session.id
        ).count()
        
        # Get linked file
        excel_file = db.query(ExcelFile).filter(
            ExcelFile.session_id == session.id
        ).first()
        
        history_items.append(SessionHistoryItem(
            session_id=session.id,
            created_at=session.created_at,
            updated_at=session.updated_at,
            first_prompt=first_message.content[:100] if first_message else None,
            last_response=last_response.content[:100] if last_response else None,
            summary=session.summary,
            filename=excel_file.filename if excel_file else None,
            message_count=message_count,
            task_count=task_count,
            is_active=session.is_active
        ))
    
    return history_items


@router.get("/session/{session_id}/full", response_model=dict)
async def get_full_session_history(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Get full session history with all messages and tasks
    
    For opening a specific chat session
    
    Requires Authorization header: Bearer <token>
    """
    
    # Verify session belongs to user
    session = db.query(ChatSession).filter(
        ChatSession.id == session_id,
        ChatSession.user_id == current_user.id
    ).first()
    
    if not session:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Session not found"
        )
    
    # Get all messages
    messages = db.query(Message).filter(
        Message.session_id == session_id
    ).order_by(Message.timestamp).all()
    
    # Get all tasks
    tasks = db.query(Task).filter(
        Task.session_id == session_id
    ).order_by(Task.created_at).all()
    
    # Get linked file
    excel_file = db.query(ExcelFile).filter(
        ExcelFile.session_id == session_id
    ).first()
    
    return {
        "session_id": session.id,
        "created_at": session.created_at,
        "updated_at": session.updated_at,
        "summary": session.summary,
        "file": {
            "file_id": excel_file.id if excel_file else None,
            "filename": excel_file.filename if excel_file else None
        },
        "messages": [
            {
                "id": msg.id,
                "role": msg.role,
                "content": msg.content,
                "timestamp": msg.timestamp
            } for msg in messages
        ],
        "tasks": [
            {
                "task_id": task.id,
                "prompt": task.prompt,
                "status": task.status,
                "progress": task.progress,
                "response": task.response,
                "created_at": task.created_at,
                "completed_at": task.completed_at
            } for task in tasks
        ]
    }
