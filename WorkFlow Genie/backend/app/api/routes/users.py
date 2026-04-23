"""
User Session APIs for WorkflowGenie
Handles listing user sessions and retrieving session messages
"""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from sqlalchemy import desc
from pydantic import BaseModel
from typing import List, Optional
from datetime import datetime

from app.core.database import get_db
from app.core.models import User, Session as ChatSession, Message, ExcelFile
from app.api.routes.auth import get_current_user

router = APIRouter(prefix="/api/users", tags=["User Sessions"])

# ============================================================================
# RESPONSE MODELS
# ============================================================================

class MessageResponse(BaseModel):
    id: str
    role: str
    content: str
    timestamp: datetime
    
    class Config:
        from_attributes = True

class SessionSummary(BaseModel):
    id: str
    created_at: datetime
    updated_at: datetime
    created_at_display: str  # "14 Apr 2026, 1:09 PM"
    updated_at_display: str
    is_active: bool
    summary: Optional[str]
    message_count: int
    file_id: Optional[str] = None
    filename: Optional[str] = None
    
    class Config:
        from_attributes = True

class SessionDetail(BaseModel):
    id: str
    created_at: datetime
    updated_at: datetime
    is_active: bool
    summary: Optional[str]
    messages: List[MessageResponse]
    
    class Config:
        from_attributes = True

# ============================================================================
# API ENDPOINTS
# ============================================================================

@router.get("/sessions", response_model=List[SessionSummary])
async def get_user_sessions(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
    active_only: bool = True,
    limit: int = 50
):
    """
    Get all sessions for current user
    
    - **active_only**: Filter only active sessions (default: true)
    - **limit**: Maximum number of sessions to return (default: 50)
    
    Returns list of sessions sorted by most recent first
    
    Requires Authorization header: Bearer <token>
    """
    
    query = db.query(ChatSession).filter(ChatSession.user_id == current_user.id)
    
    if active_only:
        query = query.filter(ChatSession.is_active == True)
    
    sessions = query.order_by(desc(ChatSession.updated_at)).limit(limit).all()
    
    # Add message count to each session
    session_summaries = []
    for session in sessions:
        message_count = db.query(Message).filter(Message.session_id == session.id).count()
        excel_file = db.query(ExcelFile).filter(ExcelFile.session_id == session.id).first()
        session_dict = {
            "id": session.id,
            "created_at": session.created_at,
            "updated_at": session.updated_at,
            "created_at_display": session.created_at.strftime("%d %b %Y, %I:%M %p"),
            "updated_at_display": session.updated_at.strftime("%d %b %Y, %I:%M %p"),
            "is_active": session.is_active,
            "summary": session.summary,
            "message_count": message_count,
            "file_id": excel_file.id if excel_file else None,
            "filename": excel_file.filename if excel_file else None,
        }
        session_summaries.append(SessionSummary(**session_dict))
    
    return session_summaries


@router.get("/sessions/{session_id}", response_model=SessionDetail)
async def get_session_detail(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Get detailed session information with all messages
    
    - **session_id**: Session ID to retrieve
    
    Returns session details with all messages in chronological order
    
    Requires Authorization header: Bearer <token>
    """
    
    # Find session
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
    
    return SessionDetail(
        id=session.id,
        created_at=session.created_at,
        updated_at=session.updated_at,
        is_active=session.is_active,
        summary=session.summary,
        messages=[MessageResponse.from_orm(msg) for msg in messages]
    )


@router.get("/sessions/{session_id}/messages", response_model=List[MessageResponse])
async def get_session_messages(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
    limit: Optional[int] = None,
    offset: int = 0
):
    """
    Get messages from a specific session
    
    - **session_id**: Session ID
    - **limit**: Maximum number of messages (optional)
    - **offset**: Number of messages to skip (default: 0)
    
    Returns messages in chronological order
    
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
    
    # Get messages
    query = db.query(Message).filter(
        Message.session_id == session_id
    ).order_by(Message.timestamp).offset(offset)
    
    if limit:
        query = query.limit(limit)
    
    messages = query.all()
    
    return [MessageResponse.from_orm(msg) for msg in messages]


@router.delete("/sessions/{session_id}")
async def delete_session(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Delete a session (soft delete - marks as inactive)
    
    - **session_id**: Session ID to delete
    
    Requires Authorization header: Bearer <token>
    """
    
    # Find session
    session = db.query(ChatSession).filter(
        ChatSession.id == session_id,
        ChatSession.user_id == current_user.id
    ).first()
    
    if not session:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Session not found"
        )
    
    # Soft delete
    session.is_active = False
    db.commit()
    
    return {
        "message": "Session deleted successfully",
        "session_id": session_id
    }


@router.post("/sessions/{session_id}/activate")
async def activate_session(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Reactivate a deleted session
    
    - **session_id**: Session ID to activate
    
    Requires Authorization header: Bearer <token>
    """
    
    # Find session
    session = db.query(ChatSession).filter(
        ChatSession.id == session_id,
        ChatSession.user_id == current_user.id
    ).first()
    
    if not session:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Session not found"
        )
    
    # Activate
    session.is_active = True
    db.commit()
    
    return {
        "message": "Session activated successfully",
        "session_id": session_id
    }