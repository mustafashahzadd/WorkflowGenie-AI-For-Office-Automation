"""
Session Management APIs
Create and manage user sessions with one-to-one file relationship
"""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from sqlalchemy import desc
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime

from app.core.database import get_db
from app.core.models import User, Session as ChatSession, ExcelFile, Message
from app.api.routes.auth import get_current_user

router = APIRouter(prefix="/api/sessions", tags=["Session Management"])

# ============================================================================
# RESPONSE MODELS
# ============================================================================

class SessionCreateResponse(BaseModel):
    session_id: str
    message: str

class SessionDetailResponse(BaseModel):
    session_id: str
    user_id: str
    created_at: datetime
    updated_at: datetime
    created_at_display: str  # Human-readable "14 Apr 2026, 1:09 PM"
    updated_at_display: str
    is_active: bool
    summary: Optional[str]
    file_id: Optional[str]
    filename: Optional[str]
    message_count: int = 0
    
    class Config:
        from_attributes = True

# ============================================================================
# API ENDPOINTS
# ============================================================================

@router.post("/create", response_model=SessionCreateResponse, status_code=status.HTTP_201_CREATED)
async def create_session(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Create new session for user
    
    Each session can have ONE file only
    
    Requires Authorization header: Bearer <token>
    """
    
    new_session = ChatSession(user_id=current_user.id)
    db.add(new_session)
    db.commit()
    db.refresh(new_session)
    
    return SessionCreateResponse(
        session_id=new_session.id,
        message="Session created successfully. Upload ONE file to this session."
    )


@router.get("/list", response_model=List[SessionDetailResponse])
async def list_sessions(
    active_only: bool = True,
    limit: int = 50,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    List all sessions for current user — newest first.
    
    - **active_only**: Only show active sessions (default: true)
    - **limit**: Max sessions to return (default: 50)
    
    Requires Authorization header: Bearer <token>
    """
    query = db.query(ChatSession).filter(ChatSession.user_id == current_user.id)
    
    if active_only:
        query = query.filter(ChatSession.is_active == True)
    
    sessions = query.order_by(desc(ChatSession.created_at)).limit(limit).all()
    
    result = []
    for session in sessions:
        excel_file = db.query(ExcelFile).filter(ExcelFile.session_id == session.id).first()
        msg_count = db.query(Message).filter(Message.session_id == session.id).count()
        result.append(SessionDetailResponse(
            session_id=session.id,
            user_id=session.user_id,
            created_at=session.created_at,
            updated_at=session.updated_at,
            created_at_display=session.created_at.strftime("%d %b %Y, %I:%M %p"),
            updated_at_display=session.updated_at.strftime("%d %b %Y, %I:%M %p"),
            is_active=session.is_active,
            summary=session.summary,
            file_id=excel_file.id if excel_file else None,
            filename=excel_file.filename if excel_file else None,
            message_count=msg_count
        ))
    
    return result


@router.get("/{session_id}", response_model=SessionDetailResponse)
async def get_session_detail(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Get session details including linked file
    
    Requires Authorization header: Bearer <token>
    """
    
    session = db.query(ChatSession).filter(
        ChatSession.id == session_id,
        ChatSession.user_id == current_user.id
    ).first()
    
    if not session:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Session not found"
        )
    
    # Get file linked to this session (should be only ONE)
    excel_file = db.query(ExcelFile).filter(ExcelFile.session_id == session_id).first()
    msg_count = db.query(Message).filter(Message.session_id == session_id).count()
    
    return SessionDetailResponse(
        session_id=session.id,
        user_id=session.user_id,
        created_at=session.created_at,
        updated_at=session.updated_at,
        created_at_display=session.created_at.strftime("%d %b %Y, %I:%M %p"),
        updated_at_display=session.updated_at.strftime("%d %b %Y, %I:%M %p"),
        is_active=session.is_active,
        summary=session.summary,
        file_id=excel_file.id if excel_file else None,
        filename=excel_file.filename if excel_file else None,
        message_count=msg_count,
    )


@router.delete("/{session_id}")
async def delete_session(
    session_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Delete session (soft delete - marks as inactive)
    
    Requires Authorization header: Bearer <token>
    """
    
    session = db.query(ChatSession).filter(
        ChatSession.id == session_id,
        ChatSession.user_id == current_user.id
    ).first()
    
    if not session:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Session not found"
        )
    
    session.is_active = False
    db.commit()
    
    return {"message": "Session deleted successfully", "session_id": session_id}
