"""
Session Management APIs
Create and manage user sessions with one-to-one file relationship
"""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional
from datetime import datetime

from app.core.database import get_db
from app.core.models import User, Session as ChatSession, ExcelFile
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
    is_active: bool
    summary: Optional[str]
    file_id: Optional[str]
    filename: Optional[str]
    
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
    
    return SessionDetailResponse(
        session_id=session.id,
        user_id=session.user_id,
        created_at=session.created_at,
        updated_at=session.updated_at,
        is_active=session.is_active,
        summary=session.summary,
        file_id=excel_file.id if excel_file else None,
        filename=excel_file.filename if excel_file else None
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
