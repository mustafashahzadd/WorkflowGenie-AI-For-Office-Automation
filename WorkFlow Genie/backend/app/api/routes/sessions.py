"""
Sessions API Routes
"""

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from sqlalchemy import func

from app.core.database import get_db
from app.core.models import Session as DBSession, Message, Operation

router = APIRouter()

@router.get("/")
async def get_sessions(db: Session = Depends(get_db)):
    """Get all active sessions"""
    
    try:
        sessions = db.query(DBSession).filter(
            DBSession.is_active == True
        ).order_by(DBSession.updated_at.desc()).all()
        
        result = []
        for session in sessions:
            # Get message and operation counts
            message_count = db.query(func.count(Message.id)).filter(
                Message.session_id == session.id
            ).scalar()
            
            operation_count = db.query(func.count(Operation.id)).filter(
                Operation.session_id == session.id
            ).scalar()
            
            # Get last message
            last_message = db.query(Message).filter(
                Message.session_id == session.id
            ).order_by(Message.timestamp.desc()).first()
            
            result.append({
                "id": session.id,
                "summary": session.summary,
                "created_at": session.created_at.isoformat(),
                "updated_at": session.updated_at.isoformat(),
                "message_count": message_count,
                "operation_count": operation_count,
                "last_message": last_message.content if last_message else None
            })
        
        return {"sessions": result}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/{session_id}")
async def clear_session(session_id: str, db: Session = Depends(get_db)):
    """Clear/deactivate a session"""
    
    try:
        session = db.query(DBSession).filter(DBSession.id == session_id).first()
        
        if not session:
            raise HTTPException(status_code=404, detail="Session not found")
        
        session.is_active = False
        session.summary = None
        db.commit()
        
        return {
            "success": True,
            "message": "Session cleared successfully"
        }
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))