"""
Chat API Routes
Handles conversation and Excel automation requests
"""

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional, List, Dict
import json
from loguru import logger

from app.core.database import get_db
from app.core.models import Session as DBSession, Message, Operation
from app.services.llm_service import llm_service
from app.services.mcp_service import mcp_service

router = APIRouter()

class MessageRequest(BaseModel):
    """Request model for sending a message"""
    session_id: Optional[str] = None
    message: str

class MessageResponse(BaseModel):
    """Response model for message"""
    session_id: str
    response: str
    operations: List[Dict] = []

@router.post("/message", response_model=MessageResponse)
async def send_message(
    request_data: MessageRequest,
    req: Request,
    db: Session = Depends(get_db)
):
    """Send a message and get AI response"""
    
    try:
        # Get or create session
        if request_data.session_id:
            session = db.query(DBSession).filter(
                DBSession.id == request_data.session_id
            ).first()
            
            if not session:
                raise HTTPException(status_code=404, detail="Session not found")
        else:
            session = DBSession(is_active=True)
            db.add(session)
            db.commit()
            db.refresh(session)
        
        # Save user message
        user_message = Message(
            session_id=session.id,
            role="user",
            content=request_data.message
        )
        db.add(user_message)
        db.commit()
        
        # Check for "forget everything" command
        if _is_forget_command(request_data.message):
            session.summary = None
            db.commit()
            
            reset_message = "I've cleared our conversation history. How can I help you with Excel automation?"
            
            assistant_message = Message(
                session_id=session.id,
                role="assistant",
                content=reset_message
            )
            db.add(assistant_message)
            db.commit()
            
            return MessageResponse(
                session_id=session.id,
                response=reset_message,
                operations=[]
            )
        
        # Get recent messages
        recent_messages = db.query(Message).filter(
            Message.session_id == session.id
        ).order_by(Message.timestamp.desc()).limit(20).all()
        
        conversation_history = [
            {"role": msg.role, "content": msg.content}
            for msg in reversed(recent_messages)
        ]
        
        # Detect Excel operation
        is_excel_operation = llm_service.detect_excel_intent(request_data.message)
        
        if is_excel_operation:
            # Handle Excel automation workflow
            result = await _handle_excel_operation(
                session_id=session.id,
                user_message=request_data.message,
                conversation_history=conversation_history,
                session_summary=session.summary,
                db=db,
                ws_manager=req.app.state.ws_manager
            )
            
            return result
        
        else:
            # Regular conversation
            response = llm_service.generate_response(
                messages=conversation_history,
                session_summary=session.summary
            )
            
            assistant_message = Message(
                session_id=session.id,
                role="assistant",
                content=response
            )
            db.add(assistant_message)
            
            # Update session summary periodically
            message_count = db.query(Message).filter(
                Message.session_id == session.id
            ).count()
            
            if message_count % 10 == 0:
                summary = llm_service.generate_session_summary(
                    [{"role": m.role, "content": m.content} for m in recent_messages]
                )
                session.summary = summary
            
            db.commit()
            
            return MessageResponse(
                session_id=session.id,
                response=response,
                operations=[]
            )
    
    except Exception as e:
        logger.error(f"Send message error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/history/{session_id}")
async def get_history(session_id: str, db: Session = Depends(get_db)):
    """Get chat history for a session"""
    
    try:
        session = db.query(DBSession).filter(DBSession.id == session_id).first()
        
        if not session:
            raise HTTPException(status_code=404, detail="Session not found")
        
        messages = db.query(Message).filter(
            Message.session_id == session_id
        ).order_by(Message.timestamp.asc()).all()
        
        return {
            "session_id": session_id,
            "summary": session.summary,
            "messages": [
                {
                    "id": msg.id,
                    "role": msg.role,
                    "content": msg.content,
                    "timestamp": msg.timestamp.isoformat(),
                    "metadata": json.loads(msg.metadata) if msg.metadata else None
                }
                for msg in messages
            ]
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Get history error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

async def _handle_excel_operation(
    session_id: str,
    user_message: str,
    conversation_history: List[Dict],
    session_summary: Optional[str],
    db: Session,
    ws_manager
):
    """Handle Excel operation workflow"""
    
    try:
        # Get available files
        files = await mcp_service.list_files()
        available_files = [f"{f['filename']} ({f['file_id']})" for f in files]
        
        # Get recent operations
        recent_ops = db.query(Operation).filter(
            Operation.session_id == session_id
        ).order_by(Operation.timestamp.desc()).limit(5).all()
        
        # Broadcast planning
        await ws_manager.broadcast({
            "type": "planning",
            "session_id": session_id,
            "message": "Planning Excel operations..."
        })
        
        # Plan operations
        plan = llm_service.plan_excel_operations(
            user_intent=user_message,
            context={
                "session_summary": session_summary,
                "available_files": available_files,
                "recent_operations": [
                    {
                        "tool": op.tool_name,
                        "params": json.loads(op.input_params),
                        "result": json.loads(op.output_result) if op.output_result else None
                    }
                    for op in recent_ops
                ]
            }
        )
        
        # Execute steps
        operation_results = []
        
        for step in plan["steps"]:
            # Broadcast step start
            await ws_manager.broadcast({
                "type": "step_start",
                "session_id": session_id,
                "step": step["step"],
                "description": step["description"]
            })
            
            # Create operation log
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
                # Execute tool
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
                
                # Stop on failure
                if not result["success"]:
                    break
            
            except Exception as e:
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
        
        # Generate natural language response
        response = llm_service.format_operation_results(
            operations=operation_results,
            user_intent=user_message
        )
        
        # Save assistant message
        assistant_message = Message(
            session_id=session_id,
            role="assistant",
            content=response,
            metadata=json.dumps({"operations": operation_results})
        )
        db.add(assistant_message)
        db.commit()
        
        # Broadcast completion
        await ws_manager.broadcast({
            "type": "operations_complete",
            "session_id": session_id,
            "total_steps": len(plan["steps"])
        })
        
        return MessageResponse(
            session_id=session_id,
            response=response,
            operations=operation_results
        )
    
    except Exception as e:
        logger.error(f"Excel operation error: {e}")
        
        error_message = f"I encountered an error: {str(e)}"
        
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
            operations=[]
        )

def _is_forget_command(message: str) -> bool:
    """Check if message is a forget command"""
    
    forget_patterns = [
        'forget everything',
        'forget all',
        'reset conversation',
        'clear history',
        'start over',
        'new conversation'
    ]
    
    message_lower = message.lower().strip()
    return any(pattern in message_lower for pattern in forget_patterns)