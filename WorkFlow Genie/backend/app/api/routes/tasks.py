# """
# Task APIs for WorkflowGenie
# ONLY returns LLM response text - MCP execution happens via /api/chat
# """

# from fastapi import APIRouter, Depends, HTTPException, status, Form
# from sqlalchemy.orm import Session
# from pydantic import BaseModel
# from typing import List, Optional
# from datetime import datetime

# from app.core.database import get_db
# from app.core.models import User, Task, Session as ChatSession, ExcelFile
# from app.api.routes.auth import get_current_user
# from app.services.llm_service import LLMService

# router = APIRouter(prefix="/api/tasks", tags=["Tasks"])

# llm_service = LLMService()

# # ============================================================================
# # MODELS
# # ============================================================================

# class TaskCreateResponse(BaseModel):
#     task_id: str
#     session_id: str
#     file_id: str
#     llm_response: str
#     message: str

# class TaskListItem(BaseModel):
#     task_id: str
#     prompt: str
#     llm_response: Optional[str]
#     created_at: datetime
    
#     class Config:
#         from_attributes = True

# # ============================================================================
# # API ENDPOINTS
# # ============================================================================

# @router.post("/execute", response_model=TaskCreateResponse, status_code=status.HTTP_201_CREATED)
# async def execute_task(
#     prompt: str = Form(...),
#     session_id: str = Form(...),
#     current_user: User = Depends(get_current_user),
#     db: Session = Depends(get_db)
# ):
#     """
#     Execute task - Returns LLM response instantly
    
#     Input:
#     - session_id: Session ID (file_id is auto-fetched from session)
#     - prompt: User's request
    
#     Returns:
#     - task_id: For tracking
#     - llm_response: Friendly text for UI
#     - file_id: For calling /api/chat to execute MCP
    
#     Note: This ONLY generates text. Call /api/chat for actual Excel operations.
#     """
    
#     # Verify session belongs to user
#     session = db.query(ChatSession).filter(
#         ChatSession.id == session_id,
#         ChatSession.user_id == current_user.id
#     ).first()
    
#     if not session:
#         raise HTTPException(
#             status_code=status.HTTP_404_NOT_FOUND,
#             detail="Session not found"
#         )
    
#     # Get file_id from session (one-to-one relationship)
#     excel_file = db.query(ExcelFile).filter(
#         ExcelFile.session_id == session_id
#     ).first()
    
#     if not excel_file:
#         raise HTTPException(
#             status_code=status.HTTP_404_NOT_FOUND,
#             detail="No file found in this session. Upload a file first."
#         )
    
#     file_id = excel_file.id
    
#     # Generate LLM friendly response (instant)
#     try:
#         llm_response_obj = llm_service.client.chat.completions.create(
#             model="gpt-4o",
#             messages=[{
#                 "role": "user",
#                 "content": f"User asked: '{prompt}'. Respond with a friendly, helpful message saying you'll help them with their Excel task. Be concise (1-2 sentences). Sound enthusiastic and professional."
#             }],
#             max_tokens=150
#         )
        
#         llm_text = llm_response_obj.choices[0].message.content
#     except Exception as e:
#         llm_text = f"Sure! I'll help you with that Excel task."
    
#     # Create task record (for history tracking)
#     new_task = Task(
#         user_id=current_user.id,
#         session_id=session_id,
#         prompt=prompt,
#         response=llm_text,
#         status="completed",  # Task itself is complete (just text generation)
#         progress=100,
#         result_file_id=file_id
#     )
    
#     db.add(new_task)
#     db.commit()
#     db.refresh(new_task)
    
#     return TaskCreateResponse(
#         task_id=new_task.id,
#         session_id=session_id,
#         file_id=file_id,
#         llm_response=llm_text,
#         message="LLM response generated. Call /api/chat to execute Excel operations."
#     )


# @router.get("/list", response_model=List[TaskListItem])
# async def list_user_tasks(
#     current_user: User = Depends(get_current_user),
#     db: Session = Depends(get_db),
#     limit: int = 50
# ):
#     """List all user tasks (LLM responses history)"""
    
#     tasks = db.query(Task).filter(
#         Task.user_id == current_user.id
#     ).order_by(Task.created_at.desc()).limit(limit).all()
    
#     return [TaskListItem(
#         task_id=t.id,
#         prompt=t.prompt[:100] + "..." if len(t.prompt) > 100 else t.prompt,
#         llm_response=t.response,
#         created_at=t.created_at
#     ) for t in tasks]


# @router.get("/{task_id}")
# async def get_task(
#     task_id: str,
#     current_user: User = Depends(get_current_user),
#     db: Session = Depends(get_db)
# ):
#     """Get specific task details"""
    
#     task = db.query(Task).filter(
#         Task.id == task_id,
#         Task.user_id == current_user.id
#     ).first()
    
#     if not task:
#         raise HTTPException(
#             status_code=status.HTTP_404_NOT_FOUND,
#             detail="Task not found"
#         )
    
#     return {
#         "task_id": task.id,
#         "session_id": task.session_id,
#         "prompt": task.prompt,
#         "llm_response": task.response,
#         "file_id": task.result_file_id,
#         "created_at": task.created_at
#     }


# @router.delete("/{task_id}")
# async def delete_task(
#     task_id: str,
#     current_user: User = Depends(get_current_user),
#     db: Session = Depends(get_db)
# ):
#     """Delete task"""
    
#     task = db.query(Task).filter(
#         Task.id == task_id,
#         Task.user_id == current_user.id
#     ).first()
    
#     if not task:
#         raise HTTPException(
#             status_code=status.HTTP_404_NOT_FOUND,
#             detail="Task not found"
#         )
    
#     db.delete(task)
#     db.commit()
    
#     return {"message": "Task deleted", "task_id": task_id}