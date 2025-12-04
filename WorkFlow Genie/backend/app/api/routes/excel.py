"""
Excel API Routes
"""

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Any

from app.core.database import get_db
from app.services.mcp_service import mcp_service

router = APIRouter()

# ✅ Define QuickUpdateRequest at module level (outside any function)
class QuickUpdateRequest(BaseModel):
    file_id: str
    sheet_name: str
    person_name: str
    field: str
    new_value: Any

@router.get("/files")
async def list_files(db: Session = Depends(get_db)):
    """List all Excel files"""
    
    try:
        files = await mcp_service.list_files()
        return {"files": files}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/files/{file_id}")
async def get_file_metadata(file_id: str, db: Session = Depends(get_db)):
    """Get metadata for a specific file"""
    
    try:
        result = await mcp_service.execute_tool(
            tool_name="get_file_metadata",
            parameters={"file_id": file_id}
        )
        
        if not result["success"]:
            raise HTTPException(status_code=404, detail=result.get("error", "File not found"))
        
        return result["data"]
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/quick-update")
async def quick_update(request: QuickUpdateRequest):
    """
    Quick update without LLM - direct Excel update
    
    Example request:
    {
      "file_id": "3cf05b87-ce48-4dd7-b463-547d91ffb2bd",
      "sheet_name": "HR",
      "person_name": "John Smith",
      "field": "Salary",
      "new_value": 62000
    }
    """
    try:
        result = await mcp_service.update_by_search(
            file_id=request.file_id,
            sheet_name=request.sheet_name,
            search_column="Name",
            search_value=request.person_name,
            update_column=request.field,
            new_value=request.new_value
        )
        return {
            "success": True,
            "data": result
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))