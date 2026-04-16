"""
File Management APIs for WorkflowGenie
Handles Excel file upload, download, and metadata
"""

from fastapi import APIRouter, Depends, HTTPException, status, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime
from pathlib import Path
from io import BytesIO
import uuid
import openpyxl
from loguru import logger

from app.core.config import settings
from app.core.database import get_db
from app.core.models import User, ExcelFile
from app.api.routes.auth import get_current_user
from app.services.mcp_service import MCPService
from app.services.rag_service import rag_service

router = APIRouter(prefix="/api/files", tags=["File Management"])
mcp_service = MCPService()

# ============================================================================
# RESPONSE MODELS
# ============================================================================

class FileUploadResponse(BaseModel):
    file_id: str
    filename: str
    filepath: str
    message: str
    rag_synced: bool = True
    rag_message: Optional[str] = None


class FileValidationResult(BaseModel):
    signature_verified: bool
    structure_verified: bool
    integrity_verified: bool


class FileSaveResponse(BaseModel):
    message: str
    file_id: str
    filename: str
    size_bytes: int
    validation: Optional[FileValidationResult] = None

class FileMetadata(BaseModel):
    file_id: str
    filename: str
    uploaded_at: datetime
    sheets: List[str]
    rows: Optional[int]
    columns: Optional[int]
    
    class Config:
        from_attributes = True

class FileListItem(BaseModel):
    file_id: str
    filename: str
    uploaded_at: datetime
    
    class Config:
        from_attributes = True

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_file_path(file_id: str) -> Path:
    """Get file path from file_id"""
    data_dir = Path("data/excel_files")
    
    # Search for file with this file_id
    for file_path in data_dir.glob(f"{file_id}_*"):
        return file_path
    
    raise HTTPException(
        status_code=status.HTTP_404_NOT_FOUND,
        detail="File not found"
    )


def refresh_rag_index_safe(action: str) -> None:
    """Sync RAG index without interrupting API success path."""
    if not settings.RAG_ENABLED:
        logger.info(f"RAG is disabled; {action} sync skipped")
        return

    try:
        rag_result = rag_service.refresh_index()
        logger.info(
            f"RAG index synced after {action} "
            f"({rag_result.get('indexed_chunks', 0)} chunks across "
            f"{rag_result.get('indexed_files', 0)} files)"
        )
    except Exception as exc:
        logger.error(f"Failed to sync RAG index after {action}: {exc}")


def validate_excel_payload(filename: str, content: bytes) -> dict:
    """Perform lightweight Excel payload validation before saving."""
    lower_name = filename.lower()

    if lower_name.endswith((".xlsx", ".xlsm")):
        signature_verified = content.startswith(b"PK\x03\x04")
    elif lower_name.endswith(".xls"):
        signature_verified = content.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1")
    else:
        signature_verified = False

    structure_verified = False
    if signature_verified and lower_name.endswith((".xlsx", ".xlsm")):
        try:
            workbook = openpyxl.load_workbook(BytesIO(content), read_only=True, data_only=True)
            workbook.close()
            structure_verified = True
        except Exception:
            structure_verified = False
    elif signature_verified and lower_name.endswith(".xls"):
        # Legacy .xls is not parseable by openpyxl; signature check is used.
        structure_verified = True

    return {
        "signature_verified": signature_verified,
        "structure_verified": structure_verified,
        "integrity_verified": signature_verified and structure_verified,
    }

# ============================================================================
# API ENDPOINTS
# ============================================================================

@router.post("/upload", response_model=FileUploadResponse)
async def upload_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    session_id: str = Form(...),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Upload Excel file to session (ONE file per session only)
    
    Requires Authorization header: Bearer <token>
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only Excel files (.xlsx, .xls, .xlsm) are allowed"
        )
    
    # Verify session exists and belongs to user
    from app.core.models import Session as ChatSession
    session = db.query(ChatSession).filter(
        ChatSession.id == session_id,
        ChatSession.user_id == current_user.id
    ).first()
    
    if not session:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Session not found. Create session first using POST /api/sessions/create"
        )
    
    # Check if session already has a file (enforce one-to-one)
    existing_file = db.query(ExcelFile).filter(ExcelFile.session_id == session_id).first()
    if existing_file:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="This session already has a file. Create a new session to upload another file."
        )
    
    # Generate file ID
    file_id = str(uuid.uuid4())
    
    # Create directory if not exists
    data_dir = Path("data/excel_files")
    data_dir.mkdir(parents=True, exist_ok=True)
    
    # Save file
    file_path = data_dir / f"{file_id}_{file.filename}"
    
    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)
    
    # Save metadata to database with session link
    excel_file = ExcelFile(
        id=file_id,
        filename=file.filename,
        filepath=str(file_path),
        session_id=session_id
    )
    
    db.add(excel_file)
    db.commit()

    rag_synced = True
    rag_message: Optional[str] = None
    if settings.RAG_ENABLED:
        background_tasks.add_task(refresh_rag_index_safe, "upload")
        rag_message = "File uploaded successfully; RAG index sync scheduled"
    else:
        rag_message = "RAG is disabled; index sync skipped"
    
    return FileUploadResponse(
        file_id=file_id,
        filename=file.filename,
        filepath=str(file_path),
        message="File uploaded successfully to session",
        rag_synced=rag_synced,
        rag_message=rag_message,
    )


@router.get("/download/{file_id}")
async def download_file(
    file_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Download Excel file by file_id
    
    Requires Authorization header: Bearer <token>
    """
    
    # Get file from database
    excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
    
    if excel_file:
        # File found in database
        file_path = Path(excel_file.filepath)
        filename = excel_file.filename
    else:
        # Fallback: search by file_id pattern in data directory
        data_dir = Path("data/excel_files")
        matching_files = list(data_dir.glob(f"{file_id}_*"))
        
        if not matching_files:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        file_path = matching_files[0]
        # Extract filename from path (remove file_id prefix)
        filename = file_path.name.replace(f"{file_id}_", "")
    
    # Check if file exists on disk
    if not file_path.exists():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found on disk"
        )
    
    # Return file for download
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
    )


@router.put("/save/{file_id}", response_model=FileSaveResponse)
async def save_file(
    file_id: str,
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Save/replace existing Excel file content by file_id.

    Requires Authorization header: Bearer <token>
    """

    if not file.filename.endswith((".xlsx", ".xls", ".xlsm")):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only Excel files (.xlsx, .xls, .xlsm) are allowed"
        )

    excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found"
        )

    content = await file.read()
    if not content:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Uploaded file is empty"
        )

    validation = validate_excel_payload(file.filename, content)
    if not validation["integrity_verified"]:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Invalid Excel file content; integrity validation failed"
        )

    old_path = Path(excel_file.filepath)
    target_dir = old_path.parent if old_path.parent else Path("data/excel_files")
    target_dir.mkdir(parents=True, exist_ok=True)
    new_path = target_dir / f"{file_id}_{file.filename}"

    with open(new_path, "wb") as f:
        f.write(content)

    if old_path != new_path and old_path.exists():
        old_path.unlink()

    excel_file.filename = file.filename
    excel_file.filepath = str(new_path)
    db.commit()

    if settings.RAG_ENABLED:
        background_tasks.add_task(refresh_rag_index_safe, "save")

    return FileSaveResponse(
        message="File saved successfully",
        file_id=file_id,
        filename=file.filename,
        size_bytes=len(content),
        validation=FileValidationResult(**validation)
    )


@router.get("/metadata/{file_id}", response_model=FileMetadata)
async def get_file_metadata(
    file_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Get file metadata (sheets, rows, columns)
    
    Requires Authorization header: Bearer <token>
    """
    
    try:
        # Get file info from database
        excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
        
        if not excel_file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found in database"
            )
        
        # Get metadata from MCP service
        metadata = await mcp_service.get_file_metadata(file_id)
        
        return FileMetadata(
            file_id=file_id,
            filename=excel_file.filename,
            uploaded_at=excel_file.uploaded_at,
            sheets=metadata.get("sheets", []),
            rows=metadata.get("total_rows"),
            columns=metadata.get("total_columns")
        )
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error reading file metadata: {str(e)}"
        )


@router.get("/info/{file_id}")
async def get_file_info(
    file_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Get basic file info by file_id (for UI display)
    
    Requires Authorization header: Bearer <token>
    """
    
    excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
    
    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found"
        )
    
    return {
        "file_id": excel_file.id,
        "filename": excel_file.filename,
        "filepath": excel_file.filepath,
        "uploaded_at": excel_file.uploaded_at,
        "session_id": excel_file.session_id,
        "exists_on_disk": Path(excel_file.filepath).exists()
    }


@router.get("/list", response_model=List[FileListItem])
async def list_user_files(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
    limit: int = 50
):
    """
    List all files uploaded by user
    
    Requires Authorization header: Bearer <token>
    """
    
    # Simple query - get all files
    all_files = db.query(ExcelFile).order_by(
        ExcelFile.uploaded_at.desc()
    ).limit(limit).all()
    
    return [FileListItem(
        file_id=f.id,
        filename=f.filename,
        uploaded_at=f.uploaded_at
    ) for f in all_files]


@router.delete("/{file_id}")
async def delete_file(
    file_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Delete file
    
    Requires Authorization header: Bearer <token>
    """
    
    # Get file from database
    excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
    
    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found"
        )
    
    # Delete physical file
    try:
        file_path = get_file_path(file_id)
        if file_path.exists():
            file_path.unlink()
    except Exception as e:
        # Continue even if file deletion fails
        pass
    
    # Delete database record
    db.delete(excel_file)
    db.commit()

    rag_synced = True
    rag_message: Optional[str] = None
    if settings.RAG_ENABLED:
        try:
            rag_result = rag_service.refresh_index()
            rag_message = (
                f"RAG index synced ({rag_result.get('indexed_chunks', 0)} chunks across "
                f"{rag_result.get('indexed_files', 0)} files)"
            )
        except Exception as exc:
            rag_synced = False
            rag_message = f"File deleted, but failed to sync RAG index: {exc}"
            logger.error(rag_message)
    else:
        rag_message = "RAG is disabled; index sync skipped"
    
    return {
        "message": "File deleted successfully",
        "file_id": file_id,
        "rag_synced": rag_synced,
        "rag_message": rag_message,
    }


@router.get("/preview/{file_id}")
async def preview_file(
    file_id: str,
    sheet_name: str = "Sheet1",
    limit: int = 10,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Preview file content
    
    Requires Authorization header: Bearer <token>
    """
    
    try:
        # Use read_data to get preview
        result = await mcp_service.read_data(
            file_id=file_id,
            sheet_name=sheet_name,
            limit=limit
        )
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "data": result.get("data", []),
            "rows_shown": result.get("rows", 0),
            "total_columns": result.get("columns", 0)
        }
    
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error reading file: {str(e)}"
        )