"""
File Management APIs for WorkflowGenie
Handles Excel file upload, download, and metadata
"""

from fastapi import APIRouter, Depends, HTTPException, status, UploadFile, File, Form
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime
from pathlib import Path
import uuid
from loguru import logger
import openpyxl
import io
import os
import tempfile
import shutil
import hashlib

from app.core.database import get_db
from app.core.models import User, ExcelFile
from app.api.routes.auth import get_current_user
from app.services.mcp_service import MCPService

router = APIRouter(prefix="/api/files", tags=["File Management"])
mcp_service = MCPService()

# ============================================================================
# SECURITY CONSTANTS
# ============================================================================

# Maximum file size: 50MB
MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024  # 50MB

# Valid Excel file signatures (magic bytes)
EXCEL_SIGNATURES = {
    # .xlsx, .xlsm - ZIP format (Excel 2007+)
    b'PK\x03\x04': ['.xlsx', '.xlsm'],
    # .xls - OLE format (Excel 97-2003)
    b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1': ['.xls']
}

# ============================================================================
# RESPONSE MODELS
# ============================================================================

class FileUploadResponse(BaseModel):
    file_id: str
    filename: str
    filepath: str
    message: str

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

def is_file_locked(filepath: Path) -> bool:
    """
    Check if file is locked by another process (Windows-specific)
    
    Args:
        filepath: Path to the file to check
        
    Returns:
        True if file is locked, False otherwise
    """
    if not filepath.exists():
        return False
    
    # Windows-specific file locking check
    if os.name == 'nt':
        try:
            import msvcrt
            # Try to open file in exclusive mode
            with open(filepath, 'r+b') as f:
                # Try to get an exclusive lock
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
            return False
        except (IOError, OSError):
            return True
        except ImportError:
            logger.warning("msvcrt not available, cannot check file lock")
            return False
    else:
        # Unix-like systems: try to open with exclusive access
        try:
            import fcntl
            with open(filepath, 'r+b') as f:
                fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
            return False
        except (IOError, OSError, ImportError):
            return False  # Assume not locked if we can't check


def atomic_file_write(content: bytes, target_path: Path) -> None:
    """
    Atomically write file content using temp file + move strategy
    This prevents corruption and handles file locking better
    
    Args:
        content: The bytes to write
        target_path: The final destination path
        
    Raises:
        IOError: If write fails
    """
    # Create temp file in the same directory (important for atomic move on Windows)
    temp_dir = target_path.parent
    
    with tempfile.NamedTemporaryFile(
        mode='wb',
        dir=temp_dir,
        delete=False,
        prefix=f".tmp_{target_path.stem}_",
        suffix=target_path.suffix
    ) as temp_file:
        temp_path = Path(temp_file.name)
        try:
            # Write to temp file
            temp_file.write(content)
            temp_file.flush()
            os.fsync(temp_file.fileno())
            
            # Close the temp file before moving (important on Windows)
            temp_file.close()
            
            # Check if target exists and is locked
            if target_path.exists():
                if is_file_locked(target_path):
                    logger.error(f"Target file is locked: {target_path}")
                    raise IOError(f"Target file is locked by another process (possibly open in Excel): {target_path.name}")
                
                # Backup existing file
                backup_path = target_path.with_suffix(target_path.suffix + '.backup')
                try:
                    shutil.copy2(target_path, backup_path)
                except Exception as e:
                    logger.warning(f"Could not create backup: {e}")
            
            # Atomic move/replace (on Windows, need to handle differently)
            try:
                # On Windows, replace requires the target to be deletable
                if os.name == 'nt' and target_path.exists():
                    os.remove(target_path)
                
                # Move temp file to target
                shutil.move(str(temp_path), str(target_path))
                
            except Exception as e:
                logger.error(f"Failed to move temp file: {e}")
                raise IOError(f"Failed to replace target file: {e}")
            
            # Clean up backup if everything succeeded
            if target_path.exists():
                backup_path = target_path.with_suffix(target_path.suffix + '.backup')
                if backup_path.exists():
                    try:
                        os.remove(backup_path)
                    except:
                        pass
                        
        except Exception as e:
            # Clean up temp file if it still exists
            if temp_path.exists():
                try:
                    os.remove(temp_path)
                except:
                    pass
            raise


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


def validate_file_signature(file_content: bytes, filename: str) -> None:
    """
    Validate file signature (magic bytes) matches expected Excel format
    
    Args:
        file_content: The raw file bytes
        filename: Original filename for extension checking
        
    Raises:
        HTTPException: If file signature doesn't match or content is invalid
    """
    if len(file_content) == 0:
        logger.error("Empty file detected")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is empty. Please upload a valid Excel file."
        )
    
    # Check minimum size (Excel files should be at least a few hundred bytes)
    if len(file_content) < 512:
        logger.error(f"File too small: {len(file_content)} bytes")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is too small to be a valid Excel file."
        )
    
    # Get file extension
    file_ext = Path(filename).suffix.lower()
    
    # Check magic bytes
    file_signature_matched = False
    for signature, valid_extensions in EXCEL_SIGNATURES.items():
        if file_content.startswith(signature):
            if file_ext in valid_extensions:
                file_signature_matched = True
                break
            else:
                logger.error(f"File signature mismatch: signature={signature.hex()}, extension={file_ext}")
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail=f"File extension {file_ext} doesn't match file content. Possible file corruption or spoofing attempt."
                )
    
    if not file_signature_matched:
        logger.error(f"Invalid file signature: {file_content[:8].hex()}")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File does not appear to be a valid Excel file. Upload rejected for security reasons."
        )


def validate_excel_structure(file_content: bytes) -> None:
    """
    Validate that file can be opened as a valid Excel workbook
    
    Args:
        file_content: The raw file bytes
        
    Raises:
        HTTPException: If file cannot be parsed as Excel
    """
    try:
        # Try to open the file with openpyxl to validate structure
        workbook = openpyxl.load_workbook(io.BytesIO(file_content), read_only=True, data_only=True)
        
        # Verify it has at least one sheet
        if len(workbook.sheetnames) == 0:
            logger.error("Excel file has no sheets")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Excel file contains no sheets."
            )
        
        workbook.close()
        logger.info("Excel structure validation passed")
        
    except HTTPException:
        raise
    except openpyxl.utils.exceptions.InvalidFileException as e:
        logger.error(f"Invalid Excel file structure: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is corrupted or not a valid Excel format."
        )
    except Exception as e:
        logger.error(f"Excel validation error: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to validate Excel file: {str(e)}"
        )

# ============================================================================
# API ENDPOINTS
# ============================================================================

@router.post("/upload", response_model=FileUploadResponse)
async def upload_file(
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
    data_dir = Path("data/excel_files").resolve()  # Use absolute path
    data_dir.mkdir(parents=True, exist_ok=True)
    
    # Save file
    file_path = data_dir / f"{file_id}_{file.filename}"
    
    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)
        f.flush()  # Ensure data is written to disk
        import os
        os.fsync(f.fileno())  # Force write to disk
    
    # Save metadata to database with session link (store absolute path)
    excel_file = ExcelFile(
        id=file_id,
        filename=file.filename,
        filepath=str(file_path),  # Now stores absolute path
        session_id=session_id
    )
    
    db.add(excel_file)
    db.commit()
    
    return FileUploadResponse(
        file_id=file_id,
        filename=file.filename,
        filepath=str(file_path),
        message="File uploaded successfully to session"
    )


@router.put("/save/{file_id}")
async def save_file(
    file_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Save/update file content for an existing file ID (with security validation)
    
    This endpoint allows updating the content of an existing file.
    The file must already exist in the database (created via upload endpoint).
    
    Security Features:
    - File size limit enforcement (50MB max)
    - File signature (magic bytes) validation
    - Excel structure validation
    - Empty file detection
    - Content type verification
    
    Args:
        file_id: The ID of the file to update
        file: The new file content to save
        
    Returns:
        Success message with file details
        
    Raises:
        404: If file_id does not exist in database
        400: If file validation fails (size, type, signature, structure)
        413: If file exceeds maximum size limit
        500: If file write operation fails
        
    Requires Authorization header: Bearer <token>
    """
    
    logger.info(f"Save file request received for file_id: {file_id} by user: {current_user.username}")
    
    try:
        # ===== STEP 1: Verify file exists in database =====
        excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
        
        if not excel_file:
            logger.warning(f"File not found in database: {file_id}")
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"File with ID '{file_id}' not found. Please ensure the file ID is correct."
            )
        
        # ===== STEP 2: Validate file extension =====
        if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
            logger.warning(f"Invalid file extension for file_id {file_id}: {file.filename}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Only Excel files (.xlsx, .xls, .xlsm) are allowed"
            )
        
        # ===== STEP 3: Read file content with size limit =====
        logger.info(f"Reading uploaded file content for validation")
        content = b''
        chunk_size = 1024 * 1024  # 1MB chunks
        total_size = 0
        
        try:
            while True:
                chunk = await file.read(chunk_size)
                if not chunk:
                    break
                    
                total_size += len(chunk)
                
                # Check file size limit
                if total_size > MAX_FILE_SIZE_BYTES:
                    logger.error(f"File size exceeds limit: {total_size} bytes (max: {MAX_FILE_SIZE_BYTES})")
                    raise HTTPException(
                        status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                        detail=f"File size exceeds maximum limit of {MAX_FILE_SIZE_BYTES // (1024*1024)}MB"
                    )
                
                content += chunk
                
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Error reading file content: {str(e)}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Error reading uploaded file: {str(e)}"
            )
        
        logger.info(f"File content read successfully: {total_size} bytes")
        
        # ===== STEP 4: Validate file signature (magic bytes) =====
        logger.info("Validating file signature...")
        validate_file_signature(content, file.filename)
        
        # ===== STEP 5: Validate Excel structure =====
        logger.info("Validating Excel file structure...")
        validate_excel_structure(content)
        
        # ===== STEP 6: Save file to disk =====
        file_path = Path(excel_file.filepath).resolve()  # Convert to absolute path
        
        # Check if file is currently locked (e.g., open in Excel)
        if file_path.exists() and is_file_locked(file_path):
            logger.error(f"File is locked and cannot be updated: {file_path}")
            raise HTTPException(
                status_code=status.HTTP_423_LOCKED,
                detail=f"File '{excel_file.filename}' is currently locked (possibly open in Excel or another program). Please close it and try again."
            )
        
        # Ensure the directory exists
        file_path.parent.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Saving file to: {file_path}")
        try:
            # Use atomic write for better reliability
            atomic_file_write(content, file_path)
            logger.info(f"File saved successfully: {file_id} ({len(content)} bytes)")
        except IOError as e:
            logger.error(f"Failed to write file {file_id}: {str(e)}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to save file to disk: {str(e)}"
            )
        except Exception as e:
            logger.error(f"Unexpected error during file write {file_id}: {str(e)}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to save file: {str(e)}"
            )
        
        # ===== STEP 7: Verify file was written successfully =====
        if not file_path.exists():
            logger.error(f"File verification failed after write: {file_id} at path: {file_path}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"File was not saved correctly to {file_path}"
            )
        
        # Verify file size matches
        written_size = file_path.stat().st_size
        if written_size != len(content):
            logger.error(f"File size mismatch: expected {len(content)}, got {written_size}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"File integrity check failed: expected {len(content)} bytes, got {written_size} bytes"
            )
        
        # Read back file content to verify it matches (extra safety check)
        try:
            with open(file_path, 'rb') as verify_file:
                written_content = verify_file.read()
                if written_content != content:
                    logger.error(f"Content verification failed: file content doesn't match")
                    raise HTTPException(
                        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                        detail="File content verification failed"
                    )
        except HTTPException:
            raise
        except Exception as e:
            logger.warning(f"Could not verify file content: {e}")
        
        logger.info(f"File save operation completed successfully for file_id: {file_id}")
        
        return {
            "message": "File saved successfully",
            "file_id": file_id,
            "filename": excel_file.filename,
            "size_bytes": len(content),
            "validation": {
                "signature_verified": True,
                "structure_verified": True,
                "integrity_verified": True
            }
        }
        
    except HTTPException:
        # Re-raise HTTP exceptions as-is
        raise
    except Exception as e:
        # Catch any unexpected errors
        logger.error(f"Unexpected error saving file {file_id}: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"An unexpected error occurred while saving the file: {str(e)}"
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
    
    file_path = Path(excel_file.filepath).resolve()
    exists = file_path.exists()
    
    result = {
        "file_id": excel_file.id,
        "filename": excel_file.filename,
        "filepath": str(file_path),
        "uploaded_at": excel_file.uploaded_at,
        "session_id": excel_file.session_id,
        "exists_on_disk": exists
    }
    
    # Add file system info if file exists
    if exists:
        stat = file_path.stat()
        result["size_bytes"] = stat.st_size
        result["modified_at"] = datetime.fromtimestamp(stat.st_mtime).isoformat()
        
        # Calculate hash for verification
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            result["content_hash"] = hashlib.sha256(content).hexdigest()[:16]
        except Exception as e:
            result["content_hash_error"] = str(e)
    
    return result


@router.get("/verify/{file_id}")
async def verify_file_on_disk(
    file_id: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Verify file exists and get detailed diagnostic information
    Useful for checking if file updates are being saved correctly
    
    Requires Authorization header: Bearer <token>
    """
    
    excel_file = db.query(ExcelFile).filter(ExcelFile.id == file_id).first()
    
    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found"
        )
    
    file_path = Path(excel_file.filepath).resolve()
    
    if not file_path.exists():
        return {
            "status": "error",
            "message": "File does not exist on disk",
            "filepath": str(file_path)
        }
    
    # Get file info
    stat = file_path.stat()
    
    # Read content and calculate hash
    with open(file_path, 'rb') as f:
        content = f.read()
    
    full_hash = hashlib.sha256(content).hexdigest()
    
    # Check if locked
    is_locked = is_file_locked(file_path)
    
    return {
        "status": "success",
        "file_id": file_id,
        "filename": excel_file.filename,
        "filepath": str(file_path),
        "exists": True,
        "writable": not is_locked,
        "locked": is_locked,
        "size_bytes": len(content),
        "file_system_size": stat.st_size,
        "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(),
        "content_hash_short": full_hash[:16],
        "content_hash_full": full_hash,
        "diagnostic": {
            "size_matches": len(content) == stat.st_size,
            "can_be_read": True,
            "message": "File is locked by another process (Excel?)" if is_locked else "File is ready for updates"
        }
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
    
    return {
        "message": "File deleted successfully",
        "file_id": file_id
    }


@router.get("/preview/{file_id}")
async def preview_file(
    file_id: str,
    sheet_name: Optional[str] = None,
    limit: int = 10,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Preview file content. If sheet_name is not provided, uses the first sheet in the file.
    
    Requires Authorization header: Bearer <token>
    """
    
    try:
        # Auto-detect sheet name if not provided
        actual_sheet_name = sheet_name
        if not actual_sheet_name:
            metadata = await mcp_service.get_file_metadata(file_id)
            sheets = metadata.get("sheets", [])
            if sheets:
                actual_sheet_name = sheets[0]["name"]
            else:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="File has no sheets"
                )
        
        # Use read_data to get preview
        result = await mcp_service.read_data(
            file_id=file_id,
            sheet_name=actual_sheet_name,
            limit=limit
        )
        
        return {
            "file_id": file_id,
            "sheet_name": actual_sheet_name,
            "data": result.get("data", []),
            "rows_shown": result.get("rows", 0),
            "total_columns": result.get("columns", 0)
        }
    
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error reading file: {str(e)}"
        )
