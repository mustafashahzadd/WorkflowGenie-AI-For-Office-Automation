"""
RAG query endpoint — read-only, completely separate from chat pipeline.
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field
from typing import Optional
from loguru import logger

from app.services.rag_service import rag_service

router = APIRouter(prefix="/api/rag", tags=["rag"])


class RAGQueryRequest(BaseModel):
    question: str = Field(..., description="Natural language question about the Excel data")
    file_id: Optional[str] = Field(None, description="Limit search to a specific file")
    top_k: int = Field(5, description="Number of chunks to retrieve")


@router.post("/query")
async def rag_query(request: RAGQueryRequest):
    """Answer a question using RAG over indexed Excel files."""
    try:
        rag_service.ensure_index_ready()
        result = rag_service.answer_query(
            question=request.question,
            file_id=request.file_id,
            top_k=request.top_k,
        )
        return result
    except Exception as e:
        logger.error(f"RAG query error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/index/rebuild")
async def rebuild_index():
    """Rebuild the RAG index from all Excel files."""
    try:
        result = rag_service.rebuild_index()
        return result
    except Exception as e:
        logger.error(f"RAG index rebuild error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/status")
async def rag_status():
    """Get RAG index status."""
    return rag_service.get_status()
