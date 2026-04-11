"""
RAG API routes for grounded Q&A over indexed Excel data.
"""

from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException
from loguru import logger
from pydantic import BaseModel, Field

from app.core.config import settings
from app.services.rag_service import rag_service

router = APIRouter()


class RAGQueryRequest(BaseModel):
    """Request model for grounded RAG query."""

    question: str = Field(..., min_length=2, description="User question")
    file_id: Optional[str] = Field(None, description="Optional file ID filter")
    sheet_name: Optional[str] = Field(None, description="Optional sheet filter")
    top_k: int = Field(default=4, ge=1, le=20, description="Number of chunks to retrieve")
    provider: Optional[str] = Field(None, description="LLM provider override: openai or claude")


class RAGSource(BaseModel):
    """Source details for retrieved context."""

    file_id: str
    filename: str
    sheet_name: str
    row_number: Optional[int]
    score: float


class RAGQueryResponse(BaseModel):
    """Response model for RAG query."""

    answer: str
    sources: List[RAGSource]
    retrieved_chunks: int
    indexed_chunks: int
    built_at: Optional[str]


class RAGRebuildRequest(BaseModel):
    """Request model for rebuilding index."""

    file_id: Optional[str] = Field(None, description="Optional file ID to limit indexing")
    sheet_name: Optional[str] = Field(None, description="Optional sheet name to limit indexing")


@router.get("/status", summary="Get RAG status", tags=["RAG"])
async def get_rag_status() -> Dict[str, Any]:
    """Return current RAG index metadata and readiness details."""

    return rag_service.get_status()


@router.post(
    "/index/rebuild",
    summary="Rebuild RAG index",
    tags=["RAG"],
)
async def rebuild_rag_index(request: RAGRebuildRequest) -> Dict[str, Any]:
    """Rebuild retrieval index from Excel files."""

    if not settings.RAG_ENABLED:
        raise HTTPException(status_code=503, detail="RAG is disabled in server configuration")

    try:
        result = rag_service.rebuild_index(file_id=request.file_id, sheet_name=request.sheet_name)
        return {
            "success": True,
            "message": "RAG index rebuilt successfully",
            **result,
        }
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        logger.error(f"RAG rebuild failed: {exc}")
        raise HTTPException(status_code=500, detail="Failed to rebuild RAG index")


@router.post(
    "/query",
    response_model=RAGQueryResponse,
    summary="Query Excel data with RAG",
    tags=["RAG"],
)
async def rag_query(request: RAGQueryRequest) -> RAGQueryResponse:
    """Retrieve relevant data chunks and generate a grounded answer."""

    if not settings.RAG_ENABLED:
        raise HTTPException(status_code=503, detail="RAG is disabled in server configuration")

    try:
        result = rag_service.answer_query(
            question=request.question,
            top_k=request.top_k,
            file_id=request.file_id,
            sheet_name=request.sheet_name,
            provider=request.provider,
        )
        return RAGQueryResponse(**result)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        logger.error(f"RAG query failed: {exc}")
        raise HTTPException(status_code=500, detail="Failed to process RAG query")
