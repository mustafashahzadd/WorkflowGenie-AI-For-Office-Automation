"""
RAG service for retrieving Excel data context and generating grounded answers.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from loguru import logger

from app.core.config import settings

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

try:
    from anthropic import Anthropic
except ImportError:
    Anthropic = None

try:
    from sentence_transformers import SentenceTransformer
except ImportError:
    SentenceTransformer = None


class RAGService:
    """Retrieval-Augmented Generation over Excel files indexed in local storage."""

    SUPPORTED_EXTENSIONS: Tuple[str, ...] = ("xlsx", "xls", "xlsm")

    def __init__(self):
        self.excel_dir: Path = settings.EXCEL_DIR
        self.index_dir: Path = settings.RAG_INDEX_DIR
        self.index_dir.mkdir(parents=True, exist_ok=True)

        self.embeddings_file: Path = self.index_dir / "embeddings.npy"
        self.metadata_file: Path = self.index_dir / "metadata.json"

        self.embedding_model = None
        self.openai_client = None
        self.claude_client = None

        self.embeddings: Optional[np.ndarray] = None
        self.documents: List[Dict[str, Any]] = []
        self.last_built_at: Optional[str] = None

    def _normalize_vectors(self, vectors: np.ndarray) -> np.ndarray:
        """Normalize vectors so cosine similarity can be computed via dot product."""

        norms = np.linalg.norm(vectors, axis=1, keepdims=True)
        norms[norms == 0] = 1.0
        return vectors / norms

    def _encode_with_sentence_transformers(self, texts: List[str]) -> np.ndarray:
        """Generate dense vectors locally with SentenceTransformers."""

        if SentenceTransformer is None:
            raise RuntimeError(
                "sentence-transformers is not installed. Install dependencies from requirements.txt."
            )

        if self.embedding_model is None:
            self.embedding_model = SentenceTransformer(settings.RAG_EMBEDDING_MODEL)

        vectors = self.embedding_model.encode(
            texts,
            convert_to_numpy=True,
            show_progress_bar=False,
            normalize_embeddings=True,
        )

        return vectors.astype(np.float32)

    def _encode_with_openai(self, texts: List[str]) -> np.ndarray:
        """Generate vectors using OpenAI embedding endpoint."""

        if OpenAI is None:
            raise RuntimeError("OpenAI SDK is not installed.")
        if not settings.OPENAI_API_KEY:
            raise RuntimeError("OPENAI_API_KEY is not configured for OpenAI embeddings.")

        if self.openai_client is None:
            self.openai_client = OpenAI(api_key=settings.OPENAI_API_KEY)

        vectors: List[List[float]] = []
        batch_size = 100

        for start in range(0, len(texts), batch_size):
            batch = texts[start : start + batch_size]
            response = self.openai_client.embeddings.create(
                model=settings.OPENAI_EMBEDDING_MODEL,
                input=batch,
            )
            vectors.extend(item.embedding for item in response.data)

        vector_array = np.array(vectors, dtype=np.float32)
        return self._normalize_vectors(vector_array)

    def _encode_texts(self, texts: List[str]) -> np.ndarray:
        """Encode texts using configured embedding provider with local fallback."""

        provider = (settings.RAG_EMBEDDING_PROVIDER or "openai").strip().lower()

        if provider == "openai":
            try:
                return self._encode_with_openai(texts)
            except Exception as exc:
                logger.warning(f"OpenAI embeddings unavailable, falling back to local model: {exc}")
                return self._encode_with_sentence_transformers(texts)

        return self._encode_with_sentence_transformers(texts)

    def _extract_file_id(self, file_path: Path) -> str:
        """Extract file_id prefix from <file_id>_<filename>.<extension>."""

        prefix, _, _ = file_path.stem.partition("_")
        return prefix

    def _display_filename(self, file_path: Path) -> str:
        """Return original filename without UUID prefix."""

        _, sep, rest = file_path.stem.partition("_")
        if sep and rest:
            return rest + file_path.suffix
        return file_path.name

    def _chunk_text(self, text: str) -> List[str]:
        """Split text into overlapping chunks for better retrieval."""

        chunk_size = max(120, settings.RAG_CHUNK_SIZE)
        overlap = max(0, min(settings.RAG_CHUNK_OVERLAP, chunk_size - 20))

        if len(text) <= chunk_size:
            return [text]

        step = chunk_size - overlap
        chunks: List[str] = []
        start = 0

        while start < len(text):
            end = start + chunk_size
            chunks.append(text[start:end])
            if end >= len(text):
                break
            start += step

        return chunks

    def _normalize_headers(self, header_values: List[Any]) -> List[str]:
        """Build stable header names and replace blank/unnamed headers."""

        normalized: List[str] = []
        seen: Dict[str, int] = {}

        for idx, raw_value in enumerate(header_values, start=1):
            header = ""
            if raw_value is not None:
                header = str(raw_value).strip()

            if not header or header.lower().startswith("unnamed:"):
                header = f"Column_{idx}"

            base_header = header
            count = seen.get(base_header, 0) + 1
            seen[base_header] = count

            if count > 1:
                header = f"{base_header}_{count}"

            normalized.append(header)

        return normalized

    def _prepare_sheet_dataframe(
        self,
        raw_df: pd.DataFrame,
    ) -> Optional[Tuple[pd.DataFrame, List[str], int]]:
        """Trim empty margins and produce a row-wise dataframe with detected headers."""

        trimmed_rows = raw_df.dropna(how="all")
        if trimmed_rows.empty:
            return None

        trimmed = trimmed_rows.dropna(axis=1, how="all")
        if trimmed.empty:
            return None

        header_row_idx = int(trimmed.index[0])
        header_values = trimmed.iloc[0].tolist()
        columns = self._normalize_headers(header_values)

        data_df = trimmed.iloc[1:].copy()
        if data_df.empty:
            return None

        data_df.columns = columns
        return data_df, columns, header_row_idx + 1

    def _load_excel_documents(
        self,
        file_path: Path,
        sheet_name: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Load and convert workbook rows into retrievable text chunks."""

        # Read raw sheet cells so tables that start away from A1 can be detected reliably.
        workbook = pd.read_excel(file_path, sheet_name=None, dtype=str, header=None)
        if not workbook:
            return []

        if sheet_name and sheet_name not in workbook:
            raise ValueError(f"Sheet '{sheet_name}' not found in {self._display_filename(file_path)}")

        documents: List[Dict[str, Any]] = []
        file_id = self._extract_file_id(file_path)
        filename = self._display_filename(file_path)

        for current_sheet, df in workbook.items():
            if sheet_name and current_sheet != sheet_name:
                continue

            prepared = self._prepare_sheet_dataframe(df)
            if not prepared:
                continue

            safe_df, columns, header_row_number = prepared
            safe_df = safe_df.fillna("")

            schema_text = (
                f"File: {filename}\n"
                f"Sheet: {current_sheet}\n"
                f"Header Row: {header_row_number}\n"
                f"Columns: {', '.join(columns)}"
            )
            documents.append(
                {
                    "file_id": file_id,
                    "filename": filename,
                    "sheet_name": current_sheet,
                    "row_number": None,
                    "text": schema_text,
                }
            )

            max_rows = min(len(safe_df.index), settings.RAG_MAX_ROWS_PER_SHEET)
            for excel_row_idx, (source_row_idx, row) in enumerate(safe_df.iterrows()):
                if excel_row_idx >= max_rows:
                    break

                parts: List[str] = []

                for col in safe_df.columns:
                    value = str(row[col]).strip()
                    if value:
                        parts.append(f"{col}: {value}")

                if not parts:
                    continue

                row_number = int(source_row_idx) + 1
                row_text = (
                    f"File: {filename}\n"
                    f"Sheet: {current_sheet}\n"
                    f"Excel Row: {row_number}\n"
                    f"Data: {' | '.join(parts)}"
                )

                for chunk in self._chunk_text(row_text):
                    documents.append(
                        {
                            "file_id": file_id,
                            "filename": filename,
                            "sheet_name": current_sheet,
                            "row_number": row_number,
                            "text": chunk,
                        }
                    )

        return documents

    def _save_index_to_disk(self) -> None:
        """Persist embeddings and metadata for warm restarts."""

        if self.embeddings is None:
            return

        np.save(self.embeddings_file, self.embeddings)

        payload = {
            "built_at": self.last_built_at,
            "embedding_provider": settings.RAG_EMBEDDING_PROVIDER,
            "embedding_model": settings.RAG_EMBEDDING_MODEL,
            "openai_embedding_model": settings.OPENAI_EMBEDDING_MODEL,
            "documents": self.documents,
        }

        self.metadata_file.write_text(
            json.dumps(payload, indent=2, ensure_ascii=True),
            encoding="utf-8",
        )

    def _clear_index(self) -> None:
        """Remove in-memory and on-disk index artifacts."""

        self.embeddings = None
        self.documents = []
        self.last_built_at = None

        for index_file in (self.embeddings_file, self.metadata_file):
            try:
                if index_file.exists():
                    index_file.unlink()
            except Exception as exc:
                logger.warning(f"Failed to remove index file {index_file}: {exc}")

        logger.info("Cleared RAG index because no Excel files are available")

    def _load_index_from_disk(self) -> bool:
        """Load previously built index files if present."""

        if not self.embeddings_file.exists() or not self.metadata_file.exists():
            return False

        try:
            embeddings = np.load(self.embeddings_file)
            payload = json.loads(self.metadata_file.read_text(encoding="utf-8"))
            documents = payload.get("documents", [])

            if embeddings.shape[0] != len(documents):
                logger.warning("RAG index metadata mismatch, rebuilding is required")
                return False

            self.embeddings = embeddings.astype(np.float32)
            self.documents = documents
            self.last_built_at = payload.get("built_at")

            logger.info(f"Loaded RAG index from disk with {len(self.documents)} chunks")
            return True
        except Exception as exc:
            logger.warning(f"Failed to load RAG index from disk: {exc}")
            return False

    def _resolve_target_files(self, file_id: Optional[str] = None) -> List[Path]:
        """Resolve file paths to index."""

        def _is_valid_candidate(path: Path) -> bool:
            # Excel creates temporary lock files prefixed with "~$" that are not readable workbooks.
            return not path.name.startswith("~$")

        if file_id:
            matches: List[Path] = []
            for extension in self.SUPPORTED_EXTENSIONS:
                matches.extend(self.excel_dir.glob(f"{file_id}_*.{extension}"))
            matches = sorted(path for path in matches if _is_valid_candidate(path))
            if not matches:
                raise ValueError(f"File with ID '{file_id}' not found")
            return matches

        files: List[Path] = []
        for extension in self.SUPPORTED_EXTENSIONS:
            files.extend(self.excel_dir.glob(f"*.{extension}"))

        return sorted(path for path in files if _is_valid_candidate(path))

    def ensure_index_ready(self) -> None:
        """Ensure index exists in memory by loading or rebuilding."""

        if self.embeddings is not None and self.documents:
            return

        if self._load_index_from_disk():
            return

        self.rebuild_index()

    def refresh_index(self) -> Dict[str, Any]:
        """Synchronize index with current Excel files after file changes."""

        file_paths = self._resolve_target_files()
        if not file_paths:
            self._clear_index()
            return {
                "indexed_chunks": 0,
                "indexed_files": 0,
                "built_at": None,
                "cleared": True,
            }

        result = self.rebuild_index()
        result["cleared"] = False
        return result

    def rebuild_index(
        self,
        file_id: Optional[str] = None,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Build or rebuild RAG index from Excel files."""

        file_paths = self._resolve_target_files(file_id=file_id)
        if not file_paths:
            raise ValueError("No Excel files found to index")

        documents: List[Dict[str, Any]] = []
        skipped_files: List[str] = []
        for file_path in file_paths:
            try:
                documents.extend(self._load_excel_documents(file_path, sheet_name=sheet_name))
            except Exception as exc:
                skipped_files.append(file_path.name)
                logger.warning(f"Skipping unreadable workbook '{file_path.name}': {exc}")

        if not documents:
            raise ValueError("No readable content found in selected Excel files")

        texts = [doc["text"] for doc in documents]
        embeddings = self._encode_texts(texts)

        if embeddings.shape[0] != len(documents):
            raise RuntimeError("Embedding count does not match indexed chunks")

        self.embeddings = embeddings
        self.documents = documents
        self.last_built_at = datetime.utcnow().isoformat()
        self._save_index_to_disk()

        unique_files = {doc["file_id"] for doc in self.documents}

        logger.info(
            f"RAG index built: {len(self.documents)} chunks across {len(unique_files)} files"
        )
        if skipped_files:
            logger.warning(f"RAG index skipped unreadable files: {', '.join(skipped_files)}")

        return {
            "indexed_chunks": len(self.documents),
            "indexed_files": len(unique_files),
            "built_at": self.last_built_at,
        }

    def _resolve_provider(self, provider: Optional[str]) -> str:
        """Resolve LLM provider for answer generation."""

        active_provider = (provider or settings.LLM_PROVIDER or "openai").strip().lower()
        if active_provider not in {"openai", "claude"}:
            raise ValueError("Invalid provider. Use 'openai' or 'claude'.")
        return active_provider

    def _fallback_answer(self, question: str, retrieved_chunks: List[Dict[str, Any]]) -> str:
        """Return a deterministic fallback answer when no LLM client is configured."""

        lines = [
            "LLM provider is not configured. Here are the closest matching data snippets:",
        ]

        for item in retrieved_chunks[:3]:
            row_label = item["row_number"] if item["row_number"] is not None else "-"
            preview = item["text"].replace("\n", " ")[:240]
            lines.append(
                f"- {item['filename']} | {item['sheet_name']} | row {row_label}: {preview}"
            )

        lines.append(f"Original question: {question}")
        return "\n".join(lines)

    def _generate_answer(
        self,
        question: str,
        context: str,
        retrieved_chunks: List[Dict[str, Any]],
        provider: Optional[str] = None,
    ) -> str:
        """Generate final grounded answer from retrieved context."""

        active_provider = self._resolve_provider(provider)

        system_prompt = (
            "You are WorkflowGenie RAG assistant. "
            "Answer strictly from the provided context chunks. "
            "If the answer is not in context, say you could not find it in indexed data."
        )
        user_prompt = (
            f"Question:\n{question}\n\n"
            f"Context Chunks:\n{context}\n\n"
            "Return a concise answer and mention the most relevant sheet/row references."
        )

        try:
            if active_provider == "openai" and OpenAI and settings.OPENAI_API_KEY:
                if self.openai_client is None:
                    self.openai_client = OpenAI(api_key=settings.OPENAI_API_KEY)

                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.2,
                    max_completion_tokens=1000,
                )
                return response.choices[0].message.content or "No answer generated."

            if active_provider == "claude" and Anthropic and settings.ANTHROPIC_API_KEY:
                if self.claude_client is None:
                    self.claude_client = Anthropic(api_key=settings.ANTHROPIC_API_KEY)

                response = self.claude_client.messages.create(
                    model=settings.CLAUDE_MODEL,
                    system=system_prompt,
                    messages=[{"role": "user", "content": user_prompt}],
                    temperature=0.2,
                    max_tokens=1000,
                )

                parts: List[str] = []
                for block in getattr(response, "content", []) or []:
                    if isinstance(block, dict):
                        if block.get("type") == "text" and block.get("text"):
                            parts.append(str(block["text"]))
                    elif getattr(block, "type", None) == "text":
                        block_text = getattr(block, "text", None)
                        if block_text:
                            parts.append(str(block_text))

                return "\n".join(parts).strip() or "No answer generated."

        except Exception as exc:
            logger.error(f"RAG answer generation failed ({active_provider}): {exc}")

        return self._fallback_answer(question, retrieved_chunks)

    def _filtered_indices(
        self,
        file_id: Optional[str] = None,
        sheet_name: Optional[str] = None,
    ) -> List[int]:
        """Get candidate indices for optional file/sheet filters."""

        indices: List[int] = []
        for idx, doc in enumerate(self.documents):
            if file_id and doc.get("file_id") != file_id:
                continue
            if sheet_name and doc.get("sheet_name") != sheet_name:
                continue
            indices.append(idx)
        return indices

    def retrieve(
        self,
        question: str,
        top_k: Optional[int] = None,
        file_id: Optional[str] = None,
        sheet_name: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Retrieve most relevant chunks for a question."""

        if not question or not question.strip():
            raise ValueError("Question cannot be empty")

        self.ensure_index_ready()

        if self.embeddings is None or not self.documents:
            raise ValueError("RAG index is empty. Rebuild index first.")

        requested_k = top_k or settings.RAG_TOP_K
        active_top_k = max(1, min(requested_k, 20))

        candidate_indices = self._filtered_indices(file_id=file_id, sheet_name=sheet_name)
        if not candidate_indices:
            raise ValueError("No indexed chunks matched requested file/sheet filters")

        query_vector = self._encode_texts([question])[0]
        candidate_matrix = self.embeddings[candidate_indices]
        scores = candidate_matrix @ query_vector

        ranked_positions = np.argsort(scores)[::-1][:active_top_k]

        results: List[Dict[str, Any]] = []
        for position in ranked_positions:
            doc_idx = candidate_indices[int(position)]
            item = dict(self.documents[doc_idx])
            item["score"] = float(scores[int(position)])
            results.append(item)

        return results

    def answer_query(
        self,
        question: str,
        top_k: Optional[int] = None,
        file_id: Optional[str] = None,
        sheet_name: Optional[str] = None,
        provider: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Run retrieval and generation for a user query."""

        retrieved = self.retrieve(
            question=question,
            top_k=top_k,
            file_id=file_id,
            sheet_name=sheet_name,
        )

        context_parts: List[str] = []
        sources: List[Dict[str, Any]] = []

        for idx, chunk in enumerate(retrieved, start=1):
            row_label = chunk["row_number"] if chunk["row_number"] is not None else "-"
            context_parts.append(
                f"[Source {idx}] {chunk['filename']} | {chunk['sheet_name']} | row {row_label}\n"
                f"{chunk['text']}"
            )
            sources.append(
                {
                    "file_id": chunk["file_id"],
                    "filename": chunk["filename"],
                    "sheet_name": chunk["sheet_name"],
                    "row_number": chunk["row_number"],
                    "score": chunk["score"],
                }
            )

        context = "\n\n".join(context_parts)
        answer = self._generate_answer(
            question=question,
            context=context,
            retrieved_chunks=retrieved,
            provider=provider,
        )

        return {
            "answer": answer,
            "sources": sources,
            "retrieved_chunks": len(retrieved),
            "indexed_chunks": len(self.documents),
            "built_at": self.last_built_at,
        }

    def get_status(self) -> Dict[str, Any]:
        """Return current RAG index status."""

        if self.embeddings is None or not self.documents:
            self._load_index_from_disk()

        indexed_files = len({doc.get("file_id") for doc in self.documents}) if self.documents else 0

        return {
            "enabled": settings.RAG_ENABLED,
            "embedding_provider": settings.RAG_EMBEDDING_PROVIDER,
            "embedding_model": settings.RAG_EMBEDDING_MODEL,
            "openai_embedding_model": settings.OPENAI_EMBEDDING_MODEL,
            "indexed_chunks": len(self.documents),
            "indexed_files": indexed_files,
            "built_at": self.last_built_at,
            "index_dir": str(self.index_dir),
        }


rag_service = RAGService()
