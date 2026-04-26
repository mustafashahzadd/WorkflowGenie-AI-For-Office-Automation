"""
Configuration settings for WorkflowGenie
"""

from pydantic_settings import BaseSettings
from typing import List, Optional
from pathlib import Path

class Settings(BaseSettings):
    """Application settings"""
    
    # API Settings
    APP_NAME: str = "WorkflowGenie"
    VERSION: str = "1.0.0"
    DEBUG: bool = True
    HOST: str = "0.0.0.0"
    PORT: int = 8000
    
    # Database
    DATABASE_URL: str = "sqlite:///./workflowgenie.db"
    
    # LLM Provider — Claude is primary
    LLM_PROVIDER: str = "claude"

    # OpenAI (kept as fallback — uncomment in llm_service if needed)
    OPENAI_API_KEY: Optional[str] = None
    OPENAI_MODEL: str = "gpt-4o"
    OPENAI_EMBEDDING_MODEL: str = "text-embedding-3-small"

    # Claude (Anthropic) — Primary provider
    ANTHROPIC_API_KEY: Optional[str] = None
    CLAUDE_MODEL: str = "claude-sonnet-4-20250514"  # default
    CLAUDE_MODEL_SONNET: str = "claude-sonnet-4-20250514"
    CLAUDE_MODEL_OPUS: str = "claude-opus-4-20250514"
    
    # CORS
    CORS_ORIGINS: List[str] = ["*"]
    
    # File Storage
    DATA_DIR: Path = Path("./data")
    EXCEL_DIR: Path = Path("./data/excel_files")

    # RAG
    RAG_INDEX_DIR: Path = Path("./data/rag_index")
    RAG_EMBEDDING_PROVIDER: str = "sentence_transformers"
    RAG_EMBEDDING_MODEL: str = "sentence-transformers/all-MiniLM-L6-v2"
    RAG_MAX_ROWS_PER_SHEET: int = 1000
    RAG_CHUNK_SIZE: int = 500
    RAG_CHUNK_OVERLAP: int = 50
    RAG_ENABLED: bool = True
    RAG_WRITE_ENABLED: bool = False
    RAG_WRITE_TOP_K: int = 3
    RAG_WRITE_SCORE_THRESHOLD: float = 0.5

    # Logging
    LOG_LEVEL: str = "INFO"
    
    class Config:
        env_file = ".env"
        case_sensitive = True

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        # Create directories if they don't exist
        self.DATA_DIR.mkdir(parents=True, exist_ok=True)
        self.EXCEL_DIR.mkdir(parents=True, exist_ok=True)
        self.RAG_INDEX_DIR.mkdir(parents=True, exist_ok=True)

# Create settings instance
settings = Settings()
