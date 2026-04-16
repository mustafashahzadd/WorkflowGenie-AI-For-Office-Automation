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
    
    # LLM Provider
    LLM_PROVIDER: str = "openai"

    # OpenAI
    OPENAI_API_KEY: Optional[str]
    OPENAI_MODEL: str = "gpt-4o"
    OPENAI_EMBEDDING_MODEL: str = "text-embedding-3-small"

    # Claude (Anthropic)
    ANTHROPIC_API_KEY: Optional[str]
    CLAUDE_MODEL: str = "claude-opus-4-6"
    
    # CORS
    CORS_ORIGINS: List[str] = ["*"]
    
    # File Storage
    DATA_DIR: Path = Path("./data")
    EXCEL_DIR: Path = Path("./data/excel_files")

    # RAG
    RAG_ENABLED: bool = True
    RAG_INDEX_DIR: Path = Path("./data/rag_index")
    RAG_EMBEDDING_PROVIDER: str = "openai"  # openai | sentence_transformers
    RAG_EMBEDDING_MODEL: str = "sentence-transformers/all-MiniLM-L6-v2"
    RAG_CHUNK_SIZE: int = 600
    RAG_CHUNK_OVERLAP: int = 80
    RAG_TOP_K: int = 4
    RAG_MAX_ROWS_PER_SHEET: int = 2000
    
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