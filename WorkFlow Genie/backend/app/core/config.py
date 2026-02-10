"""
Configuration settings for WorkflowGenie
"""

from pydantic_settings import BaseSettings
from typing import List
import os
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
    
    # OpenAI
    OPENAI_API_KEY: str
    OPENAI_MODEL: str = "gpt-4-turbo-preview"
    
    # CORS
    CORS_ORIGINS: List[str] = ["*"]
    
    # File Storage
    DATA_DIR: Path = Path("./data")
    EXCEL_DIR: Path = Path("./data/excel_files")
    
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

# Create settings instance
settings = Settings()