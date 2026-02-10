"""
Database models for WorkflowGenie
Extended with User authentication and Task tracking
"""

from sqlalchemy import Column, String, DateTime, Boolean, Integer, ForeignKey, Text, JSON
from sqlalchemy.orm import relationship
from datetime import datetime
import uuid

from app.core.database import Base

def generate_uuid():
    return str(uuid.uuid4())

# ============================================================================
# USER MODEL (NEW)
# ============================================================================

class User(Base):
    """User authentication model"""
    __tablename__ = "users"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    username = Column(String, unique=True, nullable=False, index=True)
    password_hash = Column(String, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True)
    
    # Relationships
    sessions = relationship("Session", back_populates="user", cascade="all, delete-orphan")
    tasks = relationship("Task", back_populates="user", cascade="all, delete-orphan")


# ============================================================================
# SESSION MODEL (UPDATED - linked to User)
# ============================================================================

class Session(Base):
    """Chat session model"""
    __tablename__ = "sessions"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    user_id = Column(String, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)  # Now required
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    is_active = Column(Boolean, default=True)
    summary = Column(Text, nullable=True)
    
    # Relationships
    user = relationship("User", back_populates="sessions")
    messages = relationship("Message", back_populates="session", cascade="all, delete-orphan")
    operations = relationship("Operation", back_populates="session", cascade="all, delete-orphan")
    tasks = relationship("Task", back_populates="session", cascade="all, delete-orphan")
    excel_file = relationship("ExcelFile", back_populates="session", uselist=False, cascade="all, delete-orphan")  # ONE-TO-ONE


# ============================================================================
# MESSAGE MODEL (unchanged)
# ============================================================================

class Message(Base):
    """Chat message model"""
    __tablename__ = "messages"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"))
    role = Column(String, nullable=False)
    content = Column(Text, nullable=False)
    timestamp = Column(DateTime, default=datetime.utcnow)
    meta_data = Column(Text, nullable=True)
    
    session = relationship("Session", back_populates="messages")


# ============================================================================
# OPERATION MODEL (unchanged)
# ============================================================================

class Operation(Base):
    """Excel operation log model"""
    __tablename__ = "operations"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"))
    tool_name = Column(String, nullable=False)
    input_params = Column(Text, nullable=False)
    output_result = Column(Text, nullable=True)
    status = Column(String, nullable=False)
    timestamp = Column(DateTime, default=datetime.utcnow)
    duration = Column(Integer, nullable=True)
    error_message = Column(Text, nullable=True)
    
    session = relationship("Session", back_populates="operations")


# ============================================================================
# EXCEL FILE MODEL (unchanged)
# ============================================================================

class ExcelFile(Base):
    """Excel file metadata model - ONE file per session"""
    __tablename__ = "excel_files"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    filename = Column(String, nullable=False)
    filepath = Column(String, nullable=False)
    uploaded_at = Column(DateTime, default=datetime.utcnow)
    meta_info = Column(Text, nullable=True)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"), unique=True, nullable=False)  # ONE-TO-ONE with UNIQUE constraint
    
    # Relationship
    session = relationship("Session", back_populates="excel_file")


# ============================================================================
# TASK MODEL (NEW - for async task tracking)
# ============================================================================

class Task(Base):
    """Async task tracking model"""
    __tablename__ = "tasks"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    user_id = Column(String, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"), nullable=True)
    
    # Task details
    prompt = Column(Text, nullable=False)
    uploaded_file_id = Column(String, nullable=True)  # Original file user uploaded
    result_file_id = Column(String, nullable=True)    # Modified file to download
    
    # Status tracking
    status = Column(String, default="pending")  # pending, processing, completed, failed
    progress = Column(Integer, default=0)  # 0-100
    current_step = Column(String, nullable=True)  # Current operation description
    
    # Results
    operations = Column(JSON, nullable=True)  # List of operations performed
    response = Column(Text, nullable=True)    # Final response text
    error = Column(Text, nullable=True)       # Error message if failed
    
    # Timestamps
    created_at = Column(DateTime, default=datetime.utcnow)
    started_at = Column(DateTime, nullable=True)
    completed_at = Column(DateTime, nullable=True)
    
    # Relationships
    user = relationship("User", back_populates="tasks")
    session = relationship("Session", back_populates="tasks")
