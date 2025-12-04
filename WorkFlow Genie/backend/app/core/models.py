"""
Database models for WorkflowGenie
"""

from sqlalchemy import Column, String, DateTime, Boolean, Integer, ForeignKey, Text
from sqlalchemy.orm import relationship
from datetime import datetime
import uuid

from app.core.database import Base

def generate_uuid():
    return str(uuid.uuid4())

class Session(Base):
    """Chat session model"""
    __tablename__ = "sessions"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    user_id = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    is_active = Column(Boolean, default=True)
    summary = Column(Text, nullable=True)
    
    messages = relationship("Message", back_populates="session", cascade="all, delete-orphan")
    operations = relationship("Operation", back_populates="session", cascade="all, delete-orphan")

class Message(Base):
    """Chat message model"""
    __tablename__ = "messages"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"))
    role = Column(String, nullable=False)
    content = Column(Text, nullable=False)
    timestamp = Column(DateTime, default=datetime.utcnow)
    meta_data = Column(Text, nullable=True)  # Changed from 'metadata' to 'meta_data'
    
    session = relationship("Session", back_populates="messages")

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

class ExcelFile(Base):
    """Excel file metadata model"""
    __tablename__ = "excel_files"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    filename = Column(String, nullable=False)
    filepath = Column(String, nullable=False)
    uploaded_at = Column(DateTime, default=datetime.utcnow)
    meta_info = Column(Text, nullable=True)  # Changed from 'metadata' to 'meta_info'
    session_id = Column(String, nullable=True)