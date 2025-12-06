"""
Create all database tables for WorkflowGenie
Run this after any model changes
"""

from sqlalchemy import create_engine, Column, String, DateTime, Boolean, Integer, ForeignKey, Text, JSON
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime
import uuid

Base = declarative_base()
engine = create_engine('sqlite:///workflowgenie.db')

def generate_uuid():
    return str(uuid.uuid4())

# ============================================================================
# USER MODEL
# ============================================================================

class User(Base):
    __tablename__ = "users"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    username = Column(String, unique=True, nullable=False, index=True)
    password_hash = Column(String, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True)
    
    sessions = relationship("Session", back_populates="user", cascade="all, delete-orphan")
    tasks = relationship("Task", back_populates="user", cascade="all, delete-orphan")

# ============================================================================
# SESSION MODEL
# ============================================================================

class Session(Base):
    __tablename__ = "sessions"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    user_id = Column(String, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    is_active = Column(Boolean, default=True)
    summary = Column(Text, nullable=True)
    
    user = relationship("User", back_populates="sessions")
    messages = relationship("Message", back_populates="session", cascade="all, delete-orphan")
    operations = relationship("Operation", back_populates="session", cascade="all, delete-orphan")
    tasks = relationship("Task", back_populates="session", cascade="all, delete-orphan")
    excel_file = relationship("ExcelFile", back_populates="session", uselist=False, cascade="all, delete-orphan")

# ============================================================================
# MESSAGE MODEL
# ============================================================================

class Message(Base):
    __tablename__ = "messages"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"))
    role = Column(String, nullable=False)
    content = Column(Text, nullable=False)
    timestamp = Column(DateTime, default=datetime.utcnow)
    meta_data = Column(Text, nullable=True)
    
    session = relationship("Session", back_populates="messages")

# ============================================================================
# OPERATION MODEL
# ============================================================================

class Operation(Base):
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
# EXCEL FILE MODEL (ONE-TO-ONE with Session)
# ============================================================================

class ExcelFile(Base):
    __tablename__ = "excel_files"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    filename = Column(String, nullable=False)
    filepath = Column(String, nullable=False)
    uploaded_at = Column(DateTime, default=datetime.utcnow)
    meta_info = Column(Text, nullable=True)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"), unique=True, nullable=False)  # ONE-TO-ONE
    
    session = relationship("Session", back_populates="excel_file")

# ============================================================================
# TASK MODEL
# ============================================================================

class Task(Base):
    __tablename__ = "tasks"
    
    id = Column(String, primary_key=True, default=generate_uuid)
    user_id = Column(String, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    session_id = Column(String, ForeignKey("sessions.id", ondelete="CASCADE"), nullable=False)
    
    prompt = Column(Text, nullable=False)
    uploaded_file_id = Column(String, nullable=True)
    result_file_id = Column(String, nullable=True)
    
    status = Column(String, default="pending")
    progress = Column(Integer, default=0)
    current_step = Column(String, nullable=True)
    
    operations = Column(JSON, nullable=True)
    response = Column(Text, nullable=True)
    error = Column(Text, nullable=True)
    
    created_at = Column(DateTime, default=datetime.utcnow)
    started_at = Column(DateTime, nullable=True)
    completed_at = Column(DateTime, nullable=True)
    
    user = relationship("User", back_populates="tasks")
    session = relationship("Session", back_populates="tasks")

# ============================================================================
# CREATE ALL TABLES
# ============================================================================

print("Creating database tables...")
Base.metadata.create_all(engine)
print("✅ ALL 6 tables created successfully!")

# Verify
import sqlite3
conn = sqlite3.connect('workflowgenie.db')
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()
print(f"\n📋 Tables in database: {[t[0] for t in tables]}")
conn.close()

print("\n🎉 Database ready for use!")