"""
WorkflowGenie Backend - FastAPI Server
Main entry point for the application
"""

from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
import uvicorn
from loguru import logger

from app.core.config import settings
from app.core.database import engine, Base
from app.core import models  # Import models to register them

# Import all routers (REMOVED tasks, ADDED history)
from app.api.routes import auth, users, files, chat, sessions, excel, history

# WebSocket connection manager
class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)
        logger.info(f"Client connected. Total connections: {len(self.active_connections)}")

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)
        logger.info(f"Client disconnected. Total connections: {len(self.active_connections)}")

    async def broadcast(self, message: dict):
        """Broadcast message to all connected clients"""
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except Exception as e:
                logger.error(f"Error broadcasting to client: {e}")

manager = ConnectionManager()

# Lifespan context manager
@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    logger.info("🚀 Starting WorkflowGenie Backend...")
    
    # Create database tables
    Base.metadata.create_all(bind=engine)
    logger.info("📊 Database tables created")
    
    yield
    
    # Shutdown
    logger.info("🛑 Shutting down WorkflowGenie Backend...")

# Create FastAPI app
app = FastAPI(
    title="WorkflowGenie API",
    description="AI-powered Excel automation via natural language",
    version="1.0.0",
    lifespan=lifespan
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================================
# INCLUDE ROUTERS
# ============================================================================

# Authentication & User Management
app.include_router(auth.router)          # /api/auth/register, /api/auth/login
app.include_router(users.router)         # /api/users/sessions, /api/users/sessions/{id}/messages

# File Management
app.include_router(files.router)         # /api/files/upload, /api/files/download/{id}

# Session Management
app.include_router(sessions.router)      # /api/sessions/create, /api/sessions/{id}

# History (Chat UI)
app.include_router(history.router)       # /api/history/sessions

# Chat (Main execution)
app.include_router(chat.router, prefix="/api/chat", tags=["chat"])

# Excel (Direct operations)
app.include_router(excel.router, prefix="/api/excel", tags=["excel"])

# ============================================================================
# BASIC ENDPOINTS
# ============================================================================

@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "message": "WorkflowGenie API",
        "version": "1.0.0",
        "status": "running",
        "endpoints": {
            "auth": "/api/auth",
            "users": "/api/users",
            "files": "/api/files",
            "sessions": "/api/sessions",
            "history": "/api/history",
            "chat": "/api/chat",
            "excel": "/api/excel",
            "docs": "/docs"
        }
    }

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "ok",
        "service": "WorkflowGenie Backend",
        "database": "connected"
    }

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    """WebSocket endpoint for real-time updates"""
    await manager.connect(websocket)
    try:
        while True:
            # Keep connection alive
            data = await websocket.receive_text()
            logger.debug(f"Received WebSocket message: {data}")
    except WebSocketDisconnect:
        manager.disconnect(websocket)

# Make manager available globally
app.state.ws_manager = manager

if __name__ == "__main__":
    logger.info("🚀 Starting server...")
    uvicorn.run(
        "main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG,
        log_level="info"
    )