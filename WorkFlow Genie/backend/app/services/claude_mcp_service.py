"""
Claude MCP Service
Provides the same tool execution surface as MCPService for Claude workflows.
"""

from app.services.mcp_service import MCPService


class ClaudeMCPService(MCPService):
    """Claude-specific MCP service that mirrors all MCPService tools."""


claude_mcp_service = ClaudeMCPService()
