"""
HTTP/SSE transport server for Railway deployment.

This module wraps the FastMCP server with Starlette to provide HTTP/SSE transport,
enabling deployment to Railway and integration with Claude.ai web interface.
"""

import os
import logging
import uuid
from typing import Dict
from starlette.applications import Starlette
from starlette.routing import Route
from starlette.responses import JSONResponse, Response
from starlette.requests import Request
from mcp.server.sse import SseServerTransport
import uvicorn

# Import the MCP server and register all tools
from .common import logger, mcp
from . import tools, resources  # This registers all the tools

logger.info("Initializing HTTP/SSE transport server...")

# Session management - map session IDs to transports
sessions: Dict[str, SseServerTransport] = {}


async def health_check(request: Request) -> Response:
    """
    Health check endpoint for Railway and monitoring.

    Returns:
        200 OK with server status
    """
    return JSONResponse({
        "status": "healthy",
        "server": "SharePoint MCP Railway",
        "version": "0.1.6",
        "transport": "SSE/HTTP",
        "multi_site": True,
        "active_sessions": len(sessions)
    })


async def handle_sse(request: Request) -> Response:
    """
    Server-Sent Events endpoint for MCP communication.

    This endpoint handles the MCP protocol over SSE, allowing Claude.ai
    web interface to communicate with the SharePoint MCP server.
    """
    session_id = str(uuid.uuid4())
    logger.info(f"New SSE connection from {request.client.host}, session: {session_id}")

    # Create a new transport for this session
    sse_transport = SseServerTransport(f"/messages/{session_id}")
    sessions[session_id] = sse_transport

    try:
        async with sse_transport.connect_sse(
            request.scope,
            request.receive,
            request._send
        ) as streams:
            await mcp._mcp_server.run(
                streams[0],  # read stream
                streams[1],  # write stream
                mcp._mcp_server.create_initialization_options()
            )
    finally:
        # Clean up session when connection closes
        if session_id in sessions:
            del sessions[session_id]
            logger.info(f"Session {session_id} closed")

    return Response()


async def handle_messages(request: Request) -> Response:
    """
    Handle POST messages from client.
    
    The session ID is in the URL path.
    """
    # Extract session ID from path
    path_parts = request.url.path.split('/')
    if len(path_parts) >= 3:
        session_id = path_parts[2]
    else:
        return JSONResponse({"error": "Invalid session ID"}, status_code=400)

    # Find the transport for this session
    sse_transport = sessions.get(session_id)
    if not sse_transport:
        return JSONResponse({"error": "Session not found"}, status_code=404)

    # Handle the message
    return await sse_transport.handle_post_message(request.scope, request.receive, request._send)


async def root_endpoint(request: Request) -> JSONResponse:
    """
    Root endpoint with server information.

    Returns:
        JSON with server details and usage instructions
    """
    return JSONResponse({
        "name": "SharePoint MCP Railway",
        "description": "Multi-site SharePoint MCP server for Claude.ai",
        "version": "0.1.6",
        "endpoints": {
            "sse": "/sse",
            "messages": "/messages/{session_id}",
            "health": "/health"
        },
        "features": [
            "Multi-site SharePoint access",
            "Site discovery via Microsoft Graph",
            "Document management (read/write)",
            "Advanced document processing (Word, PDF, Excel)",
            "Customer site mapping"
        ],
        "usage": "Configure in Claude.ai with SSE endpoint: https://your-domain.railway.app/sse",
        "github": "https://github.com/MagicTurtle-s/sharepoint-mcp-railway",
        "active_sessions": len(sessions)
    })


# Create Starlette application
app = Starlette(
    debug=False,
    routes=[
        Route('/', root_endpoint, methods=['GET']),
        Route('/health', health_check, methods=['GET']),
        Route('/sse', handle_sse, methods=['GET']),
        Route('/messages/{session_id:path}', handle_messages, methods=['POST']),
    ]
)


@app.on_event("startup")
async def startup_event():
    """Log server startup."""
    logger.info("=" * 60)
    logger.info("SharePoint MCP Railway Server Starting")
    logger.info("=" * 60)
    logger.info(f"Multi-site mode enabled")
    logger.info(f"SSE endpoint: /sse")
    logger.info(f"Messages endpoint: /messages/{{session_id}}")
    logger.info(f"Health check: /health")
    logger.info("=" * 60)


@app.on_event("shutdown")
async def shutdown_event():
    """Log server shutdown."""
    logger.info("SharePoint MCP Railway Server Shutting Down")


# For direct execution (development)
if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    host = os.getenv("HOST", "0.0.0.0")

    logger.info(f"Starting development server on {host}:{port}")

    uvicorn.run(
        app,
        host=host,
        port=port,
        log_level="info",
        access_log=True
    )
