"""
HTTP/SSE transport server for Railway deployment.

This module wraps the FastMCP server with Starlette to provide HTTP/SSE transport,
enabling deployment to Railway and integration with Claude.ai web interface.
"""

import os
import logging
from starlette.applications import Starlette
from starlette.routing import Route, Mount
from starlette.responses import JSONResponse, Response
from starlette.requests import Request
from mcp.server.sse import SseServerTransport
import uvicorn

# Import the MCP server and register all tools
from .common import logger, mcp
from . import tools, resources  # This registers all the tools

logger.info("Initializing HTTP/SSE transport server...")

# Create SSE transport
sse_transport = SseServerTransport("/messages")


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
        "multi_site": True
    })


async def handle_sse(request: Request) -> Response:
    """
    Server-Sent Events endpoint for MCP communication.

    This endpoint handles the MCP protocol over SSE, allowing Claude.ai
    web interface to communicate with the SharePoint MCP server.
    """
    logger.info(f"SSE connection established from {request.client.host}")

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

    return Response()


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
            "messages": "/messages",
            "health": "/health"
        },
        "features": [
            "Multi-site SharePoint access",
            "Site discovery via Microsoft Graph",
            "Document management (read/write)",
            "Advanced document processing (Word, PDF, Excel)",
            "Customer site mapping"
        ],
        "usage": "Configure in Claude.ai with SSE endpoint URL",
        "github": "https://github.com/MagicTurtle-s/sharepoint-mcp-railway"
    })


# Create Starlette application
app = Starlette(
    debug=False,
    routes=[
        Route('/', root_endpoint, methods=['GET']),
        Route('/health', health_check, methods=['GET']),
        Route('/sse', handle_sse, methods=['GET']),
        Mount('/messages', app=sse_transport.handle_post_message),
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
    logger.info(f"Messages endpoint: /messages")
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
