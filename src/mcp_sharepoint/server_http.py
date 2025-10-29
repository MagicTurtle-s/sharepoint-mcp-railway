"""
HTTP/SSE transport server for Railway deployment.

This module wraps the FastMCP server with Starlette to provide HTTP/SSE transport,
enabling deployment to Railway and integration with Claude.ai web interface.
"""

import os
import logging
import uuid
import json
from typing import Dict
from starlette.applications import Starlette
from starlette.routing import Route
from starlette.responses import JSONResponse, Response, StreamingResponse
from starlette.requests import Request
from starlette.middleware import Middleware
from starlette.middleware.cors import CORSMiddleware
from mcp.server.sse import SseServerTransport
import uvicorn
import asyncio

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

    async def send_wrapper(message):
        """Wrapper to properly handle send calls"""
        await request._send(message)

    try:
        async with sse_transport.connect_sse(
            request.scope,
            request.receive,
            send_wrapper
        ) as streams:
            await mcp._mcp_server.run(
                streams[0],  # read stream
                streams[1],  # write stream
                mcp._mcp_server.create_initialization_options()
            )
    except Exception as e:
        logger.error(f"Error in SSE connection: {e}", exc_info=True)
        raise
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
        logger.warning("Invalid session ID in message request")
        return JSONResponse({"error": "Invalid session ID"}, status_code=400)

    # Find the transport for this session
    sse_transport = sessions.get(session_id)
    if not sse_transport:
        logger.warning(f"Session {session_id} not found")
        return JSONResponse({"error": "Session not found"}, status_code=404)

    async def send_wrapper(message):
        """Wrapper to properly handle send calls"""
        await request._send(message)

    try:
        # Handle the message
        return await sse_transport.handle_post_message(request.scope, request.receive, send_wrapper)
    except Exception as e:
        logger.error(f"Error handling message for session {session_id}: {e}", exc_info=True)
        return JSONResponse({"error": "Internal server error"}, status_code=500)


async def handle_mcp(request: Request) -> Response:
    """
    Streamable HTTP transport endpoint for MCP communication.

    This is the modern transport method that Claude.ai prefers.
    Handles GET (for establishing connections), POST (for messages), and HEAD (for discovery).
    """
    if request.method == "HEAD":
        # HEAD request - endpoint discovery
        logger.info(f"HEAD request to /mcp from {request.client.host}")
        # Return 200 OK with headers indicating this is a public MCP endpoint
        return Response(
            status_code=200,
            headers={
                "X-MCP-Version": "2025-03-26",
                "X-MCP-Auth": "none",  # Indicate no authentication required
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "GET, POST, HEAD, OPTIONS",
                "Access-Control-Allow-Headers": "*",
            }
        )

    elif request.method == "GET":
        # GET request - establish a streaming connection
        logger.info(f"New Streamable HTTP GET connection from {request.client.host}")

        # Create a new transport for this connection
        sse_transport = SseServerTransport("/mcp")

        async def send_wrapper(message):
            """Wrapper to properly handle send calls"""
            await request._send(message)

        try:
            async with sse_transport.connect_sse(
                request.scope,
                request.receive,
                send_wrapper
            ) as streams:
                await mcp._mcp_server.run(
                    streams[0],  # read stream
                    streams[1],  # write stream
                    mcp._mcp_server.create_initialization_options()
                )
        except Exception as e:
            logger.error(f"Error in Streamable HTTP GET connection: {e}", exc_info=True)
            raise

        return Response()

    elif request.method == "POST":
        # POST request - handle a JSON-RPC message and return JSON response
        logger.info(f"Streamable HTTP POST message from {request.client.host}")

        try:
            # Read the JSON-RPC request body
            body = await request.body()
            request_data = json.loads(body)
            method = request_data.get('method', 'unknown')
            request_id = request_data.get('id')

            logger.info(f"Received JSON-RPC request: method={method}, id={request_id}")

            # Handle initialize request
            if method == "initialize":
                response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {
                            "tools": {},
                            "resources": {},
                            "prompts": {}
                        },
                        "serverInfo": {
                            "name": "mcp_sharepoint",
                            "version": "0.1.6"
                        }
                    }
                }
                logger.info(f"Returning initialize response")
                return JSONResponse(response)

            # Handle tools/list request
            elif method == "tools/list":
                # Get tools from the MCP server
                tools_list = []

                # FastMCP stores tools in _tool_manager
                logger.info(f"MCP type: {type(mcp)}")
                
                if hasattr(mcp, '_tool_manager'):
                    tool_manager = mcp._tool_manager
                    logger.info(f"Found _tool_manager: {type(tool_manager)}")
                    
                    # Check what the tool_manager has
                    if hasattr(tool_manager, 'list_tools'):
                        logger.info("Calling tool_manager.list_tools()")
                        tools_result = tool_manager.list_tools()
                        logger.info(f"list_tools() result type: {type(tools_result)}, count: {len(tools_result) if hasattr(tools_result, '__len__') else 'unknown'}")
                        
                        # Convert Tool objects to MCP protocol format
                        for tool in tools_result:
                            logger.info(f"Tool: {tool.name}")
                            
                            # MCP protocol requires: name, description, inputSchema
                            # The 'parameters' field contains the input schema
                            tool_dict = {
                                "name": tool.name,
                                "description": tool.description or "",
                                "inputSchema": tool.parameters if hasattr(tool, 'parameters') else {}
                            }
                            
                            logger.info(f"Tool dict: name={tool_dict['name']}, has inputSchema={bool(tool_dict['inputSchema'])}")
                            tools_list.append(tool_dict)
                    elif hasattr(tool_manager, 'tools'):
                        logger.info(f"Found tool_manager.tools")
                        for tool in tool_manager.tools:
                            tool_dict = {
                                "name": tool.name,
                                "description": tool.description if hasattr(tool, 'description') else "",
                                "inputSchema": tool.inputSchema if hasattr(tool, 'inputSchema') else {}
                            }
                            tools_list.append(tool_dict)
                    else:
                        logger.warning(f"tool_manager attributes: {[a for a in dir(tool_manager) if not a.startswith('__')]}")
                else:
                    logger.error("No _tool_manager found on mcp object")

                response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "tools": tools_list
                    }
                }
                logger.info(f"Returning {len(tools_list)} tools: {[t.get('name', 'unknown') if isinstance(t, dict) else getattr(t, 'name', 'unknown') for t in tools_list]}")
                return JSONResponse(response)

            # Handle resources/list request
            elif method == "resources/list":
                response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "resources": []
                    }
                }
                logger.info(f"Returning resources list")
                return JSONResponse(response)

            # Handle prompts/list request
            elif method == "prompts/list":
                response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "prompts": []
                    }
                }
                logger.info(f"Returning prompts list")
                return JSONResponse(response)

            # Handle tools/call request
            elif method == "tools/call":
                params = request_data.get('params', {})
                tool_name = params.get('name')
                tool_args = params.get('arguments', {})

                logger.info(f"Calling tool: {tool_name} with args: {tool_args}")

                # Call the tool through the MCP server
                if hasattr(mcp, '_mcp_server') and hasattr(mcp._mcp_server, '_tools'):
                    if tool_name in mcp._mcp_server._tools:
                        try:
                            tool_func = mcp._mcp_server._tools[tool_name]
                            result = await tool_func(**tool_args)

                            response = {
                                "jsonrpc": "2.0",
                                "id": request_id,
                                "result": {
                                    "content": [
                                        {
                                            "type": "text",
                                            "text": str(result)
                                        }
                                    ]
                                }
                            }
                            logger.info(f"Tool {tool_name} executed successfully")
                            return JSONResponse(response)
                        except Exception as e:
                            logger.error(f"Error executing tool {tool_name}: {e}", exc_info=True)
                            response = {
                                "jsonrpc": "2.0",
                                "id": request_id,
                                "error": {
                                    "code": -32603,
                                    "message": f"Tool execution failed: {str(e)}"
                                }
                            }
                            return JSONResponse(response)

                # Tool not found
                response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Tool not found: {tool_name}"
                    }
                }
                return JSONResponse(response)

            # Handle notifications (no response needed)
            elif request_id is None:
                logger.info(f"Received notification: {method}")
                return Response(status_code=204)  # No content for notifications

            # Unknown method
            else:
                logger.warning(f"Unknown method: {method}")
                response = {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Method not found: {method}"
                    }
                }
                return JSONResponse(response)

        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in POST request: {e}")
            return JSONResponse({"jsonrpc": "2.0", "error": {"code": -32700, "message": "Parse error"}, "id": None}, status_code=400)
        except Exception as e:
            logger.error(f"Error handling Streamable HTTP POST: {e}", exc_info=True)
            return JSONResponse({"jsonrpc": "2.0", "error": {"code": -32603, "message": f"Internal error: {str(e)}"}, "id": None}, status_code=500)

    else:
        return JSONResponse({"error": "Method not allowed"}, status_code=405)


async def oauth_not_required(request: Request) -> JSONResponse:
    """
    OAuth discovery endpoints - indicate that OAuth is not required.

    These endpoints are queried by claude.ai to discover if OAuth is needed.
    Returning 404 tells claude.ai that this is a public endpoint with no auth.
    """
    return JSONResponse(
        {"error": "OAuth not required - this is a public MCP endpoint"},
        status_code=404,
        headers={
            "X-MCP-Auth": "none",
            "Access-Control-Allow-Origin": "*",
        }
    )


async def root_endpoint(request: Request) -> JSONResponse:
    """
    Root endpoint with server information.

    Returns:
        JSON with server details and usage instructions
    """
    return JSONResponse({
        "name": "SharePoint MCP Railway",
        "description": "Multi-site SharePoint MCP server for Claude.ai (Public - No Auth Required)",
        "version": "0.1.6",
        "authentication": "none",
        "endpoints": {
            "mcp": "/mcp (Streamable HTTP - Recommended)",
            "sse": "/sse (Legacy SSE)",
            "messages": "/messages/{session_id} (Legacy)",
            "health": "/health"
        },
        "features": [
            "Multi-site SharePoint access",
            "Site discovery via Microsoft Graph",
            "Document management (read/write)",
            "Advanced document processing (Word, PDF, Excel)",
            "Customer site mapping"
        ],
        "usage": "Configure in Claude.ai with: https://your-domain.railway.app/mcp",
        "github": "https://github.com/MagicTurtle-s/sharepoint-mcp-railway",
        "active_sessions": len(sessions)
    })


# Create Starlette application with CORS middleware
app = Starlette(
    debug=False,
    routes=[
        Route('/', root_endpoint, methods=['GET']),
        Route('/health', health_check, methods=['GET']),
        Route('/mcp', handle_mcp, methods=['GET', 'POST', 'HEAD']),  # Streamable HTTP transport (modern)
        Route('/sse', handle_sse, methods=['GET']),  # Legacy SSE transport
        Route('/messages/{session_id:path}', handle_messages, methods=['POST']),  # Legacy SSE messages
        # OAuth discovery endpoints - return 404 to indicate no auth required
        Route('/.well-known/oauth-protected-resource/mcp', oauth_not_required, methods=['GET']),
        Route('/.well-known/oauth-authorization-server/mcp', oauth_not_required, methods=['GET']),
        Route('/.well-known/oauth-authorization-server', oauth_not_required, methods=['GET']),
        Route('/register', oauth_not_required, methods=['POST']),
    ],
    middleware=[
        Middleware(
            CORSMiddleware,
            allow_origins=["*"],  # Allow all origins for claude.ai and other MCP clients
            allow_credentials=True,
            allow_methods=["*"],  # Allow all HTTP methods
            allow_headers=["*"],  # Allow all headers
            expose_headers=["*"],  # Expose all headers to the client
        )
    ]
)


@app.on_event("startup")
async def startup_event():
    """Log server startup."""
    logger.info("=" * 60)
    logger.info("SharePoint MCP Railway Server Starting")
    logger.info("=" * 60)
    logger.info(f"Multi-site mode enabled")
    logger.info(f"Streamable HTTP endpoint: /mcp (RECOMMENDED)")
    logger.info(f"Legacy SSE endpoint: /sse")
    logger.info(f"Legacy messages endpoint: /messages/{{session_id}}")
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
