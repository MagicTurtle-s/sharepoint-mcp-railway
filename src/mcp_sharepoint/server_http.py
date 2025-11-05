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
from .oauth import init_oauth_manager, get_oauth_manager

logger.info("Initializing HTTP/SSE transport server...")

# Session management - map session IDs to transports
sessions: Dict[str, SseServerTransport] = {}

# User session management - map session IDs to user IDs
user_sessions: Dict[str, str] = {}


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

                # Call the tool through FastMCP's tool manager
                if hasattr(mcp, '_tool_manager'):
                    tool_manager = mcp._tool_manager
                    logger.info(f"Found tool_manager, looking for tool: {tool_name}")

                    # Get all tools from tool_manager
                    if hasattr(tool_manager, 'list_tools'):
                        tools_result = tool_manager.list_tools()

                        # Find the requested tool
                        tool_obj = None
                        for tool in tools_result:
                            if tool.name == tool_name:
                                tool_obj = tool
                                break

                        if tool_obj:
                            try:
                                logger.info(f"Found tool object, executing: {tool_name}")

                                # Execute the tool - FastMCP tools have a 'fn' attribute
                                if hasattr(tool_obj, 'fn'):
                                    result = await tool_obj.fn(**tool_args) if tool_obj.is_async else tool_obj.fn(**tool_args)
                                else:
                                    logger.error(f"Tool {tool_name} has no 'fn' attribute")
                                    raise Exception("Tool has no callable function")

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
                logger.error(f"Tool not found: {tool_name}")
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


async def oauth_authorize(request: Request) -> Response:
    """
    OAuth authorization endpoint - initiates OAuth flow.

    Redirects user to Microsoft Azure AD for authentication.
    """
    try:
        oauth_mgr = get_oauth_manager()

        # Generate state for CSRF protection
        state = str(uuid.uuid4())

        # Get authorization URL
        auth_url = oauth_mgr.get_authorization_url(state=state)

        logger.info(f"Redirecting to OAuth authorization: {auth_url}")

        # Redirect to Azure AD
        from starlette.responses import RedirectResponse
        return RedirectResponse(url=auth_url)

    except Exception as e:
        logger.error(f"OAuth authorization error: {e}", exc_info=True)
        return JSONResponse(
            {"error": f"OAuth authorization failed: {str(e)}"},
            status_code=500
        )


async def oauth_callback(request: Request) -> Response:
    """
    OAuth callback endpoint - handles authorization code exchange.

    Called by Azure AD after user grants consent.
    """
    try:
        oauth_mgr = get_oauth_manager()

        # Get authorization code from query parameters
        code = request.query_params.get('code')
        state = request.query_params.get('state')
        error = request.query_params.get('error')

        if error:
            logger.error(f"OAuth error: {error}")
            return JSONResponse(
                {"error": f"OAuth authorization failed: {error}"},
                status_code=400
            )

        if not code:
            return JSONResponse(
                {"error": "Missing authorization code"},
                status_code=400
            )

        # Exchange code for tokens
        token_response = oauth_mgr.acquire_token_by_authorization_code(code)

        # Extract user ID from token (use email or object ID)
        # For now, use a session-based approach
        session_id = str(uuid.uuid4())
        user_id = token_response.get('id_token_claims', {}).get('email', session_id)

        # Store tokens
        oauth_mgr.store_user_token(user_id, token_response)
        user_sessions[session_id] = user_id

        logger.info(f"OAuth successful for user: {user_id}, session: {session_id}")

        # Return success page with session info
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>OAuth Success</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    max-width: 600px;
                    margin: 50px auto;
                    padding: 20px;
                    text-align: center;
                }}
                .success {{
                    color: #28a745;
                    font-size: 24px;
                    margin-bottom: 20px;
                }}
                .session-id {{
                    background: #f5f5f5;
                    padding: 15px;
                    border-radius: 5px;
                    margin: 20px 0;
                    font-family: monospace;
                }}
                .instructions {{
                    text-align: left;
                    background: #e9ecef;
                    padding: 15px;
                    border-radius: 5px;
                    margin-top: 20px;
                }}
            </style>
        </head>
        <body>
            <div class="success">✓ Authentication Successful!</div>
            <p>You have successfully authenticated with Microsoft SharePoint.</p>

            <div class="session-id">
                <strong>Session ID:</strong><br>
                {session_id}
            </div>

            <div class="instructions">
                <h3>Next Steps:</h3>
                <ol>
                    <li>Copy your Session ID above</li>
                    <li>Use this Session ID when calling SharePoint MCP tools</li>
                    <li>The session will remain active for 1 hour</li>
                </ol>
                <p><strong>Note:</strong> Keep this Session ID secure - it provides access to your SharePoint data.</p>
            </div>

            <p style="margin-top: 30px;">
                <a href="/">Return to home</a>
            </p>
        </body>
        </html>
        """

        from starlette.responses import HTMLResponse
        return HTMLResponse(content=html)

    except Exception as e:
        logger.error(f"OAuth callback error: {e}", exc_info=True)
        return JSONResponse(
            {"error": f"OAuth callback failed: {str(e)}"},
            status_code=500
        )


async def oauth_not_required(request: Request) -> JSONResponse:
    """
    OAuth discovery endpoints - indicate that OAuth IS required.

    These endpoints are queried by claude.ai to discover if OAuth is needed.
    We now return OAuth configuration instead of 404.
    """
    return JSONResponse(
        {
            "authorization_endpoint": "/oauth/authorize",
            "token_endpoint": "/oauth/callback",
            "grant_types_supported": ["authorization_code"],
            "scopes_supported": ["Files.ReadWrite.All", "Sites.ReadWrite.All"]
        },
        headers={
            "X-MCP-Auth": "oauth",
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
        # OAuth endpoints
        Route('/oauth/authorize', oauth_authorize, methods=['GET']),
        Route('/oauth/callback', oauth_callback, methods=['GET']),
        # OAuth discovery endpoints - return OAuth configuration
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
    """Initialize OAuth manager and log server startup."""
    # Initialize OAuth manager
    client_id = os.getenv('AZURE_CLIENT_ID') or os.getenv('SHP_ID_APP')
    client_secret = os.getenv('AZURE_CLIENT_SECRET') or os.getenv('SHP_ID_APP_SECRET')
    tenant_id = os.getenv('AZURE_TENANT_ID') or os.getenv('SHP_TENANT_ID')

    # Get Railway URL or use localhost for development
    railway_url = os.getenv('RAILWAY_PUBLIC_DOMAIN')
    if railway_url:
        redirect_uri = f"https://{railway_url}/oauth/callback"
    else:
        redirect_uri = "http://localhost:8000/oauth/callback"

    try:
        init_oauth_manager(
            client_id=client_id,
            client_secret=client_secret,
            tenant_id=tenant_id,
            redirect_uri=redirect_uri
        )
        logger.info(f"OAuth manager initialized with redirect URI: {redirect_uri}")
    except Exception as e:
        logger.error(f"Failed to initialize OAuth manager: {e}")
        logger.warning("OAuth authentication will not be available")

    logger.info("=" * 60)
    logger.info("SharePoint MCP Railway Server Starting")
    logger.info("=" * 60)
    logger.info(f"Multi-site mode enabled")
    logger.info(f"OAuth authentication: ENABLED")
    logger.info(f"Streamable HTTP endpoint: /mcp (RECOMMENDED)")
    logger.info(f"Legacy SSE endpoint: /sse")
    logger.info(f"OAuth authorize: /oauth/authorize")
    logger.info(f"OAuth callback: /oauth/callback")
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
