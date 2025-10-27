"""
Railway entry point for SharePoint MCP Server.

This file exists because Railway's auto-detection looks for main.py
in the project root. It simply imports and runs the HTTP server.
"""

import os

# Import the Starlette app (package installed via setup.py)
from mcp_sharepoint.server_http import app

# For Railway/uvicorn to discover
__all__ = ['app']

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
