"""Entry point for python -m mcp_sharepoint"""
from .server import main
import asyncio

if __name__ == "__main__":
    asyncio.run(main())
