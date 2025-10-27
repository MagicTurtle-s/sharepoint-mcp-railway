"""
Setup configuration for SharePoint MCP Railway deployment.
"""
from setuptools import setup, find_packages

setup(
    name="mcp-sharepoint",
    version="0.1.6",
    description="SharePoint MCP Server for Railway deployment",
    author="MagicTurtle",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.10",
    install_requires=[
        "mcp>=1.2.1",
        "office365-rest-python-client>=2.6.1",
        "python-dotenv>=1.0.0",
        "pymupdf>=1.23.0",
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
        "python-docx>=1.1.0",
        "msgraph-sdk>=1.0.0",
        "azure-identity>=1.15.0",
        "starlette>=0.37.0",
        "uvicorn[standard]>=0.30.0",
        "sse-starlette>=2.0.0",
    ],
)
