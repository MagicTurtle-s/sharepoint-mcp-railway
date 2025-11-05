# SharePoint MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![PyPI version](https://badge.fury.io/py/mcp-sharepoint.svg)](https://badge.fury.io/py/mcp-sharepoint)

A comprehensive MCP Server for seamless integration with Microsoft SharePoint, enabling MCP clients to interact with documents, folders and other SharePoint resources. Built with efficiency and ease of use in mind, supporting both text and binary file operations. Developed by [sofias tech](https://github.com/Sofias-ai/mcp-sharepoint/).

<a href="https://glama.ai/mcp/servers/@Sofias-ai/mcp-sharepoint">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@Sofias-ai/mcp-sharepoint/badge" alt="SharePoint Server MCP server" />
</a>

## ✨ Key Features

This server provides a clean, efficient interface to SharePoint resources through the Model Context Protocol (MCP), with optimized operations for document management and content processing.

### 🛠️ Available Tools

The server implements **10 comprehensive tools** for complete SharePoint management:

#### 📁 **Folder Management**
- **`List_SharePoint_Folders`**: Lists all folders in a specified directory or root
- **`Create_Folder`**: Creates new folders in specified directories 
- **`Delete_Folder`**: Safely deletes empty folders from SharePoint
- **`Get_SharePoint_Tree`**: Gets a recursive tree view of SharePoint folder structure

#### 📄 **Document Management**  
- **`List_SharePoint_Documents`**: Fetches all documents within a specified folder with metadata
- **`Get_Document_Content`**: Retrieves and processes document content (supports text extraction from PDF, Word, Excel)
- **`Upload_Document`**: Uploads new documents to specified folders (supports both text and binary content)
- **`Upload_Document_From_Path`**: Direct file upload from local filesystem for large files
- **`Update_Document`**: Updates content of existing documents
- **`Delete_Document`**: Removes documents from specified folders

### 🎯 **Advanced Content Processing**

The server includes intelligent content extraction capabilities:

- **📊 Excel Files**: Extracts data from all sheets, converts to readable text format (first 50 rows per sheet)
- **📝 Word Documents**: Processes paragraphs and tables, maintaining structure
- **📄 PDF Files**: Full text extraction using PyMuPDF for accurate content parsing
- **📃 Text Files**: Direct processing of various text formats (JSON, XML, HTML, MD, code files)
- **🔧 Binary Support**: Base64 encoding/decoding for seamless binary file handling

## 🏗️ Architecture

The server is built with resource efficiency and maintainability in mind:

- **Efficient SharePoint API usage** with selective property loading to minimize bandwidth
- **Smart error handling** through decorators for cleaner, more reliable code
- **Clear separation of concerns** between resource management and tool implementation
- **Optimized content handling** for both text and binary files with automatic type detection
- **Configurable tree operations** with depth limits and batch processing for large directories
- **Async/await support** throughout for non-blocking operations
- **Dual authentication modes**: Application Permissions (legacy) and Delegated Permissions (OAuth 2.0)

## 🔐 Authentication Modes

The SharePoint MCP supports two authentication modes:

### Application Permissions (Legacy Mode)

Uses a service account with tenant-wide access. Simple setup but lacks per-user audit trails.

**Use case**: Single-user scenarios, development, testing

### Delegated Permissions (OAuth 2.0) - Recommended

Per-user OAuth authentication with proper audit trails in SharePoint.

**Use case**: Multi-user deployments, production environments, compliance requirements

**Features**:
- ✅ Per-user authentication and authorization
- ✅ Proper audit trails in SharePoint (operations logged as specific users)
- ✅ Automatic token refresh (12-hour access tokens, 6-month refresh tokens)
- ✅ Session-based token management
- ✅ HTTP/SSE transport support for cloud deployment

**Setup**: See [AZURE_AD_SETUP.md](AZURE_AD_SETUP.md) for detailed OAuth configuration instructions.

## Setup

1. Register an app in Azure AD with appropriate SharePoint permissions
2. Obtain the client ID and client secret for the registered app
3. Identify your SharePoint site URL and the document library path you want to work with

## Environment Variables

The server requires these environment variables:

### Required Variables
- `SHP_ID_APP`: Your Azure AD application client ID
- `SHP_ID_APP_SECRET`: Your Azure AD application client secret
- `SHP_SITE_URL`: The URL of your SharePoint site
- `SHP_DOC_LIBRARY`: Path to the document library (default: "Shared Documents/mcp_server")
- `SHP_TENANT_ID`: Your Microsoft tenant ID

### Optional Configuration Variables
- `SHP_MAX_DEPTH`: Maximum folder depth for tree operations (default: 15)
- `SHP_MAX_FOLDERS_PER_LEVEL`: Maximum folders to process per level (default: 100)
- `SHP_LEVEL_DELAY`: Delay in seconds between processing levels (default: 0.5)

## Quickstart

### Installation

```bash
pip install -e .
```

Or install from PyPI once published:

```bash
pip install mcp-sharepoint-server
```

Using uv:

```bash
uv pip install mcp-sharepoint-server
```

### Claude Desktop Integration

To integrate with Claude Desktop, update the configuration file:

On Windows: `%APPDATA%/Claude/claude_desktop_config.json`
On macOS: `~/Library/Application\ Support/Claude/claude_desktop_config.json`

#### Standard Integration

```json
"mcpServers": {
  "sharepoint": {
    "command": "mcp-sharepoint",
    "env": {
      "SHP_ID_APP": "your-app-id",
      "SHP_ID_APP_SECRET": "your-app-secret",
      "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
      "SHP_DOC_LIBRARY": "Shared Documents/your-folder",
      "SHP_TENANT_ID": "your-tenant-id",
      "SHP_MAX_DEPTH": "15",
      "SHP_MAX_FOLDERS_PER_LEVEL": "100",
      "SHP_LEVEL_DELAY": "0.5"
    }
  }
}
```

#### Using uvx

```json
"mcpServers": {
  "sharepoint": {
    "command": "uvx",
    "args": [
      "mcp-sharepoint"
    ],
    "env": {
      "SHP_ID_APP": "your-app-id",
      "SHP_ID_APP_SECRET": "your-app-secret",
      "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
      "SHP_DOC_LIBRARY": "Shared Documents/your-folder",
      "SHP_TENANT_ID": "your-tenant-id",
      "SHP_MAX_DEPTH": "15",
      "SHP_MAX_FOLDERS_PER_LEVEL": "100",
      "SHP_LEVEL_DELAY": "0.5"
    }
  }
}
```

## Development

### Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt` and `pyproject.toml`

### Local Development

1. Clone the repository
2. Create a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```
3. Install development dependencies:
   ```bash
   pip install -e .
   ```
4. Create a `.env` file with your SharePoint credentials:
   ```
   SHP_ID_APP=your-app-id
   SHP_ID_APP_SECRET=your-app-secret
   SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
   SHP_DOC_LIBRARY=Shared Documents/your-folder
   SHP_TENANT_ID=your-tenant-id
   ```
5. Run the server:
   ```bash
   python -m mcp_sharepoint
   ```

### Debugging

For debugging the MCP server, you can use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

Copyright (c) 2025 sofias tech