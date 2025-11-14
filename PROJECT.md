# SharePoint MCP Railway

A multi-site SharePoint MCP (Model Context Protocol) server for accessing 50+ customer SharePoint sites with advanced document processing capabilities. Deployed on Railway with SSE/HTTP transport for Claude.ai web interface integration.

## Project Overview

Fork of [Sofias-ai/mcp-sharepoint](https://github.com/Sofias-ai/mcp-sharepoint) enhanced with:
- **Multi-site support**: Access any of 50+ customer SharePoint sites dynamically
- **Railway deployment**: SSE/HTTP transport for Claude.ai web interface
- **Customer convenience layer**: Map friendly customer names to site URLs
- **Per-user authentication ready**: Migration path to OAuth 2.1 delegated permissions

## Tech Stack

- **Language**: Python 3.10+
- **Framework**: FastMCP (Model Context Protocol SDK)
- **Platform**: Railway (~$5/month hobby plan)
- **Authentication**: Azure AD Application Permissions (universal service account)
  - Migration path to OAuth 2.1 Delegated Permissions documented
- **Document Processing**:
  - `python-docx` - Word document extraction
  - `PyMuPDF` - PDF text extraction
  - `openpyxl` - Excel file processing
- **SharePoint Client**: `Office365-REST-Python-Client`
- **Transport**: SSE/HTTP (Streamable HTTP with Server-Sent Events)

## Key Features

### 11 SharePoint Tools

#### Site Discovery (Multi-Site)
1. **`list_sharepoint_sites`** - Discover all accessible SharePoint sites (Microsoft Graph API)

#### Folder Management
2. **`List_SharePoint_Folders`** - List folders in a directory
3. **`Create_Folder`** - Create new folders
4. **`Delete_Folder`** - Delete empty folders
5. **`Get_SharePoint_Tree`** - Recursive folder tree view

#### Document Management
6. **`List_SharePoint_Documents`** - List documents with metadata
7. **`Get_Document_Content`** - Extract content from Word/PDF/Excel files
8. **`Upload_Document`** - Upload text or binary files (base64)
9. **`Upload_Document_From_Path`** - Direct file upload from filesystem
10. **`Update_Document`** - Update existing document content
11. **`Delete_Document`** - Remove documents

### Customer Convenience Layer
- **Customer site mapping**: Friendly names → SharePoint site URLs
- **Smart discovery**: Fallback to site search if customer not in mapping
- Optional custom tools: `search_customer_documents`, `get_customer_proposals`

## Use Case

Access customer-specific SharePoint sites to:
- Read discovery notes and proposals (Word documents)
- Search across 50+ customer sites
- Upload generated proposals or documents
- Manage customer documentation folders
- Extract data from customer spreadsheets and PDFs

## Architecture

```
Claude.ai Web Interface
    ↓ (HTTPS)
Railway MCP Server (SSE/HTTP)
    ↓ (OAuth Bearer Token)
Azure AD Tenant
    ↓ (Application Permissions)
Microsoft Graph API
    ↓
SharePoint Online (50+ Sites)
```

### Multi-Site Design

**Traditional (single-site)**:
```python
SHP_SITE_URL = "https://tenant.sharepoint.com/sites/customer1"  # Hardcoded
```

**Our Implementation (multi-site)**:
```python
@mcp.tool()
async def get_document_content(
    site_url: str,  # Dynamic per request
    folder_name: str,
    file_name: str
):
    # Create context for specific site
    sp_context = get_sp_context_for_site(site_url)
    # ... rest of implementation
```

## Required MCPs

When working on this project in Claude Code, you'll need the SharePoint MCP server enabled to test and develop the tools.

**Setup:**
1. Copy `.claude/settings.json.example` to `.claude/settings.json`
2. Update the URL if using a different deployment
3. Claude Code will automatically load the SharePoint MCP when working in this project directory

This configuration is project-specific and won't affect other projects. The SharePoint MCP (~9.4k tokens) is only loaded when needed.

## Setup Instructions

### Prerequisites

1. **Azure AD App Registration** (see AZURE_SETUP.md)
   - Application permissions: `Sites.ReadWrite.All`, `Files.ReadWrite.All`
   - Admin consent granted
   - Client ID, Client Secret, Tenant ID

2. **Railway Account** (railway.app)
   - Credit card for hobby plan (~$5/month)

### Local Development

1. **Clone repository**:
   ```bash
   git clone https://github.com/MagicTurtle-s/sharepoint-mcp-railway.git
   cd sharepoint-mcp-railway
   ```

2. **Create virtual environment**:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   # or: source .venv/bin/activate  # macOS/Linux
   ```

3. **Install dependencies**:
   ```bash
   pip install -e .
   ```

4. **Configure environment** (`.env`):
   ```env
   # Azure AD Credentials
   AZURE_CLIENT_ID=your-client-id
   AZURE_CLIENT_SECRET=your-client-secret
   AZURE_TENANT_ID=your-tenant-id

   # Optional: Customer site mappings (comma-separated)
   CUSTOMER_SITES=Acme Corp:https://tenant.sharepoint.com/sites/acme-corp,Wayne Enterprises:https://tenant.sharepoint.com/sites/wayne-ent
   ```

5. **Run locally**:
   ```bash
   python -m mcp_sharepoint
   ```

6. **Test with MCP Inspector**:
   ```bash
   npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint
   ```

### Railway Deployment

1. **Push to GitHub** (already done via fork)

2. **Create Railway project**:
   ```bash
   railway login
   railway init
   railway up
   ```

3. **Configure environment variables** in Railway dashboard:
   - `AZURE_CLIENT_ID`
   - `AZURE_CLIENT_SECRET`
   - `AZURE_TENANT_ID`
   - `CUSTOMER_SITES` (optional)

4. **Get SSE endpoint**: `https://your-project.up.railway.app/sse`

5. **Configure in Claude.ai**:
   - Settings → Developer → MCP Servers
   - Add Server → URL: `https://your-project.up.railway.app/sse`

## Project Structure

```
sharepoint-mcp-railway/
├── src/
│   └── mcp_sharepoint/
│       ├── __init__.py
│       ├── server.py           # Main entry point (stdio transport)
│       ├── server_http.py      # NEW: SSE/HTTP transport for Railway
│       ├── common.py            # MODIFIED: Multi-site context management
│       ├── tools.py             # MODIFIED: All tools accept site_url
│       ├── resources.py         # MODIFIED: Dynamic site context
│       └── graph_api.py         # NEW: Microsoft Graph site discovery
├── PROJECT.md                   # This file
├── .claude/
│   └── context.md              # Architecture, patterns, gotchas
├── AZURE_SETUP.md              # Azure AD configuration guide
├── railway.toml                # Railway deployment config
├── pyproject.toml
├── requirements.txt
├── .env.example
└── README.md
```

## External Services

### Azure AD (Microsoft 365)
- **Purpose**: Authentication and authorization
- **Credentials**: Stored in Railway environment variables
- **Permissions**: Application-level (Sites.ReadWrite.All, Files.ReadWrite.All)
- **Cost**: Free (included in Microsoft 365)

### Railway
- **Purpose**: Hosting MCP server with SSE/HTTP transport
- **Deployment**: Automatic from GitHub
- **Cost**: ~$5/month (hobby plan)
- **URL**: `https://[project-name].up.railway.app`

### Microsoft SharePoint Online
- **Purpose**: Document storage across 50+ customer sites
- **Access**: Via Microsoft Graph API
- **Cost**: Included in Microsoft 365

## Related Projects

- **Neo4j MCP Railway v2**: Pattern for Railway deployment with SSE transport
- **HubSpot MCP Railway**: Similar multi-tool MCP architecture
- **MCP Utils**: Utilities for MCP server development

## Future Enhancements

### Phase 2: SharePoint Lists Support (2-3 hours)
- Create/read/update/delete list items
- Query lists with filtering
- Proposal tracking, contact management

### Phase 3: Per-User OAuth (4 hours)
- Migrate from Application Permissions → Delegated Permissions
- Add OAuth 2.1 Authorization Code flow
- Per-request user token validation
- SharePoint permissions enforced per user

### Phase 4: Advanced Features
- Document version management
- Check-in/check-out operations
- SharePoint search API integration
- Bulk operations (upload multiple files)

## Troubleshooting

### Common Issues

1. **"SHP_SITE_URL environment variable not set"**
   - Fixed: SHP_SITE_URL no longer required (multi-site support)
   - Use `site_url` parameter in tool calls

2. **Azure AD permission errors**
   - Ensure admin consent granted for application permissions
   - Verify Sites.ReadWrite.All and Files.ReadWrite.All are added

3. **Railway deployment fails**
   - Check environment variables are set
   - Verify Python version in railway.toml
   - Check logs: `railway logs`

4. **Site not found errors**
   - Use `list_sharepoint_sites` to discover available sites
   - Verify site URL format: `https://tenant.sharepoint.com/sites/sitename`

## Development Status

- ✅ Phase 1: Repository setup and documentation (Complete)
- 🔄 Phase 2: Multi-site support implementation (In Progress)
- ⏳ Phase 3: Railway deployment (Pending)
- ⏳ Phase 4: Customer convenience layer (Pending)

## License

MIT License - see [LICENSE](LICENSE) file

Original work Copyright (c) 2025 Sofias Tech
Modified work Copyright (c) 2025 MagicTurtle-s

---

**Last Updated**: 2025-10-27
**Status**: Active Development
**GitHub**: https://github.com/MagicTurtle-s/sharepoint-mcp-railway
