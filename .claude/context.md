# SharePoint MCP Railway - Claude Context

This document provides Claude-specific context for working on the SharePoint MCP Railway project.

## MCP Configuration

This project uses a project-specific MCP configuration to load the SharePoint MCP server when working in this directory.

**Setup:**
- `.claude/settings.json` (local, gitignored) contains the actual MCP server configuration
- `.claude/settings.json.example` (committed) serves as a template
- The SharePoint MCP (~9.4k tokens) is only loaded when working in this project

**When working on this project:**
- The SharePoint MCP tools are automatically available for testing
- You can use tools like `mcp__sharepoint__Get_Document_Content` directly
- This configuration doesn't affect other projects

## Project Overview

Multi-site SharePoint MCP server that enables access to 50+ customer SharePoint sites through a single Railway-deployed service. Forked from Sofias-ai/mcp-sharepoint with multi-site capabilities added.

## Architecture

### Original Design (Single-Site)
```python
# common.py - Hardcoded single site
SHP_SITE_URL = os.getenv('SHP_SITE_URL')  # One site only
sp_context = ClientContext(SHP_SITE_URL).with_credentials(credentials)

# tools.py - All tools use global context
@mcp.tool()
async def get_document_content(folder_name: str, file_name: str):
    # Uses global sp_context for hardcoded site
    return get_document_content(folder_name, file_name)
```

### Our Design (Multi-Site)
```python
# common.py - Dynamic context factory
def get_sp_context_for_site(site_url: str) -> ClientContext:
    """Create SharePoint context for any site dynamically"""
    credentials = ClientCredential(AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
    return ClientContext(site_url).with_credentials(credentials)

# tools.py - All tools accept site_url parameter
@mcp.tool()
async def get_document_content(
    site_url: str,  # NEW: Dynamic site selection
    folder_name: str,
    file_name: str
):
    sp_context = get_sp_context_for_site(site_url)
    # Use site-specific context
```

## Key Implementation Patterns

### 1. Multi-Site Context Management

**Pattern**: Create ClientContext per-request instead of singleton

```python
# BEFORE (common.py)
sp_context = ClientContext(SHP_SITE_URL).with_credentials(credentials)

# AFTER (common.py)
# Global credentials (reusable)
azure_credentials = ClientCredential(
    os.getenv('AZURE_CLIENT_ID'),
    os.getenv('AZURE_CLIENT_SECRET')
)

def get_sp_context_for_site(site_url: str) -> ClientContext:
    """
    Create SharePoint context for specific site.
    Credentials work across all sites in tenant.
    """
    return ClientContext(site_url).with_credentials(azure_credentials)
```

### 2. Tool Refactoring for Multi-Site

**Pattern**: Add site_url as first parameter to all tools

```python
# BEFORE
@mcp.tool(name="Get_Document_Content")
async def get_document_content_tool(folder_name: str, file_name: str):
    return get_document_content(folder_name, file_name)  # Uses global context

# AFTER
@mcp.tool(name="Get_Document_Content")
async def get_document_content_tool(
    site_url: str,  # NEW: First parameter
    folder_name: str,
    file_name: str
):
    return get_document_content(site_url, folder_name, file_name)
```

**All 10 tools need this refactor:**
1. List_SharePoint_Folders
2. List_SharePoint_Documents
3. Get_SharePoint_Tree
4. Get_Document_Content
5. Create_Folder
6. Upload_Document
7. Upload_Document_From_Path
8. Update_Document
9. Delete_Document
10. Delete_Folder

### 3. Microsoft Graph Site Discovery

**Pattern**: Use Microsoft Graph API for site enumeration (separate from Office365 library)

```python
# NEW: graph_api.py
from msgraph import GraphServiceClient
from azure.identity import ClientSecretCredential

async def list_sharepoint_sites(search_query: str = None) -> list:
    """
    List all SharePoint sites using Microsoft Graph API.

    Office365-REST-Python-Client doesn't have good site discovery,
    so we use Microsoft Graph SDK for this specific operation.
    """
    credential = ClientSecretCredential(
        tenant_id=AZURE_TENANT_ID,
        client_id=AZURE_CLIENT_ID,
        client_secret=AZURE_CLIENT_SECRET
    )
    graph_client = GraphServiceClient(credential)

    # Search all sites
    if search_query:
        result = await graph_client.sites.get(search=search_query)
    else:
        result = await graph_client.sites.get()

    return [
        {
            "name": site.display_name,
            "site_url": site.web_url,
            "site_id": site.id,
            "description": site.description or ""
        }
        for site in result.value
    ]
```

### 4. Customer Convenience Mapping

**Pattern**: Optional environment variable for customer-friendly names

```python
# common.py
CUSTOMER_SITES = {}
if os.getenv('CUSTOMER_SITES'):
    # Parse: "Acme Corp:https://...,Wayne Ent:https://..."
    for mapping in os.getenv('CUSTOMER_SITES').split(','):
        name, url = mapping.split(':', 1)
        CUSTOMER_SITES[name.strip()] = url.strip()

def get_site_url_for_customer(customer_name: str) -> str:
    """
    Resolve customer name to site URL.
    Falls back to Graph API search if not in mapping.
    """
    if customer_name in CUSTOMER_SITES:
        return CUSTOMER_SITES[customer_name]

    # Fallback: Search for site
    sites = list_sharepoint_sites(customer_name)
    if sites:
        return sites[0]["site_url"]

    raise ValueError(f"Customer site not found: {customer_name}")
```

### 5. Railway SSE/HTTP Transport

**Pattern**: Wrap FastMCP with Starlette for HTTP/SSE (like Neo4j MCP v2)

```python
# NEW: server_http.py
from starlette.applications import Starlette
from starlette.routing import Route
from starlette.responses import Response
from sse_starlette.sse import EventSourceResponse
from mcp.server.fastmcp import FastMCP

# Import existing mcp instance
from .common import mcp
from . import tools, resources  # Register tools

app = Starlette(routes=[
    Route('/sse', endpoint=sse_endpoint, methods=['GET']),
    Route('/health', endpoint=health_check, methods=['GET']),
])

async def sse_endpoint(request):
    """SSE endpoint for MCP over HTTP"""
    async def event_generator():
        # MCP SSE protocol implementation
        async for message in mcp.handle_sse_connection():
            yield message

    return EventSourceResponse(event_generator())

async def health_check(request):
    return Response("OK", status_code=200)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", 8000)))
```

## Common Tasks

### Add a New Tool

1. **Define in tools.py**:
   ```python
   @mcp.tool(name="Your_Tool_Name", description="What it does")
   async def your_tool(site_url: str, other_params: str):
       sp_context = get_sp_context_for_site(site_url)
       # Implementation
   ```

2. **If needed, add helper in resources.py**:
   ```python
   def your_helper_function(sp_context, ...):
       # Reusable logic
   ```

3. **Test locally**:
   ```bash
   npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint
   ```

### Deploy to Railway

1. **Commit changes**:
   ```bash
   git add .
   git commit -m "Description"
   git push origin main
   ```

2. **Railway auto-deploys** from GitHub

3. **Check logs**:
   ```bash
   railway logs
   ```

4. **Test endpoint**:
   ```bash
   curl https://your-project.up.railway.app/health
   ```

### Test Multi-Site Access

```python
# In MCP Inspector or Claude
{
  "tool": "list_sharepoint_sites",
  "arguments": {
    "search": "customer name"
  }
}

# Then use returned site_url in other tools
{
  "tool": "Get_Document_Content",
  "arguments": {
    "site_url": "https://tenant.sharepoint.com/sites/acme-corp",
    "folder_name": "Proposals",
    "file_name": "Acme_Proposal_2024.docx"
  }
}
```

## Gotchas and Known Issues

### 1. Office365 Library Limitations

**Issue**: `Office365-REST-Python-Client` doesn't have built-in site discovery

**Solution**: Use Microsoft Graph SDK (`msgraph-sdk`) for site operations, Office365 library for document operations

```python
# Use Graph for sites
from msgraph import GraphServiceClient
sites = await graph_client.sites.get()

# Use Office365 for documents
from office365.sharepoint.client_context import ClientContext
sp_context = ClientContext(site_url).with_credentials(credentials)
file = sp_context.web.get_file_by_server_relative_url(path)
```

### 2. Tenant-Level vs Site-Level Permissions

**Issue**: Application permissions are tenant-wide, not per-site

**Mitigation**:
- Use `Sites.Selected` permission in Azure AD (not `Sites.ReadWrite.All`)
- Explicitly grant access only to specific site IDs
- Document in AZURE_SETUP.md

### 3. Server Relative URLs

**Issue**: SharePoint paths are site-relative: `/sites/sitename/Shared Documents/file.docx`

**Pattern**: Always construct full server-relative URL

```python
def _get_server_relative_path(site_url: str, doc_library: str, folder: str, file: str = None):
    """
    Construct server-relative URL for SharePoint operations.

    site_url: https://tenant.sharepoint.com/sites/sitename
    Returns: /sites/sitename/Shared Documents/folder/file.docx
    """
    site_path = urlparse(site_url).path  # /sites/sitename
    path = f"{site_path}/{doc_library}/{folder}".rstrip('/')
    return f"{path}/{file}" if file else path
```

### 4. Document Library Paths

**Issue**: Different customers may use different library names

**Original**: `SHP_DOC_LIBRARY` hardcoded for entire server

**Solution**: Accept `doc_library` parameter (default to "Shared Documents")

```python
@mcp.tool()
async def get_document_content(
    site_url: str,
    folder_name: str,
    file_name: str,
    doc_library: str = "Shared Documents"  # NEW: Optional parameter
):
    path = f"{doc_library}/{folder_name}/{file_name}"
    # ...
```

### 5. Railway Environment Variables

**Issue**: Railway restarts servers, environment must be in dashboard not .env

**Checklist**:
- ✅ Set all vars in Railway dashboard
- ✅ Don't commit .env to Git
- ✅ Test with `railway run python -m mcp_sharepoint` locally
- ✅ Verify logs show correct variable loading

### 6. FastMCP vs Starlette Transport

**Issue**: FastMCP runs `stdio` by default, Railway needs HTTP

**Solution**: Two server files
- `server.py` - stdio for local/Claude Desktop (original)
- `server_http.py` - HTTP/SSE for Railway deployment

**Railway Procfile**:
```
web: python -m mcp_sharepoint.server_http
```

## Testing Strategy

### Local Testing (stdio)
```bash
# With MCP Inspector
npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint

# Direct
python -m mcp_sharepoint
```

### Railway Testing (HTTP/SSE)
```bash
# Local Railway simulation
railway run python -m mcp_sharepoint.server_http

# Test health endpoint
curl http://localhost:8000/health

# Test SSE endpoint (should keep connection open)
curl http://localhost:8000/sse
```

### Integration Testing
```python
# In Claude.ai after deployment
1. Add MCP server with Railway URL
2. Test: "List all available SharePoint sites"
3. Test: "Get the Acme Corp proposal from their site"
4. Test: "Upload a document to Wayne Enterprises discovery folder"
```

## Migration Path to Per-User OAuth

When ready to implement per-user authentication:

### Changes Required

1. **Azure AD permissions**: Application → Delegated
2. **Add OAuth flow** (server_http.py):
   ```python
   @app.route('/oauth/callback')
   async def oauth_callback(request):
       code = request.query_params['code']
       # Exchange for user token
   ```

3. **Extract user token** from requests:
   ```python
   async def get_document_content(
       site_url: str,
       ...,
       authorization: str = Header(...)  # NEW
   ):
       user_token = authorization.replace("Bearer ", "")
       credential = AccessTokenCredential(user_token)
       sp_context = ClientContext(site_url).with_credentials(credential)
   ```

### What Doesn't Change
- ✅ All tool implementations (same SharePoint API calls)
- ✅ Multi-site architecture
- ✅ Document processing logic
- ✅ Railway deployment
- ✅ Customer mapping

## Dependencies

### Core
- `fastmcp` - MCP server framework
- `office365-rest-python-client` - SharePoint operations
- `msgraph-sdk` - Microsoft Graph site discovery
- `azure-identity` - Azure AD authentication

### Document Processing
- `python-docx` - Word document extraction
- `PyMuPDF` (fitz) - PDF text extraction
- `openpyxl` - Excel file processing

### Web Server (Railway)
- `starlette` - ASGI framework
- `sse-starlette` - Server-Sent Events
- `uvicorn` - ASGI server

### Development
- `python-dotenv` - Environment variable management
- `pytest` - Testing (if added)

## Performance Considerations

### Connection Pooling
- Creating ClientContext per request is lightweight
- Credentials object is reused (created once)
- Consider connection pooling if >100 requests/minute

### Caching
- Consider caching site discovery results (TTL: 1 hour)
- Document content caching for frequently accessed files
- Folder structure caching to reduce API calls

### Rate Limiting
- Microsoft Graph: 2000 requests/minute tenant-wide
- SharePoint: No hard limit, but throttling at ~600/minute/user
- Add exponential backoff for 429 responses

## Development Status

- ✅ Phase 1: Setup & Documentation
- 🔄 Phase 2: Multi-Site Implementation (Next)
- ⏳ Phase 3: Railway Deployment
- ⏳ Phase 4: Customer Convenience Layer

## Useful Commands

```bash
# Development
python -m venv .venv
.venv\Scripts\activate
pip install -e .

# Testing
npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint

# Railway
railway login
railway link
railway run python -m mcp_sharepoint.server_http
railway logs --tail

# Git
git status
git add .
git commit -m "feat: add multi-site support"
git push origin main
```

---

**Last Updated**: 2025-10-27
**Next Steps**: Implement multi-site refactoring (Phase 2)
