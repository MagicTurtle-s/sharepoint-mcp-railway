import os, logging
from typing import Optional
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('mcp_sharepoint.log'), logging.StreamHandler()]
)
logger = logging.getLogger('mcp_sharepoint')

# Load environment variables
load_dotenv()

# Configuration - Support both old and new environment variable names
AZURE_CLIENT_ID = os.getenv('AZURE_CLIENT_ID') or os.getenv('SHP_ID_APP')
AZURE_CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET') or os.getenv('SHP_ID_APP_SECRET')
AZURE_TENANT_ID = os.getenv('AZURE_TENANT_ID') or os.getenv('SHP_TENANT_ID')

# Legacy configuration (optional for backwards compatibility)
SHP_SITE_URL = os.getenv('SHP_SITE_URL')  # No longer required for multi-site
SHP_DOC_LIBRARY = os.getenv('SHP_DOC_LIBRARY', 'Shared Documents')

# Validate Azure AD credentials (required)
if not AZURE_CLIENT_ID:
    logger.error("AZURE_CLIENT_ID (or SHP_ID_APP) environment variable not set.")
    raise ValueError("AZURE_CLIENT_ID (or SHP_ID_APP) environment variable not set.")
if not AZURE_CLIENT_SECRET:
    logger.error("AZURE_CLIENT_SECRET (or SHP_ID_APP_SECRET) environment variable not set.")
    raise ValueError("AZURE_CLIENT_SECRET (or SHP_ID_APP_SECRET) environment variable not set.")
if not AZURE_TENANT_ID:
    logger.error("AZURE_TENANT_ID (or SHP_TENANT_ID) environment variable not set.")
    raise ValueError("AZURE_TENANT_ID (or SHP_TENANT_ID) environment variable not set.")

# Log configuration mode
if SHP_SITE_URL:
    logger.info(f"Multi-site mode with legacy fallback site: {SHP_SITE_URL}")
else:
    logger.info("Multi-site mode: No default site configured")

# Initialize MCP server
mcp = FastMCP(
    name="mcp_sharepoint",
    instructions="This server provides tools to interact with SharePoint documents and folders across multiple sites. "
                 "Use list_sharepoint_sites to discover available sites, then specify site_url in other tools."
)

# Global credentials (reusable across all sites)
_global_credentials = ClientCredential(AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)

# Legacy: Global context for backwards compatibility (if SHP_SITE_URL set)
sp_context: Optional[ClientContext] = None
if SHP_SITE_URL:
    sp_context = ClientContext(SHP_SITE_URL).with_credentials(_global_credentials)
    logger.info(f"Created legacy sp_context for: {SHP_SITE_URL}")


def get_sp_context_for_site(site_url: str) -> ClientContext:
    """
    Create a SharePoint ClientContext for a specific site URL.

    This function creates a new context per request, allowing dynamic access
    to any site within the tenant. The global Azure AD credentials work across
    all sites (tenant-level Application Permissions).

    Args:
        site_url: Full SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)

    Returns:
        ClientContext configured for the specified site

    Raises:
        ValueError: If site_url is empty or invalid format

    Example:
        >>> context = get_sp_context_for_site("https://contoso.sharepoint.com/sites/acme-corp")
        >>> file = context.web.get_file_by_server_relative_url("/sites/acme-corp/Shared Documents/file.pdf")
    """
    if not site_url:
        raise ValueError("site_url is required")

    if not site_url.startswith("https://"):
        raise ValueError(f"site_url must start with https:// - got: {site_url}")

    if "sharepoint.com" not in site_url:
        raise ValueError(f"site_url must contain sharepoint.com - got: {site_url}")

    logger.debug(f"Creating SharePoint context for site: {site_url}")
    return ClientContext(site_url).with_credentials(_global_credentials)


def get_default_doc_library() -> str:
    """
    Get the default document library path.

    Returns:
        Default document library path (e.g., "Shared Documents")
    """
    return SHP_DOC_LIBRARY