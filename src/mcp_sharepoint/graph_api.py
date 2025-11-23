"""
Microsoft Graph API operations for SharePoint site discovery and management.

This module uses Microsoft Graph REST API for operations that aren't well-supported
by the Office365-REST-Python-Client library, particularly site discovery and enumeration.

Uses httpx + azure-identity instead of full msgraph-sdk to avoid Windows Long Path issues.
"""

import os
import httpx
from typing import List, Dict, Any, Optional
from azure.identity import ClientSecretCredential
from .common import logger

# Azure AD credentials (shared with Office365 client)
AZURE_CLIENT_ID = os.getenv('AZURE_CLIENT_ID') or os.getenv('SHP_ID_APP')
AZURE_CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET') or os.getenv('SHP_ID_APP_SECRET')
AZURE_TENANT_ID = os.getenv('AZURE_TENANT_ID') or os.getenv('SHP_TENANT_ID')

# Microsoft Graph API base URL
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# Validate credentials
if not all([AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID]):
    logger.warning("Azure AD credentials not fully configured for Graph API. Site discovery may not work.")
    _credential = None
else:
    # Initialize Azure credential
    _credential = ClientSecretCredential(
        tenant_id=AZURE_TENANT_ID,
        client_id=AZURE_CLIENT_ID,
        client_secret=AZURE_CLIENT_SECRET
    )


async def _get_access_token() -> str:
    """Get Microsoft Graph API access token."""
    if not _credential:
        raise RuntimeError(
            "Microsoft Graph API not configured. Please set AZURE_CLIENT_ID, "
            "AZURE_CLIENT_SECRET, and AZURE_TENANT_ID environment variables."
        )

    # Request token for Microsoft Graph scope
    token = _credential.get_token("https://graph.microsoft.com/.default")
    return token.token


async def list_sharepoint_sites(search_query: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    List all SharePoint sites accessible to the application.

    Uses Microsoft Graph API /sites endpoint to discover sites across the tenant.

    Args:
        search_query: Optional search term to filter sites by name/description

    Returns:
        List of site dictionaries with keys:
        - name: Display name of the site
        - site_url: Full web URL (https://tenant.sharepoint.com/sites/sitename)
        - site_id: Unique site identifier
        - description: Site description (may be empty)
        - web_url: Alias for site_url (for compatibility)

    Raises:
        RuntimeError: If Graph API credentials not configured
        Exception: If Graph API request fails
    """
    if not _credential:
        raise RuntimeError(
            "Microsoft Graph API not configured. Please set AZURE_CLIENT_ID, "
            "AZURE_CLIENT_SECRET, and AZURE_TENANT_ID environment variables."
        )

    try:
        logger.info(f"Listing SharePoint sites{f' matching: {search_query}' if search_query else ''}")

        # Get access token
        token = await _get_access_token()

        # Build Graph API request
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        # Construct URL with optional search
        if search_query:
            url = f"{GRAPH_API_ENDPOINT}/sites?search={search_query}"
        else:
            url = f"{GRAPH_API_ENDPOINT}/sites?$top=100"

        # Make request
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()

        # Transform to consistent format
        sites = []
        if data and 'value' in data:
            for site in data['value']:
                site_data = {
                    "name": site.get('displayName') or site.get('name') or "Unknown",
                    "site_url": site.get('webUrl'),
                    "site_id": site.get('id'),
                    "description": site.get('description', ''),
                    "web_url": site.get('webUrl'),  # Alias for compatibility
                }
                sites.append(site_data)
                logger.debug(f"Found site: {site_data['name']} ({site_data['site_url']})")

        logger.info(f"Found {len(sites)} SharePoint site(s)")
        return sites

    except httpx.HTTPStatusError as e:
        logger.error(f"HTTP error listing SharePoint sites: {e.response.status_code} - {e.response.text}")
        raise Exception(f"Failed to list SharePoint sites: HTTP {e.response.status_code}")
    except Exception as e:
        logger.error(f"Error listing SharePoint sites: {str(e)}")
        raise Exception(f"Failed to list SharePoint sites: {str(e)}")


async def get_site_by_url(site_url: str) -> Optional[Dict[str, Any]]:
    """
    Get site information by URL.

    Args:
        site_url: Full SharePoint site URL

    Returns:
        Site dictionary or None if not found
    """
    if not _credential:
        raise RuntimeError("Microsoft Graph API not configured")

    try:
        # Extract hostname and site path from URL
        # Example: https://tenant.sharepoint.com/sites/sitename
        from urllib.parse import urlparse
        parsed = urlparse(site_url)
        hostname = parsed.netloc
        site_path = parsed.path

        # Graph API format: hostname:site_path
        site_identifier = f"{hostname}:{site_path}"

        logger.debug(f"Getting site info for: {site_identifier}")

        # Get access token
        token = await _get_access_token()

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        url = f"{GRAPH_API_ENDPOINT}/sites/{site_identifier}"

        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=headers)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            site = response.json()

        if site:
            return {
                "name": site.get('displayName') or site.get('name'),
                "site_url": site.get('webUrl'),
                "site_id": site.get('id'),
                "description": site.get('description', ''),
                "web_url": site.get('webUrl'),
            }
        return None

    except httpx.HTTPStatusError as e:
        if e.response.status_code == 404:
            logger.warning(f"Site not found: {site_url}")
            return None
        logger.error(f"HTTP error getting site {site_url}: {e.response.status_code}")
        return None
    except Exception as e:
        logger.error(f"Error getting site by URL {site_url}: {str(e)}")
        return None


def validate_site_url(site_url: str) -> bool:
    """
    Validate that a site URL is properly formatted.

    Args:
        site_url: SharePoint site URL to validate

    Returns:
        True if URL is valid format, False otherwise
    """
    if not site_url:
        return False

    # Must be HTTPS
    if not site_url.startswith("https://"):
        return False

    # Must contain sharepoint.com
    if "sharepoint.com" not in site_url:
        return False

    # Should have /sites/ path (most common)
    # Note: Root site and personal sites have different patterns
    # We'll be lenient here

    return True


# Customer site mapping support
_customer_site_cache: Dict[str, str] = {}


def load_customer_site_mappings() -> Dict[str, str]:
    """
    Load customer name to site URL mappings from environment.

    Environment variable format:
    CUSTOMER_SITES=Acme Corp:https://...,Wayne Ent:https://...

    Returns:
        Dictionary mapping customer names to site URLs
    """
    global _customer_site_cache

    if _customer_site_cache:
        return _customer_site_cache

    mappings = {}
    customer_sites_env = os.getenv('CUSTOMER_SITES', '')

    if customer_sites_env:
        try:
            for mapping in customer_sites_env.split(','):
                if ':' in mapping:
                    name, url = mapping.split(':', 1)
                    name = name.strip()
                    url = url.strip()
                    if name and validate_site_url(url):
                        mappings[name] = url
                        logger.debug(f"Loaded customer mapping: {name} -> {url}")
        except Exception as e:
            logger.error(f"Error parsing CUSTOMER_SITES: {str(e)}")

    _customer_site_cache = mappings
    return mappings


async def get_site_url_for_customer(customer_name: str) -> Optional[str]:
    """
    Resolve customer name to SharePoint site URL.

    First checks customer mappings from environment, then falls back to
    searching SharePoint sites by customer name.

    Args:
        customer_name: Friendly customer name

    Returns:
        Site URL if found, None otherwise
    """
    # Check mappings first
    mappings = load_customer_site_mappings()
    if customer_name in mappings:
        logger.info(f"Found customer '{customer_name}' in mappings: {mappings[customer_name]}")
        return mappings[customer_name]

    # Fallback: Search for site by name
    logger.info(f"Customer '{customer_name}' not in mappings, searching SharePoint sites...")
    try:
        sites = await list_sharepoint_sites(customer_name)
        if sites:
            site_url = sites[0]["site_url"]
            logger.info(f"Found site for '{customer_name}': {site_url}")
            return site_url
    except Exception as e:
        logger.error(f"Error searching for customer site: {str(e)}")

    logger.warning(f"Could not find SharePoint site for customer: {customer_name}")
    return None
