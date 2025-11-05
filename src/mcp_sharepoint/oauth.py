"""
OAuth 2.0 authentication module for SharePoint MCP with delegated permissions.

This module implements per-user authentication using MSAL (Microsoft Authentication Library)
with Azure AD delegated permissions, enabling proper audit trails and user-specific access.
"""

import os
import logging
from typing import Optional, Dict, Any
from datetime import datetime, timedelta
import msal
from azure.identity import ClientSecretCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

logger = logging.getLogger('mcp_sharepoint.oauth')


class SharePointOAuthManager:
    """
    Manages OAuth 2.0 authentication for SharePoint with delegated permissions.

    Features:
    - Per-user token management
    - Automatic token refresh
    - MSAL integration for Azure AD
    - Session-based token storage
    """

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        redirect_uri: str = "http://localhost:8000/oauth/callback"
    ):
        """
        Initialize OAuth manager.

        Args:
            client_id: Azure AD application (client) ID
            client_secret: Azure AD client secret
            tenant_id: Azure AD tenant ID
            redirect_uri: OAuth redirect URI (must match Azure AD app registration)
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.redirect_uri = redirect_uri

        # MSAL authority URL
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"

        # Required scopes for SharePoint with delegated permissions
        # Using .default scope to request all delegated permissions configured in Azure AD
        # offline_access is automatically included with .default for confidential clients
        self.scopes = [
            "https://graph.microsoft.com/.default"
        ]

        # Create MSAL confidential client application
        self.msal_app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )

        # In-memory token cache (replace with Redis/database for production)
        self._token_cache: Dict[str, Dict[str, Any]] = {}

        logger.info(f"SharePointOAuthManager initialized for tenant: {tenant_id}")

    def get_authorization_url(self, state: Optional[str] = None) -> str:
        """
        Generate OAuth authorization URL for user consent.

        Args:
            state: Optional state parameter for CSRF protection

        Returns:
            Authorization URL to redirect user to
        """
        auth_url = self.msal_app.get_authorization_request_url(
            scopes=self.scopes,
            redirect_uri=self.redirect_uri,
            state=state or "default_state"
        )
        logger.info(f"Generated authorization URL: {auth_url}")
        return auth_url

    def acquire_token_by_authorization_code(
        self,
        authorization_code: str
    ) -> Dict[str, Any]:
        """
        Exchange authorization code for access token.

        Args:
            authorization_code: Authorization code from OAuth callback

        Returns:
            Token response including access_token, refresh_token, expires_in

        Raises:
            ValueError: If token acquisition fails
        """
        result = self.msal_app.acquire_token_by_authorization_code(
            code=authorization_code,
            scopes=self.scopes,
            redirect_uri=self.redirect_uri
        )

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Unknown error"))
            logger.error(f"Token acquisition failed: {error}")
            raise ValueError(f"Failed to acquire token: {error}")

        logger.info("Successfully acquired token by authorization code")
        return result

    def acquire_token_by_refresh_token(
        self,
        refresh_token: str
    ) -> Dict[str, Any]:
        """
        Refresh access token using refresh token.

        Args:
            refresh_token: Refresh token from previous authentication

        Returns:
            New token response

        Raises:
            ValueError: If token refresh fails
        """
        result = self.msal_app.acquire_token_by_refresh_token(
            refresh_token=refresh_token,
            scopes=self.scopes
        )

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Unknown error"))
            logger.error(f"Token refresh failed: {error}")
            raise ValueError(f"Failed to refresh token: {error}")

        logger.info("Successfully refreshed access token")
        return result

    def store_user_token(self, user_id: str, token_response: Dict[str, Any]) -> None:
        """
        Store user token in cache with expiration time.

        Args:
            user_id: Unique user identifier
            token_response: Token response from MSAL
        """
        expires_in = token_response.get("expires_in", 3600)
        expires_at = datetime.utcnow() + timedelta(seconds=expires_in)

        self._token_cache[user_id] = {
            "access_token": token_response["access_token"],
            "refresh_token": token_response.get("refresh_token"),
            "expires_at": expires_at,
            "token_type": token_response.get("token_type", "Bearer")
        }

        logger.info(f"Stored token for user: {user_id}, expires at: {expires_at}")

    def get_user_token(self, user_id: str) -> Optional[str]:
        """
        Get valid access token for user, refreshing if necessary.

        Args:
            user_id: Unique user identifier

        Returns:
            Valid access token or None if not authenticated
        """
        if user_id not in self._token_cache:
            logger.warning(f"No token found for user: {user_id}")
            return None

        token_data = self._token_cache[user_id]

        # Check if token is expired (with 5-minute buffer)
        if datetime.utcnow() >= token_data["expires_at"] - timedelta(minutes=5):
            logger.info(f"Token expired for user: {user_id}, refreshing...")

            if not token_data.get("refresh_token"):
                logger.error(f"No refresh token available for user: {user_id}")
                return None

            try:
                # Refresh the token
                new_token = self.acquire_token_by_refresh_token(
                    token_data["refresh_token"]
                )
                self.store_user_token(user_id, new_token)
                return new_token["access_token"]
            except ValueError as e:
                logger.error(f"Failed to refresh token for user {user_id}: {e}")
                # Remove invalid token from cache
                del self._token_cache[user_id]
                return None

        return token_data["access_token"]

    def get_sharepoint_context(
        self,
        site_url: str,
        user_id: str
    ) -> ClientContext:
        """
        Create SharePoint ClientContext with user's access token.

        Args:
            site_url: SharePoint site URL
            user_id: Unique user identifier

        Returns:
            Authenticated ClientContext

        Raises:
            ValueError: If user is not authenticated or token is invalid
        """
        access_token = self.get_user_token(user_id)

        if not access_token:
            raise ValueError(
                f"User {user_id} is not authenticated. "
                "Please complete OAuth authorization flow."
            )

        # Create context with user token
        ctx = ClientContext(site_url)
        ctx.with_access_token(access_token)

        logger.debug(f"Created SharePoint context for user {user_id} at {site_url}")
        return ctx


# Global OAuth manager instance (initialized in server_http.py)
_oauth_manager: Optional[SharePointOAuthManager] = None


def init_oauth_manager(
    client_id: str,
    client_secret: str,
    tenant_id: str,
    redirect_uri: str = "http://localhost:8000/oauth/callback"
) -> SharePointOAuthManager:
    """
    Initialize global OAuth manager instance.

    Args:
        client_id: Azure AD application ID
        client_secret: Azure AD client secret
        tenant_id: Azure AD tenant ID
        redirect_uri: OAuth redirect URI

    Returns:
        Initialized SharePointOAuthManager instance
    """
    global _oauth_manager
    _oauth_manager = SharePointOAuthManager(
        client_id=client_id,
        client_secret=client_secret,
        tenant_id=tenant_id,
        redirect_uri=redirect_uri
    )
    return _oauth_manager


def get_oauth_manager() -> SharePointOAuthManager:
    """
    Get global OAuth manager instance.

    Returns:
        SharePointOAuthManager instance

    Raises:
        RuntimeError: If OAuth manager not initialized
    """
    if _oauth_manager is None:
        raise RuntimeError(
            "OAuth manager not initialized. Call init_oauth_manager() first."
        )
    return _oauth_manager
