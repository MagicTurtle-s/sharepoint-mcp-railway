# Azure AD OAuth Configuration for SharePoint MCP

This guide walks you through configuring Azure AD for per-user OAuth authentication with the SharePoint MCP server.

## Overview

The SharePoint MCP supports two authentication modes:

1. **Application Permissions (Legacy)**: Service account with tenant-wide access
2. **Delegated Permissions (OAuth 2.0)**: Per-user authentication (Recommended)

This guide covers setting up **Delegated Permissions** for production multi-user deployments.

## Benefits of OAuth/Delegated Permissions

- ✅ **Per-user authentication**: Each user authenticates with their own credentials
- ✅ **Proper audit trails**: SharePoint logs operations under the actual user's identity
- ✅ **Granular permissions**: Users only access what they're authorized to see
- ✅ **Automatic token refresh**: Long-lived refresh tokens (6 months) with short-lived access tokens (12 hours)
- ✅ **Secure**: No shared credentials, tokens stored securely per session

## Prerequisites

- Azure AD admin access to register applications
- SharePoint Online tenant
- Railway account (for cloud deployment) or local server
- Python 3.10 or higher

## Step 1: Create Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure the registration:
   - **Name**: `SharePoint MCP Server` (or your preferred name)
   - **Supported account types**: Select one of:
     - `Accounts in this organizational directory only (Single tenant)` - Most common
     - `Accounts in any organizational directory (Multi-tenant)` - If supporting multiple tenants
   - **Redirect URI**:
     - Platform: `Web`
     - URI: `http://localhost:8000/oauth/callback` (for local testing)
     - *Note: You'll add your Railway URL later after deployment*
5. Click **Register**

## Step 2: Configure API Permissions (Delegated)

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Add these permissions:
   - `Files.ReadWrite.All` - Read and write files in all site collections
   - `Sites.ReadWrite.All` - Read and write items in all site collections
   - `offline_access` - Maintain access to data you have given it access to
6. Click **Add permissions**
7. Click **Grant admin consent for [Your Tenant]** (requires admin rights)
8. Confirm the consent

### Why Delegated vs Application Permissions?

| Feature | Delegated Permissions | Application Permissions |
|---------|----------------------|------------------------|
| Authentication | User credentials (OAuth) | App credentials (Client Secret) |
| Audit trails | Logged as specific user | Logged as service account |
| Access scope | User's permissions only | Tenant-wide access |
| Use case | Multi-user production | Development, single-user |

## Step 3: Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Configure the secret:
   - **Description**: `SharePoint MCP Production` (or your preferred name)
   - **Expires**: Choose expiration period (recommend: 24 months for production)
4. Click **Add**
5. **IMPORTANT**: Copy the **Value** immediately - you won't be able to see it again!
6. Store it securely (you'll use it as `SHP_ID_APP_SECRET`)

## Step 4: Collect Required Information

You'll need these values for your environment variables:

| Variable | Where to Find | Example |
|----------|---------------|---------|
| `SHP_ID_APP` | App registration → Overview → Application (client) ID | `12345678-1234-1234-1234-123456789abc` |
| `SHP_ID_APP_SECRET` | The secret value you just copied | `abc123~defGHI456...` |
| `SHP_TENANT_ID` | App registration → Overview → Directory (tenant) ID | `87654321-4321-4321-4321-cba987654321` |
| `SHP_SITE_URL` | Your SharePoint site URL | `https://contoso.sharepoint.com/sites/engineering` |
| `SHP_DOC_LIBRARY` | Document library path | `Shared Documents` |

## Step 5: Deploy to Railway

### 5.1 Push Code to GitHub

```bash
cd /c/Users/jonat/sharepoint-mcp-railway
git add .
git commit -m "Add OAuth support for SharePoint MCP"
git push origin main
```

### 5.2 Deploy to Railway

```bash
railway up
```

### 5.3 Set Environment Variables

Via Railway web dashboard (https://railway.app):

1. Select your project → `sharepoint-mcp-railway`
2. Go to **Variables** tab
3. Add these variables:

```
SHP_ID_APP=your-application-client-id
SHP_ID_APP_SECRET=your-client-secret-value
SHP_TENANT_ID=your-tenant-id
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_DOC_LIBRARY=Shared Documents
```

### 5.4 Get Railway Public URL

1. In Railway dashboard, go to **Settings** tab
2. Under **Domains**, click **Generate Domain**
3. Copy the generated URL (e.g., `sharepoint-mcp-railway-production-xxxx.up.railway.app`)

## Step 6: Update Azure AD Redirect URI

1. Return to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Select your SharePoint MCP app
3. Go to **Authentication**
4. Under **Web** platform → **Redirect URIs**, click **Add URI**
5. Add your Railway URL with OAuth callback path:
   ```
   https://sharepoint-mcp-railway-production-xxxx.up.railway.app/oauth/callback
   ```
   *(Replace `xxxx` with your actual Railway subdomain)*
6. Click **Save**

## Step 7: Configure Claude Code CLI

Add the SharePoint MCP server to your Claude Code configuration:

```bash
claude mcp add sharepoint-oauth https://sharepoint-mcp-railway-production-xxxx.up.railway.app/mcp
```

Verify it was added:

```bash
claude mcp list
```

## Step 8: Test OAuth Flow

### 8.1 Test the Authorization Endpoint

Open your browser and navigate to:

```
https://sharepoint-mcp-railway-production-xxxx.up.railway.app/oauth/authorize
```

You should be redirected to Microsoft login page.

### 8.2 Complete OAuth Flow

1. Sign in with your Microsoft account
2. Consent to the requested permissions
3. You should be redirected to the callback page showing "OAuth authorization successful!"
4. The server stores your access and refresh tokens

### 8.3 Test with Claude Code

```bash
claude "List the folders in SharePoint"
```

Claude Code should use the OAuth-authenticated session to access SharePoint on your behalf.

## OAuth Token Flow

### Initial Authentication

```
User → Claude Code → MCP Server (no token) → /oauth/authorize
User → Azure AD (login) → Authorization Code
Authorization Code → MCP Server → Exchange for tokens
Azure AD → Access Token + Refresh Token → MCP Server
MCP Server → Store tokens → SharePoint API call → Response
```

### Subsequent Requests

```
User → Claude Code → MCP Server (has token)
  ├─ Token valid → SharePoint API call → Response
  └─ Token expired → Refresh with refresh_token → New access token → SharePoint API call → Response
```

## Token Lifecycle

| Token Type | Lifetime | Purpose | Storage |
|------------|----------|---------|---------|
| **Authorization Code** | 10 minutes | Exchange for tokens | Not stored (single-use) |
| **Access Token** | 12 hours | Authenticate SharePoint API calls | In-memory cache (per user) |
| **Refresh Token** | 6 months | Obtain new access tokens | In-memory cache (per user) |

### Automatic Token Refresh

The OAuth manager automatically refreshes expired access tokens using the refresh token. This happens transparently when making SharePoint API calls.

## Troubleshooting

### Error: "AADSTS50011: The redirect URI specified in the request does not match"

**Solution**: Ensure the redirect URI in Azure AD exactly matches your Railway URL:
- Azure AD: `https://sharepoint-mcp-railway-production-xxxx.up.railway.app/oauth/callback`
- No trailing slashes
- HTTPS (not HTTP)
- Exact subdomain match

### Error: "Insufficient privileges to complete the operation"

**Causes**:
1. Admin consent not granted for API permissions
2. User doesn't have SharePoint permissions

**Solutions**:
1. In Azure AD → API permissions → Grant admin consent
2. Verify user has SharePoint site access

### Error: "Token expired" or "Invalid token"

**Cause**: Access token expired and refresh failed

**Solutions**:
1. Re-authenticate: Visit `/oauth/authorize` again
2. Check refresh token hasn't expired (6 months)
3. Verify client secret is still valid

### Error: "OAuth manager not initialized"

**Cause**: Missing environment variables

**Solution**: Verify all required variables are set in Railway:
- `SHP_ID_APP`
- `SHP_ID_APP_SECRET`
- `SHP_TENANT_ID`

## Security Best Practices

1. **Rotate client secrets regularly**: Set expiration dates and rotate before expiry
2. **Use HTTPS only**: Never use HTTP for OAuth callbacks in production
3. **Validate redirect URIs**: Only register exact URLs you control
4. **Least privilege**: Only request permissions you need
5. **Monitor access**: Review Azure AD sign-in logs regularly
6. **Token storage**: In production, consider persistent storage (e.g., Redis) instead of in-memory cache

## Migration from Application Permissions

If you're currently using Application Permissions (service account), here's how to migrate:

### Before Migration

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "mcp-sharepoint",
      "env": {
        "SHP_ID_APP": "app-id",
        "SHP_ID_APP_SECRET": "app-secret",
        "SHP_SITE_URL": "https://tenant.sharepoint.com/sites/site",
        "SHP_TENANT_ID": "tenant-id"
      }
    }
  }
}
```

### After Migration

```bash
# Add to Claude Code
claude mcp add sharepoint-oauth https://sharepoint-mcp-railway-production-xxxx.up.railway.app/mcp
```

### Benefits After Migration

- ✅ Each user sees only their authorized content
- ✅ SharePoint audit logs show actual users, not service account
- ✅ No shared credentials to manage
- ✅ Better compliance and governance

## Support

For issues or questions:
- GitHub Issues: https://github.com/MagicTurtle-s/sharepoint-mcp-railway/issues
- Original Project: https://github.com/Sofias-ai/mcp-sharepoint

## References

- [Microsoft Identity Platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [Microsoft Graph Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [OAuth 2.0 Authorization Code Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow)
- [Model Context Protocol](https://modelcontextprotocol.io)

The SharePoint MCP now supports two authentication modes:
1. **Application Permissions** (legacy) - Service account with tenant-wide access
2. **Delegated Permissions** (recommended) - Per-user OAuth with proper audit trails

## Prerequisites

- Azure AD tenant with SharePoint Online
- Azure AD app registration created
- Admin access to Azure Portal

## Step 1: Update API Permissions

### Remove Application Permissions (if present)

1. Navigate to **Azure Portal** → **Azure Active Directory** → **App registrations**
2. Select your SharePoint MCP app
3. Go to **API permissions**
4. Remove any existing Application permissions:
   - `Sites.FullControl.All` (Application)
   - `Files.ReadWrite.All` (Application)

### Add Delegated Permissions

1. Click **Add a permission**
2. Select **Microsoft Graph**
3. Choose **Delegated permissions**
4. Add the following permissions:
   - `Files.ReadWrite.All` - Read and write files in all site collections
   - `Sites.ReadWrite.All` - Read and write items and lists in all site collections
   - `offline_access` - Maintain access to data you have given it access to

5. Click **Add permissions**

### Grant Admin Consent

1. Click **Grant admin consent for [Your Tenant]**
2. Confirm the consent prompt

## Step 2: Configure Authentication

### Add Redirect URIs

1. Go to **Authentication** in your app registration
2. Click **Add a platform** → **Web**
3. Add the following redirect URIs:
   - **Local testing**: `http://localhost:8000/oauth/callback`
   - **Railway production**: `https://your-railway-app.up.railway.app/oauth/callback`

4. Under **Implicit grant and hybrid flows**, ensure:
   - ☑ ID tokens (used for implicit and hybrid flows)

5. Click **Save**

### Configure Token Settings

1. Go to **Token configuration**
2. Add optional claims for ID token:
   - `email` - User's email address
   - `preferred_username` - User's UPN

3. Click **Save**

## Step 3: Update Environment Variables

Your Railway deployment needs these environment variables:

```env
# Azure AD Credentials (required)
AZURE_CLIENT_ID=<your-app-client-id>
AZURE_CLIENT_SECRET=<your-app-client-secret>
AZURE_TENANT_ID=<your-tenant-id>

# Railway will automatically provide this
RAILWAY_PUBLIC_DOMAIN=<auto-populated>
```

**Note**: Do NOT set a static redirect URI in environment variables - the server dynamically constructs it from `RAILWAY_PUBLIC_DOMAIN`.

## Step 4: Test OAuth Flow

### Local Testing

1. Install dependencies:
   ```bash
   cd C:\Users\jonat\sharepoint-mcp-railway
   pip install -e .
   ```

2. Create a `.env` file:
   ```env
   AZURE_CLIENT_ID=<your-app-id>
   AZURE_CLIENT_SECRET=<your-app-secret>
   AZURE_TENANT_ID=<your-tenant-id>
   ```

3. Start the server:
   ```bash
   python -m mcp_sharepoint.server_http
   ```

4. Open browser to `http://localhost:8000/oauth/authorize`
5. Complete the OAuth consent flow
6. You should see a success page with your session ID

### Production Testing (Railway)

1. Deploy to Railway:
   ```bash
   railway up
   ```

2. Set environment variables in Railway dashboard
3. Open `https://your-railway-app.up.railway.app/oauth/authorize`
4. Complete OAuth flow
5. Save the session ID displayed on success page

## Step 5: Using OAuth with Claude Code CLI

### Add to Claude Code

```bash
claude mcp add --transport http --scope user sharepoint https://your-railway-app.up.railway.app/mcp
```

### First Use - OAuth Setup

When Claude Code first uses the SharePoint MCP:

1. It will detect that OAuth is required
2. You'll be prompted to visit the authorization URL
3. Complete the OAuth flow in your browser
4. Copy the session ID from the success page
5. Provide the session ID back to Claude Code

### Subsequent Use

After initial OAuth setup:
- Your tokens are stored securely in the server
- Access tokens are automatically refreshed (12-hour lifespan)
- Refresh tokens last 6 months
- All SharePoint operations use your user identity

## Security Considerations

### Token Storage

- Access tokens: Stored in-memory on the server (12-hour lifespan)
- Refresh tokens: Stored in-memory on the server (6-month lifespan)
- Session IDs: Used to map users to their tokens
- **Important**: Server restart will clear all tokens

### User Identity

All SharePoint operations will be performed as the authenticated user:
- Proper audit trails in SharePoint
- Respects user's SharePoint permissions
- No over-privileged service account

### Token Refresh

The OAuth manager automatically refreshes tokens:
- Access tokens expire after 12 hours
- Refresh tokens used to get new access tokens
- Automatic refresh happens transparently
- Users only need to re-authenticate if refresh token expires (6 months)

## Troubleshooting

### "Insufficient privileges" error

**Cause**: App permissions not granted or delegated permissions missing

**Solution**:
1. Verify delegated permissions are added
2. Grant admin consent
3. Wait 5-10 minutes for permissions to propagate

### "Redirect URI mismatch" error

**Cause**: The redirect URI in the OAuth request doesn't match Azure AD configuration

**Solution**:
1. Check that your Railway URL matches the redirect URI in Azure AD
2. Ensure redirect URI format: `https://your-app.up.railway.app/oauth/callback`
3. No trailing slashes

### "Token refresh failed" error

**Cause**: Refresh token expired or invalid

**Solution**:
1. User needs to re-authenticate
2. Visit `/oauth/authorize` endpoint again
3. Complete OAuth flow to get new tokens

### "OAuth manager not initialized" warning

**Cause**: Environment variables missing or server startup failed

**Solution**:
1. Verify all required environment variables are set
2. Check server logs for startup errors
3. Restart the Railway deployment

## Architecture Notes

### How OAuth Works in SharePoint MCP

1. **Initial Setup** (one-time):
   - User visits `/oauth/authorize`
   - Redirected to Azure AD login
   - User grants consent
   - Azure AD redirects to `/oauth/callback` with auth code
   - Server exchanges auth code for tokens
   - Server stores tokens with session ID
   - User receives session ID

2. **Tool Execution** (ongoing):
   - Claude Code calls SharePoint tool with user context
   - Server retrieves user's access token by session ID
   - If token expired, automatically refresh using refresh token
   - Create SharePoint context with user's access token
   - Execute SharePoint operation as user
   - Return results to Claude Code

3. **Token Lifecycle**:
   - Access token: 12 hours → Auto-refresh
   - Refresh token: 6 months → Requires re-authentication

### Backward Compatibility

The implementation maintains backward compatibility:
- Tools accept optional `user_id` parameter
- If `user_id` provided → Use OAuth
- If `user_id` not provided → Fall back to app credentials (legacy mode)

This allows gradual migration and testing without breaking existing integrations.

## Next Steps

After completing Azure AD setup:

1. ✅ Test OAuth flow locally
2. ✅ Deploy to Railway
3. ✅ Add to Claude Code CLI
4. ✅ Test all SharePoint operations
5. ✅ Verify audit trails in SharePoint

## Reference

- [Microsoft Identity Platform - Delegated Permissions](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
- [Microsoft Graph - Files API](https://learn.microsoft.com/en-us/graph/api/resources/onedrive)
- [SharePoint Online - OAuth](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs)
