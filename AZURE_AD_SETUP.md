# Azure AD App Configuration for SharePoint MCP with OAuth

This guide explains how to configure your Azure AD app registration to support per-user OAuth authentication with delegated permissions.

## Overview

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
