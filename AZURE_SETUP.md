# Azure AD Application Setup Guide

This guide walks through creating and configuring an Azure AD application for the SharePoint MCP server.

## Prerequisites

- Microsoft 365 / Azure AD tenant
- Admin access to Azure Portal
- SharePoint Online access

## Step 1: Create Azure AD App Registration

1. **Navigate to Azure Portal**:
   - Go to [portal.azure.com](https://portal.azure.com)
   - Sign in with admin account

2. **Open App Registrations**:
   - Search for "Azure Active Directory" or "Microsoft Entra ID"
   - In the left menu, click **App registrations**
   - Click **+ New registration**

3. **Register the Application**:
   - **Name**: `SharePoint MCP Server - Railway`
   - **Supported account types**: `Accounts in this organizational directory only (Single tenant)`
   - **Redirect URI**: Leave blank (not needed for service-to-service)
   - Click **Register**

4. **Save the Application (client) ID**:
   - On the Overview page, copy the **Application (client) ID**
   - This is your `AZURE_CLIENT_ID`
   - Also copy the **Directory (tenant) ID**
   - This is your `AZURE_TENANT_ID`

## Step 2: Create Client Secret

1. **Navigate to Certificates & secrets**:
   - In your app registration, click **Certificates & secrets** in the left menu

2. **Create New Client Secret**:
   - Click **+ New client secret**
   - **Description**: `Railway Production`
   - **Expires**: `24 months` (or your preferred duration)
   - Click **Add**

3. **Copy the Secret Value**:
   - **IMPORTANT**: Copy the **Value** immediately (not the Secret ID)
   - This is your `AZURE_CLIENT_SECRET`
   - **You cannot view this again** after leaving the page
   - If you lose it, you'll need to create a new secret

## Step 3: Configure API Permissions

### Option A: All Sites Access (Easiest)

**Use this if you want the MCP server to access all SharePoint sites in your tenant.**

1. **Navigate to API permissions**:
   - Click **API permissions** in the left menu

2. **Add Microsoft Graph permissions**:
   - Click **+ Add a permission**
   - Select **Microsoft Graph**
   - Select **Application permissions** (not Delegated)

3. **Add these permissions**:
   - Search and select **`Sites.ReadWrite.All`**
     - Allows read and write access to all SharePoint sites
   - Search and select **`Files.ReadWrite.All`**
     - Allows read and write access to all files

4. **Grant Admin Consent**:
   - Click **Grant admin consent for [Your Organization]**
   - Click **Yes** to confirm
   - You should see green checkmarks next to each permission

### Option B: Selected Sites Only (More Secure)

**Use this if you want to limit access to specific SharePoint sites only.**

1. **Add Microsoft Graph permissions**:
   - Click **+ Add a permission**
   - Select **Microsoft Graph**
   - Select **Application permissions**

2. **Add these permissions**:
   - Search and select **`Sites.Selected`**
     - Allows access only to explicitly granted sites
   - Search and select **`Files.ReadWrite.All`**

3. **Grant Admin Consent**:
   - Click **Grant admin consent for [Your Organization]**
   - Click **Yes**

4. **Grant Site-Specific Access** (requires PowerShell):

```powershell
# Install required module
Install-Module PnP.PowerShell -Scope CurrentUser

# Connect to SharePoint admin
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive

# Grant permissions to specific sites
# Get your app's service principal ID first
$appId = "YOUR-APPLICATION-CLIENT-ID"
$permission = "write"  # or "read" for read-only

# Grant access to each site
Grant-PnPAzureADAppSitePermission -AppId $appId -DisplayName "SharePoint MCP Server" -Site "https://yourtenant.sharepoint.com/sites/acme-corp" -Permissions $permission

Grant-PnPAzureADAppSitePermission -AppId $appId -DisplayName "SharePoint MCP Server" -Site "https://yourtenant.sharepoint.com/sites/wayne-ent" -Permissions $permission

# Repeat for each of your 50+ sites...
```

**Pros of Sites.Selected**:
- More secure - principle of least privilege
- Explicit control over which sites the app can access
- Easier to audit and manage

**Cons of Sites.Selected**:
- Requires PowerShell setup for each site
- More maintenance when adding new customer sites
- Must run PowerShell commands for each new site

## Step 4: Verify Configuration

Your App Registration should now have:

✅ **Overview**:
- Application (client) ID → `AZURE_CLIENT_ID`
- Directory (tenant) ID → `AZURE_TENANT_ID`

✅ **Certificates & secrets**:
- Client secret value → `AZURE_CLIENT_SECRET`

✅ **API permissions** (Option A):
- Microsoft Graph → Sites.ReadWrite.All (Application) ✓ Granted
- Microsoft Graph → Files.ReadWrite.All (Application) ✓ Granted

Or **API permissions** (Option B):
- Microsoft Graph → Sites.Selected (Application) ✓ Granted
- Microsoft Graph → Files.ReadWrite.All (Application) ✓ Granted
- Site-specific permissions granted via PowerShell

## Step 5: Test the Configuration

Before deploying to Railway, test locally:

1. **Create `.env` file**:
```env
AZURE_CLIENT_ID=your-application-client-id
AZURE_CLIENT_SECRET=your-client-secret-value
AZURE_TENANT_ID=your-tenant-id
```

2. **Test authentication**:
```python
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient

credential = ClientSecretCredential(
    tenant_id="your-tenant-id",
    client_id="your-client-id",
    client_secret="your-client-secret"
)

# This should not raise an error
graph_client = GraphServiceClient(credential)
print("✓ Authentication successful!")
```

3. **Test SharePoint access**:
```python
import asyncio
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

credentials = ClientCredential("your-client-id", "your-client-secret")
ctx = ClientContext("https://yourtenant.sharepoint.com/sites/yoursite").with_credentials(credentials)

# Try to get site info
site = ctx.web
ctx.load(site)
ctx.execute_query()

print(f"✓ Connected to: {site.title}")
```

## Step 6: Deploy to Railway

Once verified, add these environment variables to Railway:

```
AZURE_CLIENT_ID=your-application-client-id
AZURE_CLIENT_SECRET=your-client-secret-value
AZURE_TENANT_ID=your-tenant-id
```

See [RAILWAY_DEPLOYMENT.md](RAILWAY_DEPLOYMENT.md) for full deployment instructions.

## Troubleshooting

### "AADSTS7000215: Invalid client secret provided"

**Cause**: Client secret is incorrect or expired

**Fix**:
1. Generate new client secret in Azure Portal
2. Update `AZURE_CLIENT_SECRET` environment variable
3. Redeploy

### "AADSTS70001: Application not found in the directory"

**Cause**: Application (client) ID is incorrect

**Fix**:
1. Verify `AZURE_CLIENT_ID` matches the one in Azure Portal
2. Check you're using the right tenant

### "Insufficient privileges to complete the operation"

**Cause**: Missing or not-consented permissions

**Fix**:
1. Verify `Sites.ReadWrite.All` permission is added
2. Ensure admin consent is granted (green checkmarks)
3. Wait 5-10 minutes for permissions to propagate

### "Access denied" when accessing specific site

**Cause**: Using `Sites.Selected` but site permission not granted

**Fix**:
1. Run PowerShell command to grant site access (see Option B above)
2. Or switch to `Sites.ReadWrite.All` permission

### PowerShell "Connect-PnPOnline" fails

**Cause**: Module not installed or authentication issue

**Fix**:
```powershell
# Uninstall old version
Uninstall-Module SharePointPnPPowerShellOnline -Force
Uninstall-Module PnP.PowerShell -Force

# Install latest
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Try connecting again
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive
```

## Security Best Practices

### 1. Use Short-Lived Secrets

- Set client secret expiration to 6-12 months (not 24)
- Set calendar reminder to rotate before expiration
- Use Azure Key Vault for production deployments

### 2. Monitor App Usage

- Enable audit logging in Azure AD
- Review sign-in logs regularly
- Watch for unusual access patterns

### 3. Implement IP Restrictions (Optional)

In Azure Portal → App Registration → Authentication:
- Configure Conditional Access policies
- Restrict to Railway IP ranges if possible

### 4. Use Managed Identity (Advanced)

For production, consider:
- Azure Managed Identity instead of client secrets
- Requires hosting on Azure (App Service, Functions)
- Eliminates secret management

### 5. Regular Permission Review

- Audit permissions quarterly
- Remove unused permissions
- Update to Sites.Selected if using Sites.ReadWrite.All

## Client Secret Rotation

Client secrets expire. Here's how to rotate without downtime:

### Before Expiration:

1. **Create new secret** in Azure Portal
2. **Test new secret** locally
3. **Update Railway** environment variable with new secret
4. **Verify deployment** works with new secret
5. **Delete old secret** after confirming

### After Expiration (Emergency):

1. **Create new secret immediately**
2. **Update Railway** environment variable
3. **Trigger redeployment**
4. **Verify** with health check endpoint

## Multi-Tenant Setup (Advanced)

If you need to access SharePoint from multiple Microsoft 365 tenants:

1. Create separate app registrations in each tenant
2. Store credentials per-tenant:
   ```
   TENANT1_CLIENT_ID=...
   TENANT1_CLIENT_SECRET=...

   TENANT2_CLIENT_ID=...
   TENANT2_CLIENT_SECRET=...
   ```
3. Modify code to select credentials based on site URL tenant

## Appendix: Permission Scopes Explained

| Permission | Type | Description | Use Case |
|------------|------|-------------|----------|
| `Sites.ReadWrite.All` | Application | Read/write all site collections | Full tenant access, simplest setup |
| `Sites.Selected` | Application | Access only granted sites | Secure, explicit per-site access |
| `Files.ReadWrite.All` | Application | Read/write all files | Required for document operations |
| `Sites.Read.All` | Application | Read-only sites | If only reading documents |

**Application vs Delegated**:
- **Application**: App acts on its own behalf (our use case)
- **Delegated**: App acts on behalf of signed-in user (not applicable)

## Support Resources

- **Azure AD Documentation**: [docs.microsoft.com/azure/active-directory](https://docs.microsoft.com/en-us/azure/active-directory/)
- **Microsoft Graph Permissions**: [docs.microsoft.com/graph/permissions-reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- **PnP PowerShell**: [pnp.github.io/powershell](https://pnp.github.io/powershell/)

---

**Setup Time**: ~15 minutes for Option A, ~30 minutes for Option B
**Last Updated**: 2025-10-27
