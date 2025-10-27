# Railway Deployment Guide

This guide covers deploying the SharePoint MCP server to Railway for use with Claude.ai web interface.

## Prerequisites

1. **Railway Account**: Sign up at [railway.app](https://railway.app)
2. **GitHub Account**: Repository must be on GitHub
3. **Azure AD Application**: See [AZURE_SETUP.md](AZURE_SETUP.md) for setup instructions
4. **Credit Card**: Railway requires payment info (~$5/month for hobby plan)

## Deployment Steps

### 1. Fork/Push Repository to GitHub

This repository is already set up at: `https://github.com/MagicTurtle-s/sharepoint-mcp-railway`

### 2. Create Railway Project

```bash
# Install Railway CLI (optional)
npm install -g @railway/cli

# Login to Railway
railway login

# Link to your GitHub repository
# Or use Railway dashboard: New Project → Deploy from GitHub
```

**Via Railway Dashboard:**
1. Go to [railway.app/new](https://railway.app/new)
2. Click "Deploy from GitHub repo"
3. Select `MagicTurtle-s/sharepoint-mcp-railway`
4. Click "Deploy Now"

### 3. Configure Environment Variables

In Railway dashboard, go to your project → Variables tab:

**Required Variables:**
```
AZURE_CLIENT_ID=your-azure-app-client-id
AZURE_CLIENT_SECRET=your-azure-app-client-secret
AZURE_TENANT_ID=your-microsoft-tenant-id
```

**Optional Variables:**
```
# Default document library (defaults to "Shared Documents")
SHP_DOC_LIBRARY=Shared Documents

# Customer site mappings (comma-separated)
# Format: CustomerName:https://site-url,AnotherName:https://site-url
CUSTOMER_SITES=Acme Corp:https://tenant.sharepoint.com/sites/acme,Wayne Ent:https://tenant.sharepoint.com/sites/wayne

# Performance tuning (optional)
SHP_MAX_DEPTH=15
SHP_MAX_FOLDERS_PER_LEVEL=100
SHP_LEVEL_DELAY=0.5
```

### 4. Deploy

Railway will automatically:
1. Detect Python project
2. Install dependencies from `requirements.txt`
3. Run the command from `Procfile`: `uvicorn mcp_sharepoint.server_http:app --host 0.0.0.0 --port $PORT`
4. Assign a public URL

**Deployment typically takes 2-3 minutes.**

### 5. Verify Deployment

Once deployed, Railway will provide a URL like:
```
https://sharepoint-mcp-production.up.railway.app
```

**Test the endpoints:**

```bash
# Health check
curl https://your-project.up.railway.app/health

# Root info
curl https://your-project.up.railway.app/

# SSE endpoint (for Claude.ai)
https://your-project.up.railway.app/sse
```

Expected health check response:
```json
{
  "status": "healthy",
  "server": "SharePoint MCP Railway",
  "version": "0.1.6",
  "transport": "SSE/HTTP",
  "multi_site": true
}
```

### 6. Configure in Claude.ai

1. Go to Claude.ai → Settings → Developer → MCP Servers
2. Click "Add Server"
3. Select "URL" connection type
4. Enter your Railway SSE URL:
   ```
   https://your-project.up.railway.app/sse
   ```
5. Save and restart Claude

## Using the MCP Server

### Example 1: Discover Sites

```
User: "List all SharePoint sites"

Claude calls: List_SharePoint_Sites()
Returns: [{name: "Acme Corp Site", site_url: "https://...", ...}, ...]
```

### Example 2: Read a Document

```
User: "Get the proposal from Acme Corp site"

Claude calls: List_SharePoint_Sites(search="Acme")
Returns: site_url

Claude calls: List_SharePoint_Documents(
  site_url="https://tenant.sharepoint.com/sites/acme-corp",
  folder_name="Proposals"
)
Returns: List of documents

Claude calls: Get_Document_Content(
  site_url="https://tenant.sharepoint.com/sites/acme-corp",
  folder_name="Proposals",
  file_name="Acme_Proposal_2024.docx"
)
Returns: Full text content extracted from Word document
```

### Example 3: Upload a Document

```
User: "Upload this proposal to Wayne Enterprises site"

Claude calls: Upload_Document(
  site_url="https://tenant.sharepoint.com/sites/wayne-ent",
  folder_name="Proposals",
  file_name="New_Proposal.docx",
  content="[document content]",
  is_base64=false
)
```

## Monitoring

### Railway Dashboard

- **Logs**: View real-time logs in Railway dashboard
- **Metrics**: CPU, memory, network usage
- **Deployments**: History of all deployments

### Check Server Status

```bash
# View logs
railway logs

# Check deployment status
railway status
```

### Health Monitoring

Set up monitoring with Railway's health checks:
- **Endpoint**: `/health`
- **Expected**: 200 OK
- **Interval**: 60 seconds

## Troubleshooting

### Deployment Fails

**Check build logs in Railway dashboard:**
- Python version issues → Verify `runtime.txt` specifies Python 3.10+
- Missing dependencies → Check `requirements.txt` is complete
- Import errors → Verify all modules in `src/mcp_sharepoint/` are valid

### Server Starts but SSE Connection Fails

1. **Verify environment variables** are set correctly in Railway
2. **Check Azure AD permissions**:
   - Application permissions granted
   - Admin consent approved
3. **Test health endpoint**: `curl https://your-url/health`
4. **Check Railway logs** for authentication errors

### "401 Unauthorized" from SharePoint

- **Cause**: Azure AD credentials incorrect or expired
- **Fix**:
  1. Regenerate client secret in Azure Portal
  2. Update `AZURE_CLIENT_SECRET` in Railway
  3. Redeploy

### "403 Forbidden" from SharePoint

- **Cause**: Insufficient permissions
- **Fix**:
  1. Verify `Sites.ReadWrite.All` permission granted
  2. Ensure admin consent is approved
  3. Check site-specific permissions if using `Sites.Selected`

### High Memory Usage

- **Cause**: Processing large files or many concurrent requests
- **Fix**:
  1. Upgrade Railway plan (more RAM)
  2. Implement file size limits
  3. Add request throttling

## Cost Estimation

**Railway Hobby Plan (~$5/month):**
- 500 hours of usage
- $5 base + usage fees
- Shared CPU, 512MB RAM
- Sufficient for personal/small team use

**Railway Pro Plan (~$20/month):**
- Priority support
- More resources
- Better for production use with multiple users

## Updating the Server

Railway automatically deploys on git push:

```bash
# Make changes locally
git add .
git commit -m "feat: add new feature"
git push origin main

# Railway auto-deploys (takes 2-3 minutes)
```

## Environment Variables Reference

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `AZURE_CLIENT_ID` | Yes | - | Azure AD app client ID |
| `AZURE_CLIENT_SECRET` | Yes | - | Azure AD app secret |
| `AZURE_TENANT_ID` | Yes | - | Microsoft tenant ID |
| `SHP_DOC_LIBRARY` | No | "Shared Documents" | Default document library |
| `CUSTOMER_SITES` | No | - | Customer name → site URL mappings |
| `SHP_MAX_DEPTH` | No | 15 | Max folder tree depth |
| `SHP_MAX_FOLDERS_PER_LEVEL` | No | 100 | Max folders per level |
| `SHP_LEVEL_DELAY` | No | 0.5 | Delay between levels (seconds) |
| `PORT` | Auto | 8000 | Railway sets automatically |

## Security Best Practices

1. **Never commit secrets to Git**
   - Use Railway environment variables
   - Keep `.env` in `.gitignore`

2. **Rotate secrets regularly**
   - Azure AD client secrets expire (max 2 years)
   - Set calendar reminder to rotate

3. **Use Sites.Selected permission**
   - Instead of `Sites.ReadWrite.All`
   - Explicitly grant access only to needed sites
   - See AZURE_SETUP.md for instructions

4. **Monitor access logs**
   - Review Railway logs regularly
   - Watch for unusual access patterns

5. **Implement rate limiting**
   - Consider adding request throttling for production
   - Protect against abuse

## Support

- **Railway Issues**: [railway.app/help](https://railway.app/help)
- **Project Issues**: [GitHub Issues](https://github.com/MagicTurtle-s/sharepoint-mcp-railway/issues)
- **Azure AD Help**: [Azure Portal Support](https://portal.azure.com)

## Next Steps

1. ✅ Deploy to Railway
2. ✅ Verify health endpoint
3. ✅ Configure in Claude.ai
4. ✅ Test with a simple query
5. → Add customer site mappings
6. → Monitor usage and costs
7. → Scale as needed

---

**Deployed**: Ready for Railway deployment
**Last Updated**: 2025-10-27
