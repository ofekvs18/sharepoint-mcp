# SharePoint MCP Server

A Model Context Protocol (MCP) server that enables AI assistants like Claude to interact with SharePoint, providing OAuth authentication, file search, and folder structure exploration capabilities.

## Features

- 🔐 **OAuth 2.0 Authentication**: Secure browser-based login with Microsoft
- 🔍 **File Search**: Search files by name, content, or both
- 📁 **Folder Structure**: Explore SharePoint folder hierarchies
- 📄 **File Content**: Retrieve file contents
- ⏰ **Recent Files**: List recently modified files
- 🛡️ **Permission-Based**: Only accesses sites the user grants permission to
- 🚀 **Quick Testing**: No Azure AD setup required for local testing

## Prerequisites

- Node.js 18 or higher
- A Microsoft 365 account with SharePoint access
- An Azure AD application registration (for production use)

## Installation

```bash
npm install
```

## Quick Start: Local Testing (No Azure Setup Required)

Perfect for testing immediately without Azure AD configuration.

### 1. Configure Claude Desktop

Edit your Claude Desktop configuration file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

Add this configuration (replace with your actual absolute path):

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/absolute/path/to/sharepoint-mcp/index.js"]
    }
  }
}
```

**Important**: Use the full absolute path to `index.js`!

### 2. Restart Claude Desktop

Completely quit and restart Claude Desktop for the configuration to take effect.

### 3. Authenticate

In Claude Desktop, simply say:

```
Authenticate with SharePoint
```

No Client ID or Tenant ID needed! Claude will:
1. Open your browser to Microsoft login
2. Ask you to sign in with your Microsoft 365 account
3. Store the access token

**How it works**: Uses Microsoft Graph Explorer's public client ID (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) with tenant ID `common`, which works for any Microsoft 365 account.

### 4. Set Your SharePoint Site

```
Set my SharePoint site to: https://yourtenant.sharepoint.com/sites/yoursite
```

### 5. Start Using It!

```
Search for files containing "quarterly report"
Show me the folder structure
What are the most recent files?
```

## Production Setup: With Your Own Azure AD App

For production use, custom permissions, or enterprise deployments, set up your own Azure AD application.

### Step 1: Register Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Fill in the details:
   - **Name**: SharePoint MCP
   - **Redirect URI**: Web → `http://localhost:3000/callback`
4. Click **Register**

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `Sites.Read.All`
   - `Files.Read.All`
   - `offline_access`
4. Click **Grant admin consent**

### Step 3: Get Your Credentials

From your app registration overview page, copy:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 4: Authenticate with Your Credentials

In Claude Desktop:

```
Authenticate with SharePoint using:
- Client ID: [your-client-id]
- Tenant ID: [your-tenant-id]
```

## Usage Examples

### Search for Files

```
"Find all files containing 'budget' in my SharePoint"
"Search for Excel files modified this week"
"What documents mention the Marketing campaign?"
```

### Explore Structure

```
"Show me the folder structure of my Documents library"
"What folders exist under /Projects?"
"List all folders in the Marketing site"
```

### Get File Information

```
"What are the 10 most recently modified files?"
"When was the budget spreadsheet last updated?"
"List recent files"
```

### Read Content

```
"Read the content of project-plan.txt"
"Get the content of meeting-notes.txt"
```

## Available Tools

### `authenticate_sharepoint`

Authenticates with SharePoint using OAuth 2.0.

**Parameters:**
- `clientId` (optional): Azure AD Application (client) ID. If omitted, uses public client ID for testing.
- `tenantId` (optional): Azure AD Tenant ID. If omitted, uses `common` for multi-tenant.

### `set_site_url`

Sets the SharePoint site URL.

**Parameters:**
- `siteUrl` (required): Full SharePoint site URL (e.g., `https://yourtenant.sharepoint.com/sites/yoursite`)

### `search_files`

Searches for files by filename or content.

**Parameters:**
- `query` (required): Search query string
- `searchType` (optional): "filename", "content", or "both" (default: "both")
- `maxResults` (optional): Maximum results to return (default: 20)

### `get_folder_structure`

Retrieves folder structure from SharePoint.

**Parameters:**
- `folderPath` (optional): Relative folder path to start from (default: root)
- `depth` (optional): Traversal depth 1-5 (default: 2)

### `get_file_content`

Retrieves the content of a specific file.

**Parameters:**
- `fileId` (required): File ID from search results
- `filePath` (optional): Full file path (alternative to fileId)

### `list_recent_files`

Lists recently modified files.

**Parameters:**
- `limit` (optional): Number of files to return (default: 10, max: 50)

## Testing Without Claude Desktop

You can test the MCP server independently:

```bash
npm test
```

This opens an interactive menu to test all server functionality.

## Troubleshooting

### "MCP server not found"

- Verify the path in your Claude Desktop config is absolute
- Check that Node.js is installed: `node --version`
- Completely restart Claude Desktop (not just close the window)
- Try running the server manually: `node /path/to/index.js`

### "Authentication failed"

- **Local testing**: Make sure you're using a Microsoft 365 account (not personal Microsoft account)
- **Production**: Verify Client ID and Tenant ID are correct
- Check that redirect URI is exactly: `http://localhost:3000/callback`
- Ensure API permissions are granted and admin consent is clicked

### "Permission denied"

- Click "Grant admin consent" in Azure AD
- Verify your account has SharePoint access
- Check that you've added all required permissions: Sites.Read.All, Files.Read.All, offline_access

### "Port 3000 already in use"

- Close any applications using port 3000
- Or modify `redirectUri` in the code to use a different port

### "Token expired"

- Tokens expire after 1 hour
- Re-authenticate by saying "Authenticate with SharePoint" in Claude

### Browser doesn't open

- Check Claude Desktop MCP logs (Help → View Logs → MCP)
- Copy the URL from logs and open manually in browser
- Check Windows Firewall/macOS Firewall isn't blocking Node.js

## Security

- **OAuth 2.0**: Uses standard Microsoft OAuth 2.0 authentication flow
- **Token Storage**: Access tokens stored in memory only (not persisted to disk)
- **Read-Only**: Only requests read permissions (Sites.Read.All, Files.Read.All)
- **User Consent**: Users must explicitly grant access through Microsoft login
- **Permission Respect**: Respects SharePoint site permissions
- **Token Expiry**: Tokens expire after 1 hour for security

### Security Notes for Production

The current implementation is suitable for development and testing. For production use:

- ⚠️ Tokens are stored in memory and expire after 1 hour
- ⚠️ No refresh token implementation yet (must re-authenticate)
- ⚠️ Local callback server runs on localhost only
- ✅ Read-only permissions by design
- ✅ No data is logged or persisted

## Limitations

- **Token Expiry**: Access tokens expire after 1 hour, requiring re-authentication
- **Read-Only**: Cannot write, edit, or delete files (by design for safety)
- **API Limits**: Subject to Microsoft Graph API rate limits
- **File Size**: Large files (>100MB) may be slow to retrieve
- **Single Session**: One active session at a time

## Local vs Production Setup Comparison

| Feature | Local Testing | Production (Azure AD) |
|---------|---------------|----------------------|
| Azure AD Setup | ❌ Not needed | ✅ Required |
| Client ID | Uses Microsoft's public ID | Your custom app ID |
| Tenant Control | Any Microsoft 365 account | Your tenant only |
| Setup Time | 5 minutes | 15 minutes |
| Best For | Testing, development | Production, enterprise |
| Refresh Tokens | ❌ No | ❌ Not yet (future) |

## Contributing

Contributions are welcome! Please submit issues or pull requests.

## License

MIT

## Support

For issues, questions, or contributions, please open an issue on GitHub.
