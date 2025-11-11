# SharePoint MCP Server

A Model Context Protocol (MCP) server that enables AI assistants like Claude to interact with SharePoint, providing OAuth authentication, file search, and folder structure exploration capabilities.

## Features

- 🔐 **OAuth 2.0 Authentication**: Secure browser-based login with Microsoft
- 🔍 **File Search**: Search files by name, content, or both
- 📁 **Folder Structure**: Explore SharePoint folder hierarchies
- 📄 **File Content**: Retrieve file contents
- ⏰ **Recent Files**: List recently modified files
- 🛡️ **Permission-Based**: Only accesses sites the user grants permission to

## Prerequisites

- Node.js 18 or higher
- A Microsoft 365 account with SharePoint access
- An Azure AD application registration (see setup below)

## Quick Start

### 🚀 NEW: Local Testing (No Azure Setup!)

Want to test immediately without Azure AD setup? See **[LOCAL_SETUP.md](./LOCAL_SETUP.md)**

**TL;DR for local testing:**
1. `npm install`
2. Configure Claude Desktop config file
3. In Claude: "Authenticate with SharePoint" (no Client ID needed!)
4. Set your site URL and start searching

### 📚 Full Production Setup

See **[START_HERE.md](./START_HERE.md)** for complete setup instructions with your own Azure AD app.

**TL;DR for production:**
1. Register an Azure AD app and get Client ID + Tenant ID
2. `npm install`
3. Configure Claude Desktop config file
4. Restart Claude
5. Authenticate through Claude with your credentials

Full instructions in [START_HERE.md](./START_HERE.md)

## Azure AD Application Setup

### Step 1: Register Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Fill in the details:
   - **Name**: SharePoint MCP
   - **Redirect URI**: Web → `http://localhost:3000/callback`
4. Click **Register**

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**
3. Add: `Sites.Read.All`, `Files.Read.All`, `offline_access`
4. Click **Grant admin consent**

### Step 3: Get Your Credentials

From your app registration overview page, note:
- **Application (client) ID**
- **Directory (tenant) ID**

## Installation

```bash
cd sharepoint-mcp
npm install
```

## Configuration for Claude Desktop

Edit your Claude Desktop configuration file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
**Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

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

**Important**: Use the FULL absolute path!

## Usage

### 1. Authenticate

```
In Claude: "Authenticate with SharePoint using:
- Client ID: [your-client-id]
- Tenant ID: [your-tenant-id]"
```

### 2. Set Site URL

```
In Claude: "Set my SharePoint site to: https://yourtenant.sharepoint.com/sites/yoursite"
```

### 3. Search and Explore

```
"Search for files containing 'quarterly report'"
"Show me the folder structure"
"What are the 10 most recently modified files?"
```

## Available Tools

### `authenticate_sharepoint`
Authenticates with SharePoint using OAuth 2.0.

**Parameters:**
- `clientId` (required): Azure AD Application (client) ID
- `tenantId` (required): Azure AD Tenant ID

### `set_site_url`
Sets the SharePoint site URL.

**Parameters:**
- `siteUrl` (required): Full SharePoint site URL

### `search_files`
Searches for files by filename or content.

**Parameters:**
- `query` (required): Search query string
- `searchType` (optional): "filename", "content", or "both"
- `maxResults` (optional): Maximum results (default: 20)

### `get_folder_structure`
Retrieves folder structure.

**Parameters:**
- `folderPath` (optional): Relative folder path
- `depth` (optional): Traversal depth 1-5 (default: 2)

### `get_file_content`
Retrieves file content.

**Parameters:**
- `fileId` (required): File ID from search results

### `list_recent_files`
Lists recently modified files.

**Parameters:**
- `limit` (optional): Number of files (default: 10)

## Troubleshooting

See [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) for detailed solutions.

**Common issues:**
- "MCP server not found" → Check config path is absolute
- "Authentication failed" → Verify Client ID and Tenant ID
- "Permission denied" → Grant admin consent in Azure

## Security

- Access tokens stored in memory only (expire after 1 hour)
- Read-only permissions
- Users must explicitly grant access
- Respects SharePoint permissions

See [SECURITY.md](./SECURITY.md) for complete security information.

## Limitations

- Tokens expire after 1 hour (no auto-refresh yet)
- Read-only (no write operations)
- Subject to Microsoft Graph API limits
- Large files may be slow to retrieve

## License

MIT

## Documentation

- **[START_HERE.md](./START_HERE.md)** - Entry point and navigation
- **[QUICKSTART.md](./QUICKSTART.md)** - 5-minute setup guide
- **[TROUBLESHOOTING.md](./TROUBLESHOOTING.md)** - Common issues
- **[SECURITY.md](./SECURITY.md)** - Security considerations
- **[ARCHITECTURE.md](./ARCHITECTURE.md)** - System design

## Contributing

Contributions welcome! Please submit issues or pull requests.
