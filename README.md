# SharePoint MCP Server

A Model Context Protocol (MCP) server that enables AI assistants like Claude to interact with SharePoint, providing OAuth authentication, file search, and folder structure exploration capabilities.

## Features

- 🔐 **OAuth 2.0 Authentication**: Secure browser-based login with Microsoft
- 🔍 **Comprehensive File Search**:
  - **Filename search** (fast, OneDrive API)
  - **Content search** (deep, searches inside files)
  - **Auto mode** (tries Graph API, falls back to content search)
  - Supports 40+ file types for content search (code, text, config files)
- 📁 **Folder Structure**: Explore OneDrive folder hierarchies
- 📄 **File Content**: Retrieve file contents
- ⏰ **Recent Files**: List recently modified files
- 🔄 **Shared Files**: Search and list files shared with you
- 🛡️ **Permission-Based**: Only accesses files the user grants permission to

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

**Quick filename search (default):**
```
"Search for files named 'quarterly report'"
```

**Deep content search (ALL file types including Office docs):**
```
"Search for files containing 'quarterly report' in their content with searchDepth='content'"
"Search for 'budget' in Word documents with searchDepth='content' and fileTypes=['docx']"
```
*Searches inside all files including Office documents (.docx, .xlsx, .pptx, .pdf) using text extraction.*

**Smart search (Graph API for enterprise):**
```
"Search for 'API documentation' using searchDepth='auto'"
```
*Uses Microsoft Graph Search API (best for work/enterprise accounts with advanced indexing).*

**Search specific file types:**
```
"Search for 'function' in JavaScript files only with fileTypes=['js', 'jsx']"
```

**Include shared files:**
```
"Search for 'budget' with includeShared=true"
```

**Other operations:**
```
"Show me the folder structure"
"What are the 10 most recently modified files?"
"List files shared with me"
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

### `search_my_files`
Searches for files in your OneDrive by filename or content with multiple search strategies.

**Parameters:**
- `query` (required): Search query string
- `searchDepth` (optional): Search strategy
  - `"filename"` (default): Fast filename-only search using OneDrive API
  - `"content"`: **Comprehensive search with text extraction** - searches inside ALL files including Office documents
  - `"auto"`: Uses Microsoft Graph Search API (best for enterprise accounts)
- `maxResults` (optional): Maximum results to return (default: 20)
- `includeShared` (optional): Include files shared with you (default: false)
- `fileTypes` (optional): Array of file extensions to search (e.g., `['js', 'md', 'docx', 'pdf']`)

**Examples:**
- Quick filename search: `searchDepth: "filename"`
- **Search Office documents**: `searchDepth: "content"` with `fileTypes: ['docx']` ✅
- Search code files: `searchDepth: "content"` with `fileTypes: ['js', 'py']`
- Enterprise search: `searchDepth: "auto"` (uses Graph Search API)

**How it works:**
- **Plain text files** (.txt, .md, .js, etc.): Direct content download and search
- **Office documents** (.docx, .xlsx, .pptx, .pdf): Automatic text extraction via Graph API, then search
- **All file types** supported: 40+ plain text formats + Word, Excel, PowerPoint, PDF

**Supported file types:**
- Plain text: js, py, java, txt, md, log, json, yml, ini, csv, sql, html, css, etc.
- Office docs: docx, doc, xlsx, xls, pptx, ppt, pdf

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
