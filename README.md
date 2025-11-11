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

**Deep content search (text files + limited Office support):**
```
"Search for files containing 'quarterly report' in their content with searchDepth='content'"
"Search for 'budget' in code files with searchDepth='content' and fileTypes=['js', 'py', 'md']"
```
*Best for plain text files (.txt, .md, .js, etc.). Office document text extraction has limited availability.*

**Smart search (recommended for Office docs):**
```
"Search for 'API documentation' using searchDepth='auto'"
"Search for 'budget' in Word documents with searchDepth='auto'"
```
*Uses Microsoft Graph Search API - best option for searching Word, Excel, PowerPoint, PDF files.*

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
  - `"content"`: Downloads and searches file contents (best for plain text files)
  - `"auto"`: **Recommended** - Uses Microsoft Graph Search API (works for all file types)
- `maxResults` (optional): Maximum results to return (default: 20)
- `includeShared` (optional): Include files shared with you (default: false)
- `fileTypes` (optional): Array of file extensions to search (e.g., `['js', 'md', 'docx', 'pdf']`)

**Examples:**
- Quick filename search: `searchDepth: "filename"`
- Search code/text files: `searchDepth: "content"` with `fileTypes: ['js', 'py', 'md']` ✅
- **Search Office documents**: `searchDepth: "auto"` ⭐ (most reliable for .docx, .xlsx, .pdf)

**How it works:**
- **Plain text files** (.txt, .md, .js, etc.): Content mode downloads and searches directly
- **Office documents** (.docx, .xlsx, .pptx, .pdf):
  - Content mode tries text extraction (limited availability)
  - **Auto mode recommended** - uses Graph Search API (more reliable)

**Supported file types:**
- **Full support** (content mode): js, py, java, txt, md, log, json, yml, ini, csv, sql, html, css, etc.
- **Limited support** (content mode): docx, xlsx, pptx, pdf (text extraction may not work for all accounts)
- **Best support** (auto mode): All file types via Graph Search API

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
