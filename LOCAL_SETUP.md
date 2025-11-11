# 🚀 Local Setup - No Azure AD Required!

This guide shows you how to run the SharePoint MCP server locally for testing **without setting up an Azure AD app registration**. Perfect for quick testing and development!

## ✨ What's New?

The server now uses Microsoft's public Graph Explorer client ID by default, so you can:
- ✅ Test immediately without Azure AD setup
- ✅ Login with your Microsoft 365 account
- ✅ Access any SharePoint site you have permissions to
- ✅ No configuration files needed

## 📋 Prerequisites

- Node.js 18 or higher
- A Microsoft 365 account with SharePoint access
- 5 minutes

## ⚡ Quick Start

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure Claude Desktop

Edit your Claude Desktop config file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

Add this configuration (replace with your actual path):

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/FULL/PATH/TO/sharepoint-mcp/index.js"]
    }
  }
}
```

**Important**: Use the absolute path! For example:
- macOS: `/Users/yourname/projects/sharepoint-mcp/index.js`
- Windows: `C:\\Users\\yourname\\projects\\sharepoint-mcp\\index.js`

### 3. Restart Claude Desktop

Completely quit and restart Claude Desktop for the config to take effect.

### 4. Authenticate in Claude

In Claude Desktop, simply say:

```
Authenticate with SharePoint
```

That's it! No Client ID or Tenant ID needed. Claude will:
1. Call the `authenticate_sharepoint` tool (with no parameters!)
2. Open your browser to Microsoft login
3. Ask you to login with your Microsoft 365 account
4. Store the access token

### 5. Set Your SharePoint Site

After authentication, tell Claude your SharePoint site URL:

```
Set my SharePoint site to: https://yourtenant.sharepoint.com/sites/yoursite
```

### 6. Start Using It!

Now you can ask Claude:
- "Search for files containing 'budget'"
- "Show me the folder structure"
- "What are the most recent files?"
- "Read the content of [filename]"

## 🎯 Example Session

```
You: Authenticate with SharePoint

Claude: [Opens browser, you login]

Claude: Successfully authenticated with SharePoint using default public client ID! 🎉

You: Set my SharePoint site to https://contoso.sharepoint.com/sites/myteam

Claude: SharePoint site URL set to: https://contoso.sharepoint.com/sites/myteam

You: Search for files containing "quarterly report"

Claude: [Returns search results...]
```

## 🔧 How It Works

1. **Default Client ID**: Uses Microsoft Graph Explorer's public client ID (`14d82eec-204b-4c2f-b7e8-296a70dab67e`)
2. **Multi-Tenant**: Uses tenant ID `common`, which works for any Microsoft 365 account
3. **OAuth 2.0 Flow**: Opens browser, you login, redirects back to local server
4. **Token Storage**: Access token stored in memory (expires after 1 hour)

## 🆚 Difference from Full Setup

| Feature | Local Setup (This Guide) | Full Setup (Azure AD) |
|---------|-------------------------|---------------------|
| Azure AD Setup | ❌ Not needed | ✅ Required |
| Client ID | Uses Microsoft's public ID | Your custom app ID |
| Tenant Control | Any Microsoft 365 account | Your tenant only |
| Setup Time | 5 minutes | 15+ minutes |
| Best For | Testing, development | Production, enterprise |

## 🔐 Security Notes

**This local setup is for testing only!**

- ✅ Uses official Microsoft public client ID
- ✅ Standard OAuth 2.0 flow
- ✅ Token stored in memory only
- ⚠️ Token expires after 1 hour (no refresh yet)
- ⚠️ Not suitable for production use
- ⚠️ You must login again after token expires

For production use, set up your own Azure AD app (see [README.md](./README.md)).

## 🐛 Troubleshooting

### "MCP server not found"
- Check that the path in your Claude config is absolute
- Try running `node /path/to/sharepoint-mcp/index.js` manually to test

### "Port 3000 already in use"
- Close any applications using port 3000
- Or modify `redirectUri` in the code to use a different port

### "Authentication failed"
- Make sure you're logging in with a Microsoft 365 account (not personal Microsoft account)
- Check that your account has SharePoint access
- Try clearing browser cookies and trying again

### "Token expired"
- Just authenticate again by saying "Authenticate with SharePoint"
- Tokens expire after 1 hour currently

### Browser doesn't open
- Check the Claude Desktop MCP logs (Help → View Logs → MCP)
- Copy the URL from logs and open manually
- Make sure the `open` npm package is installed

## 📊 What Can You Do?

Once authenticated, you have access to these tools:

1. **authenticate_sharepoint** - Login (no parameters needed!)
2. **set_site_url** - Specify your SharePoint site
3. **search_files** - Search by filename or content
4. **get_folder_structure** - Browse folders
5. **get_file_content** - Read file contents
6. **list_recent_files** - See recently modified files

## 🚀 Advanced: Using Custom Azure AD App

If you want more control (custom permissions, refresh tokens, etc.), you can still pass your own Client ID:

```
Authenticate with SharePoint using:
- Client ID: your-client-id-here
- Tenant ID: your-tenant-id-here
```

See [README.md](./README.md) for full Azure AD setup instructions.

## 📝 Notes

- The public client ID is officially provided by Microsoft for Graph Explorer
- This is the same ID used by Microsoft's own Graph Explorer tool
- It's safe and intended for development/testing scenarios
- All authentication happens through official Microsoft OAuth endpoints

## ✅ Next Steps

After setup:
1. Try the example queries above
2. Explore your SharePoint content through Claude
3. Check [README.md](./README.md) for all available tools
4. See [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) if you hit issues

## 💡 Tips

- Keep Claude Desktop open while authenticating
- Check MCP logs if something doesn't work
- The browser will auto-close after successful login
- You only need to authenticate once per session (until token expires)

---

**Ready to try it?** Follow the Quick Start steps above and start searching SharePoint in 5 minutes!
