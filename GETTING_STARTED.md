# GETTING STARTED

## You're almost there! Here's what to do next:

### Step 1: Install Dependencies (30 seconds)
```bash
cd "C:\Users\ofek\Downloads\gitRepos\sharepoint-mcp"
npm install
```

### Step 2: Register Azure AD App (3 minutes)
1. Open https://portal.azure.com
2. Go to: Azure Active Directory → App registrations → New registration
3. App Name: "SharePoint MCP"  
4. Redirect URI: Web → `http://localhost:3000/callback`
5. Click "Register"
6. Add API Permissions:
   - Microsoft Graph → Delegated permissions
   - Add: Sites.Read.All, Files.Read.All, offline_access
   - Click "Grant admin consent"
7. **Copy these values** (you'll need them):
   - Application (client) ID: `___________________________`
   - Directory (tenant) ID: `___________________________`

### Step 3: Configure Claude Desktop (1 minute)

Edit this file (create if it doesn't exist):
- Windows: `C:\Users\ofek\AppData\Roaming\Claude\claude_desktop_config.json`

Paste this content (update the path if needed):
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": [
        "C:\\Users\\ofek\\Downloads\\gitRepos\\sharepoint-mcp\\index.js"
      ]
    }
  }
}
```

### Step 4: Restart Claude (10 seconds)

1. Completely quit Claude Desktop (not just close the window)
2. Re-open Claude Desktop
3. Wait for it to fully load

### Step 5: Test It! (1 minute)

In Claude, type:
```
Authenticate with SharePoint using:
- Client ID: [paste the client ID you copied]
- Tenant ID: [paste the tenant ID you copied]
```

A browser will open - sign in with your Microsoft account.

Then:
```
Set my SharePoint site to: https://[yourtenant].sharepoint.com/sites/[yoursite]
```

Finally:
```
Search for files in my SharePoint
```

## 🎉 Success!

If you see search results, you're all set!

## 📚 What's Next?

- Read **START_HERE.md** for detailed documentation
- Read **SETUP.md** for more setup details  
- Run `npm test` to test the server independently
- Ask Claude: "What can you do with my SharePoint?"

## ⚠️ Troubleshooting

**Problem: "MCP server not found"**
- Make sure you completely restarted Claude Desktop
- Verify the path in the config file is correct
- Check Node.js is installed: `node --version`

**Problem: "Authentication failed"**
- Double-check you copied the correct Client ID and Tenant ID
- Verify redirect URI is exactly: `http://localhost:3000/callback`
- Ensure you clicked "Grant admin consent" in Azure

**Problem: Browser doesn't open**
- Port 3000 might be in use
- Try closing other applications
- Check Windows Firewall isn't blocking Node.js

## 🆘 Need Help?

1. Check **SETUP.md** for detailed instructions
2. Check **START_HERE.md** for navigation
3. Ask Claude: "I'm having trouble with SharePoint MCP setup"

---

**Total setup time: ~5 minutes** 🚀
