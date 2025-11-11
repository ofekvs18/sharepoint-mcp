# 🚀 START HERE - SharePoint MCP Setup

Welcome! This is your complete SharePoint MCP (Model Context Protocol) server that lets Claude interact with your SharePoint files and folders.

## 📋 What You Need

Before starting, make sure you have:
- ✅ Node.js 18 or higher installed
- ✅ A Microsoft 365 account with SharePoint access
- ✅ 15 minutes for setup

## 🎯 Quick Decision Guide

**Choose your path:**

### ⚡ I just want to test it NOW (5 minutes, no Azure setup)
**➡️ Read: [LOCAL_SETUP.md](./LOCAL_SETUP.md)** 🆕

Test immediately using a public client ID. No Azure AD setup needed! Perfect for quick testing.

### 👤 I want to use it properly (10-15 minutes)
**➡️ Read: [QUICKSTART.md](./QUICKSTART.md)**

Full setup with your own Azure AD app. Better for regular use and production.

### 📚 I want to understand everything
**➡️ Read in order:**
1. [PROJECT_SUMMARY.md](./PROJECT_SUMMARY.md) - What this is and what it does
2. [ARCHITECTURE.md](./ARCHITECTURE.md) - How it works under the hood
3. [README.md](./README.md) - Complete documentation
4. [SECURITY.md](./SECURITY.md) - Security model and best practices

### 🔧 I'm having problems
**➡️ Read: [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)**

Step-by-step solutions for common issues.

### 🏢 I want to deploy this for my team/org
**➡️ Read in order:**
1. [SECURITY.md](./SECURITY.md) - **READ THIS FIRST**
2. [ARCHITECTURE.md](./ARCHITECTURE.md) - System design
3. [README.md](./README.md) - Full documentation

⚠️ **Important:** Current version is NOT production-ready. See SECURITY.md for requirements.

## ⚡ Super Quick Setup (TL;DR)

### Option A: Test NOW (No Azure setup) 🆕
```bash
# 1. Install dependencies
npm install

# 2. Configure Claude Desktop config
# Edit: ~/Library/Application Support/Claude/claude_desktop_config.json (macOS)
# Or: %APPDATA%/Claude/claude_desktop_config.json (Windows)

# 3. Restart Claude Desktop

# 4. In Claude, say:
"Authenticate with SharePoint"
# No Client ID needed! Browser opens, you login, done.
```

Full details in [LOCAL_SETUP.md](./LOCAL_SETUP.md)

### Option B: Full Setup (With your Azure AD app)
```bash
# 1. Install dependencies
npm install

# 2. Setup Azure AD
# Go to portal.azure.com → Azure AD → App registrations
# Create app, add permissions: Sites.Read.All, Files.Read.All, offline_access
# Copy Client ID and Tenant ID

# 3. Configure Claude Desktop
# Same as above

# 4. Restart Claude Desktop

# 5. In Claude, say:
"Authenticate with SharePoint using Client ID: xxx and Tenant ID: yyy"
```

Full details in [QUICKSTART.md](./QUICKSTART.md)

## 📁 File Guide

### 🔴 Critical Files (You Need These)
- **index.js** - The main MCP server (don't modify unless you know what you're doing)
- **package.json** - Dependencies (run `npm install`)

### 🟢 Start Here
- **START_HERE.md** - This file, your entry point
- **QUICKSTART.md** - Fastest setup path
- **PROJECT_SUMMARY.md** - Overview of everything

### 🔵 Documentation
- **README.md** - Complete documentation
- **ARCHITECTURE.md** - How it works
- **SECURITY.md** - Security considerations
- **TROUBLESHOOTING.md** - Problem solving

### 🟡 Utilities
- **test.js** - Test the server without Claude Desktop
- **claude_desktop_config.example.json** - Configuration template
- **.env.example** - Environment variables reference

### 🟤 System Files
- **.gitignore** - Git exclusions (protects secrets)

## 🎓 Learning Path

### Beginner (Just want to use it)
```
START_HERE.md → QUICKSTART.md → Use it! → TROUBLESHOOTING.md (if needed)
```

### Intermediate (Want to understand)
```
START_HERE.md → PROJECT_SUMMARY.md → README.md → Try it → Explore ARCHITECTURE.md
```

### Advanced (Want to modify/deploy)
```
All docs → SECURITY.md (carefully) → ARCHITECTURE.md → index.js source code
```

## 💡 What Can I Do With This?

Once set up, you can ask Claude:

**Search for files:**
- "Find all files containing 'quarterly report' in my SharePoint"
- "Search for Excel files modified this week"
- "What documents mention the Marketing campaign?"

**Explore structure:**
- "Show me the folder structure of my Documents library"
- "What folders exist under /Projects?"
- "List all folders in the Marketing site"

**Get file info:**
- "What are the 10 most recently modified files?"
- "When was the budget spreadsheet last updated?"
- "Who created the meeting notes document?"

**Read content:**
- "Read the content of project-plan.txt"
- "What's in the latest status report?"

## ⚠️ Important Limitations

Current version:
- ✅ Read files and folders
- ✅ Search by name and content
- ❌ Cannot write/edit/delete files (by design - read-only for safety)
- ❌ Tokens expire after 1 hour (need to re-authenticate)
- ❌ Not production-ready (see SECURITY.md)

## 🆘 Need Help?

**Having issues?**
1. Check [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) first
2. Run `npm test` to test independently
3. Check Claude Desktop logs
4. Verify Azure AD setup

**Common first-time issues:**
- "MCP not found" → Check config file path is absolute
- "Authentication failed" → Verify Client ID and Tenant ID
- "Permission denied" → Grant admin consent in Azure
- "Site not found" → Check SharePoint URL format

## 🎯 Next Steps

1. **Right now:** Open [QUICKSTART.md](./QUICKSTART.md) and follow the steps
2. **After setup:** Try the example queries above
3. **If curious:** Read [ARCHITECTURE.md](./ARCHITECTURE.md) to understand how it works
4. **If problems:** Check [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)

## 📊 Project Stats

- **Total Lines:** 2,580+ lines of code and documentation
- **Main Server:** 641 lines (index.js)
- **Documentation:** 6 comprehensive guides
- **Setup Time:** 5-15 minutes
- **Dependencies:** 4 packages

## 🔐 Security Note

This tool:
- Uses OAuth 2.0 for authentication
- Only requests read permissions
- Stores tokens in memory only
- Respects SharePoint permissions
- Cannot modify or delete files

See [SECURITY.md](./SECURITY.md) for complete security information.

## ✨ Ready to Start?

**Open [QUICKSTART.md](./QUICKSTART.md) now and get started!**

You'll be searching SharePoint through Claude in about 5 minutes.

---

**Questions?** All answers are in the documentation files listed above.

**Problems?** Check TROUBLESHOOTING.md first.

**Want to contribute?** Read ARCHITECTURE.md to understand the codebase.
