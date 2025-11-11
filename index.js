#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import axios from "axios";
import open from "open";

// Default public Azure AD client ID (Microsoft Graph Explorer)
// This allows testing without setting up your own Azure AD app
const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const DEFAULT_TENANT_ID = "common"; // Multi-tenant, works for any Microsoft 365 account

// OneDrive MCP Server
class SharePointMCP {
  constructor() {
    this.server = new Server(
      {
        name: "onedrive-mcp",
        version: "1.0.0",
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    // Store authentication tokens
    this.authTokens = {
      accessToken: null,
      refreshToken: null,
      expiresAt: null,
    };

    this.setupHandlers();
  }

  setupHandlers() {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: "authenticate_sharepoint",
          description:
            "Authenticate with SharePoint using OAuth 2.0 Device Code Flow. Shows a code to enter at microsoft.com/devicelogin. No Azure AD setup required - uses default public client ID for testing! No admin consent needed.",
          inputSchema: {
            type: "object",
            properties: {
              clientId: {
                type: "string",
                description: "Azure AD Application (client) ID (optional - uses public default if not provided)",
              },
              tenantId: {
                type: "string",
                description: "Azure AD Tenant ID (optional - uses 'common' if not provided, works for any Microsoft 365 account)",
              },
            },
            required: [],
          },
        },
        {
          name: "search_my_files",
          description:
            "Search for files in your OneDrive by filename or content. Optionally include files shared with you.",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description: "Search query string",
              },
              maxResults: {
                type: "number",
                description: "Maximum number of results to return",
                default: 20,
              },
              includeShared: {
                type: "boolean",
                description: "Include files shared with you in search results",
                default: false,
              },
              searchDepth: {
                type: "string",
                description: "Search depth: 'filename' (fast, names only), 'content' (downloads and searches file contents - best for code/text files, limited for Office docs), 'auto' (RECOMMENDED for Office docs - uses Graph Search API)",
                enum: ["filename", "content", "auto"],
                default: "filename",
              },
              fileTypes: {
                type: "array",
                description: "Filter by file extensions (e.g., ['txt', 'js', 'md', 'docx', 'pdf']). Works with all searchDepth modes.",
                items: {
                  type: "string",
                },
              },
            },
            required: ["query"],
          },
        },
        {
          name: "list_my_files",
          description:
            "List files and folders in your OneDrive (optionally in a specific folder)",
          inputSchema: {
            type: "object",
            properties: {
              folderPath: {
                type: "string",
                description: "Folder path (leave empty for root/recent files)",
                default: "",
              },
              limit: {
                type: "number",
                description: "Number of items to return",
                default: 20,
              },
            },
          },
        },
        {
          name: "get_file_content",
          description:
            "Retrieve the content of a specific file from your OneDrive or shared files",
          inputSchema: {
            type: "object",
            properties: {
              fileId: {
                type: "string",
                description: "File ID from search results",
              },
              driveId: {
                type: "string",
                description: "Drive ID (optional, required for shared files from other drives)",
              },
            },
            required: ["fileId"],
          },
        },
        {
          name: "inspect_file_metadata",
          description:
            "Get detailed metadata about a file including all IDs and paths for debugging",
          inputSchema: {
            type: "object",
            properties: {
              fileId: {
                type: "string",
                description: "File ID from search results",
              },
              driveId: {
                type: "string",
                description: "Drive ID (optional, if file is from another drive)",
              },
            },
            required: ["fileId"],
          },
        },
        {
          name: "list_recent_files",
          description:
            "List your recently accessed or modified files in OneDrive",
          inputSchema: {
            type: "object",
            properties: {
              limit: {
                type: "number",
                description: "Number of files to return",
                default: 10,
              },
            },
          },
        },
        {
          name: "list_shared_files",
          description:
            "List files that have been shared with you by others",
          inputSchema: {
            type: "object",
            properties: {
              limit: {
                type: "number",
                description: "Number of items to return",
                default: 20,
              },
            },
          },
        },
      ],
    }));

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case "authenticate_sharepoint":
            return await this.authenticateSharePoint(args);
          case "search_my_files":
            return await this.searchMyFiles(args);
          case "list_my_files":
            return await this.listMyFiles(args);
          case "get_file_content":
            return await this.getFileContent(args);
          case "inspect_file_metadata":
            return await this.inspectFileMetadata(args);
          case "list_recent_files":
            return await this.listRecentFiles(args);
          case "list_shared_files":
            return await this.listSharedFiles(args);
          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Error: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  async authenticateSharePoint(args) {
    // Handle empty strings by treating them as undefined
    const clientId = args.clientId?.trim() || DEFAULT_CLIENT_ID;
    const tenantId = args.tenantId?.trim() || DEFAULT_TENANT_ID;

    // Use Device Code Flow - no redirect URI needed!
    // This is perfect for local testing and CLI apps
    const deviceCodeUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`;
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    // Scopes that don't require admin consent
    const scopes = [
      "User.Read",
      "Files.Read",
      "offline_access",
    ].join(" ");

    try {
      // Step 1: Request device code
      const deviceCodeResponse = await axios.post(
        deviceCodeUrl,
        new URLSearchParams({
          client_id: clientId,
          scope: scopes,
        }),
        {
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
        }
      );

      const {
        device_code,
        user_code,
        verification_uri,
        expires_in,
        interval = 5,
        message,
      } = deviceCodeResponse.data;

      console.error("\n=== SharePoint Authentication ===");
      if (clientId === DEFAULT_CLIENT_ID) {
        console.error("Using default public client ID (no Azure AD setup needed!)");
      }
      console.error("\nTo sign in, use a web browser to open the page:");
      console.error(`  ${verification_uri}`);
      console.error("\nAnd enter the code:");
      console.error(`  ${user_code}`);
      console.error("\nWaiting for you to authenticate...");
      console.error("================================\n");

      // Try to open the browser automatically
      try {
        await open(verification_uri);
      } catch (err) {
        // Silent fail - user can open manually
      }

      // Step 2: Poll for token
      const pollUntil = Date.now() + expires_in * 1000;

      while (Date.now() < pollUntil) {
        await new Promise(resolve => setTimeout(resolve, interval * 1000));

        try {
          const tokenResponse = await axios.post(
            tokenUrl,
            new URLSearchParams({
              client_id: clientId,
              grant_type: "urn:ietf:params:oauth:grant-type:device_code",
              device_code: device_code,
            }),
            {
              headers: {
                "Content-Type": "application/x-www-form-urlencoded",
              },
            }
          );

          const { access_token, refresh_token, expires_in } = tokenResponse.data;

          // Store tokens
          this.authTokens.accessToken = access_token;
          this.authTokens.refreshToken = refresh_token;
          this.authTokens.expiresAt = Date.now() + expires_in * 1000;

          console.error("\n✅ Authentication successful!\n");

          const successMessage = clientId === DEFAULT_CLIENT_ID
            ? "Successfully authenticated with OneDrive using default public client ID! 🎉\n\n" +
              "Access token stored in memory.\n\n" +
              "You can now:\n" +
              "- Search your files: 'search_my_files'\n" +
              "- List files: 'list_my_files'\n" +
              "- See recent files: 'list_recent_files'"
            : "Successfully authenticated with OneDrive! Access token stored.";

          return {
            content: [
              {
                type: "text",
                text: successMessage,
              },
            ],
          };
        } catch (pollError) {
          // Check for specific errors
          const errorCode = pollError.response?.data?.error;

          if (errorCode === "authorization_pending") {
            // User hasn't completed auth yet, continue polling
            continue;
          } else if (errorCode === "slow_down") {
            // Need to slow down polling
            await new Promise(resolve => setTimeout(resolve, interval * 1000));
            continue;
          } else if (errorCode === "authorization_declined") {
            throw new Error("Authentication was declined by user");
          } else if (errorCode === "expired_token") {
            throw new Error("Authentication code expired. Please try again.");
          } else {
            // Unknown error, rethrow
            throw pollError;
          }
        }
      }

      throw new Error("Authentication timeout - please try again");
    } catch (error) {
      const errorMessage = error.response?.data?.error_description || error.message;
      throw new Error(`Authentication failed: ${errorMessage}`);
    }
  }

  async ensureAuthenticated() {
    if (!this.authTokens.accessToken) {
      throw new Error("Not authenticated. Please run 'authenticate_sharepoint' first.");
    }

    // Check if token expired
    if (Date.now() >= this.authTokens.expiresAt) {
      throw new Error("Access token expired. Please re-authenticate.");
    }
  }

  // Helper function to construct full file path from Graph API item
  constructFullPath(item) {
    if (!item.parentReference || !item.parentReference.path) {
      // If no parent reference, return just the name (root level item)
      return `/${item.name}`;
    }

    // parentReference.path format: "/drive/root:" or "/drive/root:/path/to/folder"
    // We need to remove the "/drive/root:" prefix
    let parentPath = item.parentReference.path.replace(/^\/drive\/root:?/, '');

    // Ensure we don't have double slashes
    if (!parentPath) {
      parentPath = '';
    }

    // Construct full path
    const fullPath = parentPath ? `${parentPath}/${item.name}` : `/${item.name}`;

    return fullPath;
  }

  // Helper: Check if file is Office document
  isOfficeDocument(filename) {
    const officeExtensions = ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt', 'pdf'];
    const ext = filename.split('.').pop().toLowerCase();
    return officeExtensions.includes(ext);
  }

  // Helper: Check if file is searchable based on extension
  isSearchableFile(filename, allowedTypes = null) {
    const searchableExtensions = [
      // Plain text files
      'txt', 'md', 'js', 'jsx', 'ts', 'tsx', 'json', 'xml', 'html', 'htm',
      'css', 'scss', 'sass', 'py', 'java', 'c', 'cpp', 'h', 'cs', 'php',
      'rb', 'go', 'rs', 'sh', 'bash', 'yml', 'yaml', 'toml', 'ini', 'cfg',
      'log', 'csv', 'sql', 'r', 'swift', 'kt', 'dart', 'vue', 'svelte',
      // Office documents (with text extraction)
      'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt', 'pdf'
    ];

    const ext = filename.split('.').pop().toLowerCase();

    // If specific file types are requested, check against those
    if (allowedTypes && allowedTypes.length > 0) {
      return allowedTypes.includes(ext) && searchableExtensions.includes(ext);
    }

    return searchableExtensions.includes(ext);
  }

  // Helper: Extract text from Office document
  async extractOfficeText(fileId, driveId = null, filename = '') {
    try {
      console.error(`Attempting to extract text from Office doc: ${filename}`);

      // Method 1: Try HTML conversion (works for some accounts)
      const htmlEndpoint = driveId
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content?format=html`
        : `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content?format=html`;

      try {
        console.error(`  Trying HTML conversion for ${filename}...`);
        const htmlResponse = await axios.get(htmlEndpoint, {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          responseType: 'text',
          timeout: 20000,
          validateStatus: (status) => status === 200, // Only accept 200
        });

        // Strip HTML tags to get plain text
        const text = htmlResponse.data
          .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
          .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
          .replace(/<[^>]+>/g, ' ')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .replace(/&[a-z]+;/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

        if (text && text.length > 10) {
          console.error(`  ✓ Successfully extracted ${text.length} chars from ${filename}`);
          return text;
        }
      } catch (htmlError) {
        console.error(`  ✗ HTML conversion failed for ${filename}: ${htmlError.response?.status || htmlError.message}`);
      }

      // Method 2: Try getting file metadata with description (sometimes contains indexed text)
      const metadataEndpoint = driveId
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}?$select=name,description,file`
        : `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}?$select=name,description,file`;

      try {
        console.error(`  Trying metadata extraction for ${filename}...`);
        const metadataResponse = await axios.get(metadataEndpoint, {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          timeout: 10000,
        });

        const description = metadataResponse.data.description;
        if (description && description.length > 0) {
          console.error(`  ✓ Found description text (${description.length} chars) for ${filename}`);
          return description;
        }
      } catch (metadataError) {
        console.error(`  ✗ Metadata extraction failed for ${filename}`);
      }

      // If all methods fail, return empty (file will be skipped)
      console.error(`  ✗ All extraction methods failed for ${filename}`);
      return '';

    } catch (error) {
      console.error(`  ✗ Error extracting text from ${filename}: ${error.message}`);
      return '';
    }
  }

  // Helper: Download and search file content
  async searchFileContent(fileId, query, driveId = null, filename = '') {
    try {
      let content = '';

      // Check if this is an Office document
      if (filename && this.isOfficeDocument(filename)) {
        // Use Office text extraction
        content = await this.extractOfficeText(fileId, driveId, filename);

        if (!content) {
          // Text extraction failed, skip this file silently
          return { found: false, error: 'Office text extraction not available for this file' };
        }
      } else {
        // Regular plain text file - download directly
        const endpoint = driveId
          ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`
          : `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`;

        const response = await axios.get(endpoint, {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          responseType: 'text',
          maxContentLength: 10 * 1024 * 1024, // Limit to 10MB
          timeout: 30000,
        });

        content = response.data;
      }

      const queryLower = query.toLowerCase();

      // Search for query in content (case-insensitive)
      if (content.toLowerCase().includes(queryLower)) {
        // Find context around matches
        const lines = content.split(/\r?\n/);
        const matchingLines = [];

        for (let i = 0; i < lines.length; i++) {
          if (lines[i].toLowerCase().includes(queryLower)) {
            matchingLines.push({
              lineNumber: i + 1,
              content: lines[i].trim().substring(0, 300),
            });

            // Limit to first 5 matches per file
            if (matchingLines.length >= 5) break;
          }
        }

        return {
          found: true,
          matches: matchingLines,
          preview: matchingLines[0]?.content.substring(0, 200) || '',
        };
      }

      return { found: false };
    } catch (error) {
      // If file is too large, binary, or inaccessible, skip it
      if (error.response?.status === 404) {
        return { found: false, error: 'File not found or inaccessible' };
      }
      return { found: false, error: error.message };
    }
  }

  // Helper: Try Microsoft Graph Search API (works for work accounts)
  async searchWithGraphAPI(query, maxResults) {
    try {
      const response = await axios.post(
        'https://graph.microsoft.com/v1.0/search/query',
        {
          requests: [
            {
              entityTypes: ['driveItem'],
              query: {
                queryString: query,
              },
              from: 0,
              size: maxResults,
            },
          ],
        },
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
            'Content-Type': 'application/json',
          },
        }
      );

      const hits = response.data.value[0]?.hitsContainers[0]?.hits || [];

      return {
        success: true,
        results: hits.map((hit) => {
          const resource = hit.resource;
          return {
            id: resource.id,
            name: resource.name,
            path: this.constructFullPath(resource),
            webUrl: resource.webUrl,
            size: resource.size,
            lastModified: resource.lastModifiedDateTime,
            author: resource.createdBy?.user?.displayName,
            type: resource.folder ? 'folder' : 'file',
            source: 'graphSearch',
            summary: hit.summary || '',
            relevanceScore: hit.rank,
          };
        }),
      };
    } catch (error) {
      // Graph Search API not available (personal account or insufficient permissions)
      return { success: false, error: error.message };
    }
  }

  // Helper: Check if file should be skipped (system files, caches, etc.)
  shouldSkipFile(filename, filepath) {
    const skipPatterns = [
      /^~\$/,                    // Temp files (~$)
      /^\./,                     // Hidden files
      /\.tmp$/i,                 // Temp files
      /\.bak$/i,                 // Backup files
      /\.log$/i,                 // Log files (unless searching for logs)
      /node_modules/i,           // Node modules
      /\.git\//i,                // Git directory
      /package-lock\.json$/i,    // Lock files
      /yarn\.lock$/i,
      /\.min\.js$/i,             // Minified files
      /\.min\.css$/i,
      /\.map$/i,                 // Source maps
      /\.cache/i,                // Cache files
      /thumbs\.db$/i,            // Windows thumbnail cache
      /\.DS_Store$/i,            // Mac system files
    ];

    // Check filename patterns
    for (const pattern of skipPatterns) {
      if (pattern.test(filename) || (filepath && pattern.test(filepath))) {
        return true;
      }
    }

    return false;
  }

  // Helper: Calculate relevance score for a file
  getFileRelevanceScore(file, query) {
    let score = 0;
    const fileName = file.name.toLowerCase();
    const queryLower = query.toLowerCase();
    const queryWords = queryLower.split(/\s+/).filter(w => w.length > 2);

    // Exact filename match - very high score
    if (fileName === queryLower) {
      score += 100;
    }

    // Filename contains full query
    if (fileName.includes(queryLower)) {
      score += 50;
    }

    // Filename contains query words
    for (const word of queryWords) {
      if (fileName.includes(word)) {
        score += 10;
      }
    }

    // File path contains query
    const filePath = this.constructFullPath(file).toLowerCase();
    if (filePath.includes(queryLower)) {
      score += 20;
    }

    // Recently modified files get bonus
    if (file.lastModifiedDateTime) {
      const modifiedDate = new Date(file.lastModifiedDateTime);
      const daysSinceModified = (Date.now() - modifiedDate.getTime()) / (1000 * 60 * 60 * 24);
      if (daysSinceModified < 30) score += 5;
      if (daysSinceModified < 7) score += 10;
    }

    return score;
  }

  // Helper: Recursively get all files from drive
  async getAllFilesRecursively(folderId = 'root', maxFiles = 5000) {
    const allFiles = [];
    const processedIds = new Set();
    const foldersToProcess = [folderId];

    while (foldersToProcess.length > 0 && allFiles.length < maxFiles) {
      const currentFolderId = foldersToProcess.shift();

      try {
        const endpoint = currentFolderId === 'root'
          ? 'https://graph.microsoft.com/v1.0/me/drive/root/children'
          : `https://graph.microsoft.com/v1.0/me/drive/items/${currentFolderId}/children`;

        const response = await axios.get(endpoint, {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          params: {
            $top: 1000,
          },
        });

        const items = response.data.value || [];

        for (const item of items) {
          if (processedIds.has(item.id)) continue;
          processedIds.add(item.id);

          if (item.folder) {
            // Skip system folders
            if (this.shouldSkipFile(item.name, '')) {
              continue;
            }
            // Add folder to be processed
            foldersToProcess.push(item.id);
          } else {
            // Add file to results (filtering happens later)
            allFiles.push(item);
            if (allFiles.length >= maxFiles) break;
          }
        }
      } catch (err) {
        console.error(`Error processing folder ${currentFolderId}:`, err.message);
        // Continue with other folders
      }
    }

    return allFiles;
  }

  // Helper: Perform comprehensive content search
  async searchWithContentAnalysis(query, maxResults, includeShared, fileTypes) {
    const allResults = [];

    console.error('Recursively scanning all files in drive...');

    // Get ALL files from user's drive recursively
    let files = await this.getAllFilesRecursively('root', 5000);

    console.error(`Found ${files.length} files. Filtering and prioritizing...`);

    // Filter out system files and non-searchable files
    files = files.filter(file => {
      if (file.folder) return false;
      if (this.shouldSkipFile(file.name, this.constructFullPath(file))) return false;
      return true;
    });

    console.error(`After filtering: ${files.length} searchable files.`);

    // Count Office documents that will be processed with text extraction
    const officeDocCount = files.filter(f => this.isOfficeDocument(f.name)).length;
    if (officeDocCount > 0) {
      console.error(`ℹ️  Found ${officeDocCount} Office documents - attempting text extraction`);
      console.error(`   Note: Office text extraction has limited availability. For best results, use searchDepth="auto"`);
    }

    // Calculate relevance scores and sort by relevance
    const scoredFiles = files.map(file => ({
      file,
      score: this.getFileRelevanceScore(file, query)
    }));

    // Sort by score (highest first)
    scoredFiles.sort((a, b) => b.score - a.score);

    // Separate files into high-relevance and low-relevance
    const highRelevanceFiles = scoredFiles.filter(sf => sf.score > 0);
    const lowRelevanceFiles = scoredFiles.filter(sf => sf.score === 0);

    console.error(`High-relevance files: ${highRelevanceFiles.length}, Low-relevance: ${lowRelevanceFiles.length}`);

    // Search high-relevance files first
    let filesProcessed = 0;
    let filesSearched = 0;

    // Process high-relevance files (always search content)
    for (const { file, score } of highRelevanceFiles) {
      if (allResults.length >= maxResults) break;

      filesProcessed++;
      const fileName = file.name;
      const queryLower = query.toLowerCase();

      // Check filename match
      const nameMatch = fileName.toLowerCase().includes(queryLower);

      // For high-relevance files, always search content if searchable
      let contentMatch = null;
      if (this.isSearchableFile(fileName, fileTypes)) {
        filesSearched++;
        const driveId = file.parentReference?.driveId;
        contentMatch = await this.searchFileContent(file.id, query, driveId, fileName);
      }

      if (nameMatch || (contentMatch && contentMatch.found)) {
        const result = {
          id: file.id,
          name: fileName,
          path: this.constructFullPath(file),
          webUrl: file.webUrl,
          size: file.size,
          lastModified: file.lastModifiedDateTime,
          author: file.createdBy?.user?.displayName,
          type: 'file',
          source: 'myDrive',
          matchType: nameMatch && contentMatch?.found ? 'both' : (nameMatch ? 'filename' : 'content'),
          contentMatches: contentMatch?.matches || [],
          preview: contentMatch?.preview || '',
          relevanceScore: score,
        };

        const driveId = file.parentReference?.driveId;
        if (driveId) {
          result.driveId = driveId;
        }

        allResults.push(result);
      }

      if (filesProcessed % 20 === 0) {
        console.error(`Searched ${filesSearched} files (${filesProcessed} processed), found ${allResults.length} matches...`);
      }
    }

    // Only search low-relevance files if we haven't found enough results
    if (allResults.length < maxResults && lowRelevanceFiles.length > 0) {
      console.error(`Not enough matches in high-relevance files. Searching broader set...`);

      // Limit how many low-relevance files we search (max 200)
      const maxLowRelevanceToSearch = Math.min(200, lowRelevanceFiles.length);

      for (let i = 0; i < maxLowRelevanceToSearch && allResults.length < maxResults; i++) {
        const { file } = lowRelevanceFiles[i];
        filesProcessed++;

        const fileName = file.name;

        // Only search content for searchable files
        if (this.isSearchableFile(fileName, fileTypes)) {
          filesSearched++;
          const driveId = file.parentReference?.driveId;
          const contentMatch = await this.searchFileContent(file.id, query, driveId, fileName);

          if (contentMatch && contentMatch.found) {
            const result = {
              id: file.id,
              name: fileName,
              path: this.constructFullPath(file),
              webUrl: file.webUrl,
              size: file.size,
              lastModified: file.lastModifiedDateTime,
              author: file.createdBy?.user?.displayName,
              type: 'file',
              source: 'myDrive',
              matchType: 'content',
              contentMatches: contentMatch.matches || [],
              preview: contentMatch.preview || '',
              relevanceScore: 0,
            };

            if (driveId) {
              result.driveId = driveId;
            }

            allResults.push(result);
          }
        }

        if (filesProcessed % 20 === 0) {
          console.error(`Searched ${filesSearched} files (${filesProcessed} processed), found ${allResults.length} matches...`);
        }
      }
    }

    console.error(`Content search complete. Searched ${filesSearched} files, found ${allResults.length} matches.`);

    // Include shared files if requested
    if (includeShared && allResults.length < maxResults) {
      try {
        const sharedResponse = await axios.get(
          'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe',
          {
            headers: {
              Authorization: `Bearer ${this.authTokens.accessToken}`,
            },
            params: {
              $top: 100,
            },
          }
        );

        const sharedFiles = sharedResponse.data.value || [];

        for (const item of sharedFiles) {
          if (item.folder || item.remoteItem?.folder) continue;

          const fileName = item.name;
          const fileId = item.remoteItem?.id || item.id;
          const driveId = item.remoteItem?.parentReference?.driveId;
          const queryLower = query.toLowerCase();

          const nameMatch = fileName.toLowerCase().includes(queryLower);

          let contentMatch = null;
          if (this.isSearchableFile(fileName, fileTypes)) {
            // For shared files, we need to pass the driveId
            contentMatch = await this.searchFileContent(fileId, query, driveId, fileName);
          }

          if (nameMatch || (contentMatch && contentMatch.found)) {
            const itemToUse = item.remoteItem || item;
            const result = {
              id: fileId,
              name: fileName,
              path: this.constructFullPath(itemToUse),
              webUrl: item.remoteItem?.webUrl || item.webUrl,
              size: item.size,
              lastModified: item.lastModifiedDateTime,
              author: item.remoteItem?.createdBy?.user?.displayName,
              type: 'file',
              source: 'sharedWithMe',
              sharedBy: item.remoteItem?.createdBy?.user?.displayName,
              matchType: nameMatch && contentMatch?.found ? 'both' : (nameMatch ? 'filename' : 'content'),
              contentMatches: contentMatch?.matches || [],
              preview: contentMatch?.preview || '',
            };

            // Include driveId for shared files
            if (driveId) {
              result.driveId = driveId;
            }

            allResults.push(result);

            if (allResults.length >= maxResults) break;
          }
        }
      } catch (err) {
        // Continue without shared files if it fails
      }
    }

    return allResults;
  }

  async searchMyFiles(args) {
    await this.ensureAuthenticated();

    const {
      query,
      maxResults = 20,
      includeShared = false,
      searchDepth = 'filename',
      fileTypes = null
    } = args;

    try {
      let allFiles = [];
      let searchMethod = 'filename';

      // Choose search strategy based on searchDepth
      if (searchDepth === 'auto') {
        // Try Graph API first (best for work accounts)
        console.error('Attempting Microsoft Graph Search API...');
        const graphResult = await this.searchWithGraphAPI(query, maxResults);

        if (graphResult.success) {
          allFiles = graphResult.results;
          searchMethod = 'graphAPI';
          console.error('Graph API search successful!');
        } else {
          console.error('Graph API not available, falling back to content search...');
          // Fall back to content search
          allFiles = await this.searchWithContentAnalysis(query, maxResults, includeShared, fileTypes);
          searchMethod = 'contentAnalysis';
        }
      } else if (searchDepth === 'content') {
        // Deep content search
        console.error('Performing comprehensive content search...');
        allFiles = await this.searchWithContentAnalysis(query, maxResults, includeShared, fileTypes);
        searchMethod = 'contentAnalysis';
      } else {
        // Default: filename-only search using OneDrive API
        const response = await axios.get(
          `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${encodeURIComponent(query)}')`,
          {
            headers: {
              Authorization: `Bearer ${this.authTokens.accessToken}`,
            },
            params: {
              $top: maxResults,
            },
          }
        );

        let items = response.data.value || [];
        allFiles = items.map((item) => ({
          id: item.id,
          name: item.name,
          path: this.constructFullPath(item),
          webUrl: item.webUrl,
          size: item.size,
          lastModified: item.lastModifiedDateTime,
          author: item.createdBy?.user?.displayName,
          type: item.folder ? "folder" : "file",
          source: "myDrive",
          matchType: "filename",
        }));

        // If includeShared is true, also search through shared files
        if (includeShared) {
          try {
            const sharedResponse = await axios.get(
              "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe",
              {
                headers: {
                  Authorization: `Bearer ${this.authTokens.accessToken}`,
                },
                params: {
                  $top: 100,
                },
              }
            );

            const sharedItems = sharedResponse.data.value || [];
            const matchingShared = sharedItems
              .filter((item) =>
                item.name.toLowerCase().includes(query.toLowerCase())
              )
              .slice(0, Math.floor(maxResults / 2))
              .map((item) => {
                const itemToUse = item.remoteItem || item;
                return {
                  id: item.remoteItem?.id || item.id,
                  name: item.name,
                  path: this.constructFullPath(itemToUse),
                  webUrl: item.remoteItem?.webUrl || item.webUrl,
                  size: item.size,
                  lastModified: item.lastModifiedDateTime,
                  author: item.remoteItem?.createdBy?.user?.displayName,
                  type: item.folder || item.remoteItem?.folder ? "folder" : "file",
                  source: "sharedWithMe",
                  sharedBy: item.remoteItem?.createdBy?.user?.displayName,
                  matchType: "filename",
                };
              });

            allFiles = [...allFiles, ...matchingShared];
          } catch (sharedError) {
            console.error("Shared files search failed:", sharedError.message);
          }
        }

        searchMethod = 'oneDriveAPI';
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                query: query,
                searchMethod: searchMethod,
                searchDepth: searchDepth,
                resultCount: allFiles.length,
                includeShared: includeShared,
                files: allFiles,
                note: searchMethod === 'contentAnalysis'
                  ? 'Results include content matches. Check contentMatches field for line numbers and previews.'
                  : searchMethod === 'graphAPI'
                  ? 'Using Microsoft Graph Search API (comprehensive, includes content search)'
                  : 'Using filename-only search. Use searchDepth="content" or "auto" for comprehensive search.',
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Search failed: ${error.response?.data?.error?.message || error.message}`);
    }
  }

  async listMyFiles(args) {
    await this.ensureAuthenticated();

    const { folderPath = "", limit = 20 } = args;

    try {
      // If no folder path, get recent files
      let endpoint;
      if (!folderPath) {
        endpoint = "https://graph.microsoft.com/v1.0/me/drive/recent";
      } else {
        endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}:/children`;
      }

      const response = await axios.get(endpoint, {
        headers: {
          Authorization: `Bearer ${this.authTokens.accessToken}`,
        },
        params: {
          $top: limit,
        },
      });

      const items = response.data.value || [];
      const files = items.map((item) => ({
        id: item.id,
        name: item.name,
        path: this.constructFullPath(item),
        type: item.folder ? "folder" : "file",
        size: item.size,
        lastModified: item.lastModifiedDateTime,
        webUrl: item.webUrl,
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                path: folderPath || "recent files",
                count: files.length,
                items: files,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to list files: ${error.response?.data?.error?.message || error.message}`);
    }
  }

  async inspectFileMetadata(args) {
    await this.ensureAuthenticated();

    const { fileId, driveId } = args;

    if (!fileId) {
      throw new Error("fileId is required");
    }

    try {
      // Try multiple endpoints to get file metadata
      const results = {
        fileId: fileId,
        driveId: driveId || null,
        attempts: [],
      };

      // Attempt 1: Try with me/drive/items
      try {
        const endpoint1 = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`;
        const response1 = await axios.get(endpoint1, {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        });

        results.attempts.push({
          method: 'me/drive/items/{id}',
          status: 'SUCCESS',
          data: response1.data,
        });
      } catch (err1) {
        results.attempts.push({
          method: 'me/drive/items/{id}',
          status: 'FAILED',
          error: err1.response?.status || err1.message,
        });
      }

      // Attempt 2: Try with drives/{driveId}/items if driveId provided
      if (driveId) {
        try {
          const endpoint2 = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}`;
          const response2 = await axios.get(endpoint2, {
            headers: {
              Authorization: `Bearer ${this.authTokens.accessToken}`,
            },
          });

          results.attempts.push({
            method: 'drives/{driveId}/items/{id}',
            status: 'SUCCESS',
            data: response2.data,
          });
        } catch (err2) {
          results.attempts.push({
            method: 'drives/{driveId}/items/{id}',
            status: 'FAILED',
            error: err2.response?.status || err2.message,
          });
        }
      }

      // Attempt 3: Search for the file to get its real location
      try {
        const searchResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${fileId}')`,
          {
            headers: {
              Authorization: `Bearer ${this.authTokens.accessToken}`,
            },
            params: {
              $top: 5,
            },
          }
        );

        results.attempts.push({
          method: 'search by id',
          status: 'SUCCESS',
          data: searchResponse.data.value,
        });
      } catch (err3) {
        results.attempts.push({
          method: 'search by id',
          status: 'FAILED',
          error: err3.response?.status || err3.message,
        });
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(results, null, 2),
          },
        ],
      };
    } catch (error) {
      throw new Error(
        `Failed to inspect file metadata: ${error.message}`
      );
    }
  }

  async getFileContent(args) {
    await this.ensureAuthenticated();

    const { fileId, driveId } = args;

    if (!fileId) {
      throw new Error("fileId is required");
    }

    try {
      // Construct the correct endpoint based on whether it's a shared file
      let endpoint;
      if (driveId) {
        // Shared file - use drives/{driveId}/items/{itemId}
        endpoint = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`;
      } else {
        // Regular file - use me/drive/items/{itemId}
        endpoint = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`;
      }

      const response = await axios.get(endpoint, {
        headers: {
          Authorization: `Bearer ${this.authTokens.accessToken}`,
        },
        responseType: "text",
      });

      return {
        content: [
          {
            type: "text",
            text: response.data,
          },
        ],
      };
    } catch (error) {
      throw new Error(
        `Failed to get file content: ${error.response?.data?.error?.message || error.message}`
      );
    }
  }

  async listRecentFiles(args) {
    await this.ensureAuthenticated();

    const { limit = 10 } = args;

    try {
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/me/drive/recent",
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          params: {
            $top: limit,
          },
        }
      );

      const items = response.data.value || [];
      const files = items
        .filter((item) => !item.folder)
        .map((item) => ({
          id: item.id,
          name: item.name,
          path: this.constructFullPath(item),
          size: item.size,
          lastModified: item.lastModifiedDateTime,
          lastAccessed: item.lastAccessedDateTime,
          webUrl: item.webUrl,
        }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                recentFiles: files,
                count: files.length,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(
        `Failed to list recent files: ${error.response?.data?.error?.message || error.message}`
      );
    }
  }

  async listSharedFiles(args) {
    await this.ensureAuthenticated();

    const { limit = 20 } = args;

    try {
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe",
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          params: {
            $top: limit,
          },
        }
      );

      const items = response.data.value || [];
      const files = items.map((item) => {
        // For shared items, use remoteItem if available for path construction
        const itemToUse = item.remoteItem || item;
        return {
          id: item.remoteItem?.id || item.id,
          name: item.name,
          path: this.constructFullPath(itemToUse),
          type: item.folder || item.remoteItem?.folder ? "folder" : "file",
          size: item.size,
          lastModified: item.lastModifiedDateTime,
          webUrl: item.remoteItem?.webUrl || item.webUrl,
          sharedBy: item.remoteItem?.createdBy?.user?.displayName,
          sharedDateTime: item.remoteItem?.shared?.sharedDateTime,
        };
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                sharedFiles: files,
                count: files.length,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(
        `Failed to list shared files: ${error.response?.data?.error?.message || error.message}`
      );
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("OneDrive MCP server running on stdio");
  }
}

// Start server
const server = new SharePointMCP();
server.run().catch(console.error);
