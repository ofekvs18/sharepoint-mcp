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
                description: "Search depth: 'filename' (fast, names only), 'content' (comprehensive, searches file contents), 'auto' (tries Graph API first, falls back to content search)",
                enum: ["filename", "content", "auto"],
                default: "filename",
              },
              fileTypes: {
                type: "array",
                description: "Filter by file extensions (e.g., ['txt', 'js', 'md']). Only applies to content search.",
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

  // Helper: Check if file is searchable based on extension
  isSearchableFile(filename, allowedTypes = null) {
    const searchableExtensions = [
      'txt', 'md', 'js', 'jsx', 'ts', 'tsx', 'json', 'xml', 'html', 'htm',
      'css', 'scss', 'sass', 'py', 'java', 'c', 'cpp', 'h', 'cs', 'php',
      'rb', 'go', 'rs', 'sh', 'bash', 'yml', 'yaml', 'toml', 'ini', 'cfg',
      'log', 'csv', 'sql', 'r', 'swift', 'kt', 'dart', 'vue', 'svelte'
    ];

    const ext = filename.split('.').pop().toLowerCase();

    // If specific file types are requested, check against those
    if (allowedTypes && allowedTypes.length > 0) {
      return allowedTypes.includes(ext);
    }

    return searchableExtensions.includes(ext);
  }

  // Helper: Download and search file content
  async searchFileContent(fileId, query, driveId = null) {
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
        responseType: 'text',
        maxContentLength: 10 * 1024 * 1024, // Limit to 10MB
        timeout: 30000, // 30 second timeout
      });

      const content = response.data;
      const queryLower = query.toLowerCase();

      // Search for query in content (case-insensitive)
      if (content.toLowerCase().includes(queryLower)) {
        // Find context around matches
        const lines = content.split('\n');
        const matchingLines = [];

        for (let i = 0; i < lines.length; i++) {
          if (lines[i].toLowerCase().includes(queryLower)) {
            matchingLines.push({
              lineNumber: i + 1,
              content: lines[i].trim(),
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
            // Add folder to be processed
            foldersToProcess.push(item.id);
          } else {
            // Add file to results
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
    const files = await this.getAllFilesRecursively('root', 5000);

    console.error(`Found ${files.length} files. Searching through content...`);

    // Search through files
    let filesProcessed = 0;
    for (const file of files) {
      if (file.folder) continue; // Skip folders

      filesProcessed++;
      if (filesProcessed % 50 === 0) {
        console.error(`Processed ${filesProcessed}/${files.length} files, found ${allResults.length} matches...`);
      }

      const fileName = file.name;
      const queryLower = query.toLowerCase();

      // Check filename match first
      const nameMatch = fileName.toLowerCase().includes(queryLower);

      // For content search, check if file is searchable
      let contentMatch = null;
      if (this.isSearchableFile(fileName, fileTypes)) {
        // Pass driveId if available (for files from other drives)
        const driveId = file.parentReference?.driveId;
        contentMatch = await this.searchFileContent(file.id, query, driveId);
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
        };

        // Include driveId if it's from a different drive
        if (driveId) {
          result.driveId = driveId;
        }

        allResults.push(result);

        if (allResults.length >= maxResults) break;
      }
    }

    console.error(`Content search complete. Found ${allResults.length} matching files.`);

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
            contentMatch = await this.searchFileContent(fileId, query, driveId);
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
