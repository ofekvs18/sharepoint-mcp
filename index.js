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
            "Retrieve the content of a specific file from your OneDrive",
          inputSchema: {
            type: "object",
            properties: {
              fileId: {
                type: "string",
                description: "File ID from search results",
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

  async searchMyFiles(args) {
    await this.ensureAuthenticated();

    const { query, maxResults = 20, includeShared = false } = args;

    try {
      // Use OneDrive Drive API search - works for both personal and work accounts
      // The Microsoft Graph Search API (/search/query) only works for work accounts
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
      let allFiles = items.map((item) => ({
        id: item.id,
        name: item.name,
        path: item.webUrl,
        size: item.size,
        lastModified: item.lastModifiedDateTime,
        author: item.createdBy?.user?.displayName,
        type: item.folder ? "folder" : "file",
        source: "myDrive",
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
                $top: 100, // Get more shared files to search through
              },
            }
          );

          const sharedItems = sharedResponse.data.value || [];
          // Filter shared items by query (simple name matching)
          const matchingShared = sharedItems
            .filter((item) =>
              item.name.toLowerCase().includes(query.toLowerCase())
            )
            .slice(0, Math.floor(maxResults / 2)) // Limit shared results
            .map((item) => ({
              id: item.remoteItem?.id || item.id,
              name: item.name,
              path: item.remoteItem?.webUrl || item.webUrl,
              size: item.size,
              lastModified: item.lastModifiedDateTime,
              author: item.remoteItem?.createdBy?.user?.displayName,
              type: item.folder || item.remoteItem?.folder ? "folder" : "file",
              source: "sharedWithMe",
              sharedBy: item.remoteItem?.createdBy?.user?.displayName,
            }));

          allFiles = [...allFiles, ...matchingShared];
        } catch (sharedError) {
          // If shared search fails, continue with just myDrive results
          console.error("Shared files search failed:", sharedError.message);
        }
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                query: query,
                resultCount: allFiles.length,
                includeShared: includeShared,
                files: allFiles,
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

    const { fileId } = args;

    if (!fileId) {
      throw new Error("fileId is required");
    }

    try {
      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`,
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
          responseType: "text",
        }
      );

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
      const files = items.map((item) => ({
        id: item.remoteItem?.id || item.id,
        name: item.name,
        type: item.folder || item.remoteItem?.folder ? "folder" : "file",
        size: item.size,
        lastModified: item.lastModifiedDateTime,
        webUrl: item.remoteItem?.webUrl || item.webUrl,
        sharedBy: item.remoteItem?.createdBy?.user?.displayName,
        sharedDateTime: item.remoteItem?.shared?.sharedDateTime,
      }));

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
