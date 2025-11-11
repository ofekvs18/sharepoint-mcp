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

// SharePoint MCP Server
class SharePointMCP {
  constructor() {
    this.server = new Server(
      {
        name: "sharepoint-mcp",
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
      siteUrl: null,
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
          name: "set_site_url",
          description:
            "Set the SharePoint site URL to work with (e.g., https://yourtenant.sharepoint.com/sites/yoursite)",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "Full SharePoint site URL",
              },
            },
            required: ["siteUrl"],
          },
        },
        {
          name: "search_files",
          description:
            "Search for files in SharePoint by filename or content using Microsoft Graph API",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description: "Search query string",
              },
              searchType: {
                type: "string",
                enum: ["filename", "content", "both"],
                description: "Type of search to perform",
                default: "both",
              },
              maxResults: {
                type: "number",
                description: "Maximum number of results to return",
                default: 20,
              },
            },
            required: ["query"],
          },
        },
        {
          name: "get_folder_structure",
          description:
            "Get the folder structure of a SharePoint site or specific folder path",
          inputSchema: {
            type: "object",
            properties: {
              folderPath: {
                type: "string",
                description: "Relative folder path (leave empty for root)",
                default: "",
              },
              depth: {
                type: "number",
                description: "Depth of folder traversal (1-5)",
                default: 2,
              },
            },
          },
        },
        {
          name: "get_file_content",
          description:
            "Retrieve the content of a specific file from SharePoint",
          inputSchema: {
            type: "object",
            properties: {
              fileId: {
                type: "string",
                description: "File ID from search results",
              },
              filePath: {
                type: "string",
                description: "Alternate: full file path",
              },
            },
          },
        },
        {
          name: "list_recent_files",
          description:
            "List recently modified files in the SharePoint site",
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
      ],
    }));

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case "authenticate_sharepoint":
            return await this.authenticateSharePoint(args);
          case "set_site_url":
            return await this.setSiteUrl(args);
          case "search_files":
            return await this.searchFiles(args);
          case "get_folder_structure":
            return await this.getFolderStructure(args);
          case "get_file_content":
            return await this.getFileContent(args);
          case "list_recent_files":
            return await this.listRecentFiles(args);
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
            ? "Successfully authenticated with SharePoint using default public client ID! 🎉\n\n" +
              "Access token stored in memory.\n\n" +
              "Next step: Use 'set_site_url' to specify your SharePoint site URL.\n" +
              "Example: https://yourtenant.sharepoint.com/sites/yoursite"
            : "Successfully authenticated with SharePoint! Access token stored. Use 'set_site_url' to specify your SharePoint site.";

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

  async setSiteUrl(args) {
    const { siteUrl } = args;

    // Validate URL format
    if (!siteUrl.startsWith("https://") || !siteUrl.includes(".sharepoint.com")) {
      throw new Error(
        "Invalid SharePoint URL. Expected format: https://yourtenant.sharepoint.com/sites/yoursite"
      );
    }

    this.authTokens.siteUrl = siteUrl;

    return {
      content: [
        {
          type: "text",
          text: `SharePoint site URL set to: ${siteUrl}`,
        },
      ],
    };
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

  async ensureSiteUrl() {
    if (!this.authTokens.siteUrl) {
      throw new Error("Site URL not set. Please run 'set_site_url' first.");
    }
  }

  async searchFiles(args) {
    await this.ensureAuthenticated();
    await this.ensureSiteUrl();

    const { query, searchType = "both", maxResults = 20 } = args;

    // Build search query
    let searchQuery = query;
    if (searchType === "filename") {
      searchQuery = `filename:${query}`;
    } else if (searchType === "content") {
      searchQuery = `"${query}"`;
    }

    try {
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/search/query",
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
            "Content-Type": "application/json",
          },
          data: {
            requests: [
              {
                entityTypes: ["driveItem"],
                query: {
                  queryString: searchQuery,
                },
                from: 0,
                size: maxResults,
              },
            ],
          },
        }
      );

      const results = response.data.value[0]?.hitsContainers[0]?.hits || [];

      const files = results.map((hit) => ({
        id: hit.resource.id,
        name: hit.resource.name,
        path: hit.resource.webUrl,
        size: hit.resource.size,
        lastModified: hit.resource.lastModifiedDateTime,
        author: hit.resource.createdBy?.user?.displayName,
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                query: searchQuery,
                resultCount: files.length,
                files: files,
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

  async getFolderStructure(args) {
    await this.ensureAuthenticated();
    await this.ensureSiteUrl();

    const { folderPath = "", depth = 2 } = args;

    // Extract site info from URL
    const siteUrl = new URL(this.authTokens.siteUrl);
    const pathParts = siteUrl.pathname.split("/").filter((p) => p);

    if (pathParts[0] !== "sites" || !pathParts[1]) {
      throw new Error("Invalid site URL format");
    }

    const siteName = pathParts[1];

    try {
      // Get site ID
      const siteResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/sites/${siteUrl.hostname}:/sites/${siteName}`,
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        }
      );

      const siteId = siteResponse.data.id;

      // Get drive
      const driveResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`,
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        }
      );

      const driveId = driveResponse.data.id;

      // Recursive function to get folder structure
      const getFolders = async (path, currentDepth) => {
        if (currentDepth > depth) return null;

        const endpoint = path
          ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${path}:/children`
          : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;

        const response = await axios.get(endpoint, {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        });

        const items = response.data.value;
        const structure = [];

        for (const item of items) {
          if (item.folder) {
            const children =
              currentDepth < depth
                ? await getFolders(
                    path ? `${path}/${item.name}` : item.name,
                    currentDepth + 1
                  )
                : null;

            structure.push({
              name: item.name,
              type: "folder",
              itemCount: item.folder.childCount,
              children: children,
            });
          } else {
            structure.push({
              name: item.name,
              type: "file",
              size: item.size,
              lastModified: item.lastModifiedDateTime,
            });
          }
        }

        return structure;
      };

      const structure = await getFolders(folderPath, 1);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                siteName: siteName,
                rootPath: folderPath || "root",
                depth: depth,
                structure: structure,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(
        `Failed to get folder structure: ${error.response?.data?.error?.message || error.message}`
      );
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
        `https://graph.microsoft.com/v1.0/drives/items/${fileId}/content`,
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
    await this.ensureSiteUrl();

    const { limit = 10 } = args;

    const siteUrl = new URL(this.authTokens.siteUrl);
    const pathParts = siteUrl.pathname.split("/").filter((p) => p);
    const siteName = pathParts[1];

    try {
      const siteResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/sites/${siteUrl.hostname}:/sites/${siteName}`,
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        }
      );

      const siteId = siteResponse.data.id;

      const driveResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`,
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        }
      );

      const driveId = driveResponse.data.id;

      const filesResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$orderby=lastModifiedDateTime desc&$top=${limit}`,
        {
          headers: {
            Authorization: `Bearer ${this.authTokens.accessToken}`,
          },
        }
      );

      const files = filesResponse.data.value
        .filter((item) => !item.folder)
        .map((item) => ({
          id: item.id,
          name: item.name,
          size: item.size,
          lastModified: item.lastModifiedDateTime,
          author: item.lastModifiedBy?.user?.displayName,
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

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("SharePoint MCP server running on stdio");
  }
}

// Start server
const server = new SharePointMCP();
server.run().catch(console.error);
