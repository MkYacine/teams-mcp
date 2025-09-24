#!/usr/bin/env node
// src/index.ts
import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";
import { registerCallTools } from "./tools/call.js";

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const TOKEN_PATH = join(homedir(), ".msgraph-mcp-auth.json");
const TENANT_ID = "common";

interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

interface TokenResponse {
  access_token: string;
  refresh_token: string;
  expires_in: number;
  token_type: string;
  scope: string;
}

interface StoredAuthInfo {
  clientId: string;
  authenticated: boolean;
  timestamp: string;
  expiresAt: string;
  accessToken: string;
  refreshToken: string;
}

// Manual OAuth 2.0 Device Code Flow
async function getDeviceCode(): Promise<DeviceCodeResponse> {
  const scopes = [
    "User.Read",
    "User.ReadBasic.All",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Read.All",
    "ChannelMessage.Send",
    "TeamMember.Read.All",
    "Chat.ReadBasic",
    "Chat.ReadWrite",
    "offline_access",
  ];

  const response = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      scope: scopes.join(" "),
    }),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to get device code: ${response.status} ${error}`);
  }

  return await response.json();
}

async function pollForToken(deviceCode: string, interval: number): Promise<TokenResponse> {
  const maxAttempts = 100; // Prevent infinite polling
  let attempts = 0;

  while (attempts < maxAttempts) {
    await new Promise(resolve => setTimeout(resolve, interval * 1000));

    const response = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: CLIENT_ID,
        device_code: deviceCode,
      }),
    });

    const result = await response.json();

    if (response.ok) {
      return result as TokenResponse;
    }

    // Handle expected errors during polling
    if (result.error === "authorization_pending") {
      attempts++;
      continue;
    }

    if (result.error === "slow_down") {
      interval += 5; // Increase polling interval
      attempts++;
      continue;
    }

    // Handle terminal errors
    if (result.error === "authorization_declined") {
      throw new Error("User declined the authorization request");
    }

    if (result.error === "expired_token") {
      throw new Error("Device code expired. Please try again.");
    }

    throw new Error(`Authentication failed: ${result.error} - ${result.error_description}`);
  }

  throw new Error("Authentication timed out");
}

// Authentication functions
async function authenticate() {
  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));

  try {
    // Step 1: Get device code
    console.log("📱 Initiating device code authentication...");
    const deviceCodeInfo = await getDeviceCode();

    // Step 2: Show user instructions
    console.log("\n📱 Please complete authentication:");
    console.log(`🌐 Visit: ${deviceCodeInfo.verification_uri}`);
    console.log(`🔑 Enter code: ${deviceCodeInfo.user_code}`);
    console.log("\n⏳ Waiting for you to complete authentication...");

    // Step 3: Poll for token
    const tokenResponse = await pollForToken(deviceCodeInfo.device_code, deviceCodeInfo.interval);

    // Step 4: Calculate expiration time
    const expiresAt = new Date(Date.now() + (tokenResponse.expires_in * 1000));

    // Step 5: Save authentication info with both tokens
    const authInfo: StoredAuthInfo = {
      clientId: CLIENT_ID,
      authenticated: true,
      timestamp: new Date().toISOString(),
      expiresAt: expiresAt.toISOString(),
      accessToken: tokenResponse.access_token,
      refreshToken: tokenResponse.refresh_token,
    };

    await fs.writeFile(TOKEN_PATH, JSON.stringify(authInfo, null, 2));

    console.log("\n✅ Authentication successful!");
    console.log(`💾 Credentials saved to: ${TOKEN_PATH}`);
    console.log(`⏰ Access token expires: ${expiresAt.toLocaleString()}`);
    console.log("\n🚀 You can now use the MCP server in Cursor!");
    console.log("   The server will automatically refresh tokens as needed.");
    
  } catch (error) {
    console.error(
      "\n❌ Authentication failed:",
      error instanceof Error ? error.message : String(error)
    );
    process.exit(1);
  }
}

async function checkAuth() {
  try {
    const data = await fs.readFile(TOKEN_PATH, "utf8");
    const authInfo: StoredAuthInfo = JSON.parse(data);

    if (authInfo.authenticated && authInfo.clientId) {
      console.log("✅ Authentication found");
      console.log(`📅 Authenticated on: ${authInfo.timestamp}`);

      const expiresAt = new Date(authInfo.expiresAt);
      const now = new Date();

      if (expiresAt > now) {
        console.log(`⏰ Access token expires: ${expiresAt.toLocaleString()}`);
        console.log("🎯 Ready to use with MCP server!");
      } else {
        console.log(`⚠️  Access token expired at: ${expiresAt.toLocaleString()}`);
        console.log("🔄 Server will automatically refresh token when needed");
        console.log("🎯 Ready to use with MCP server!");
      }
      return true;
    }
  } catch (_error) {
    console.log("❌ No authentication found");
    return false;
  }
  return false;
}

async function logout() {
  try {
    await fs.unlink(TOKEN_PATH);
    console.log("✅ Successfully logged out");
    console.log("🔄 Run 'npx @floriscornel/teams-mcp@latest authenticate' to re-authenticate");
  } catch (_error) {
    console.log("ℹ️  No authentication to clear");
  }
}

// MCP Server setup
async function startMcpServer() {
  // Create MCP server
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.3.3",
  });

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();

  // Register all tools
  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);
  registerCallTools(server, graphService);

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Graph MCP Server started");
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const command = args[0];

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      await authenticate();
      return;
    case "check":
      await checkAuth();
      return;
    case "logout":
      await logout();
      return;
    case "help":
    case "--help":
    case "-h":
      console.log("Microsoft Graph MCP Server");
      console.log("");
      console.log("Usage:");
      console.log(
        "  npx @floriscornel/teams-mcp@latest authenticate # Authenticate with Microsoft"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest check        # Check authentication status"
      );
      console.log("  npx @floriscornel/teams-mcp@latest logout       # Clear authentication");
      console.log("  npx @floriscornel/teams-mcp@latest              # Start MCP server (default)");
      return;
    case undefined:
      // No command = start MCP server
      await startMcpServer();
      return;
    default:
      console.error(`Unknown command: ${command}`);
      console.error("Use --help to see available commands");
      process.exit(1);
  }
}

// Handle uncaught errors
process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});