// src/services/graph.ts
import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { Client } from "@microsoft/microsoft-graph-client";

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

interface StoredAuthInfo {
  clientId: string;
  authenticated: boolean;
  timestamp: string;
  expiresAt: string;
  accessToken: string;
  refreshToken: string;
}

interface RefreshTokenResponse {
  access_token: string;
  refresh_token: string;
  expires_in: number;
  token_type: string;
  scope: string;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private readonly authPath = join(homedir(), ".msgraph-mcp-auth.json");
  private isInitialized = false;
  private authInfo: StoredAuthInfo | undefined;
  private readonly clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
  private readonly tenantId = "common";

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async refreshAccessToken(): Promise<boolean> {
    if (!this.authInfo?.refreshToken) {
      console.error("No refresh token available");
      return false;
    }

    try {
      const response = await fetch(`https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({
          grant_type: "refresh_token",
          client_id: this.clientId,
          refresh_token: this.authInfo.refreshToken,
          scope: [
            "User.Read",
            "User.ReadBasic.All",
            "Team.ReadBasic.All",
            "Channel.ReadBasic.All",
            "ChannelMessage.Read.All",
            "ChannelMessage.Send",
            "TeamMember.Read.All",
            "Chat.ReadBasic",
            "Chat.ReadWrite",
          ].join(" "),
        }),
      });

      if (!response.ok) {
        const error = await response.text();
        console.error(`Token refresh failed: ${response.status} ${error}`);
        return false;
      }

      const tokenResponse: RefreshTokenResponse = await response.json();

      // Calculate new expiration time
      const newExpiresAt = new Date(Date.now() + (tokenResponse.expires_in * 1000));

      // Update stored auth info
      this.authInfo = {
        ...this.authInfo,
        accessToken: tokenResponse.access_token,
        refreshToken: tokenResponse.refresh_token, // Azure typically rotates refresh tokens
        expiresAt: newExpiresAt.toISOString(),
        timestamp: new Date().toISOString(), // Update last refresh time
      };

      // Save updated tokens to file
      await fs.writeFile(this.authPath, JSON.stringify(this.authInfo, null, 2));

      console.log(`✅ Token refreshed successfully. New expiration: ${newExpiresAt.toLocaleString()}`);
      return true;

    } catch (error) {
      console.error("Error refreshing token:", error);
      return false;
    }
  }

  private async ensureValidToken(): Promise<boolean> {
    if (!this.authInfo) {
      return false;
    }

    const expiresAt = new Date(this.authInfo.expiresAt);
    const now = new Date();
    
    // Refresh if token expires within the next 5 minutes
    const refreshThreshold = new Date(now.getTime() + (5 * 60 * 1000));

    if (expiresAt <= refreshThreshold) {
      console.log("🔄 Access token expired or expiring soon, attempting refresh...");
      return await this.refreshAccessToken();
    }

    return true;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      const authData = await fs.readFile(this.authPath, "utf8");
      this.authInfo = JSON.parse(authData);

      if (this.authInfo?.authenticated && this.authInfo?.accessToken) {
        // Ensure we have a valid token
        const hasValidToken = await this.ensureValidToken();
        
        if (!hasValidToken) {
          console.log("❌ Unable to refresh token. Please re-authenticate with: npx @floriscornel/teams-mcp@latest authenticate");
          return;
        }

        // Create Graph client with dynamic token provider
        this.client = Client.initWithMiddleware({
          authProvider: {
            getAccessToken: async () => {
              // Always check token validity before returning it
              const isValid = await this.ensureValidToken();
              if (!isValid || !this.authInfo?.accessToken) {
                throw new Error("Unable to obtain valid access token");
              }
              return this.authInfo.accessToken;
            },
          },
        });

        this.isInitialized = true;
      }
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
    }
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    if (!this.client) {
      return { isAuthenticated: false };
    }

    try {
      const me = await this.client.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? undefined,
        displayName: me?.displayName ?? undefined,
        expiresAt: this.authInfo?.expiresAt,
      };
    } catch (error) {
      // If API call fails, try to refresh token once more
      if (this.authInfo) {
        const refreshed = await this.refreshAccessToken();
        if (refreshed) {
          try {
            const me = await this.client.api("/me").get();
            return {
              isAuthenticated: true,
              userPrincipalName: me?.userPrincipalName ?? undefined,
              displayName: me?.displayName ?? undefined,
              expiresAt: this.authInfo?.expiresAt,
            };
          } catch (retryError) {
            console.error("Error getting user info after token refresh:", retryError);
          }
        }
      }
      
      console.error("Error getting user info:", error);
      return { isAuthenticated: false };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client) {
      throw new Error(
        "Not authenticated. Please run the authentication CLI tool first: npx @floriscornel/teams-mcp@latest authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }
}