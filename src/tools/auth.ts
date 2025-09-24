// src/tools/auth.ts
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(server: McpServer, graphService: GraphService) {
  // Authentication status tool
  server.tool(
    "auth_status",
    "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
    {},
    async () => {
      const status = await graphService.getAuthStatus();
      
      if (status.isAuthenticated) {
        let statusText = `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`;
        
        if (status.expiresAt) {
          const expiresAt = new Date(status.expiresAt);
          const now = new Date();
          const timeUntilExpiry = expiresAt.getTime() - now.getTime();
          
          if (timeUntilExpiry > 0) {
            const hours = Math.floor(timeUntilExpiry / (1000 * 60 * 60));
            const minutes = Math.floor((timeUntilExpiry % (1000 * 60 * 60)) / (1000 * 60));
            statusText += `\n⏰ Access token expires in ${hours}h ${minutes}m (${expiresAt.toLocaleString()})`;
            statusText += `\n🔄 Token will be automatically refreshed when needed`;
          } else {
            statusText += `\n⚠️ Access token expired at ${expiresAt.toLocaleString()}`;
            statusText += `\n🔄 Token will be refreshed on next API call`;
          }
        }
        
        return {
          content: [
            {
              type: "text",
              text: statusText,
            },
          ],
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: "❌ Not authenticated. Please run: npx @floriscornel/teams-mcp@latest authenticate",
            },
          ],
        };
      }
    }
  );
}