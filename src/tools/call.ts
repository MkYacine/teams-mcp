import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { exec } from "node:child_process";
import { promisify } from "node:util";
import type { GraphService } from "../services/graph.js";
import robot from "robotjs";

const execAsync = promisify(exec);

export function registerCallTools(server: McpServer, _: GraphService) {
  // Start Teams call with a user
  server.tool(
    "start_teams_call",
    "Start a Microsoft Teams call to a specific user using their user ID. This will launch the Teams call dialog where you can confirm the call.",
    {
      userId: z.string().describe("User ID obtained from user search tools (format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx or email address)"),
      callType: z.enum(["audio", "video"]).default("audio").describe("Type of call to initiate - 'audio' for audio-only call or 'video' for video call with camera enabled"),
    },
    async ({ userId, callType = "audio" }) => {
      try {
        const baseUrl = `https://teams.microsoft.com/l/call/0/0?users=8:orgid:${userId}`;
        const teamsUrl = callType === "video" ? `${baseUrl}&withVideo=true` : baseUrl;
        
        // Cross-platform URL launching with proper shell handling
        let command: string;
        
        if (process.platform === 'win32') {
          command = `start "" "${teamsUrl}"`;
        } else if (process.platform === 'darwin') {
          command = `open "${teamsUrl}"`;
        } else {
          command = `xdg-open "${teamsUrl}"`;
        }
        
        // Use appropriate shell for platform
        const shellOptions = process.platform === 'win32' 
          ? { shell: process.env.ComSpec || 'cmd.exe' }
          : { shell: '/bin/sh' };
        
        await execAsync(command, shellOptions);

        // Wait for Teams to load and show confirmaton popup
        await new Promise(resolve => setTimeout(resolve, 2000));
        // Simulate enter to confirm call
        robot.keyTap('enter');
        
        return {
          content: [
            {
              type: "text",
              text: `✅ Teams ${callType} call initiated to user: ${userId}\n🌐 URL: ${teamsUrl}\n⏳ Check your Teams app for the call dialog.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error starting call: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Join Teams meeting using meeting URL
  server.tool(
    "join_teams_meeting",
    "Join a Microsoft Teams meeting by extracting the meeting URL from Outlook calendar appointments using the Outlook Python API (win32com.client). You must use Python code to access the calendar appointment body and extract the Teams meeting URL that is wrapped in Outlook safe links. The tool will automatically handle safe link extraction and URL processing.",
    {
      meetingUrl: z.string().describe("Teams meeting URL extracted from Outlook calendar appointment body using win32com.client Python API. Must be the raw URL from appointment.Body property, typically wrapped in Outlook safe links starting with 'https://can01.safelinks.protection.outlook.com/ap/t-' or similar. Example format: 'https://can01.safelinks.protection.outlook.com/ap/t-59584e83/?url=https%3A%2F%2Fteams.microsoft.com%2Fl%2Fmeetup-join%2F...' - the tool will extract and process the actual Teams URL automatically."),
    },
    async ({ meetingUrl }) => {
      try {
        let cleanUrl = meetingUrl;
        
        // Extract from Outlook safe links if present
        const safeLinkMatch = cleanUrl.match(/(?:url=|&url=)([^&\s]+)/);
        if (safeLinkMatch) {
          // Only decode the safe link wrapper, not the Teams URL itself
          cleanUrl = decodeURIComponent(safeLinkMatch[1]);
        }
        
        // Validate that this looks like a Teams meeting URL
        if (!cleanUrl.includes('teams.microsoft.com/l/meetup-join')) {
          throw new Error(`Invalid Teams meeting URL format. Expected teams.microsoft.com/l/meetup-join but got: ${cleanUrl}`);
        }
        
        console.log(`Extracted Teams URL: ${cleanUrl}`);
        
        // Cross-platform URL launching
        let command: string;
        
        if (process.platform === 'win32') {
          command = `start "" "${cleanUrl}"`;
        } else if (process.platform === 'darwin') {
          command = `open "${cleanUrl}"`;
        } else {
          command = `xdg-open "${cleanUrl}"`;
        }
        
        const shellOptions = process.platform === 'win32' 
          ? { shell: process.env.ComSpec || 'cmd.exe' }
          : { shell: '/bin/sh' };
        
        await execAsync(command, shellOptions);

        // Wait for Teams to load and show pre-join screen
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        // Simulate 7 tabs + enter to join meeting
        for (let i = 0; i < 7; i++) {
          robot.keyTap('tab');
          await new Promise(resolve => setTimeout(resolve, 100)); // Small delay between tabs
        }
        robot.keyTap('enter');
        
        return {
          content: [
            {
              type: "text",
              text: `Teams meeting joined successfully.\nProcessed URL: ${cleanUrl}\nCheck your Teams app for the meeting window.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `Error joining meeting: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

}