// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerChatTools } from './tools/chatTools.js';

// Create server instance
const server = new McpServer({
  name: "ms-teams-chat",
  version: "1.0.0",
  description: "Microsoft Teams Chat Integration"
});

// Register tools
registerChatTools(server);

// Start the server
async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("Chat MCP Server running on stdio");
  } catch (error) {
    process.exit(1);
  }
}

// Start the server
main().catch(() => {
  process.exit(1);
});