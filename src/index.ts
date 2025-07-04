// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerChatTools } from './tools/chatTools.js';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load environment variables from .env file in the project root
const envPath = path.resolve(__dirname, '../.env');
dotenv.config({ path: envPath });

// Check for required environment variables
const requiredEnvVars = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"];
const missingVars = requiredEnvVars.filter((varName) => !process.env[varName]);

if (missingVars.length > 0) {
  // Error messages removed for production
  process.exit(1);
}

// Environment variables loaded

// Create server instance
const server = new McpServer({
  name: "ms-teams-chat",
  version: "1.0.0",
  description: "Microsoft Teams Chat Integration"
});

// Register tools
registerChatTools(server);

// Store the original console methods
const originalConsole = {
  log: console.log,
  error: console.error,
  warn: console.warn,
  info: console.info,
  debug: console.debug
};

// Suppress all console output during server initialization
function suppressConsole() {
  console.log = () => {};
  console.error = () => {};
  console.warn = () => {};
  console.info = () => {};
  console.debug = () => {};
}

// Restore original console methods
function restoreConsole() {
  console.log = originalConsole.log;
  console.error = originalConsole.error;
  console.warn = originalConsole.warn;
  console.info = originalConsole.info;
  console.debug = originalConsole.debug;
}

// Start the server
async function main() {
  try {
    // Suppress console output during server initialization
    suppressConsole();
    
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    // Restore console after successful connection
    restoreConsole();
  } catch (error) {
    // Restore console before exiting on error
    restoreConsole();
    process.exit(1);
  }
}

// Start the server
main().catch(() => {
  process.exit(1);
});