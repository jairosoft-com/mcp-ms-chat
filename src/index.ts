import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Env } from "./interface/chatInterfaces";
import { createChat, setAuthToken } from "./tools/chatTools";

// Define our MCP agent with tools
export class MyMCP extends McpAgent {
	server = new McpServer({
		name: "Microsoft Chat Fetcher",
		version: "1.0.0",
	});

	async init() {
		// Get tool definitions by calling the factory functions
		const createChatTool = createChat();

		// Register the createChat tool
		this.server.tool(
			createChatTool.name,
			createChatTool.schema,
			createChatTool.handler
		);

    }
}

export default {
    fetch(request: Request, env: Env, ctx: ExecutionContext) {
        const url = new URL(request.url);
        const tokenFromUrl = url.searchParams.get('token');
        const authToken = tokenFromUrl || env.AUTH_TOKEN;
        
        console.log('Auth token received:', authToken ? `${authToken.substring(0, 10)}...` : 'No token found');
		
		setAuthToken(authToken);

		if (url.pathname === "/sse" || url.pathname === "/sse/message") {
			return MyMCP.serveSSE("/sse").fetch(request, env, ctx);
		}

		if (url.pathname === "/mcp") {
			return MyMCP.serve("/mcp").fetch(request, env, ctx);
		}

		return new Response("Not found", { status: 404 });
	},
};