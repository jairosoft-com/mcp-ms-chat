# Microsoft Teams Chat MCP Server

This project provides an MCP (Model Context Protocol) server for interacting with Microsoft Teams chats. It allows AI assistants to list chats, view chat details, and send messages to Microsoft Teams.

## Features

- List all your Microsoft Teams chats
- View detailed chat information including members
- Send messages to group chats
- Secure authentication using Microsoft Graph API tokens

## Prerequisites

- Node.js 16+ and npm
- Microsoft 365 account with Teams access
- Azure AD App Registration with Microsoft Graph API permissions

## Configuration

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd mcp-ms-chat
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file in the root directory with the following variables:
   ```
   AUTH_TOKEN=your_microsoft_graph_token
   NODE_ENV=development
   ```

4. To get a Microsoft Graph token:
   - Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
   - Sign in with your Microsoft 365 account
   - Request the following permissions:
     - Chat.Read
     - Chat.ReadWrite
     - ChatMessage.Send
   - Copy the access token and use it as `AUTH_TOKEN`

## Running Locally

```bash
# Development mode with hot-reload
npm run dev

# Production build
npm run build
npm start
```

The server will be available at `http://localhost:8787`

## Available MCP Tools

### 1. listChats
List all your Microsoft Teams chats with details.

**Parameters:**
- `top` (number, optional): Number of chats to return (default: 50, max: 100)
- `skip` (number, optional): Number of chats to skip for pagination
- `includeMembers` (boolean, optional): Whether to include member information (default: true)

### 2. sendMessage
Send a message to a Microsoft Teams chat.

**Parameters:**
- `chatId` (string, required): The ID of the chat to send the message to
- `content` (string, required): The message content
- `contentType` (string, optional): 'text' or 'html' (default: 'text')
- `importance` (string, optional): 'normal', 'high', or 'urgent' (default: 'normal')
- `subject` (string, optional): Message subject

### 3. createChat
Create a new chat.

**Parameters:**
- `chatType` (string, required): 'oneOnOne', 'group', 'meeting', or 'unknown'
- `topic` (string, optional): Chat topic/subject
- `members` (array, required): Array of member objects with `id`, `displayName`, and `email`

## Connecting Clients

### Cloudflare AI Playground
1. Go to [Cloudflare AI Playground](https://playground.ai.cloudflare.com/)
2. Enter your MCP server URL (e.g., `http://localhost:8787/sse` for local development)
3. Start using the MCP tools directly in the playground

### Claude Desktop
1. Open Claude Desktop
2. Go to Settings > Developer > Edit Config
3. Add your MCP server configuration:

```json
{
  "mcpServers": {
    "teams-chat": {
      "command": "npx",
      "args": [
        "mcp-remote",
        "http://localhost:8787/sse?token=your_auth_token"
      ]
    }
  }
}
```

## Deployment

### Cloudflare Workers
[![Deploy to Workers](https://deploy.workers.cloudflare.com/button)](https://deploy.workers.cloudflare.com/?url=https://github.com/yourusername/mcp-ms-chat)

Or deploy using Wrangler:
```bash
npm run deploy
```

### Docker
See [Docker README](README-docker.md) for container deployment options.

## Security Notes

- Never commit your `.env` file to version control
- Use environment variables for sensitive information
- Regularly rotate your Microsoft Graph tokens
- The server requires a valid token for all operations

## Troubleshooting

- **Invalid Token**: Ensure your Microsoft Graph token is valid and has the correct permissions
- **CORS Issues**: When running locally, ensure your client is configured to allow requests to your server
- **Rate Limiting**: Microsoft Graph API has rate limits; implement proper error handling in your client

## License

[MIT](LICENSE)

```json
{
  "mcpServers": {
    "calculator": {
      "command": "npx",
      "args": [
        "mcp-remote",
        "http://localhost:8787/sse?token=your-token" // or remote-mcp-server-authless.your-account.workers.dev/sse?token=your-token
      ]
    }
  }
}
```

Restart Claude and you should see the tools become available.
