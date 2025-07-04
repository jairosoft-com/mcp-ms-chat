# MCP Server for Microsoft Teams Chat

A Node.js implementation of the Model Context Protocol (MCP) server for managing Microsoft Teams chats, built with TypeScript for type safety and better developer experience.

## Features

- Create new chats in Microsoft Teams
- Send messages to existing chats
- Microsoft Graph API integration
- TypeScript support with comprehensive type definitions
- Simple token-based authentication
- Error handling and logging
- MCP protocol implementation

## Prerequisites

- Node.js 18.x or later
- npm 9.x or later
- TypeScript 5.0 or later
- Microsoft 365 account with appropriate permissions
- Microsoft Graph API access token with required permissions:
  - Chat.Create
  - Chat.ReadWrite
  - User.Read

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/mcp-ms-chat.git
cd mcp-ms-chat
```

### 2. Install dependencies

```bash
npm install
```

### 3. Build the project

```bash
npm run build
```

## Authentication

This service requires a valid Microsoft Graph API access token with the following permissions:
- `Chat.Create` - For creating new chats
- `Chat.ReadWrite` - For reading and updating chats
- `User.Read` - For basic user information

### Obtaining an Access Token

You can obtain an access token using one of these methods:

1. **Microsoft Graph Explorer**:
   - Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
   - Sign in with your Microsoft 365 account
   - Request the required permissions
   - Copy the access token

2. **Azure Portal**:
   - Register an application in Azure AD
   - Configure the required API permissions
   - Use the OAuth 2.0 flow to obtain a token

## Usage

All API endpoints require an `access_token` parameter with a valid Microsoft Graph API token.

Example request:
```
create-chat --accessToken "your_access_token_here" --topic "Team Discussion" --chatType "group" --members '[{"id":"user1@example.com"}]'
```

## Building the Project

To compile TypeScript to JavaScript:

```bash
npm run build
```

This will compile the TypeScript files and output them to the `build` directory.

## Usage

### Building and Running the Project

### Install Dependencies

```bash
npm install
```

### Build the Project

To compile TypeScript to JavaScript:

```bash
npm run build
```

### Run in Development Mode

```bash
npm run dev
```

### Run in Production Mode

```bash
npm start
```

## Available Tools

The server provides the following MCP tools for managing Teams chats:

### 1. **create-chat** - Create a new chat in Microsoft Teams

**Parameters:**
- `topic`: (Required) The topic/subject of the chat
- `chatType`: (Optional) Type of chat - 'oneOnOne', 'group', 'meeting', or 'unknown' (default: 'group')
- `members`: (Required) Array of members to add to the chat
  - `id`: (Required) The user ID or email of the member
  - `roles`: (Optional) Array of roles for the member (e.g., ['owner'])
- `message`: (Optional) Initial message to send to the chat
  - `content`: (Required) The message content
  - `contentType`: (Optional) 'text', 'html', or 'content' (default: 'text')

**Example Request:**
```json
{
  "topic": "Project Discussion",
  "chatType": "group",
  "members": [
    {
      "id": "user1@example.com",
      "roles": ["owner"]
    },
    {
      "id": "user2@example.com"
    }
  ],
  "message": {
    "content": "Hello, let's discuss the project!",
    "contentType": "text"
  }
}
```

**Example Response:**
```
# Project Discussion

**Type:** Group Chat
**Created:** 4/1/2025, 2:30:45 PM

**Members:**
- John Doe
- Jane Smith

[Open in Teams](https://teams.microsoft.com/...)

**Chat ID:** 19:meeting_...
```

## Development

### Project Structure

```
src/
  ├── interfaces/    # TypeScript interfaces
  │   └── chat.ts   # Chat-related interfaces
  ├── schemas/      # Zod validation schemas
  │   └── chatSchemas.ts  # Chat validation schemas
  ├── services/     # Business logic and API clients
  │   ├── authService.ts  # Authentication service
  │   └── chatService.ts  # Chat service for Microsoft Graph API
  ├── tools/        # MCP tool implementations
  │   └── chatTools.ts    # Chat-related MCP tools
  └── index.ts      # Server entry point
```

### Environment Variables

- `AZURE_TENANT_ID`: Your Azure AD tenant ID
- `AZURE_CLIENT_ID`: Your Azure AD application (client) ID
- `AZURE_CLIENT_SECRET`: Your Azure AD application client secret
- `USER_ID`: (Optional) The user ID or 'me' for the current user

### Testing

To test the chat functionality:

1. Ensure all environment variables are set correctly
2. Start the server in development mode:
   ```bash
   npm run dev
   ```
3. Use an MCP client to interact with the server and test the chat functionality

## Troubleshooting

### Common Issues
1. Access token is required

2. **API Permissions**
   - Ensure the Azure AD app has admin consent for the required permissions
   - The application needs `Chat.Create` as an Application permission (not Delegated)

3. **Rate Limiting**
   - The Microsoft Graph API has rate limits
   - Implement proper error handling and retry logic in your client



## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview)
- [Model Context Protocol](https://github.com/modelcontextprotocol)
- [TypeScript](https://www.typescriptlang.org/)

- Source code is in the `src` directory
- Built files are output to the `build` directory
- The project uses TypeScript for type safety
