# MCP Server for Microsoft Teams Chat

A Node.js implementation of the Model Context Protocol (MCP) server for managing Microsoft Teams chats, built with TypeScript for type safety and better developer experience.

## Features

- Create new chats in Microsoft Teams
- Send messages to existing chats
- Microsoft Graph API integration
- TypeScript support with comprehensive type definitions
- Environment-based configuration
- Error handling and logging
- MCP protocol implementation

## Prerequisites

- Node.js 18.x or later
- npm 9.x or later
- TypeScript 5.0 or later
- Microsoft 365 account with appropriate permissions
- Azure AD application with required API permissions:
  - Chat.Create (Application)
  - User.Read (delegated)
  - offline_access (for refresh tokens)

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/mcp-server-nodejs.git
cd mcp-server-nodejs
```

### 2. Install dependencies

```bash
npm install
```

### 3. Set up Azure AD Application

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to "Azure Active Directory" > "App registrations" > "New registration"
3. Enter a name for your application and select the appropriate account type
4. After registration, note down the following from the "Overview" page:
   - Application (client) ID
   - Directory (tenant) ID
5. Go to "Certificates & secrets" and create a new client secret
6. Go to "API permissions" and add the following Microsoft Graph API permissions:
   - **Application Permissions**:
     - `Chat.Create` - Required for creating new chats
     - `Chat.Read.All` - Required for reading chat details after creation (recommended)
   - **Delegated Permissions**:
     - `User.Read` - Required for basic user information
     - `offline_access` - Required for refresh tokens

> **Note on Permissions**:
> - `Chat.Read.All` is required to read full chat details after creation. Without it, the server will only return basic chat information (ID and type). For full functionality, it's recommended to request this permission from your Azure AD administrator.
   - Chat.Create (Application)
   - User.Read (Delegated)
   - offline_access
7. Grant admin consent for the permissions
   - User.Read (delegated)
   - offline_access
7. Grant admin consent for the permissions

### 4. Configure Environment Variables

Create a new `.env` file in the project root and add the following environment variables:

```
# Azure AD Application Settings
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_ID=your_client_id_here
AZURE_CLIENT_SECRET=your_client_secret_here

# User settings (optional, defaults to 'me' which is the current user)
USER_ID=me  # or specific user email/ID
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

1. **Authentication Errors**
   - Ensure your Azure AD application has the `Chat.Create` (Application) permission
   - Verify your client secret hasn't expired
   - Check that your tenant ID and client ID are correct

2. **Missing Environment Variables**
   - The server will fail to start if required variables are missing
   - Double-check your `.env` file and ensure all required variables are set

3. **API Permissions**
   - Ensure the Azure AD app has admin consent for the required permissions
   - The application needs `Chat.Create` as an Application permission (not Delegated)

4. **Rate Limiting**
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
