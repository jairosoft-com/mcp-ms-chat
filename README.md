# MCP Server for Microsoft Teams Chat

A Node.js implementation of the Model Context Protocol (MCP) server for managing Microsoft Teams chats, built with TypeScript for type safety and better developer experience. This implementation uses Server-Sent Events (SSE) for real-time communication, making it ideal for containerized environments.

## Features

- Create new chats in Microsoft Teams
- Send and receive messages in real-time
- Server-Sent Events (SSE) for efficient streaming
- Microsoft Graph API integration
- TypeScript support with comprehensive type definitions
- Simple token-based authentication
- Error handling and logging
- MCP protocol implementation
- Docker-ready for easy deployment

## Prerequisites

- Node.js 18.x or later
- npm 9.x or later
- TypeScript 5.0 or later
- Docker (optional, for containerization)
- Microsoft 365 account with appropriate permissions
- Microsoft Graph API access token with required permissions:
  - Chat.Create
  - Chat.ReadWrite
  - User.Read
  - Chat.ReadBasic

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

### 4. Start the server

```bash
# Development
npm run dev

# Production
npm start

# Using Docker
docker build -t mcp-ms-chat .
docker run -p 3000:3000 mcp-ms-chat
```

## Authentication

This service requires a valid Microsoft Graph API access token with the following permissions:
- `Chat.Create` - For creating new chats
- `Chat.ReadWrite` - For reading and updating chats
- `User.Read` - For basic user information
- `Chat.ReadBasic` - For basic chat information

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

## API Endpoints

### Server-Sent Events (SSE)

The server uses SSE for real-time communication. Connect to the following endpoint:

```
GET /api/events
```

### REST API Endpoints

All API endpoints require an `Authorization` header with a valid Microsoft Graph API token.

#### Create a Chat
```
POST /api/chats
Content-Type: application/json
Authorization: Bearer <access_token>

{
  "topic": "Team Discussion",
  "chatType": "group",
  "members": [
    {"id": "user1@example.com"},
    {"id": "user2@example.com"}
  ]
}
```

#### Send a Message
```
POST /api/chats/{chatId}/messages
Content-Type: application/json
Authorization: Bearer <access_token>

{
  "content": "Hello, team!",
  "contentType": "text"
}
```

#### Get Chat Messages
```
GET /api/chats/{chatId}/messages?top=50&skip=0
Authorization: Bearer <access_token>
```

## Testing with Postman

### 1. Setup
1. Import the Postman collection from `postman/collection.json`
2. Set up environment variables in Postman:
   - `baseUrl`: Your server URL (e.g., `http://localhost:3000`)
   - `accessToken`: Your Microsoft Graph API access token

### 2. Sample Requests

#### Create a Chat
```
POST {{baseUrl}}/api/chats
Content-Type: application/json
Authorization: Bearer {{accessToken}}

{
  "topic": "Project Sync",
  "chatType": "group",
  "members": [
    {"id": "team1@example.com", "roles": ["owner"]},
    {"id": "member1@example.com", "roles": ["guest"]}
  ]
}
```

#### Send a Rich Text Message
```
POST {{baseUrl}}/api/chats/19:meeting_NDQ4M2E4ZWMtYjYyMy00YjA2LWI0Y2ItYmYzY2MxNzNlY2Y4@thread.v2/messages
Content-Type: application/json
Authorization: Bearer {{accessToken}}

{
  "content": "<h1>Important Update</h1><p>Please review the <strong>Q2 Report</strong> before our meeting tomorrow.</p>",
  "contentType": "html"
}
```

### 3. Testing SSE

1. Open a new tab in Postman
2. Set request to `GET {{baseUrl}}/api/events`
3. Add `Accept: text/event-stream` header
4. Send the request
5. In another tab, send a message to see real-time updates

## Development

### Building the Project

```bash
# Install dependencies
npm install

# Build TypeScript to JavaScript
npm run build

# Watch for changes (development)
npm run dev
```

### Environment Variables

Create a `.env` file in the root directory with the following variables:

```env
PORT=3000
NODE_ENV=development
LOG_LEVEL=debug
```

## Docker Deployment

### Building the Image

```bash
docker build -t mcp-ms-chat .
```

### Running the Container

```bash
docker run -d \
  -p 3000:3000 \
  -e PORT=3000 \
  -e NODE_ENV=production \
  -e LOG_LEVEL=info \
  --name mcp-ms-chat \
  mcp-ms-chat
```

### Docker Compose

```yaml
version: '3.8'
services:
  mcp-chat:
    build: .
    ports:
      - "3000:3000"
    environment:
      - PORT=3000
      - NODE_ENV=production
      - LOG_LEVEL=info
    restart: unless-stopped
```

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

**Members Sample Payload:**

Here's an example of how to structure the `members` array when creating a chat:

```json
[
  {
    "id": "user1@example.com",
    "roles": ["owner"]
  },
  {
    "id": "user2@example.com",
    "roles": []
  }
]
```

**Notes about members:**
- At least one member is required
- The first member is typically the creator/owner of the chat
- For one-on-one chats, specify exactly two members
- For group chats, you can specify two or more members
- The `roles` array is optional and defaults to `["owner"]` if not specified

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

### 2. **send-message** - Send a message to an existing chat

**Parameters:**
- `chatId`: (Required) The ID of the chat to send the message to
- `content`: (Required) The content of the message (max 5000 characters)
- `contentType`: (Optional) The content type of the message - 'text' or 'html' (default: 'text')
- `messageMetadata`: (Optional) Additional metadata for the message

**Examples:**

1. **Basic Text Message**
```json
{
  "chatId": "19:meeting_NDQ4M2E4ZWMtYjYyMy00YjA2LWI0Y2ItYmYzY2MxNzNlY2Y4@thread.v2",
  "content": "Hello team! Just checking in with a quick update.",
  "contentType": "text"
}
```

2. **HTML Formatted Message**
```json
{
  "chatId": "19:meeting_NDQ4M2E4ZWMtYjYyMy00YjA2LWI0Y2ItYmYzY2MxNzNlY2Y4@thread.v2",
  "content": "<h1>Meeting Reminder</h1><p>Don't forget about our <strong>team meeting</strong> today at <span style='color:blue'>2:00 PM</span>.</p><p>Agenda:</p><ul><li>Project updates</li><li>Q2 Planning</li><li>Team feedback</li></ul>",
  "contentType": "html"
}
```

3. **Message with Metadata**
```json
{
  "chatId": "19:meeting_NDQ4M2E4ZWMtYjYyMy00YjA2LWI0Y2ItYmYzY2MxNzNlY2Y4@thread.v2",
  "content": "Please review the latest project documentation.",
  "contentType": "text",
  "messageMetadata": {
    "priority": "high",
    "tags": ["important", "documentation"],
    "mentions": ["user1@example.com", "user2@example.com"],
    "source": "automated-notification",
    "referenceId": "doc-12345"
  }
}
```

4. **Message with Rich Card**
```json
{
  "chatId": "19:meeting_NDQ4M2E4ZWMtYjYyMy00YjA2LWI0Y2ItYmYzY2MxNzNlY2Y4@thread.v2",
  "content": "Here's the latest project update",
  "contentType": "html",
  "messageMetadata": {
    "card": {
      "title": "Project Update: Q2 2025",
      "subtitle": "Milestone Achievements",
      "text": "We've successfully completed 85% of our Q2 goals!",
      "images": ["https://example.com/project-update.png"],
      "buttons": [
        {
          "type": "openUrl",
          "title": "View Dashboard",
          "value": "https://example.com/dashboard"
        }
      ]
    }
  }
}
```

**Notes:**
- `contentType: 'text'` is best for plain text messages
- `contentType: 'html'` allows for rich text formatting
- `messageMetadata` can be used to attach additional context or functionality to messages
- The total message size (including metadata) must not exceed 28KB

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
