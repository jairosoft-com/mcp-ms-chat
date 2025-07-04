# Microsoft Graph Chat API - MCP Server Implementation

## Overview
This document outlines the essential chat features for a personal assistant using Microsoft Graph Chat API as an MCP server. The implementation focuses on core chat functionality with delegated permissions and direct input-based authentication.

## Required Azure AD Configuration
Users will need to provide the following information, which can be obtained from their Azure AD admin:

1. **Tenant ID** - The unique identifier of the Azure AD tenant
2. **Client ID** - The application (client) ID registered in Azure AD
3. **Client Secret** - A client secret for the registered application
4. **Access Token** - (Optional) Pre-authenticated access token
5. **User Principal Name (UPN)** - The email of the user performing the actions

## Core Features

### 1. List Chats
**Permission:** Chat.ReadBasic
**Description:** List all available chats with basic information
**Inputs:**
- Azure credentials
- Optional: Filter parameters (unread, date range, etc.)

### 2. Read Chat Messages
**Permission:** Chat.Read
**Description:** Read messages from a specific chat
**Inputs:**
- Azure credentials
- Chat ID
- Optional: Message count, pagination

### 3. Send Message
**Permission:** ChatMessage.Send
**Description:** Send a message to a chat
**Inputs:**
- Azure credentials
- Chat ID
- Message content
- Optional: Mentions, attachments

### 4. Create Chat
**Permission:** Chat.Create
**Description:** Create a new chat (1:1 or group)
**Inputs:**
- Azure credentials
- Participant emails
- Optional: Chat topic, initial message

### 5. Delete Chat
**Permission:** Chat.ManageDeletion.All
**Description:** Delete a chat
**Inputs:**
- Azure credentials
- Chat ID

### 6. Search Messages
**Permission:** Chat.Read
**Description:** Search across chat messages
**Inputs:**
- Azure credentials
- Search query
- Optional: Date range, chat filter

## Authentication

### Input Schema
```typescript
interface ChatRequest {
  // Required credentials (either accessToken or all credentials)
  accessToken?: string;
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  userPrincipalName: string;
  
  // Feature-specific inputs
  chatId?: string;
  message?: string;
  participants?: string[];
  // ... other operation-specific fields
}
```

### Authentication Flow
1. Accept Azure AD credentials as direct inputs
2. If accessToken is provided, use it directly
3. Otherwise, use client credentials to obtain a token
4. Validate all required permissions are present
5. Execute the requested chat operation

## Error Handling

All errors will be thrown using `console.error` with the following structure:
```typescript
{
  error: string;           // Short error code
  message: string;        // User-friendly message
  details?: any;          // Additional error details
  remediation?: string;   // Suggested remediation steps
}
```

## Docker Deployment

The MCP server will be containerized using a Dockerfile based on the reference implementation:
```dockerfile
FROM node:18.19.1-alpine3.19

WORKDIR /app
COPY package*.json ./
RUN npm ci --only=production --legacy-peer-deps
COPY build/ ./build/

# Make the entry point executable
RUN chmod +x ./build/index.js

# Set the entry point
ENTRYPOINT ["node", "./build/index.js"]
```

## Next Steps

1. Implement the authentication service with input validation
2. Create MCP schemas for each chat operation
3. Implement chat service methods with proper error handling
4. Add comprehensive logging
5. Set up Docker build and test pipeline
- Tenant ID
- Client ID
- Client Secret
- Redirect URI
- User Principal Name (UPN)
- Chat ID (for deletion/recovery)

### 3. Read Chat Messages (Chat.Read)
**Permission Scope:** Delegated
**Description:** Read user chat messages
**Required Inputs:**
- Tenant ID
- Client ID
- Client Secret
- Redirect URI
- User Principal Name (UPN)
- Chat ID

### 4. Read Basic Chat Information (Chat.ReadBasic)
**Permission Scope:** Delegated
**Description:** Read names and members of user chat threads
**Required Inputs:**
- Tenant ID
- Client ID
- Client Secret
- Redirect URI
- User Principal Name (UPN)

### 5. Read and Write Chat Messages (Chat.ReadWrite)
**Permission Scope:** Delegated
**Description:** Read and write user chat messages
**Required Inputs:**
- Tenant ID
- Client ID
- Client Secret
- Redirect URI
- User Principal Name (UPN)
- Chat ID
- Message content (for writing)

### 6. Full Chat Management (Chat.ReadWrite.All)
**Permission Scope:** Delegated
**Description:** Read and write all chat messages
**Required Inputs:**
- Tenant ID
- Client ID
- Client Secret
- Redirect URI
- User Principal Name (UPN)
- Chat ID (for specific operations)
- Operation-specific parameters

## User Flow

1. **Initialization**
   - User provides Azure AD configuration
   - System validates the configuration
   - If validation fails, display helpful error message with instructions

2. **Authentication**
   - System initiates OAuth 2.0 authorization code flow
   - User authenticates with Microsoft
   - System receives and stores access token

3. **Feature Execution**
   - User selects desired chat operation
   - System performs the operation using the stored token
   - Results are displayed to the user

## Error Handling

For each operation, the system will:
1. Validate all required inputs
2. Check for valid authentication
3. Handle Microsoft Graph API errors gracefully
4. Provide clear, actionable error messages

## Implementation Questions

Before proceeding with implementation, I need the following information:

1. **User Interface**
   - Would you prefer a web-based UI or a command-line interface?
   - Should we implement a wizard for initial Azure AD configuration?

2. **Authentication**
   - Should we support token caching to avoid frequent logins?
   - What should be the token refresh strategy?

3. **Error Handling**
   - Should we implement retry logic for failed API calls?
   - What level of error detail should be shown to end users?

4. **Logging**
   - What level of logging is required?
   - Should logs include any sensitive information?

5. **Deployment**
   - Will this be deployed as a standalone application or integrated into an existing system?
   - What are the target platforms (Windows, Linux, macOS)?

## Next Steps

1. Review and provide feedback on this PRD
2. Answer the implementation questions above
3. Approve the scope and requirements
4. Begin implementation of the first feature

Would you like me to proceed with implementing any specific feature first, or would you like to review and refine this PRD further?
