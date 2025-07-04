import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createChatSchema, listChatsSchema, type CreateChatInput, type ListChatsInput } from '../schemas/chatSchemas.js';
import type { ChatMember, Chat } from '../interfaces/chat.js';
import ChatService from '../services/chatService.js';
import { z } from 'zod';

// Type for the chat creation response
interface ChatCreationResponse {
  content: Array<{ type: string; text: string }>;
  metadata?: {
    chatId: string;
    webUrl?: string;
  };
}

/**
 * Registers the chat tool with the MCP server
 * @param server - The MCP server instance
 */
export function registerChatTools(server: McpServer): void {
  // Register the List Chats tool with the server
  server.tool(
    'list-chats',
    'List all available chats in Microsoft Teams with basic information',
    listChatsSchema.shape,
    /**
     * @example
     * // Basic usage (just show all chats with messages)
     * {
     *   "accessToken": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
     * }
     * 
     * @example
     * // Find chats with a specific participant
     * {
     *   "accessToken": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...",
     *   "filter": "members/any(m:m/userPrincipalName eq 'sgeraldino@jairosoft.com')",
     *   "expand": ["members($select=id,displayName,userPrincipalName)"]
     * }
     * 
     * @example
     * // Find unread group chats, sorted by most recent
     * {
     *   "accessToken": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...",
     *   "filter": "chatType eq 'group' and lastMessagePreview/isRead eq false",
     *   "orderby": ["lastMessagePreview/createdDateTime desc"]
     * }
     */
    async (args: z.infer<typeof listChatsSchema>) => {
      // Log the incoming request (without sensitive data)
      const safeArgs = args ? { ...args as Record<string, unknown> } : {};
      if ('accessToken' in safeArgs && safeArgs.accessToken) safeArgs.accessToken = '***REDACTED***';

      console.log('Listing chats with parameters:', JSON.stringify(safeArgs, null, 2));

      try {
        // The schema validation is handled by the MCP server
        const listOptions = args as ListChatsInput;

        // Default values for optional parameters
        const defaultOptions = {
          top: 50, // Default to 50 chats per page
          skip: 0,  // Start from the beginning
          filter: undefined, // Temporarily removing filter to see all chats
          orderby: ['lastMessagePreview/createdDateTime desc'], // Sort by most recent message
          select: [
            'id',
            'topic',
            'chatType',
            'createdDateTime',
            'lastMessagePreview',
            'webUrl',
            'createdDateTime',
            'lastUpdatedDateTime'
          ],
          expand: ['members($select=id,displayName,userPrincipalName)']
        };

        // Merge provided options with defaults
        const options = {
          ...defaultOptions,
          ...listOptions,
          // Handle array options separately to avoid overwriting with undefined
          orderby: listOptions.orderby?.length ? listOptions.orderby : defaultOptions.orderby,
          select: listOptions.select?.length ? listOptions.select : defaultOptions.select,
          expand: listOptions.expand?.length ? listOptions.expand : defaultOptions.expand
        };

        // Initialize the chat service with the provided credentials
        console.log('Initializing chat service...');
        const chatService = new ChatService({
          accessToken: listOptions.accessToken
        });

        console.log('Fetching chats with options:', {
          ...options,
          accessToken: '***REDACTED***' // Don't log the actual token
        });
        
        // Log token details for debugging (without exposing the full token)
        if (listOptions.accessToken) {
          try {
            const tokenParts = listOptions.accessToken.split('.');
            if (tokenParts.length === 3) {
              const payload = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
              console.log('Token details:', {
                scopes: payload.scp || payload.scopes || 'no scopes found',
                appId: payload.appid || 'no app id',
                upn: payload.upn || 'no upn',
                tenantId: payload.tid || 'no tenant id'
              });
            }
          } catch (e) {
            console.log('Could not parse token for debugging');
          }
        }

        const response = await chatService.listChats({
          top: options.top,
          skip: options.skip,
          filter: options.filter,
          orderby: Array.isArray(options.orderby) ? options.orderby.join(',') : options.orderby,
          select: options.select,
          expand: options.expand
        });

        // Format the response for the MCP inspector
        const formattedChats = response.value.map(chat => ({
          id: chat.id,
          topic: chat.topic || 'No topic',
          type: chat.chatType || 'unknown',
          created: chat.createdDateTime,
          lastUpdated: (chat as any).lastMessagePreview?.createdDateTime ?? chat.createdDateTime,
          webUrl: chat.webUrl,
          memberCount: chat.members?.length || 0,
          lastMessage: chat.lastMessagePreview ? {
            from: chat.lastMessagePreview.from?.user?.displayName || 'Unknown',
            content: chat.lastMessagePreview.body?.content || 'No content',
            created: chat.lastMessagePreview.createdDateTime
          } : null,
          members: chat.members?.map(member => ({
            id: member.id,
            displayName: member.displayName || 'Unknown',
            userPrincipalName: member.userPrincipalName || '',
            roles: member.roles || []
          })) || []
        }));

        return {
          content: [{
            type: 'text',
            text: `Found ${response.value.length} chats`
          }],
          metadata: {
            chats: formattedChats,
            count: response.value.length,
            nextLink: response['@odata.nextLink']
          }
        };
      } catch (error: unknown) {
        // Handle authentication errors specifically
        if (error instanceof Error && error.message.includes('AuthenticationError')) {
          const authError = {
            error: 'AUTHENTICATION_FAILED',
            message: 'Authentication failed when trying to list chats',
            details: error.message,
            remediation: 'Please ensure your access token is valid and has the Chat.ReadBasic permission.'
          };
          console.error('Authentication error listing chats:', authError);
          throw new Error(JSON.stringify(authError));
        }

        // Handle permission errors
        if (error instanceof Error && error.message.includes('permission')) {
          const permissionError = {
            error: 'INSUFFICIENT_PERMISSIONS',
            message: 'Insufficient permissions to list chats',
            details: error.message,
            remediation: 'Please ensure your access token includes the Chat.ReadBasic permission.'
          };
          console.error('Permission error listing chats:', permissionError);
          throw new Error(JSON.stringify(permissionError));
        }

        // Handle rate limiting
        if (error instanceof Error && 
            (error.message.includes('429') || error.message.includes('Too Many Requests'))) {
          const rateLimitError = {
            error: 'RATE_LIMIT_EXCEEDED',
            message: 'Too many requests to the Microsoft Graph API',
            details: error.message,
            remediation: 'Please wait a few minutes before trying again.'
          };
          console.error('Rate limit exceeded when listing chats:', rateLimitError);
          throw new Error(JSON.stringify(rateLimitError));
        }

        // Handle other errors
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        const stack = error instanceof Error ? error.stack : undefined;
        
        const genericError = {
          error: 'LIST_CHATS_FAILED',
          message: 'Failed to list chats',
          details: errorMessage,
          remediation: 'Please check your network connection and try again. If the problem persists, contact support.'
        };
        
        console.error('Error listing chats:', {
          ...genericError,
          stack
        });
        
        throw new Error(JSON.stringify(genericError));
      }
    }
  );

  // Register the Create Chat tool with the server
  server.tool(
    'create-chat',
    'Create a new chat in Microsoft Teams',
    createChatSchema.shape,
    async (args: z.infer<typeof createChatSchema>) => {
      // Log the incoming request (without sensitive data)
      const safeArgs = args ? { ...args as Record<string, unknown> } : {};
      if ('clientSecret' in safeArgs) safeArgs.clientSecret = '***REDACTED***';
      if ('accessToken' in safeArgs && safeArgs.accessToken) safeArgs.accessToken = '***REDACTED***';

      console.log('Creating chat with parameters:', JSON.stringify(safeArgs, null, 2));

      try {
        // The schema validation is handled by the MCP server
        const chatData = args as CreateChatInput;

        // Extract auth fields
        const { 
          accessToken,
          ...chatPayload
        } = chatData as CreateChatInput & { 
          accessToken: string;
        };

        // Validate required access token
        if (!accessToken) {
          console.error('Validation error: Missing required access token');
          throw new Error('access_token needed');
        }

        // Initialize the chat service with the provided credentials
        console.log('Initializing chat service...');
        const chatService = new ChatService({
          accessToken
        });

        // Prepare chat payload with proper typing
        const chatRequest = {
          topic: typeof chatPayload.topic === 'string' ? chatPayload.topic : 'New Chat',
          chatType: (chatPayload.chatType && 
                    ['oneOnOne', 'group', 'meeting', 'unknown'].includes(chatPayload.chatType as string)) 
                    ? chatPayload.chatType as 'oneOnOne' | 'group' | 'meeting' | 'unknown' 
                    : 'group',
          members: (Array.isArray(chatPayload.members) ? chatPayload.members : []) as ChatMember[]
        };

        console.log('Creating chat with payload:', { 
          topic: chatRequest.topic,
          chatType: chatRequest.chatType,
          memberCount: chatRequest.members.length
        });

        const chat = await chatService.createChat(chatRequest);
        
        console.log('Successfully created chat:', { 
          chatId: chat.id, 
          topic: chat.topic,
          createdDateTime: chat.createdDateTime,
          webUrl: chat.webUrl
        });

        // Format the response
        const response = formatChatResponse(chat);
        
        return {
          content: [{
            type: 'text',
            text: response,
          }],
          metadata: {
            chatId: chat.id,
            webUrl: chat.webUrl
          }
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error creating chat:', {
          error: errorMessage,
          stack: error instanceof Error ? error.stack : undefined
        });
        
        return {
          content: [{
            type: 'text',
            text: `❌ Failed to create chat: ${errorMessage}\n\nPlease check your authentication credentials and try again.`,
          }],
        };
      }
    }
  );
}

/**
 * Formats a chat object into a readable string
 */
function formatChatResponse(chat: Chat): string {
  const parts: string[] = [];
  
  // Add chat information
  parts.push(`# ${chat.topic || 'Untitled Chat'}`);
  
  if (chat.chatType) {
    parts.push(`**Type:** ${formatChatType(chat.chatType)}`);
  }
  
  if (chat.createdDateTime) {
    parts.push(`**Created:** ${new Date(chat.createdDateTime).toLocaleString()}`);
  }
  
  // Add members if available
  if (chat.members && chat.members.length > 0) {
    const memberList = chat.members
      .map(member => `- ${member.displayName || member.userId || 'Unknown user'}`)
      .join('\n');
    
    parts.push('\n**Members:**');
    parts.push(memberList);
  }
  
  // Add web URL if available
  if (chat.webUrl) {
    parts.push(`\n[Open in Teams](${chat.webUrl})`);
  }
  
  // Add chat ID at the end
  parts.push(`\n**Chat ID:** ${chat.id}`);
  
  return parts.join('\n');
}

/**
 * Formats chat type for display
 */
function formatChatType(chatType: string): string {
  switch (chatType) {
    case 'oneOnOne':
      return '1:1 Chat';
    case 'group':
      return 'Group Chat';
    case 'meeting':
      return 'Meeting Chat';
    default:
      return chatType;
  }
}

export default {
  registerChatTools,
};
