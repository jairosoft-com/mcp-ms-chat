import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { 
  createChatSchema, 
  listChatsSchema, 
  type CreateChatInput, 
  type ListChatsInput 
} from '../schemas/chatSchemas.js';
import {
  type ChatMessage
} from '../interfaces/chat.js';
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
  // List Chats Tool
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
          expand: ['members']
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

        console.error('Fetched chats:', response.value);

        // Format the response for the MCP inspector
        const formattedChats = response.value.map(chat =>
           ({
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
            email: member.email || '',
            userPrincipalName: member.userPrincipalName || '',
            roles: member.roles || []
          })) || []
        }));

        // console.error('Formatted chats:', formattedChats);

        // Format member list for display
        const formatMemberList = (members: any[] = []) => {
          if (!members.length) return 'No members';
          return members.map(m => 
            `${m.displayName || 'Unknown'}${m.email ? ` (${m.email})` : ''}${m.roles?.length ? ` [${m.roles.join(', ')}]` : ''}`
          ).join('\n  ');
        };

        // Create a tabulated view of chats with expandable member details
        const formatChatTable = (chats: any[]) => {
          if (chats.length === 0) return 'No chats found';
          
          // Define column widths
          const columns = {
            index: 8,       // [1]  
            topic: 30,      // Team Collaboration Channel
            type: 15,       // group
            members: 8,     // 5
            lastMessage: 40  // John: Let's discuss the...
          };
          
          // Create header
          const header = [
            'ID'.padEnd(columns.index),
            'TOPIC'.padEnd(columns.topic),
            'TYPE'.padEnd(columns.type),
            'MEMBERS'.padStart(columns.members),
            'LAST MESSAGE'.padEnd(columns.lastMessage)
          ].join(' | ');
          
          // Create separator
          const separator = '='.repeat(header.length);
          
          // Create rows with expandable member details
          const rows: string[] = [];
          
          chats.forEach((chat, index) => {
            const lastMessage = chat.lastMessage 
              ? `${chat.lastMessage.from}: ${chat.lastMessage.content.substring(0, 38)}${chat.lastMessage.content.length > 38 ? '...' : ''}`
              : 'No messages';
            
            // Main row
            rows.push([
              `[${index + 1}]`.padEnd(columns.index),
              (chat.topic || 'No topic').substring(0, columns.topic - 3).padEnd(columns.topic) + (chat.topic && chat.topic.length > columns.topic - 3 ? '...' : ''),
              chat.type.padEnd(columns.type),
              String(chat.memberCount).padStart(columns.members),
              lastMessage.padEnd(columns.lastMessage)
            ].join(' | '));
            
            // Member details row (collapsed by default)
            if (chat.members?.length) {
              const memberDetails = formatMemberList(chat.members);
              rows.push(`  ${' '.repeat(columns.index + columns.topic + columns.type + 6)}Members: ${memberDetails}`);
            }
            
            // Add a thin separator between chats
            rows.push('-'.repeat(header.length));
          });
          
          // Remove the last separator if it exists
          if (rows[rows.length - 1] === '-'.repeat(header.length)) {
            rows.pop();
          }
          
          return [
            `Found ${chats.length} chats (${response.value.length} total):\n`,
            separator,
            header,
            separator,
            ...rows
          ].join('\n');
        };
        
        return {
          content: [{
            type: 'text',
            text: formatChatTable(formattedChats),
            _meta: {
              chats: formattedChats,
              count: response.value.length,
              hasMore: !!response['@odata.nextLink'],
              nextLink: response['@odata.nextLink']
            }
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

  // List Messages Tool
  server.tool(
    'list-messages',
    'List messages in a specific chat with advanced filtering options',
    {
      accessToken: z.string().describe('Microsoft Graph API access token'),
      chatId: z.string().describe('ID of the chat to list messages from'),
      from: z.string().optional().describe('Filter messages from a specific user (email or ID). Use "me" for current user'),
      isRead: z.boolean().optional().describe('Filter read/unread messages'),
      importance: z.enum(['low', 'normal', 'high']).optional().describe('Filter by importance level'),
      contains: z.string().optional().describe('Filter messages containing specific text'),
      afterDateTime: z.string().optional().describe('Filter messages after this date (ISO 8601 format)'),
      beforeDateTime: z.string().optional().describe('Filter messages before this date (ISO 8601 format)'),
      top: z.number().int().min(1).max(1000).optional().default(50).describe('Number of messages to return (1-1000)'),
      skip: z.number().int().min(0).optional().default(0).describe('Number of messages to skip'),
      orderBy: z.string().optional().default('createdDateTime desc').describe('Sort order (e.g., "createdDateTime desc")'),
    },
    async (args) => {
      try {
        const chatService = new ChatService({
          accessToken: args.accessToken
        });

        const messages = await chatService.getChatMessages(args.chatId, {
          from: args.from,
          isRead: args.isRead,
          importance: args.importance,
          contains: args.contains,
          afterDateTime: args.afterDateTime,
          beforeDateTime: args.beforeDateTime,
          top: args.top,
          skip: args.skip,
          orderBy: args.orderBy
        });

        // Format messages for display
        const formattedMessages = messages.value.map((msg, index) => ({
          id: msg.id,
          index: index + 1,
          from: msg.from?.user?.displayName || 'Unknown',
          content: msg.body?.content || '',
          created: msg.createdDateTime,
          isRead: msg.isRead,
          importance: msg.importance || 'normal',
          chatId: args.chatId
        }));

        // Create a tabulated view of messages
        const formatMessageTable = (msgs: typeof formattedMessages) => {
          if (msgs.length === 0) return 'No messages found';
          
          // Define column widths
          const columns = {
            index: 8,       // [1]  
            from: 25,       // John Doe
            content: 50,    // Message preview...
            date: 20,       // 2023-01-01 12:00 PM
            status: 10      // [Unread]
          };
          
          // Create header
          const header = [
            '#'.padEnd(columns.index),
            'FROM'.padEnd(columns.from),
            'MESSAGE'.padEnd(columns.content),
            'DATE'.padEnd(columns.date),
            'STATUS'.padEnd(columns.status)
          ].join(' | ');
          
          // Create separator
          const separator = '='.repeat(header.length);
          
          // Create rows
          const rows = msgs.map(msg => [
            `[${msg.index}]`.padEnd(columns.index),
            (msg.from || 'Unknown').substring(0, columns.from - 3).padEnd(columns.from) + (msg.from && msg.from.length > columns.from - 3 ? '...' : ''),
            (msg.content || '').substring(0, columns.content - 3).replace(/\n/g, ' ').padEnd(columns.content) + (msg.content && msg.content.length > columns.content - 3 ? '...' : ''),
            new Date(msg.created).toLocaleString().padEnd(columns.date),
            (msg.isRead ? '' : '[Unread]').padEnd(columns.status)
          ].join(' | '));
          
          return [
            `Showing ${msgs.length} messages:\n`,
            separator,
            header,
            separator,
            ...rows
          ].join('\n');
        };

        return {
          content: [{
            type: 'text',
            text: formatMessageTable(formattedMessages),
            _meta: {
              messages: formattedMessages,
              count: messages.value.length,
              hasMore: !!messages['@odata.nextLink'],
              nextLink: messages['@odata.nextLink']
            }
          }]
        };
      } catch (error) {
        console.error('Error listing messages:', error);
        throw new Error(`Failed to list messages: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    }
  );

  // Define interface for the formatted chat result
  interface FormattedChatResult {
    index: number;
    id: string;
    topic: string;
    type: string;
    created: string;
    lastUpdated: string;
    webUrl: string;
  }

  // Search Chats by Name or Topic Tool
  const searchChatsSchema = z.object({
    accessToken: z.string().describe('Microsoft Graph API access token'),
    searchTerm: z.string().describe('Search term to match against chat names or topics'),
    top: z.number().int().min(1).max(100).optional().default(20).describe('Maximum number of results to return (1-100)'),
    skip: z.number().int().min(0).optional().default(0).describe('Number of results to skip for pagination')
  });

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
        console.error("chat: ", chat)
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
    console.log("members: ", chat)
    const memberList = chat.members
      .map(member => `- ${member.displayName || member.id || 'Unknown user'}`)
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
