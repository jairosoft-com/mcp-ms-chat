import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createChatSchema, type CreateChatInput } from '../schemas/chatSchemas.js';
import type { ChatMember } from '../interfaces/chat.js';
import ChatService from '../services/chatService.js';
import type { Chat } from '../interfaces/chat.js';
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
          const errorMsg = 'Missing required access token';
          console.error('Validation error:', { message: errorMsg });
          
          return {
            content: [{
              type: 'text',
              text: `❌ ${errorMsg}. Please provide a valid access token with the required Microsoft Graph API permissions.`,
            }],
          };
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
