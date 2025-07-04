import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createChatSchema } from '../schemas/chatSchemas.js';
import { createChat } from '../services/chatService.js';
import type { Chat } from '../interfaces/chat.js';

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
    async (args: unknown) => {
      try {
        // The schema validation is handled by the MCP server
        const chatData = args as Parameters<typeof createChat>[0];
        
        // Call the chat service to create the chat
        const chat = await createChat(chatData);
        
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
        
        return {
          content: [{
            type: 'text',
            text: `❌ Failed to create chat: ${errorMessage}`,
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
