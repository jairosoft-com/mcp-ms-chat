import { Client } from "@microsoft/microsoft-graph-client";
import { getAzureCredentials } from "./authService.js";
import { 
  Chat, 
  ChatMessage, 
  ChatMessageCollectionResponse,
  ChatCollectionResponse,
  ListChatsOptions
} from "../interfaces/chat.js";

// Initialize Microsoft Graph client with Azure AD credentials
function getAuthenticatedClient() {
  const credentials = getAzureCredentials();
  
  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const tokenResponse = await credentials.getToken("https://graph.microsoft.com/.default");
        return tokenResponse.token;
      }
    }
  });

  return client;
}

/**
 * Create a new chat
 * @param chatData Chat data including members and optional message
 * @returns The created chat
 */
export async function createChat(chatData: {
  topic: string;
  chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
  members: Array<{ id: string; roles?: string[] }>;
}): Promise<Chat> {
  try {
    const client = getAuthenticatedClient();
    
    // Format members for the API
    const members = chatData.members.map(member => ({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${member.id}`,
      roles: member.roles || ['owner']
    }));
    
    // Create the chat
    const chatResponse = await client
      .api('/chats')
      .post({
        chatType: chatData.chatType,
        topic: chatData.topic,
        members
      });
    
    // Try to get the chat details, but fall back to basic info if we don't have permission
    try {
      return await getChat(chatResponse.id);
    } catch (getChatError) {
      console.warn('Warning: Could not fetch full chat details. You may need additional permissions like Chat.Read.All.', getChatError);
      // Return minimal chat info with just the ID and basic details
      return {
        id: chatResponse.id,
        topic: chatData.topic || '',
        chatType: chatData.chatType,
        webUrl: '',
        createdDateTime: new Date().toISOString(),
        lastUpdatedDateTime: new Date().toISOString()
      } as Chat;
    }
    
    return chatResponse as Chat;
  } catch (error) {
    console.error('Error creating chat:', error);
    throw new Error(`Failed to create chat: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Send a message to a chat
 * @param chatId The ID of the chat
 * @param messageData The message data
 * @returns The sent message
 */
export async function sendChatMessage(
  chatId: string,
  messageData: {
    content: string;
    contentType?: 'text' | 'html' | 'content';
    messageMetadata?: Record<string, any>;
  }
): Promise<ChatMessage> {
  try {
    const client = getAuthenticatedClient();
    
    const message = {
      body: {
        content: messageData.content,
        contentType: messageData.contentType || 'text',
      },
      ...(messageData.messageMetadata && { messageMetadata: messageData.messageMetadata })
    };
    
    const response = await client
      .api(`/chats/${chatId}/messages`)
      .post(message);
      
    return response as ChatMessage;
  } catch (error) {
    console.error('Error sending chat message:', error);
    throw new Error(`Failed to send message: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Get a chat by ID
 * @param chatId The ID of the chat to retrieve
 * @returns The requested chat
 */
export async function getChat(chatId: string): Promise<Chat> {
  try {
    const client = getAuthenticatedClient();
    
    const response = await client
      .api(`/chats/${chatId}`)
      .expand('members')
      .get();
      
    return response as Chat;
  } catch (error) {
    console.error('Error getting chat:', error);
    throw new Error(`Failed to get chat: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * List chats with optional filtering and pagination
 * @param options Query options
 * @returns Paginated list of chats
 */
export async function listChats(
  options: ListChatsOptions = {}
): Promise<ChatCollectionResponse> {
  try {
    const { top, skip, filter, orderby, select, expand } = options;
    const client = getAuthenticatedClient();
    
    let request = client.api('/me/chats');
    
    // Apply query parameters if provided
    if (top) request = request.top(top);
    if (skip) request = request.skip(skip);
    if (filter) request = request.filter(filter);
    if (orderby) request = request.orderby(orderby);
    if (select) request = request.select(select);
    if (expand) request = request.expand(expand);
    
    const response = await request.get();
    return response as ChatCollectionResponse;
  } catch (error) {
    console.error('Error listing chats:', error);
    throw new Error(`Failed to list chats: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Get messages from a chat
 * @param chatId The ID of the chat
 * @param top Maximum number of messages to return
 * @returns Collection of chat messages
 */
export async function getChatMessages(
  chatId: string,
  top: number = 50
): Promise<ChatMessageCollectionResponse> {
  try {
    const client = getAuthenticatedClient();
    
    const response = await client
      .api(`/chats/${chatId}/messages`)
      .top(top)
      .orderby('createdDateTime desc')
      .get();
      
    return response as ChatMessageCollectionResponse;
  } catch (error) {
    console.error('Error getting chat messages:', error);
    throw new Error(`Failed to get chat messages: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

export default {
  createChat,
  sendChatMessage,
  getChat,
  listChats,
  getChatMessages,
};
