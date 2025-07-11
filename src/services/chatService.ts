import { Client } from "@microsoft/microsoft-graph-client";
import { AuthConfig } from "./authService.js";
import { AuthenticationError } from "../errors/authError.js";
import { 
  Chat, 
  ChatCollectionResponse, 
  ChatMember, 
  ChatMessage, 
  ChatMessageCollectionResponse, 
  ListChatsOptions, 
  ListMessagesOptions
} from "../interfaces/chat.js";

/**
 * Interface for chat service configuration
 */
interface ChatServiceConfig extends AuthConfig {
  /**
   * Optional user principal name for context
   */
  userPrincipalName?: string;
}

/**
 * ChatService class for handling Microsoft Teams chat operations
 */
class ChatService {
  private config: Omit<ChatServiceConfig, 'accessToken'> & { accessToken: string };
  private tokenExpiry: number = 0;

  constructor(config: ChatServiceConfig) {
    if (!config.accessToken) {
      throw new Error('accessToken is required in ChatService configuration');
    }
    // We know accessToken is defined here due to the check above
    this.config = config as ChatServiceConfig & { accessToken: string };
  }

  /**
   * Gets an authenticated Microsoft Graph client
   * @private
   */
  private async getAuthenticatedClient(): Promise<Client> {
    try {
      // Create the client with the access token
      // We know accessToken is defined because it's required in the constructor
      return Client.init({
        authProvider: (done) => {
          done(null, this.config.accessToken);
        }
      });
    } catch (error: unknown) {
      if (error instanceof AuthenticationError) {
        error.log();
        throw error;
      }
      
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      const helpMessage = `
We encountered an issue initializing the Microsoft Graph client. This might be due to:

• Network connectivity issues
• Invalid or expired authentication credentials
• Missing required API permissions
• Service availability issues

Please try the following:
1. Verify your internet connection
2. Check that all authentication credentials are correct and up to date
3. Ensure your Azure AD application has the required API permissions:
   - Chat.ReadWrite
   - Chat.Create
   - User.Read
4. Try again in a few minutes if this is a temporary issue

If the problem persists, please contact your system administrator with the error details.
`;

      const authError = new AuthenticationError(
        'GRAPH_CLIENT_ERROR',
        'Microsoft Graph Connection Failed',
        helpMessage,
        { 
          originalError: errorMessage,
          documentation: 'https://learn.microsoft.com/en-us/graph/sdks/choose-authentication-providers'
        }
      );
      authError.log();
      throw authError;
    }
  }

  /**
   * Create a new chat
   * @param chatData Chat data including members and optional message
   * @returns The created chat
   */
  public async createChat(chatData: {
    topic: string;
    chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
    members: ChatMember[];
  }): Promise<Chat> {
    try {
      const client = await this.getAuthenticatedClient();
      
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
        return await this.getChat(chatResponse.id);
      } catch (getChatError) {
        console.warn('Warning: Could not fetch full chat details. You may need additional permissions like Chat.ReadAll.', getChatError);
        // Return minimal chat info with just the ID and basic details
        return {
          id: chatResponse.id,
          topic: chatData.topic,
          chatType: chatData.chatType,
          createdDateTime: new Date().toISOString(),
          lastUpdatedDateTime: new Date().toISOString()
        };
      }
    } catch (error) {
      if (error instanceof AuthenticationError) {
        throw error;
      }
      
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
  public async sendMessage(
    chatId: string,
    messageData: {
      content: string;
      contentType?: 'text' | 'html' | 'content';
      messageMetadata?: Record<string, any>;
    }
  ): Promise<ChatMessage> {
    try {
      const client = await this.getAuthenticatedClient();
      
      const response = await client
        .api(`/chats/${chatId}/messages`)
        .post({
          body: {
            content: messageData.content,
            contentType: messageData.contentType || 'text',
            ...(messageData.messageMetadata && { messageMetadata: messageData.messageMetadata })
          }
        });

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
  public async getChat(chatId: string): Promise<Chat> {
    try {
      const client = await this.getAuthenticatedClient();
      
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
  public async listChats(
    options: ListChatsOptions = {}
  ): Promise<ChatCollectionResponse> {
    try {
      const client = await this.getAuthenticatedClient();
      
      let request = client.api('/me/chats');
      
      // Apply pagination
      if (options.top) {
        request = request.top(options.top);
      }

      if (options.skip) {
        request = request.top(options.skip);
      }
            
      // Apply filters if provided
      if (options.filter) {
        request = request.filter(options.filter);
      }
      
      // Apply ordering if provided
      if (options.orderby) {
        request = request.orderby(options.orderby);
      }

      // Apply select if provided
      if (options.select) {
        request = request.select(options.select);
      }

      // Apply expand if provided
      if (options.expand) {
        request = request.expand(options.expand);
      }
      
      console.error("request: ", request);
      const response = await request.get();
      return response as ChatCollectionResponse;
    } catch (error) {
      console.error('Error listing chats:', error);
      throw new Error(`Failed to list chats: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Get messages from a chat with advanced filtering and pagination
   * @param chatId The ID of the chat
   * @param options Message listing options
   * @returns Collection of chat messages
   */
  public async getChatMessages(
    chatId: string,
    options: ListMessagesOptions = {}
  ): Promise<ChatMessageCollectionResponse> {
    try {
      const client = await this.getAuthenticatedClient();
      
      // Start building the request
      let request = client
        .api(`/chats/${chatId}/messages`)
        .header('Prefer', 'outlook.timezone="UTC"');
      
      // Apply pagination
      if (options.top) {
        request = request.top(Math.min(options.top, 1000)); // Enforce max page size
      } else {
        request = request.top(50); // Default page size
      }
      
      if (options.skip) {
        request = request.skip(options.skip);
      }
      
      // Build filter string (excluding isRead which we'll handle client-side)
      const filterParts: string[] = [];
      
      if (options.from) {
        const fromValue = options.from === 'me' 
          ? this.config.userPrincipalName || 'me' 
          : options.from;
        filterParts.push(`from/emailAddress/address eq '${encodeURIComponent(fromValue)}'`);
      }
      
      if (options.importance) {
        filterParts.push(`importance eq '${options.importance}'`);
      }
      
      if (options.afterDateTime) {
        filterParts.push(`createdDateTime ge ${new Date(options.afterDateTime).toISOString()}`);
      }
      
      if (options.beforeDateTime) {
        filterParts.push(`createdDateTime le ${new Date(options.beforeDateTime).toISOString()}`);
      }
      
      if (options.contains) {
        filterParts.push(`contains(body/content, '${options.contains.replace(/'/g, "''")}')`);
      }
      
      // Apply filter if we have any conditions
      if (filterParts.length > 0) {
        request = request.filter(filterParts.join(' and '));
      }
      
      // Apply sorting
      if (options.orderBy) {
        request = request.orderby(options.orderBy);
      } else {
        // Default sort order
        request = request.orderby('createdDateTime desc');
      }
      
      // Apply select fields if specified
      if (options.select && options.select.length > 0) {
        request = request.select(options.select);
      }
      
      // Apply expand if specified
      if (options.expand && options.expand.length > 0) {
        request = request.expand(options.expand);
      }
      
      // Execute the request
      const response = await request.get();
      
      // Handle isRead filtering client-side if needed
      if (options.isRead !== undefined) {
        response.value = response.value.filter((msg: ChatMessage) => 
          options.isRead ? msg.isRead : !msg.isRead
        );
        
        // Update the @odata.count to reflect the filtered count
        if (response['@odata.count'] !== undefined) {
          response['@odata.count'] = response.value.length;
        }
      }
      
      return response as ChatMessageCollectionResponse;
    } catch (error) {
      console.error('Error getting chat messages:', error);
      throw new Error(`Failed to get chat messages: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }


  
  /**
   * Get recent messages from all chats
   * @param options Message listing options
   * @returns Collection of recent messages
   */
  public async getRecentMessages(
    options: Omit<ListMessagesOptions, 'afterDateTime' | 'beforeDateTime'> & { 
      days?: number 
    } = {}
  ): Promise<ChatMessageCollectionResponse> {
    try {
      // Default to last 7 days if not specified
      const days = options.days || 7;
      const afterDateTime = new Date();
      afterDateTime.setDate(afterDateTime.getDate() - days);
      
      // Define an extended message type that includes chat information
      interface ExtendedChatMessage extends ChatMessage {
        chatId: string;
        chatTopic: string;
        chatType: string;
      }
      
      // Get all chats
      const chats = await this.listChats({ top: 100 });
      
      // Get recent messages from each chat
      const allMessages: ExtendedChatMessage[] = [];
      
      for (const chat of chats.value) {
        try {
          const messages = await this.getChatMessages(chat.id, {
            ...options,
            afterDateTime: afterDateTime.toISOString(),
            top: options.top || 20 // Limit per chat to avoid too many requests
          });
          
          // Add chat info to each message with proper typing
          const messagesWithChatInfo: ExtendedChatMessage[] = messages.value.map(msg => ({
            ...msg,
            chatId: chat.id,
            chatTopic: chat.topic || 'No Topic',
            chatType: chat.chatType || 'unknown'
          }));
          
          allMessages.push(...messagesWithChatInfo);
        } catch (error) {
          console.warn(`Error getting messages for chat ${chat.id}:`, error);
          // Continue with next chat
        }
      }
      
      // Sort all messages by creation time (newest first)
      allMessages.sort((a, b) => 
        new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime()
      );
      
      // Apply pagination
      const skip = options.skip || 0;
      const top = options.top || 50;
      const paginatedMessages = allMessages.slice(skip, skip + top);
      
      return {
        value: paginatedMessages,
        '@odata.nextLink': allMessages.length > skip + top 
          ? `skip=${skip + top}&top=${top}` 
          : undefined
      };
    } catch (error) {
      console.error('Error getting recent messages:', error);
      throw new Error(`Failed to get recent messages: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Delete a chat
   * @param chatId The ID of the chat to delete
   */
  public async deleteChat(chatId: string): Promise<void> {
    try {
      const client = await this.getAuthenticatedClient();
      await client.api(`/chats/${chatId}`).delete();
    } catch (error) {
      console.error('Error deleting chat:', error);
      throw new Error(`Failed to delete chat: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
}

export default ChatService;
