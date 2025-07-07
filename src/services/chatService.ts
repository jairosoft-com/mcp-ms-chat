import { Client } from "@microsoft/microsoft-graph-client";
import { AuthConfig } from "./authService.js";
import { AuthenticationError } from "../errors/authError.js";
import { Chat, ChatCollectionResponse, ChatMember, ChatMessage, ChatMessageCollectionResponse, ListChatsOptions } from "../interfaces/chat.js";

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
      
      // Apply filters if provided
      if (options.filter) {
        request = request.filter(options.filter);
      }
      
      // Apply pagination
      if (options.top) {
        request = request.top(options.top);
      }
      
      if (options.skip) {
        request = request.skip(options.skip);
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
   * Get messages from a chat
   * @param chatId The ID of the chat
   * @param top Maximum number of messages to return
   * @returns Collection of chat messages
   */
  public async getChatMessages(
    chatId: string,
    top: number = 50
  ): Promise<ChatMessageCollectionResponse> {
    try {
      const client = await this.getAuthenticatedClient();
      
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
