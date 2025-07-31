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
          try {
            if (!this.config.accessToken) {
              throw new Error('Access token is missing or invalid');
            }
            // Log the first 10 characters of the token for debugging (don't log the whole token for security)
            console.log('Using access token starting with:', this.config.accessToken.substring(0, 10) + '...');
            done(null, this.config.accessToken);
          } catch (error) {
            console.error('Error in authProvider:', error);
            done(error as Error, null);
          }
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
  /**
   * Extracts tenant ID from JWT token
   */
  private getTenantIdFromToken(token: string): string | null {
    try {
      const base64Url = token.split('.')[1];
      const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
      const payload = JSON.parse(Buffer.from(base64, 'base64').toString());
      return payload.tid || null;
    } catch (error) {
      console.error('Error extracting tenant ID from token:', error);
      return null;
    }
  }

  /**
   * Gets the current user's UPN from the access token
   */
  private async getCurrentUserUpn(): Promise<string> {
    try {
      const client = await this.getAuthenticatedClient();
      const me = await client.api('/me').get();
      return me.userPrincipalName || me.mail || '';
    } catch (error) {
      console.error('Error getting current user info:', error);
      throw new Error('Could not determine current user information');
    }
  }

  public async createChat(chatData: {
    topic: string;
    chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
    members: ChatMember[];
  }): Promise<Chat> {
    try {
      console.log('Creating chat with data:', {
        topic: chatData.topic,
        chatType: chatData.chatType,
        memberCount: chatData.members?.length || 0
      });
  
      if (!chatData.members?.length) {
        throw new Error('At least one member is required to create a chat');
      }
  
      const client = await this.getAuthenticatedClient();
      
      // Get current user's UPN
      const me = await client.api('/me').get();
      const currentUserUpn = me.userPrincipalName;
      if (!currentUserUpn) {
        throw new Error('Could not determine current user information');
      }
      console.log('Current user UPN:', currentUserUpn);
  
      // Format members for the API
      const members = await Promise.all(chatData.members.map(async (member) => {
        if (!member.id) {
          throw new Error('Member ID is required');
        }
  
        // Handle both email and user ID formats
        let userId = member.id;
        if (!userId.includes('@')) {
          // If it's not an email, try to get the user by ID
          try {
            const user = await client.api(`/users/${userId}`).get();
            userId = user.userPrincipalName || userId;
          } catch (error) {
            console.warn(`Could not find user with ID ${userId}, using as-is`);
          }
        }
  
        return {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}`,
          roles: Array.isArray(member.roles) && member.roles.length > 0 ? member.roles : ['guest']
        };
      }));
  
      // Add current user if not already in members
      const currentUserExists = members.some(m => 
        m['user@odata.bind'].includes(encodeURIComponent(currentUserUpn))
      );
  
      if (!currentUserExists) {
        members.unshift({
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(currentUserUpn)}`,
          roles: ['owner']
        });
      }
  
      // Create the chat
      const response = await client.api('/chats').post({
        chatType: chatData.chatType,
        topic: chatData.topic,
        members
      });
  
      if (!response?.id) {
        throw new Error('Invalid response from Microsoft Graph API');
      }
  
      console.log('Chat created successfully with ID:', response.id);
      return {
        id: response.id,
        topic: response.topic || chatData.topic,
        chatType: response.chatType || chatData.chatType,
        createdDateTime: response.createdDateTime || new Date().toISOString(),
        lastUpdatedDateTime: response.lastUpdatedDateTime || new Date().toISOString(),
        webUrl: response.webUrl || '',
        members: response.members || members.map(m => ({
          id: m['user@odata.bind'].split('/').pop() || '',
          roles: m.roles
        }))
      };
  
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      const statusCode = (error as any)?.statusCode;
      const errorCode = (error as any)?.code;
      
      console.error('Error creating chat:', {
        error: errorMessage,
        statusCode,
        code: errorCode,
        timestamp: new Date().toISOString()
      });
  
      // Provide more specific error messages for common scenarios
      if (statusCode === 401) {
        throw new Error('Authentication failed. Please check your access token and permissions.');
      } else if (statusCode === 403) {
        throw new Error('Insufficient permissions. The app needs Chat.Create and Chat.ReadWrite permissions.');
      } else if (statusCode === 404) {
        throw new Error('One or more users could not be found. Please verify the user IDs and try again.');
      } else if (errorMessage.includes('TenantNotFound')) {
        throw new Error('The specified tenant was not found. Please verify your Azure AD tenant configuration.');
      }
  
      throw new Error(`Failed to create chat: ${errorMessage}`);
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
      
      // Apply pagination - Microsoft Graph has a max of 50 items per page
      const top = options.top ? Math.min(options.top, 50) : 50;
      request = request.top(top);

      // Skip is not directly supported in Microsoft Graph v1.0 for /me/chats
      // We'll use it as an offset for server-side pagination if needed
      if (options.skip && options.skip > 0) {
        // Note: This is a simplified approach. For production, you might want to implement
        // proper pagination using @odata.nextLink from the response
        console.warn('Skip parameter may not work as expected with Microsoft Graph API');
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
      
      const response = await request.get();
      
      // Transform the response to match our ChatCollectionResponse type
      const result: ChatCollectionResponse = {
        value: response.value || [],
        '@odata.nextLink': response['@odata.nextLink']
      };
      
      return result;
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
   * Get messages from a chat with pagination support
   * @param chatId The ID of the chat to get messages from
   * @param options Options for pagination and filtering
   * @returns Array of chat messages
   */
  public async getMessages(
    chatId: string,
    options: {
      top?: number;
      skip?: number;
      filter?: string;
      orderBy?: string;
      select?: string[];
      expand?: string[];
    } = {}
  ): Promise<ChatMessage[]> {
    try {
      const client = await this.getAuthenticatedClient();
      
      // Build the query parameters
      const query: Record<string, string | number> = {};
      if (options.top) query.$top = options.top;
      if (options.skip) query.$skip = options.skip;
      if (options.filter) query.$filter = options.filter;
      if (options.orderBy) query.$orderby = options.orderBy;
      if (options.select) query.$select = options.select.join(',');
      if (options.expand) query.$expand = options.expand.join(',');

      // Make the API request
      const response = await client
        .api(`/chats/${chatId}/messages`)
        .query(query)
        .get();

      // Return the messages array from the response
      return response.value || [];
    } catch (error) {
      console.error('Error fetching chat messages:', error);
      throw new Error(
        `Failed to fetch messages: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }
}

export default ChatService;
