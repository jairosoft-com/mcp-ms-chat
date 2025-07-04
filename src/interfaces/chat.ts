/**
 * Represents a chat member
 */
export interface ChatMember {
  /** The member's ID (user principal name or Azure AD user ID) */
  id: string;
  
  /** The member's display name */
  displayName?: string;
  
  /** The member's roles in the chat */
  roles?: string[];
  
  /** The member's user principal name (email) */
  userPrincipalName?: string;
  
  /** The member's user ID */
  userId?: string;
}

/**
 * Represents a chat message
 */
export interface ChatMessage {
  /** The message content */
  content: string;
  
  /** The message type */
  contentType?: 'text' | 'html' | 'content';
  
  /** The message ID */
  id?: string;
  
  /** The message creation timestamp */
  createdDateTime?: string;
  
  /** The sender's user ID */
  from?: {
    user: {
      id: string;
      displayName?: string;
    };
  };
}

/**
 * Represents a chat in Microsoft Teams
 */
export interface Chat {
  /** The chat ID */
  id: string;
  
  /** The chat topic */
  topic?: string;
  
  /** The chat creation timestamp */
  createdDateTime?: string;
  
  /** The chat type */
  chatType?: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
  
  /** The web URL for the chat */
  webUrl?: string;
  
  /** The last message in the chat */
  lastMessagePreview?: {
    id?: string;
    createdDateTime?: string;
    isDeleted?: boolean;
    messageType?: 'message' | 'systemEventMessage' | 'unknown';
  };
  
  /** The members in the chat */
  members?: Array<{
    id: string;
    displayName?: string;
    userId?: string;
    email?: string;
    roles?: string[];
  }>;
}

/**
 * Represents a chat message collection response
 */
export interface ChatMessageCollectionResponse {
  /** Array of chat messages */
  value: ChatMessage[];
  
  /** URL to the next page of results, if available */
  '@odata.nextLink'?: string;
}

/**
 * Represents a chat collection response
 */
export interface ChatCollectionResponse {
  /** Array of chats */
  value: Chat[];
  
  /** URL to the next page of results, if available */
  '@odata.nextLink'?: string;
}

/**
 * Represents options for listing chats
 */
export interface ListChatsOptions {
  /** Maximum number of items to return */
  top?: number;
  
  /** Skip the first n items */
  skip?: number;
  
  /** Filter items by property values */
  filter?: string;
  
  /** Include count of items */
  count?: boolean;
  
  /** Order items by property values */
  orderby?: string[];
  
  /** Select properties to be returned */
  select?: string[];
  
  /** Expand related entities */
  expand?: string[];
}
