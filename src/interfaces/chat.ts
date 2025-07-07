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
 * Represents a chat in Microsoft Teams
 */
export interface Chat {
  id: string;
  topic?: string;
  chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
  createdDateTime: string;
  lastUpdatedDateTime: string;
  webUrl?: string;
  members?: Array<{
    id: string;
    displayName?: string;
    userId?: string;
    email?: string;
  }>;
  lastMessagePreview?: {
    id: string;
    createdDateTime: string;
    body: {
      contentType: 'text' | 'html' | 'content';
      content: string;
    };
    from?: {
      user: {
        id: string;
        displayName?: string;
      };
    };
  };
}

export interface ChatMessage {
  id: string;
  createdDateTime: string;
  messageType: 'message' | 'systemEventMessage';
  body: {
    contentType: 'text' | 'html' | 'content';
    content: string;
  };
  from?: {
    user: {
      id: string;
      displayName?: string;
    };
  };
}

export interface ChatMessageCollectionResponse {
  value: ChatMessage[];
  '@odata.nextLink'?: string;
}

export interface ChatCollectionResponse {
  value: Chat[];
  '@odata.nextLink'?: string;
}

export interface ListChatsOptions {
  filter?: string;
  top?: number;
  skip?: number;
  orderby?: string;
  select?: string[];
  expand?: string[];
}