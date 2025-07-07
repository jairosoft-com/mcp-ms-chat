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
    email?: string;
    roles?: string[];
    userPrincipalName?: string;
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

/**
 * Represents a chat message
 */
export interface ChatMessage {
  /** Unique identifier for the message */
  id: string;
  
  /** The date and time when the message was sent */
  createdDateTime: string;
  
  /** The type of message */
  messageType: 'message' | 'systemEventMessage' | 'system' | 'typing';
  
  /** The content of the message */
  body: {
    /** The content type of the message body */
    contentType: 'text' | 'html' | 'content';
    /** The content of the message */
    content: string;
  };
  
  /** The sender of the message */
  from?: {
    user: {
      /** The sender's ID */
      id: string;
      /** The sender's display name */
      displayName?: string;
      /** The sender's user principal name (email) */
      userPrincipalName?: string;
    };
  };
  
  /** The importance/priority of the message */
  importance?: 'normal' | 'high' | 'urgent' | 'low';
  
  /** The subject of the message */
  subject?: string;
  
  /** The URL for the message in Microsoft Teams */
  webUrl?: string;
  
  /** The ID of the chat the message belongs to */
  chatId?: string;
  
  /** The ID of the reply-to message */
  replyToId?: string;
  
  /** Collection of replies to the message */
  replies?: ChatMessage[];
  
  /** Whether the message has been read by the current user */
  isRead?: boolean;
  
  /** The date and time when the message was last modified */
  lastModifiedDateTime?: string;
  
  /** The MIME type of the message content */
  mimeType?: string;
  
  /** The policy violation information */
  policyViolation?: any;
  
  /** The reactions to the message */
  reactions?: MessageReaction[];
  
  /** The mentions in the message */
  mentions?: Mention[];
  
  /** The attachments in the message */
  attachments?: Attachment[];

  chatTopic: string;
}

/**
 * Represents a reaction to a message
 */
export interface MessageReaction {
  /** The type of reaction */
  reactionType: string;
  /** The date and time when the reaction was added */
  createdDateTime: string;
  /** The user who added the reaction */
  user: {
    id: string;
    displayName?: string;
    userPrincipalName?: string;
  };
}

/**
 * Represents a mention in a message
 */
export interface Mention {
  /** The ID of the mention */
  id?: number;
  /** The text of the mention */
  mentionText?: string;
  /** The user or application that was mentioned */
  mentioned?: {
    application?: any;
    device?: any;
    user?: {
      id: string;
      displayName?: string;
      userPrincipalName?: string;
    };
  };
}

/**
 * Represents a file or item attachment in a message
 */
export interface Attachment {
  /** The ID of the attachment */
  id: string;
  /** The MIME type of the attachment */
  contentType: string;
  /** The name of the attachment */
  name?: string;
  /** The size of the attachment in bytes */
  size?: number;
  /** The URL to download the attachment */
  contentUrl?: string;
  /** The thumbnail URL of the attachment if it's an image */
  thumbnailUrl?: string;
  /** The content of the attachment if it's inline */
  content?: string;
  /** The content ID of the attachment if it's inline */
  contentId?: string;
  /** The last modified date and time of the attachment */
  lastModifiedDateTime?: string;
  /** The ID of the message that the attachment is attached to */
  messageId?: string;
  /** The ID of the chat that the attachment is in */
  chatId?: string;
  /** The ID of the team that the attachment is in */
  teamId?: string;
  /** The ID of the channel that the attachment is in */
  channelId?: string;
  /** The ID of the attachment in SharePoint */
  sharepointIds?: any;
  /** The URL to the attachment in OneDrive for Business or SharePoint */
  sharepointUrl?: string;
  /** The URL to the attachment in OneDrive for Business or SharePoint */
  webUrl?: string;
  /** The URL to download the attachment */
  downloadUrl?: string;
  /** The URL to the attachment in OneDrive for Business or SharePoint */
  self?: string;
  /** The URL to download the attachment from Microsoft Graph */
  downloadGraphUrl?: string;
  
  /** Media-related metadata for the attachment */
  mediaMetadata?: {
    /** The content type of the media */
    contentType?: string;
    /** The ETag of the media */
    etag?: string;
    /** The last modified date of the media */
    lastModifiedDateTime?: string;
    /** The content disposition of the media */
    contentDisposition?: string;
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

/**
 * Options for listing chats with pagination and filtering
 */
export interface ListChatsOptions {
  /** OData filter string */
  filter?: string;
  
  /** Maximum number of items to return */
  top?: number;
  
  /** Number of items to skip */
  skip?: number;
  
  /** Sorting order (e.g., 'createdDateTime desc') */
  orderby?: string;
  
  /** Properties to include in the response */
  select?: string[];
  
  /** Related entities to include in the response */
  expand?: string[];
}

/**
 * Options for listing messages with advanced filtering and pagination
 */
export interface ListMessagesOptions {
  /** 
   * Filter messages by sender ID or user principal name (email). 
   * Can be 'me' for current user or a specific user ID/email.
   */
  from?: string;
  
  /** 
   * Filter messages sent after this date/time (ISO 8601 format)
   * Example: '2023-01-01T00:00:00Z'
   */
  afterDateTime?: string;
  
  /** 
   * Filter messages sent before this date/time (ISO 8601 format)
   * Example: '2023-12-31T23:59:59Z'
   */
  beforeDateTime?: string;
  
  /** 
   * Filter messages that contain the specified text in the content
   */
  contains?: string;
  
  /** 
   * Filter to only unread messages
   */
  isRead?: boolean;
  
  /** 
   * Filter by importance level
   */
  importance?: 'low' | 'normal' | 'high';
  
  /** 
   * Include message replies in the results
   */
  includeReplies?: boolean;
  
  /** 
   * Include deleted messages in the results
   */
  includeDeleted?: boolean;
  
  /** 
   * Maximum number of items to return (1-1000, default: 50)
   */
  top?: number;
  
  /** 
   * Number of items to skip (for pagination)
   */
  skip?: number;
  
  /** 
   * Sorting order (e.g., 'createdDateTime desc', 'from/emailAddress/name')
   */
  orderBy?: string;
  
  /** 
   * Properties to include in the response
   */
  select?: string[];
  
  /** 
   * Related entities to include in the response
   */
  expand?: string[];
}