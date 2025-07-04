import { z } from 'zod';
import { Chat } from '../interfaces/chat.js';

/**
 * Base authentication schema that requires only an access token
 */
const authSchema = z.object({
  /**
   * Access token obtained through OAuth 2.0 authentication
   * @example "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
   */
  accessToken: z.string({
    required_error: 'Access token is required',
    invalid_type_error: 'Access token must be a string',
  })
  .min(1, 'Access token cannot be empty')
  .describe('A valid OAuth 2.0 access token obtained through user authentication')
});

/**
 * Schema for creating a new chat
 */
export const createChatSchema = authSchema.extend({
  /**
   * The topic or subject of the chat
   * @example "Project Discussion"
   */
  topic: z.string({
    required_error: 'Chat topic is required',
    invalid_type_error: 'Chat topic must be a string',
  })
  .min(1, 'Chat topic cannot be empty')
  .max(100, 'Chat topic cannot exceed 100 characters')
  .describe('The topic or subject of the chat'),
  
  /**
   * The type of chat to create
   * @default "group"
   */
  chatType: z.enum(['oneOnOne', 'group', 'meeting'], {
    required_error: 'Chat type is required',
    invalid_type_error: 'Chat type must be one of: oneOnOne, group, meeting',
  })
  .default('group')
  .describe('The type of chat to create'),
  
  /**
   * Array of members to add to the chat
   * @example [{"id": "user1@example.com", "roles": ["owner"]}]
   */
  members: z.array(
    z.object({
      /**
       * The email address or user ID of the member to add
       * @example "user@example.com"
       */
      id: z.string()
        .min(1, 'Member ID cannot be empty')
        .describe('The email address or user ID of the member to add'),
      
      /**
       * Array of roles for this member
       * @default ["owner"]
       */
      roles: z.array(z.string())
        .optional()
        .default(['owner'])
        .describe('Array of roles for this member (e.g., ["owner"])')
    })
  )
  .min(1, 'You must specify at least one member for the chat')
  .describe('Array of members to add to the chat')
});

/**
 * Schema for listing chats
 */
export const listChatsSchema = authSchema.extend({
  /**
   * Maximum number of items to return
   */
  top: z.number().int().positive().max(50).optional(),
  
  /**
   * Skip the first n items
   */
  skip: z.number().int().nonnegative().optional(),
  
  /**
   * Filter items by property values
   */
  filter: z.string().optional(),
  
  /**
   * Order items by property values
   */
  orderby: z.array(z.string()).optional(),
  
  /**
   * Select properties to be returned
   */
  select: z.array(z.string()).optional(),
  
  /**
   * Expand related entities
   */
  expand: z.array(z.string()).optional(),
});

/**
 * Export types for use in the application
 */
export type CreateChatInput = z.infer<typeof createChatSchema>;
export type ListChatsInput = z.infer<typeof listChatsSchema>;
