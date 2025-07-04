import { z } from 'zod';
import { Chat } from '../interfaces/chat.js';

/**
 * Schema for creating a new chat
 */
export const createChatSchema = z.object({
  /** The chat topic */
  topic: z.string().min(1, 'Chat topic is required'),
  
  /** The chat type */
  chatType: z.enum(['oneOnOne', 'group', 'meeting', 'unknown']).default('group'),
  
  /** Array of member IDs to add to the chat */
  members: z.array(
    z.object({
      id: z.string().min(1, 'Member ID is required'),
      roles: z.array(z.string()).optional(),
    })
  ).min(1, 'At least one member is required'),
});

/**
 * Schema for listing chats
 */
export const listChatsSchema = z.object({
  /** Maximum number of items to return */
  top: z.number().int().positive().max(50).optional(),
  
  /** Skip the first n items */
  skip: z.number().int().nonnegative().optional(),
  
  /** Filter items by property values */
  filter: z.string().optional(),
  
  /** Order items by property values */
  orderby: z.array(z.string()).optional(),
  
  /** Select properties to be returned */
  select: z.array(z.string()).optional(),
  
  /** Expand related entities */
  expand: z.array(z.string()).optional(),
});

/**
 * Export types for use in the application
 */
export type CreateChatInput = z.infer<typeof createChatSchema>;
export type ListChatsInput = z.infer<typeof listChatsSchema>;
