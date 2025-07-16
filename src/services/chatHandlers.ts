import { IncomingMessage, ServerResponse } from 'http';
import { z } from 'zod';
import ChatService from './chatService.js';
import { ChatMessage, ChatMember } from '../interfaces/chat.js';

/**
 * Handler for listing all chats
 */
/**
 * Handler for creating a new chat
 */
export async function handleCreateChat(req: IncomingMessage, res: ServerResponse): Promise<void> {
  if (req.method !== 'POST') {
    res.writeHead(405, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Method not allowed' }));
    return;
  }

  let body = '';
  req.on('data', chunk => {
    body += chunk.toString();
  });

  return new Promise((resolve) => {
    req.on('end', async () => {
      try {
        const chatData = JSON.parse(body);
        
        // Validate the chat data
        const authHeader = req.headers['authorization'];
        if (!authHeader || !authHeader.startsWith('Bearer ')) {
          res.writeHead(401, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Authorization header with Bearer token is required' }));
          return resolve();
        }
        const accessToken = authHeader.split(' ')[1];
        
        const schema = z.object({
          chatType: z.enum(['group', 'oneOnOne', 'meeting']),
          topic: z.string().min(1, 'Topic is required'),
          members: z.array(z.object({
            id: z.string().min(1, 'Member ID is required'),
            userIdentityType: z.enum(['aadUser', 'guest', 'externalUser']).default('aadUser'),
            roles: z.array(z.string()).default(['guest'])
          })).min(1, 'At least one member is required')
        });
        
        const validatedData = schema.parse(chatData);
        
        const chatService = new ChatService({ accessToken });
        const chat = await chatService.createChat(validatedData);
        
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          success: true,
          data: chat
        }));
        
      } catch (error) {
        console.error('Error creating chat:', error);
        let statusCode = 500;
        let errorMessage = 'Failed to create chat';
        let errorDetails: any = { message: error instanceof Error ? error.message : 'Unknown error' };

        if (error instanceof z.ZodError) {
          statusCode = 400;
          errorDetails.issues = error.issues;
        } else if ((error as any).statusCode) {
          // Handle Microsoft Graph API errors
          statusCode = (error as any).statusCode;
          try {
            errorDetails = {
              ...errorDetails,
              ...(error as any).body ? JSON.parse((error as any).body) : {},
              statusCode: (error as any).statusCode,
              code: (error as any).code,
              requestId: (error as any).requestId
            };
            
            // Provide more user-friendly messages for common errors
            if ((error as any).code === 'Unauthorized' || statusCode === 401) {
              errorMessage = 'Authentication failed. Please check your access token.';
            } else if ((error as any).code === 'TenantNotFound') {
              errorMessage = 'The specified tenant was not found. Please verify your tenant ID.';
            }
          } catch (parseError) {
            console.error('Error parsing error response:', parseError);
          }
        }

        res.writeHead(statusCode, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ 
          error: errorMessage,
          details: errorDetails
        }));
      } finally {
        resolve();
      }
    });
  });
}

/**
 * Handler for listing all chats
 */
export async function handleListChats(req: IncomingMessage, res: ServerResponse): Promise<void> {
  return new Promise<void>(async (resolve) => {
    try {
      const authHeader = req.headers['authorization'];
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        res.writeHead(401, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Authorization header with Bearer token is required' }));
        return Promise.resolve();
      }
      const accessToken = authHeader.split(' ')[1];
      const query = new URL(req.url || '', 'http://localhost').searchParams;

      const chatService = new ChatService({ accessToken });
      // Microsoft Graph API has a maximum limit of 50 items per request
      const chats = await chatService.listChats({ top: 50 });
      
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        success: true,
        data: chats
      }));
      
    } catch (error) {
      console.error('Error listing chats:', error);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ 
        error: 'Failed to list chats',
        details: error instanceof Error ? error.message : 'Unknown error'
      }));
    } finally {
      resolve();
    }
  });
}

/**
 * Handler for sending a message
 */
export async function handleSendMessage(req: IncomingMessage, res: ServerResponse): Promise<void> {
  if (req.method !== 'POST') {
    res.writeHead(405, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Method not allowed' }));
    return;
  }

  try {
    // Read the request body
    const chunks: Buffer[] = [];
    for await (const chunk of req) {
      chunks.push(chunk);
    }
    const body = Buffer.concat(chunks).toString();
    
    // Parse and validate the request body
    const messageData = JSON.parse(body);
    
    const authHeader = req.headers['authorization'];
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      res.writeHead(401, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Authorization header with Bearer token is required' }));
      return;
    }
    const accessToken = authHeader.split(' ')[1];
    
    const schema = z.object({
      chatId: z.string().min(1, 'Chat ID is required'),
      content: z.string().min(1, 'Message content is required'),
      contentType: z.enum(['text', 'html', 'content']).optional().default('text'),
      messageMetadata: z.record(z.any()).optional()
    });
    
    const validation = schema.safeParse(messageData);
    
    if (!validation.success) {
      const errorMessages = validation.error.issues.map(issue => 
        `${issue.path.join('.')}: ${issue.message}`
      ).join(', ');
      
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ 
        error: 'Validation failed',
        details: errorMessages
      }));
      return;
    }
    
    // Initialize the chat service and send the message
    const chatService = new ChatService({ accessToken });

    const message = await chatService.sendMessage(
      validation.data.chatId,
      {
        content: validation.data.content,
        contentType: validation.data.contentType,
        ...(validation.data.messageMetadata && { messageMetadata: validation.data.messageMetadata })
      }
    );
    
    // Send success response
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      success: true,
      data: message
    }));
    
  } catch (error) {
    console.error('Error sending message:', error);
    const statusCode = error instanceof z.ZodError ? 400 : 500;
    res.writeHead(statusCode, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      error: 'Failed to send message',
      details: error instanceof Error ? error.message : 'Unknown error',
      ...(error instanceof Error && error.stack ? { stack: error.stack } : {})
    }));
  }
}

export async function handleGetMessages(req: IncomingMessage, res: ServerResponse): Promise<void> {
  if (req.method !== 'GET') {
    res.writeHead(405, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Method not allowed' }));
    return;
  }

  const authHeader = req.headers['authorization'];
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.writeHead(401, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Authorization header with Bearer token is required' }));
    return;
  }
  const accessToken = authHeader.split(' ')[1];
  
  const { searchParams } = new URL(req.url || '', `http://${req.headers.host}`);
  const chatId = searchParams.get('chatId');
  const top = parseInt(searchParams.get('top') || '50', 10);
  
  try {
    // Validate query parameters
    if (!chatId) {
      throw new Error('chatId is required');
    }
    
    const chatService = new ChatService({ accessToken });
    const messages = await chatService.getMessages(chatId, { top });
    
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      success: true,
      data: messages
    }));
    
  } catch (error) {
    console.error('Error fetching messages:', error);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      error: 'Failed to fetch messages',
      details: error instanceof Error ? error.message : 'Unknown error'
    }));
  }
}
