import { z } from "zod";
import { GraphUser, ChatResponse, ChatListResponse, SendMessageRequest, SendMessageResponse } from "../interface/chatInterfaces";
    
export async function listChats(token: string, options?: {
    top?: number;
    skip?: number;
    filter?: string;
    orderBy?: string;
    expand?: string[];
}): Promise<ChatListResponse> {
    if (!token) {
        throw new Error("Authentication token not found. Please set the AUTH_TOKEN environment variable.");
    }

    // Build query parameters
    const queryParams = new URLSearchParams();
    if (options?.top) queryParams.append('$top', options.top.toString());
    if (options?.skip) queryParams.append('$skip', options.skip.toString());
    if (options?.filter) queryParams.append('$filter', options.filter);
    if (options?.orderBy) queryParams.append('$orderby', options.orderBy);
    if (options?.expand?.length) {
        queryParams.append('$expand', options.expand.join(','));
    }

    const url = `https://graph.microsoft.com/v1.0/me/chats?${queryParams.toString()}`;
    
    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
        },
    });

    if (!response.ok) {
        const errorData = await response.json() as { error?: { message?: string } };
        throw new Error(`Failed to list chats: ${errorData?.error?.message || response.statusText}`);
    }

    return await response.json() as ChatListResponse;
}

export function listChatsTool(token: string) {
    return {
        token,
        name: "listChats",
        schema: {
            top: z.number().optional().describe("Number of chats to return per page (default: 50, max: 100)"),
            skip: z.number().optional().describe("Number of chats to skip for pagination"),
            filter: z.string().optional().describe("OData filter query to filter the results"),
            orderBy: z.string().optional().describe("OData orderBy query to sort the results"),
            includeMembers: z.boolean().optional().default(true).describe("Whether to include member information for each chat"),
            expand: z.array(z.string()).optional().default(['members']).describe("Array of relationships to expand in the response (e.g., ['members', 'lastMessagePreview'])")
        },
        handler: async ({
            top,
            skip,
            filter,
            orderBy,
            includeMembers = true,
            expand = ['members']
        }: {
            top?: number;
            skip?: number;
            filter?: string;
            orderBy?: string;
            includeMembers?: boolean;
            expand?: string[];
        }) => {
            try {
                const result = await listChats(token, {
                    top,
                    skip,
                    filter,
                    orderBy,
                    expand
                });

                // Format the response for MCP with detailed information in the main text
                const formattedChats = await Promise.all(result.value.map(async (chat) => {
                    const chatData: any = {
                        id: chat.id,
                        topic: chat.topic || 'No topic',
                        chatType: chat.chatType,
                        createdDateTime: chat.createdDateTime,
                        lastUpdatedDateTime: chat.lastUpdatedDateTime,
                        webUrl: chat.webUrl,
                        isHiddenForAllMembers: chat.isHiddenForAllMembers || false
                    };

                    // Include members if requested and available
                    if (includeMembers && chat.members) {
                        chatData.members = chat.members.map((member: any) => ({
                            id: member.userId || member.email,
                            displayName: member.displayName || 'Unknown',
                            roles: member.roles || [],
                            email: member.email || ''
                        }));
                    }

                    return chatData;
                }));

                // Create a detailed text response that includes all chat information
                let responseText = `Found ${result.value.length} chats\n\n`;
                
                if (formattedChats.length > 0) {
                    responseText += "Chat Details:\n";
                    responseText += "=".repeat(50) + "\n\n";
                    
                    formattedChats.forEach((chat, index) => {
                        responseText += `${index + 1}. Chat ID: ${chat.id}\n`;
                        responseText += `   Topic: ${chat.topic}\n`;
                        responseText += `   Type: ${chat.chatType}\n`;
                        responseText += `   Created: ${chat.createdDateTime}\n`;
                        responseText += `   Last Updated: ${chat.lastUpdatedDateTime || 'N/A'}\n`;
                        responseText += `   Hidden: ${chat.isHiddenForAllMembers}\n`;
                        
                        if (chat.webUrl) {
                            responseText += `   Web URL: ${chat.webUrl}\n`;
                        }
                        
                        if (chat.members && chat.members.length > 0) {
                            responseText += `   Members (${chat.members.length}):\n`;
                            chat.members.forEach((member: any, memberIndex: number) => {
                                responseText += `     ${memberIndex + 1}. ${member.displayName}`;
                                if (member.email) {
                                    responseText += ` (${member.email})`;
                                }
                                responseText += `\n        ID: ${member.id}\n`;
                                if (member.roles && member.roles.length > 0) {
                                    responseText += `        Roles: ${member.roles.join(', ')}\n`;
                                }
                            });
                        }
                        
                        responseText += "\n" + "-".repeat(30) + "\n\n";
                    });
                } else {
                    responseText += "No chats found matching the criteria.\n";
                }

                // Add pagination info if available
                if (result["@odata.count"]) {
                    responseText += `\nTotal Count: ${result["@odata.count"]}\n`;
                }
                if (result["@odata.nextLink"]) {
                    responseText += `Next Page Available: Yes\n`;
                }

                return {
                    content: [{
                        type: "text" as const,
                        text: responseText
                    }]
                };
            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
                return {
                    content: [{
                        type: "text" as const,
                        text: `Error listing chats: ${errorMessage}`
                    }],
                    isError: true
                };
            }
        }
    };
}

export async function sendMessage(token: string, chatId: string, message: SendMessageRequest): Promise<SendMessageResponse> {
    if (!token) {
        throw new Error("Authentication token not found. Please set the AUTH_TOKEN environment variable.");
    }

    const url = `https://graph.microsoft.com/v1.0/chats/${chatId}/messages`;
    
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(message)
    });

    if (!response.ok) {
        const errorData = await response.json() as { error?: { message?: string } };
        throw new Error(`Failed to send message: ${errorData?.error?.message || response.statusText}`);
    }

    return await response.json() as SendMessageResponse;
}

export function sendMessageTool(token: string) {
    return {
        token,
        name: "sendMessage",
        schema: {
            chatId: z.string().describe("The ID of the chat to send the message to"),
            content: z.string().describe("The content of the message"),
            contentType: z.enum(['text', 'html']).default('text').describe("The content type of the message"),
            subject: z.string().optional().describe("The subject of the message"),
            importance: z.enum(['normal', 'high', 'urgent']).default('normal').describe("The importance of the message")
        },
        handler: async ({
            chatId,
            content,
            contentType = 'text',
            subject,
            importance = 'normal'
        }: {
            chatId: string;
            content: string;
            contentType?: 'text' | 'html';
            subject?: string;
            importance?: 'normal' | 'high' | 'urgent';
        }) => {
            try {
                const messageRequest: SendMessageRequest = {
                    body: {
                        contentType,
                        content
                    },
                    importance,
                    ...(subject && { subject })
                };

                const result = await sendMessage(token, chatId, messageRequest);

                return {
                    content: [{
                        type: "text" as const,
                        text: `✅ Message sent successfully!\n` +
                              `📅 Sent at: ${new Date(result.createdDateTime).toLocaleString()}\n` +
                              (result.webUrl ? `🔗 View message: ${result.webUrl}\n` : '')
                    }]
                };
            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
                return {
                    content: [{
                        type: "text" as const,
                        text: `❌ Failed to send message: ${errorMessage}`
                    }]
                };
            }
        }
    };
}

export function createChat(token: string) {
    return {
        token,
        name: "createChat",
        schema: {
            chatType: z.enum(['oneOnOne', 'group', 'meeting', 'unknown']).describe("Type of chat to create"),
            topic: z.string().optional().describe("Optional topic/subject for the chat"),
            members: z.array(
                z.object({
                    id: z.string().describe("User ID or email address of the member"),
                    displayName: z.string().optional().describe("Display name of the member"),
                    email: z.string().email().optional().describe("Email address of the member")
                })
            ).min(1, "At least one member is required")
             .describe("List of members to include in the chat")
        },
        handler: async ({ chatType, topic, members }: { 
            chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
            topic?: string;
            members: Array<{ id: string; displayName?: string; email?: string }>;
        }) => {
            try {
                if (!token) {
                    throw new Error("Authentication token not found. Please configure the AUTH_TOKEN environment variable in your MCP server configuration.");
                }

                // First, get the current user's information
                const meResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Content-Type': 'application/json',
                    },
                });

                if (!meResponse.ok) {
                    const errorData = await meResponse.json() as { error?: { message?: string } };
                    throw new Error(`Failed to get current user info: ${errorData?.error?.message || meResponse.statusText}`);
                }

                const me = await meResponse.json() as GraphUser;
                const myEmail = me.mail || me.userPrincipalName;

                if (!myEmail) {
                    throw new Error('Could not determine your email address from the authentication token');
                }

                // Format members for Microsoft Graph API, including the current user as an owner
                const chatMembers = [
                    // Current user as owner
                    {
                        '@odata.type': '#microsoft.graph.aadUserConversationMember',
                        roles: ['owner'],
                        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${myEmail}`
                    },
                    // Other members
                    ...members.map(member => ({
                        '@odata.type': '#microsoft.graph.aadUserConversationMember',
                        roles: ['owner'],
                        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${member.id}`
                    }))
                ];

                // Prepare the request body
                const requestBody: any = {
                    chatType,
                    members: chatMembers
                };

                // Add topic if provided
                if (topic) {
                    requestBody.topic = topic;
                }

                // Make the API request to create chat
                const response = await fetch('https://graph.microsoft.com/v1.0/chats', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(requestBody)
                });

                if (!response.ok) {
                    const errorData = await response.json() as { error?: { message?: string } };
                    throw new Error(`Microsoft Graph API error: ${errorData?.error?.message || response.statusText}`);
                }

                const chatData = await response.json() as ChatResponse;

                return {
                    content: [{
                        type: "text" as const,
                        text: `Successfully created chat with ID: ${chatData.id}`,
                        _meta: {
                            chatId: chatData.id,
                            webUrl: chatData.webUrl
                        }
                    }]
                };

            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
                return {
                    content: [{
                        type: "text" as const,
                        text: `Error creating chat: ${errorMessage}`,
                        _meta: {}
                    }],
                    isError: true
                };
            }
        }
    };
}
