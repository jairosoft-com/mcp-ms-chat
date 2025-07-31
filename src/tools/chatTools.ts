import { z } from "zod";
import { GraphUser, ChatResponse } from "../interface/chatInterfaces";
    
// Shared authentication token
let currentAuthToken: string | undefined;

// Function to set the authentication token
export function setAuthToken(token: string | undefined) {
    currentAuthToken = token;
}

export function createChat() {
    return {
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
                if (!currentAuthToken) {
                    throw new Error("Authentication token not found. Please configure the AUTH_TOKEN environment variable in your MCP server configuration.");
                }

                // First, get the current user's information
                const meResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
                    headers: {
                        'Authorization': `Bearer ${currentAuthToken}`,
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
                        'Authorization': `Bearer ${currentAuthToken}`,
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
