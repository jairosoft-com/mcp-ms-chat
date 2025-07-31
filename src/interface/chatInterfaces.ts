export interface GraphUser {
    id: string;
    mail?: string;
    userPrincipalName: string;
    displayName?: string;
    [key: string]: any;
}

export interface ChatResponse {
    id: string;
    webUrl?: string;
    chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknown';
    createdDateTime: string;
    topic?: string;
    [key: string]: any;
}

// Define environment variable types
export interface Env {
    AUTH_TOKEN?: string;
}