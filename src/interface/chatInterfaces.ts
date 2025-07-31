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
    lastUpdatedDateTime?: string;
    topic?: string | null;
    isHiddenForAllMembers?: boolean;
    members?: any[];
    chatViewpoint?: {
        isHidden: boolean;
        lastMessageReadDateTime: string;
    };
    [key: string]: any;
}

export interface ChatListResponse {
    "@odata.context": string;
    "@odata.count"?: number;
    "@odata.nextLink"?: string;
    value: ChatResponse[];
}

// Define environment variable types
export interface Env {
    AUTH_TOKEN?: string;
}