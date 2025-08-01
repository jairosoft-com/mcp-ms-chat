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

export interface ChatMessage {
    id?: string;
    replyToId?: string;
    etag?: string;
    messageType: 'message' | 'systemEventMessage' | 'unknownFutureValue';
    createdDateTime: string;
    lastModifiedDateTime?: string;
    lastEditedDateTime?: string;
    deletedDateTime?: string;
    subject?: string | null;
    summary?: string | null;
    importance: 'normal' | 'high' | 'urgent';
    locale: string;
    webUrl?: string;
    from?: {
        application?: any;
        device?: any;
        user?: {
            id: string;
            displayName?: string;
            userIdentityType?: 'aadUser' | 'onPremiseAadUser' | 'anonymousGuest' | 'federatedUser' | 'personalMicrosoftAccountUser' | 'skypeUser' | 'phoneUser' | 'unknownFutureValue';
        };
    };
    body: {
        contentType: 'text' | 'html';
        content: string;
    };
    attachments?: any[];
    mentions?: any[];
    reactions?: any[];
}

export interface SendMessageRequest {
    body: {
        contentType: 'text' | 'html';
        content: string;
    };
    subject?: string | null;
    importance?: 'normal' | 'high' | 'urgent';
}

export interface SendMessageResponse {
    id: string;
    createdDateTime: string;
    webUrl?: string;
}