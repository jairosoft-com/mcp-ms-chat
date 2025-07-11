import { AuthenticationError } from "../errors/authError.js";

/**
 * Interface for authentication configuration
 * Only requires an access token for authentication
 */
export interface AuthConfig {
  /**
   * Access token obtained through OAuth 2.0 authentication
   */
  accessToken: string;
  
  /**
   * (Optional) Token expiration time in milliseconds since epoch
   * If not provided, token will be treated as valid
   */
  expiresIn?: number;
}

/**
 * Validates the authentication configuration
 * @param config Authentication configuration
 * @throws {AuthenticationError} If access token is missing
 */
function validateAuthConfig(config: Partial<AuthConfig>): void {
  if (!config.accessToken) {
    throw new AuthenticationError(
      'MISSING_ACCESS_TOKEN',
      'Access Token Required',
      'access_token needed'
    );
  }
}

/**
 * Gets an authenticated token for Microsoft Graph API
 * @param config Authentication configuration
 * @returns Promise<string> Access token
 */
export async function authenticate(config: AuthConfig): Promise<string> {
  try {
    validateAuthConfig(config);
    
    // Check if token is expired
    if (config.expiresIn && Date.now() >= config.expiresIn - 300000) {
      throw new AuthenticationError(
        'TOKEN_EXPIRED',
        'Access Token Expired',
        'The provided access token has expired. Please obtain a new token and try again.'
      );
    }
    
    return config.accessToken;
  } catch (error: unknown) {
    if (error instanceof AuthenticationError) {
      throw error;
    }
    
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    
    throw new AuthenticationError(
      'AUTHENTICATION_FAILED',
      'Authentication Failed',
      `Failed to authenticate with Microsoft Graph: ${errorMessage}`,
      { originalError: errorMessage }
    );
  }
}

export default {
  authenticate
};
