/**
 * Custom error class for authentication-related errors
 */
export class AuthenticationError extends Error {
  public readonly code: string;
  public readonly details?: Record<string, unknown>;
  public readonly remediation?: string;

  constructor(
    code: string,
    message: string,
    remediation?: string,
    details?: Record<string, unknown>
  ) {
    super(message);
    this.name = 'AuthenticationError';
    this.code = code;
    this.remediation = remediation;
    this.details = details;

    // Set the prototype explicitly for proper instanceof checks
    Object.setPrototypeOf(this, AuthenticationError.prototype);
  }

  /**
   * Converts the error to a JSON-serializable object
   */
  public toJSON() {
    return {
      error: this.code,
      message: this.message,
      ...(this.remediation && { remediation: this.remediation }),
      ...(this.details && { details: this.details }),
      stack: this.stack
    };
  }

  /**
   * Logs the error to the console with appropriate formatting
   */
  public log() {
    console.error(`[AuthenticationError: ${this.code}] ${this.message}`);
    if (this.remediation) {
      console.error(`Remediation: ${this.remediation}`);
    }
    if (this.details) {
      console.error('Details:', this.details);
    }
    if (process.env.NODE_ENV === 'development' && this.stack) {
      console.error(this.stack);
    }
  }
}
