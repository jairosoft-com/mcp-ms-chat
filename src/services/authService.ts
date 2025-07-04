import { ClientSecretCredential } from "@azure/identity";
import dotenv from "dotenv";

dotenv.config();

/**
 * Get Azure AD credentials using client secret flow
 * @returns ClientSecretCredential instance
 */
export function getAzureCredentials(): ClientSecretCredential {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "Missing required Azure AD credentials. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables."
    );
  }

  return new ClientSecretCredential(tenantId, clientId, clientSecret);
}

/**
 * Validates if all required Azure AD environment variables are set
 * @throws Error if any required environment variable is missing
 */
export function validateAuthConfig(): void {
  const requiredVars = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"];
  const missingVars = requiredVars.filter((varName) => !process.env[varName]);

  if (missingVars.length > 0) {
    throw new Error(
      `Missing required environment variables: ${missingVars.join(
        ", "
      )}. Please check your .env file.`
    );
  }
}
