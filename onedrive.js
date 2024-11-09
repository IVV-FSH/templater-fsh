import { DeviceCodeCredential } from "@azure/identity";  // Device Code Flow from Azure SDK
import { Client } from "@microsoft/microsoft-graph-client";  // MS Graph SDK
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/src/authentication/azureTokenCredentials/TokenCredentialAuthenticationProvider.js";
import dotenv from "dotenv";

dotenv.config();

// Load environment variables
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;

if (!tenantId || !clientId) {
    console.error("Missing necessary environment variables (TENANT_ID, CLIENT_ID).");
    process.exit(1);
}

// Set up the Device Code Credential to use Device Code Flow
const credential = new DeviceCodeCredential({
  tenantId,
  clientId,
  userPromptCallback: (info) => {
    console.log(info.message);  // Show device code info to the user
  }
});

// Create an AuthenticationProvider using the Device Code Credential
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['User.Read'],  // Define required scopes/permissions
});

// Initialize the Microsoft Graph client with the authentication provider
const graphClient = Client.initWithMiddleware({
  authProvider
});

// Test a simple API request to get the signed-in user's profile
const testConnection = async () => {
  try {
    const user = await graphClient.api('/me').get();  // Access the signed-in user's data
    console.log("Connected to Microsoft Graph API successfully!");
    console.log("User details:", user);
  } catch (error) {
    console.error("Error connecting to Microsoft Graph API:", error);
  }
};

// Execute the test connection function
testConnection();
