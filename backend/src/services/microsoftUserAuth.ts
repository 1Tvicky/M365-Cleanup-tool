import { ConfidentialClientApplication } from "@azure/msal-node";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

/**
 * Delegated auth-code flow for OPERATOR sign-in ("Office 365" button) — identifies the CloudFuze
 * staff member, scoped to User.Read only. Deliberately a separate MSAL app instance from
 * graph/client.ts, which holds application permissions against CUSTOMER tenants. Never share the
 * client id/secret between the two.
 *
 * Built lazily (not at module load): MSAL throws synchronously if the client secret is empty, and
 * this feature is optional — an environment with no OPERATOR_AZURE_* configured should still be
 * able to start the server and use password login, not crash on import.
 */
let msalApp: ConfidentialClientApplication | null = null;

function getMsalApp(): ConfidentialClientApplication {
  if (!config.oauth.operatorAzure.clientId || !config.oauth.operatorAzure.clientSecret) {
    throw new ApiError(503, "OFFICE365_LOGIN_NOT_CONFIGURED", "Office 365 sign-in isn't configured on this deployment yet");
  }
  if (!msalApp) {
    msalApp = new ConfidentialClientApplication({
      auth: {
        clientId: config.oauth.operatorAzure.clientId,
        clientSecret: config.oauth.operatorAzure.clientSecret,
        authority: `https://login.microsoftonline.com/${config.oauth.operatorAzure.tenant}`,
      },
    });
  }
  return msalApp;
}

const SCOPES = ["User.Read", "openid", "email", "profile"];

export async function getMicrosoftAuthUrl(state: string): Promise<string> {
  return getMsalApp().getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: config.oauth.operatorAzure.redirectUri,
    state,
  });
}

export interface MicrosoftIdentity {
  email: string;
  displayName: string;
}

export async function resolveMicrosoftIdentity(code: string): Promise<MicrosoftIdentity> {
  const result = await getMsalApp().acquireTokenByCode({
    code,
    scopes: SCOPES,
    redirectUri: config.oauth.operatorAzure.redirectUri,
  });

  const email = result?.account?.username;
  if (!email) throw new Error("Microsoft identity response missing account username/email");

  return {
    email,
    displayName: result.account?.name ?? email,
  };
}
