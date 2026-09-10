import { google } from "googleapis";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

/**
 * The Add Clouds connect flow's identity step for Google Workspace — a standard 3-legged OAuth
 * exchange used ONLY to identify the connecting admin (email + display name), mirroring
 * m365ConnectAuth.ts's delegated-identity-only role. Requests no Drive/Directory scopes: actual
 * data access always goes through services/googleWorkspaceAuth.ts's service-account domain-wide
 * delegation, verified separately once this identity step confirms who's connecting. See
 * docs/google-workspace-integration.md.
 *
 * Built lazily, same reasoning as m365ConnectAuth.ts / graph/client.ts: an unconfigured
 * deployment should start and serve everything else, not crash on import.
 */

function getOAuthClient() {
  const { clientId, clientSecret, redirectUri } = config.googleWorkspaceOAuth;
  if (!clientId || !clientSecret || !redirectUri) {
    throw new ApiError(503, "GOOGLE_WORKSPACE_CONNECT_NOT_CONFIGURED", "Google Workspace cloud connections aren't configured on this deployment yet");
  }
  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

const CONNECT_SCOPES = ["openid", "email", "profile"];

export function getGoogleConnectAuthorizeUrl(state: string): string {
  return getOAuthClient().generateAuthUrl({
    access_type: "online", // identity-only — no refresh token needed, this token is never stored or reused
    scope: CONNECT_SCOPES,
    state,
    prompt: "select_account",
  });
}

export interface GoogleDelegatedIdentity {
  adminEmail: string;
  adminDisplayName: string;
}

/** Exchanges the authorization code and resolves the connecting admin's identity — never used for data access. */
export async function exchangeGoogleConnectCode(code: string): Promise<GoogleDelegatedIdentity> {
  const client = getOAuthClient();
  const { tokens } = await client.getToken(code);
  if (!tokens.id_token) {
    throw new Error("Google connect code exchange did not return an id_token");
  }

  const ticket = await client.verifyIdToken({ idToken: tokens.id_token, audience: config.googleWorkspaceOAuth.clientId });
  const payload = ticket.getPayload();
  if (!payload?.email) {
    throw new Error("Google identity token missing email claim");
  }

  return {
    adminEmail: payload.email,
    adminDisplayName: payload.name ?? payload.email,
  };
}
