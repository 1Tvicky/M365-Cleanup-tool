import { OAuth2Client } from "google-auth-library";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

// google-auth-library doesn't validate its constructor args, so an empty clientId silently
// produces a consent URL missing client_id — the browser reaches Google's real screen and Google
// shows its own confusing "Access blocked: Authorization Error" page. Guard here instead, the
// same way graph/client.ts and microsoftUserAuth.ts guard their MSAL clients, so a deployment
// with no Google OAuth configured fails with our own clean error before ever leaving this app.
function requireGoogleConfigured(): void {
  if (!config.oauth.google.clientId || !config.oauth.google.clientSecret) {
    throw new ApiError(503, "GOOGLE_LOGIN_NOT_CONFIGURED", "Google sign-in isn't configured on this deployment yet");
  }
}

const client = new OAuth2Client(
  config.oauth.google.clientId,
  config.oauth.google.clientSecret,
  config.oauth.google.redirectUri
);

export function getGoogleAuthUrl(state: string): string {
  requireGoogleConfigured();
  return client.generateAuthUrl({
    access_type: "online",
    scope: ["openid", "email", "profile"],
    state,
  });
}

export interface GoogleIdentity {
  email: string;
  displayName: string;
  emailVerified: boolean;
}

/** Exchanges the OAuth `code` for tokens and returns the verified identity from the ID token. */
export async function resolveGoogleIdentity(code: string): Promise<GoogleIdentity> {
  requireGoogleConfigured();
  const { tokens } = await client.getToken(code);
  if (!tokens.id_token) throw new Error("Google response did not include an id_token");

  const ticket = await client.verifyIdToken({
    idToken: tokens.id_token,
    audience: config.oauth.google.clientId,
  });
  const payload = ticket.getPayload();
  if (!payload?.email) throw new Error("Google identity token missing email claim");

  return {
    email: payload.email,
    displayName: payload.name ?? payload.email,
    emailVerified: payload.email_verified ?? false,
  };
}
