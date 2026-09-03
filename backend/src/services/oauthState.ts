import { createHash, randomBytes, randomUUID } from "node:crypto";
import jwt from "jsonwebtoken";
import { config } from "../config/index.js";
import { connection as redis } from "../jobs/queue.js";
import type { CloudType, ConnectState } from "../types/connections.js";

const STATE_TTL_SECONDS = 10 * 60;

/**
 * PKCE (RFC 7636). Technically redundant here since this is a confidential client (the code
 * exchange happens server-side with a client secret, which already prevents the "stolen
 * authorization code" attack PKCE defends against for public clients) — included anyway per spec,
 * as defense-in-depth, and because it costs nothing.
 */
function generatePkcePair(): { verifier: string; challenge: string } {
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

export interface ConnectAttempt {
  state: string;
  authorizeParams: { state: string; codeChallenge: string; codeChallengeMethod: "S256" };
}

/**
 * Issues a signed `state` (binds this OAuth round-trip to cloudType + the initiating operator,
 * short expiry) and stashes the PKCE verifier server-side (Redis, same TTL) keyed by the state's
 * nonce — the popup/browser never needs to carry the verifier, only the signed state.
 */
export async function startConnectAttempt(cloudType: CloudType, operatorId: string): Promise<ConnectAttempt> {
  const nonce = randomUUID();
  const payload: ConnectState = { cloudType, operatorId, nonce, iat: Math.floor(Date.now() / 1000) };
  const state = jwt.sign(payload, config.session.jwtSecret, {
    issuer: config.session.issuer,
    expiresIn: STATE_TTL_SECONDS,
  });

  const { verifier, challenge } = generatePkcePair();
  await redis.set(`oauth-pkce:${nonce}`, verifier, "EX", STATE_TTL_SECONDS);

  return { state, authorizeParams: { state, codeChallenge: challenge, codeChallengeMethod: "S256" } };
}

export class InvalidOAuthStateError extends Error {}

/** Validates + decodes `state` and retrieves (and consumes) the matching PKCE verifier. Throws InvalidOAuthStateError on any mismatch/expiry. */
export async function consumeConnectAttempt(state: string): Promise<{ claims: ConnectState; codeVerifier: string }> {
  let claims: ConnectState;
  try {
    claims = jwt.verify(state, config.session.jwtSecret, { issuer: config.session.issuer }) as unknown as ConnectState;
  } catch {
    throw new InvalidOAuthStateError("state is invalid or expired");
  }

  const key = `oauth-pkce:${claims.nonce}`;
  const codeVerifier = await redis.get(key);
  if (!codeVerifier) {
    throw new InvalidOAuthStateError("PKCE verifier expired or already used");
  }
  await redis.del(key); // single-use — a replayed callback with the same state can't re-exchange

  return { claims, codeVerifier };
}
