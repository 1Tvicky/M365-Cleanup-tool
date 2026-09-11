import { readFileSync } from "node:fs";
import { google } from "googleapis";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

/**
 * Google Workspace domain-wide-delegation auth layer — the Google equivalent of graph/client.ts,
 * but structurally different: Microsoft's app-only client-credentials model has ONE token per
 * customer tenant (any Graph call for that tenant reuses it); Google's domain-wide delegation
 * requires impersonating a SPECIFIC user per call (a JWT "subject"), so there is no single
 * per-tenant client — every Drive/Directory call needs its own impersonated client for whichever
 * user it's acting as. See docs/google-workspace-integration.md.
 */

const DIRECTORY_SCOPES = ["https://www.googleapis.com/auth/admin.directory.user.readonly"];
const DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"];

interface ServiceAccountKey {
  client_email: string;
  private_key: string;
}

let cachedKey: ServiceAccountKey | null = null;

/**
 * Lazily loads/parses the service account JSON key. Built lazily (like graph/client.ts's
 * getMsalApp) since this whole feature is optional until GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY_*
 * is configured on a given deployment — importing this module must never crash a server that
 * hasn't set up Google Workspace support yet.
 */
function getServiceAccountKey(): ServiceAccountKey {
  if (cachedKey) return cachedKey;
  const raw = config.googleWorkspace.serviceAccountKeyJson || (config.googleWorkspace.serviceAccountKeyPath ? readKeyFile(config.googleWorkspace.serviceAccountKeyPath) : "");
  if (!raw) {
    throw new ApiError(503, "GOOGLE_WORKSPACE_NOT_CONFIGURED", "Google Workspace access isn't configured on this deployment yet");
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new ApiError(500, "GOOGLE_WORKSPACE_KEY_INVALID", "The configured Google Workspace service account key isn't valid JSON");
  }
  const key = parsed as Partial<ServiceAccountKey>;
  if (!key.client_email || !key.private_key) {
    throw new ApiError(500, "GOOGLE_WORKSPACE_KEY_INVALID", "The configured Google Workspace service account key is missing client_email/private_key");
  }
  cachedKey = { client_email: key.client_email, private_key: key.private_key };
  return cachedKey;
}

function readKeyFile(path: string): string {
  return readFileSync(path, "utf8");
}

/** Instance-reuse cache only — the underlying google-auth-library JWT client refreshes its own access token internally, this just avoids rebuilding+re-authorizing a client on every call for the same (subject, scopes). */
const clientCache = new Map<string, InstanceType<typeof google.auth.JWT>>();

async function getImpersonatedClient(subjectEmail: string, scopes: string[]): Promise<InstanceType<typeof google.auth.JWT>> {
  const cacheKey = `${subjectEmail}::${scopes.join(",")}`;
  const cached = clientCache.get(cacheKey);
  if (cached) return cached;

  const key = getServiceAccountKey();
  const client = new google.auth.JWT({
    email: key.client_email,
    key: key.private_key,
    scopes,
    subject: subjectEmail,
  });
  await client.authorize();
  clientCache.set(cacheKey, client);
  return client;
}

/** A Directory API client impersonating the given admin — used for domain-wide user listing/lookup. */
export async function getDirectoryClientAs(adminEmail: string) {
  const auth = await getImpersonatedClient(adminEmail, DIRECTORY_SCOPES);
  return google.admin({ version: "directory_v1", auth });
}

/** A Drive API client impersonating the given user — used for that user's own My Drive content, and (with supportsAllDrives) Shared Drives. */
export async function getDriveClientAs(userEmail: string) {
  const auth = await getImpersonatedClient(userEmail, DRIVE_SCOPES);
  return google.drive({ version: "v3", auth });
}

// Full mail scope only — https://mail.google.com/ is the only Gmail scope that supports deleting
// messages (gmail.readonly/gmail.modify can list/label but not permanently delete or trash).
const GMAIL_SCOPES = ["https://mail.google.com/"];

/** A Gmail API client impersonating the given user — used for that user's own mailbox. */
export async function getGmailClientAs(userEmail: string) {
  const auth = await getImpersonatedClient(userEmail, GMAIL_SCOPES);
  return google.gmail({ version: "v1", auth });
}

// Chat's admin-level space/membership enumeration and per-member message access are deliberately
// separate scope sets (see graph/googleChatEnumeration.ts's header comment for why) — two distinct
// impersonated-client helpers rather than one. chat.admin.delete (added alongside the read-only
// admin scopes) is required for spaces.delete({useAdminAccess:true}) — verified against the Chat
// API reference; this must also be authorized as an OAuth scope for this service account in the
// Workspace admin console's domain-wide delegation settings before whole-Space deletion will work
// (see docs/google-workspace-integration.md), same as every other scope this app impersonates with.
const CHAT_ADMIN_SCOPES = [
  "https://www.googleapis.com/auth/chat.admin.spaces.readonly",
  "https://www.googleapis.com/auth/chat.admin.memberships.readonly",
  "https://www.googleapis.com/auth/chat.admin.delete",
];
const CHAT_MESSAGE_SCOPES = ["https://www.googleapis.com/auth/chat.messages"];

/** A Chat API client impersonating the domain admin, with admin-level scopes — used to enumerate every space and space membership in the domain (spaces.search/members.list with useAdminAccess:true), and to delete a whole Space (spaces.delete with useAdminAccess:true). Never used for message read/delete — see getChatClientAsMember. */
export async function getChatAdminClientAs(adminEmail: string) {
  const auth = await getImpersonatedClient(adminEmail, CHAT_ADMIN_SCOPES);
  return google.chat({ version: "v1", auth });
}

/** A Chat API client impersonating one HUMAN member of a space — used to read/delete messages in that space. Chat's messages.list/messages.delete have no admin-access bypass (verified against the API docs): only a real member's own delegated auth can read/delete a space's messages, so cleanup must impersonate a member, not the admin. See jobs/googleChatCleanupExecution.ts for how the member is chosen. */
export async function getChatClientAsMember(memberEmail: string) {
  const auth = await getImpersonatedClient(memberEmail, CHAT_MESSAGE_SCOPES);
  return google.chat({ version: "v1", auth });
}

export class GoogleDelegationError extends Error {}

/**
 * Confirms domain-wide delegation has actually been granted for `domain` by attempting a real
 * Directory API call impersonating `adminEmail`, and returns the identity Google's own API reports
 * back — never the submitted domain string itself. This is the Google equivalent of
 * exchangeM365ConnectCode deriving identity from the token's `tid` claim rather than trusting a
 * frontend-supplied tenant id: here there's no OAuth callback to derive identity from, so the
 * equivalent "don't trust the client" move is verifying the claimed domain by successfully calling
 * an API that only succeeds if delegation was genuinely granted for that exact domain/admin.
 */
export async function verifyDomainDelegation(domain: string, adminEmail: string): Promise<{ customerId: string; primaryDomain: string; adminDisplayName: string | null }> {
  let directory;
  try {
    directory = await getDirectoryClientAs(adminEmail);
  } catch (err) {
    console.warn(`[google][auth] failed to build impersonated client domain=${domain}`);
    throw new GoogleDelegationError("Could not impersonate the given admin — check the admin email and that domain-wide delegation is configured.");
  }

  try {
    const res = await directory.users.get({ userKey: adminEmail });
    const user = res.data;
    if (!user.customerId || !user.primaryEmail) {
      throw new GoogleDelegationError("Google did not return a customer id for this admin.");
    }
    const domainFromEmail = user.primaryEmail.split("@")[1] ?? domain;
    return {
      customerId: user.customerId,
      primaryDomain: domainFromEmail,
      adminDisplayName: user.name?.fullName ?? null,
    };
  } catch (err) {
    if (err instanceof GoogleDelegationError) throw err;
    console.warn(`[google][auth] delegation verification failed domain=${domain}`);
    throw new GoogleDelegationError(
      "Google denied access to this domain. Make sure a Workspace super admin has authorized this app's service account in Admin Console → Security → API controls → Domain-wide delegation, with the required scopes, and that the admin email is correct."
    );
  }
}

export function invalidateSubjectCache(subjectEmail: string): void {
  for (const key of clientCache.keys()) {
    if (key.startsWith(`${subjectEmail}::`)) clientCache.delete(key);
  }
}
