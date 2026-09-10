import type { gmail_v1 } from "googleapis";

/**
 * Gmail calls specific to Add Clouds / Manage Clouds enumeration — the Google equivalent of
 * graph/cloudEnumeration.ts's getUserMailSummary. Domain user listing itself reuses
 * graph/googleDriveEnumeration.ts's listDomainUsers/getGoogleUserById (Directory API, not
 * Gmail-specific) — this module only has the one Gmail-only call.
 */

export interface GmailMailboxSummary {
  itemCount: number;
}

/**
 * users.getProfile gives an exact messagesTotal with one call — no pagination, no enumeration.
 * Unlike My Drive's about.get, it has NO storage-bytes field at all (verified against the Gmail
 * API reference) — Gmail exposes no cheap per-mailbox storage figure, lagged or otherwise. Storage
 * is therefore always "not meaningful" for Gmail, same convention already used for Outlook mailboxes
 * (graph/cloudEnumeration.ts's getUserMailSummary) — show "—", never invent a value.
 */
export async function getMailboxSummary(gmail: gmail_v1.Gmail, userEmail: string): Promise<GmailMailboxSummary | null> {
  try {
    const res = await gmail.users.getProfile({ userId: userEmail });
    return { itemCount: res.data.messagesTotal ?? 0 };
  } catch (err) {
    // No Gmail mailbox provisioned for this user — not a sync failure, same convention as getUserDriveUsage.
    if ((err as { code?: number })?.code === 404) return null;
    throw err;
  }
}
