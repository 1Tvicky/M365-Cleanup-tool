import type { Client } from "@microsoft/microsoft-graph-client";
import { ApiError } from "../../types/index.js";

/**
 * Teams Data Dump writes. Team/channel creation both use the already-granted Group.ReadWrite.All
 * application permission (docs/azure-ad-app-registration.md) — confirmed against Microsoft's own
 * permission tables for `PUT /groups/{id}/team` ("create team from group") and `POST /teams/{id}
 * /channels` ("create channel"), both of which list Group.ReadWrite.All as an accepted application
 * permission. No new Azure AD consent is required for teams/channels.
 *
 * Channel MESSAGES are deliberately not created here — see postChannelMessageUnsupported below.
 */

export interface CreatedTeam {
  groupId: string;
  displayName: string;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Creates the M365 Group backing a Team, with `ownerUserId` as its sole owner (Graph requires at
 * least one owner to later provision a Team on top of it — application-permission group creation
 * accepts exactly one member in the owners collection, per Microsoft's documented constraint).
 */
async function createGroup(client: Client, displayName: string, mailNickname: string, ownerUserId: string, visibility: "Private" | "Public", description?: string): Promise<string> {
  const group: any = await client.api("/groups").post({
    displayName,
    description,
    mailNickname,
    mailEnabled: true,
    securityEnabled: false,
    visibility,
    groupTypes: ["Unified"],
    "owners@odata.bind": [`https://graph.microsoft.com/v1.0/users/${ownerUserId}`],
  });
  return group.id;
}

/**
 * Provisions a Team on top of a just-created group. Microsoft's own documented behavior: a group
 * created less than ~15 minutes ago can 404 here because directory replication hasn't finished —
 * retrying a few times with a delay (their own recommendation) resolves this without the caller
 * needing to understand Azure AD replication timing.
 */
async function createTeamFromGroup(client: Client, groupId: string): Promise<void> {
  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await client.api(`/groups/${groupId}/team`).put({});
      return;
    } catch (err) {
      const status = (err as { statusCode?: number })?.statusCode;
      if (status === 404 && attempt < maxAttempts) {
        await sleep(10_000);
        continue;
      }
      throw err;
    }
  }
}

export async function createTeam(
  client: Client,
  displayName: string,
  mailNickname: string,
  ownerUserId: string,
  options?: { visibility?: "private" | "public"; description?: string }
): Promise<CreatedTeam> {
  const graphVisibility = options?.visibility === "public" ? "Public" : "Private";
  const groupId = await createGroup(client, displayName, mailNickname, ownerUserId, graphVisibility, options?.description);
  await createTeamFromGroup(client, groupId);
  return { groupId, displayName };
}

export interface CreatedChannel {
  id: string;
  displayName: string;
  membershipType: "standard" | "private";
}

/**
 * `membershipType: "private"` is what actually makes an explicit channel-member list meaningful
 * (spec §17: "do not assume Team membership and channel membership are always identical") — every
 * standard channel automatically includes the whole team's membership by Microsoft's own design, so
 * addChannelMember below is only ever called for private channels; a caller asking for explicit
 * channel members on a standard channel gets that documented instead of a silently-ignored call.
 */
export async function createChannel(client: Client, teamId: string, displayName: string, options?: { description?: string; membershipType?: "standard" | "private"; ownerUserId?: string }): Promise<CreatedChannel> {
  const membershipType = options?.membershipType ?? "standard";
  const body: Record<string, unknown> = { displayName, description: options?.description, membershipType };
  if (membershipType === "private" && options?.ownerUserId) {
    body.members = [
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${options.ownerUserId}`,
      },
    ];
  }
  const channel: any = await client.api(`/teams/${teamId}/channels`).post(body);
  return { id: channel.id, displayName: channel.displayName, membershipType };
}

/** Adds an existing tenant user as a team member (spec §14: "use existing tenant users as members. Do not create external users"). */
export async function addTeamMember(client: Client, teamId: string, userId: string): Promise<void> {
  await client.api(`/teams/${teamId}/members`).post({
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles: [],
    "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
  });
}

/** Explicit channel membership — only meaningful (and only ever called) for private channels; see createChannel's membershipType note. */
export async function addChannelMember(client: Client, teamId: string, channelId: string, userId: string): Promise<void> {
  await client.api(`/teams/${teamId}/channels/${channelId}/members`).post({
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles: [],
    "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
  });
}

/**
 * Channel message/reply creation is NOT implemented as a live Graph call, on purpose — this is a
 * Microsoft Graph platform restriction, not an app-imposed limit (spec §21: "do not simulate
 * unsupported Microsoft Graph operations as successful"). `ChannelMessage.Send` is a delegated-only
 * permission; application-permission POSTs to channel messages are restricted by Microsoft to
 * "migration mode" team imports (a team explicitly put into a migration state via a separate,
 * heavier provisioning flow), which this app does not implement. Calling this always throws a clear
 * 501 rather than silently no-opping or faking a message id.
 */
export function postChannelMessageUnsupported(): never {
  throw new ApiError(
    501,
    "UNSUPPORTED_OPERATION",
    "Posting Teams channel messages is not supported under application permissions outside a migration-mode team import. Teams and channels are created for real; messages/replies are configured for planning purposes only. See docs/data-dump-api.md."
  );
}
