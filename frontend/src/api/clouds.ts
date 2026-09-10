import { rawFetch } from "./client";

export type CloudType = "onedrive" | "sharepoint" | "teams" | "outlook" | "google_my_drive" | "shared_drive" | "google_chat" | "gmail";
export type ConnectionStatus = "connecting" | "active" | "error" | "needs_reauth" | "disconnected";

export interface ManageCloudsRow {
  id: string;
  cloudType: CloudType;
  iconKey: CloudType;
  displayName: string;
  adminEmail: string;
  adminDisplayName: string | null;
  tenantDomain: string;
  totalUsers: number;
  processedUsers: number;
  addedUsers: number;
  notAddedUsers: number;
  percent: number;
  status: ConnectionStatus;
  multiUser: true;
  connectedAt: string | null;
  lastSyncedAt: string | null;
  lastError: string | null;
}

export interface ConnectionUserRow {
  id: string;
  graphUserId: string;
  upn: string;
  displayName: string | null;
  storageUsedBytes: number;
  itemCount: number;
  syncStatus: "pending" | "synced" | "failed";
  lastSyncedAt: string | null;
  errorMessage: string | null;
}

/** One resource available to sync — a live Graph id, not a connection_users row id (the browse step must work even before anything's ever been synced). */
export interface AvailableResourceRow {
  id: string;
  displayName: string;
  secondary?: string;
}

export type SyncJobResourceStatus = "pending" | "processing" | "completed" | "failed" | "cancelled";

/** One selected resource's progress within a resource-scoped sync run. */
export interface SyncJobResourceRow {
  id: string;
  graphResourceId: string;
  displayName: string;
  secondary: string | null;
  status: SyncJobResourceStatus;
  errorMessage: string | null;
  startedAt: string | null;
  completedAt: string | null;
}

export interface ResyncSkippedResource {
  id: string;
  reason: "not_found" | "already_syncing";
}

export type M365ConnectMessage =
  | { type: "m365-connect-complete"; status: "success"; connectionId: string; cloudType: CloudType }
  | { type: "m365-connect-complete"; status: "error"; cloudType: CloudType | null; reason: string };

export type GoogleConnectMessage =
  | { type: "google-connect-complete"; status: "success"; connectionId: string; cloudType: CloudType }
  | { type: "google-connect-complete"; status: "error"; cloudType: CloudType | null; reason: string };

export function listManageClouds(): Promise<{ connections: ManageCloudsRow[] }> {
  return rawFetch("/api/clouds/manage");
}

export function listConnectionUsers(
  connectionId: string,
  opts: { status?: "failed" | "synced" | "pending"; cursor?: string; limit?: number } = {}
): Promise<{ users: ConnectionUserRow[]; nextCursor: string | null }> {
  const params = new URLSearchParams();
  if (opts.status) params.set("status", opts.status);
  if (opts.cursor) params.set("cursor", opts.cursor);
  if (opts.limit) params.set("limit", String(opts.limit));
  const qs = params.toString();
  return rawFetch(`/api/clouds/${connectionId}/users${qs ? `?${qs}` : ""}`);
}

/** Not a rawFetch call — the export is a CSV file download (Content-Disposition: attachment), which a plain navigation/anchor click handles natively (same pattern as api/cleaning.ts's cleanupReportUrl). */
export function exportConnectionUsersUrl(connectionId: string): string {
  return `/api/clouds/${connectionId}/users/export`;
}

export function initCloudConnect(cloudType: CloudType): Promise<{ authorizeUrl: string; state: string }> {
  return rawFetch(`/api/clouds/${cloudType}/connect/init`, { method: "POST" });
}

/**
 * Google Workspace's own connect-init endpoint — same shape as initCloudConnect above (returns a
 * popup authorizeUrl), just a different base path since Google's OAuth client/callback are
 * entirely separate from the M365 one. The popup only confirms identity; actual data access is
 * verified server-side via domain-wide delegation before the connection is created — see
 * routes/googleConnections.ts.
 */
export function initGoogleCloudConnect(cloudType: CloudType): Promise<{ authorizeUrl: string; state: string }> {
  return rawFetch(`/api/google-clouds/${cloudType}/connect/init`, { method: "POST" });
}

/**
 * Omit `resourceIds` (or pass none) for the plain "sync everything" action the existing Resync icon
 * uses; pass a non-empty array to scope the run to just those resources (the "Sync Selected" flow).
 * `skipped` reports any requested id the backend rejected (not found for this connection, or already
 * being synced by another in-flight job) — the caller should surface those, never assume acceptance.
 */
export function resyncCloudConnection(
  connectionId: string,
  resourceIds?: string[]
): Promise<{ jobId: string; status: "queued"; acceptedCount: number; skipped: ResyncSkippedResource[] }> {
  return rawFetch(`/api/clouds/${connectionId}/resync`, {
    method: "POST",
    body: resourceIds ? JSON.stringify({ resourceIds }) : undefined,
  });
}

/** The browse step for resource-level sync — every resource this workload currently has in Microsoft 365 (live, cached briefly server-side), not just what's already synced. Same page/pageSize convention as api/cleaning.ts's discovery routes, so it plugs directly into DiscoveryTable. */
export function listAvailableResources(
  connectionId: string,
  opts: { search?: string; page?: number; pageSize?: number } = {}
): Promise<{ resources: AvailableResourceRow[]; total: number; page: number; pageSize: number }> {
  const params = new URLSearchParams();
  if (opts.search) params.set("search", opts.search);
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return rawFetch(`/api/clouds/${connectionId}/available-resources?${params.toString()}`);
}

/** Per-resource progress for one resource-scoped sync run — empty for a legacy/tenant-wide run (pre-this-change history, or a plain Resync), the caller falls back to the aggregate total_users/processed_users bar in that case. */
export function getSyncJobResources(connectionId: string, jobId: string): Promise<{ resources: SyncJobResourceRow[] }> {
  return rawFetch(`/api/clouds/${connectionId}/sync-jobs/${jobId}/resources`);
}

export function disconnectCloudConnection(connectionId: string): Promise<void> {
  return rawFetch(`/api/clouds/${connectionId}`, { method: "DELETE" });
}

/**
 * Opens the Microsoft admin-consent popup and resolves once it posts back a result (or is closed
 * without completing). Mirrors the reference product's popup-based connect flow — see
 * docs/azure-ad-app-registration.md §4a.
 */
export function openConnectPopup(authorizeUrl: string, cloudType: CloudType): Promise<M365ConnectMessage> {
  return new Promise((resolve) => {
    // Name must be unique per in-flight connect: window.open reuses/navigates an existing window
    // with the same name instead of opening a new one, which would silently hijack an already-open
    // popup (e.g. OneDrive's) the moment a second connect (e.g. SharePoint) was started.
    const windowName = `m365-connect-${cloudType}-${Date.now()}`;
    const popup = window.open(authorizeUrl, windowName, "width=500,height=680,menubar=no,toolbar=no");

    function cleanup() {
      window.removeEventListener("message", onMessage);
      clearInterval(pollClosed);
    }

    function onMessage(event: MessageEvent) {
      // Validate by window identity, not origin: the popup is served from the backend's origin
      // (different from this frontend's own origin in dev, and not guaranteed to match in
      // production either), so comparing event.origin to window.location.origin here would always
      // fail. We already hold a reference to the exact window we opened — trust messages from it.
      if (event.source !== popup) return;
      const data = event.data as M365ConnectMessage;
      if (data?.type !== "m365-connect-complete") return;
      cleanup();
      resolve(data);
    }

    window.addEventListener("message", onMessage);

    // The admin can also just close the popup manually without Microsoft ever redirecting back —
    // that never fires postMessage, so poll for the window closing as a fallback.
    const pollClosed = window.setInterval(() => {
      if (popup?.closed) {
        cleanup();
        resolve({ type: "m365-connect-complete", status: "error", cloudType: null, reason: "closed" });
      }
    }, 500);
  });
}

/**
 * Same shape as openConnectPopup above (own copy, not a shared parameterized helper — see this
 * file's own convention of a dedicated function per concern), for Google's OAuth popup instead of
 * Microsoft's. Each Google workload tile connects independently — its own popup, own OAuth
 * consent — even if another Google workload is already connected for the same admin/domain.
 */
export function openGoogleConnectPopup(authorizeUrl: string, cloudType: CloudType): Promise<GoogleConnectMessage> {
  return new Promise((resolve) => {
    const windowName = `google-connect-${cloudType}-${Date.now()}`;
    const popup = window.open(authorizeUrl, windowName, "width=500,height=680,menubar=no,toolbar=no");

    function cleanup() {
      window.removeEventListener("message", onMessage);
      clearInterval(pollClosed);
    }

    function onMessage(event: MessageEvent) {
      if (event.source !== popup) return;
      const data = event.data as GoogleConnectMessage;
      if (data?.type !== "google-connect-complete") return;
      cleanup();
      resolve(data);
    }

    window.addEventListener("message", onMessage);

    const pollClosed = window.setInterval(() => {
      if (popup?.closed) {
        cleanup();
        resolve({ type: "google-connect-complete", status: "error", cloudType: null, reason: "closed" });
      }
    }, 500);
  });
}
