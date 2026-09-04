import { rawFetch } from "./client";

export type CloudType = "onedrive" | "sharepoint" | "teams";
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

export type M365ConnectMessage =
  | { type: "m365-connect-complete"; status: "success"; connectionId: string; cloudType: CloudType }
  | { type: "m365-connect-complete"; status: "error"; cloudType: CloudType | null; reason: string };

export function listManageClouds(): Promise<{ connections: ManageCloudsRow[] }> {
  return rawFetch("/api/clouds/manage");
}

export function listConnectionUsers(
  connectionId: string,
  opts: { status?: "failed" | "synced" | "pending"; cursor?: string } = {}
): Promise<{ users: ConnectionUserRow[]; nextCursor: string | null }> {
  const params = new URLSearchParams();
  if (opts.status) params.set("status", opts.status);
  if (opts.cursor) params.set("cursor", opts.cursor);
  const qs = params.toString();
  return rawFetch(`/api/clouds/${connectionId}/users${qs ? `?${qs}` : ""}`);
}

export function initCloudConnect(cloudType: CloudType): Promise<{ authorizeUrl: string; state: string }> {
  return rawFetch(`/api/clouds/${cloudType}/connect/init`, { method: "POST" });
}

export function resyncCloudConnection(connectionId: string): Promise<{ jobId: string; status: "queued" }> {
  return rawFetch(`/api/clouds/${connectionId}/resync`, { method: "POST" });
}

export function disconnectCloudConnection(connectionId: string): Promise<void> {
  return rawFetch(`/api/clouds/${connectionId}`, { method: "DELETE" });
}

/**
 * Opens the Microsoft admin-consent popup and resolves once it posts back a result (or is closed
 * without completing). Mirrors the reference product's popup-based connect flow — see
 * docs/azure-ad-app-registration.md §4a.
 */
export function openConnectPopup(authorizeUrl: string): Promise<M365ConnectMessage> {
  return new Promise((resolve) => {
    const popup = window.open(authorizeUrl, "m365-connect", "width=500,height=680,menubar=no,toolbar=no");

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
