import type { Client } from "@microsoft/microsoft-graph-client";

export interface SiteSummary {
  id: string;
  webUrl: string;
  displayName: string;
  storageUsedBytes: number;
}

export interface DocumentLibrary {
  id: string;
  name: string;
  sizeBytes: number;
}

export async function listSites(client: Client, search: string, limit: number): Promise<SiteSummary[]> {
  const res = await client.api("/sites").search(search || "*").top(limit).get();
  return (res.value as any[]).map((s) => ({
    id: s.id,
    webUrl: s.webUrl,
    displayName: s.displayName ?? s.name,
    storageUsedBytes: 0, // filled from /reports/getSiteUsageDetail in preview, not on every list call
  }));
}

export async function listDocumentLibraries(client: Client, siteId: string): Promise<DocumentLibrary[]> {
  const res = await client.api(`/sites/${siteId}/drives`).get();
  return (res.value as any[]).map((d) => ({
    id: d.id,
    name: d.name,
    sizeBytes: d.quota?.used ?? 0,
  }));
}

/** Deletes a document library (drive). Recycle-bin recoverability depends on tenant retention — see docs/rollback-safety.md. */
export async function deleteDocumentLibrary(client: Client, siteId: string, driveId: string): Promise<void> {
  await client.api(`/sites/${siteId}/drives/${driveId}`).delete();
}
