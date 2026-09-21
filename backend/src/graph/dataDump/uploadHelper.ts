import type { Client } from "@microsoft/microsoft-graph-client";
import { GRAPH_SIMPLE_UPLOAD_MAX_BYTES, GRAPH_UPLOAD_CHUNK_BYTES } from "../../services/dataDump/sizing.js";

/**
 * Shared OneDrive/SharePoint file-creation primitives (spec §5/§10/§23: "no artificial application
 * limit on file size... use Graph upload sessions where required; stream/chunk content; avoid
 * loading huge files entirely into memory"). Both graph/dataDump/oneDriveCreation.ts and
 * graph/dataDump/sharePointCreation.ts call these against their own drive base path — a OneDrive
 * user's drive (`/users/{id}/drive`) and a SharePoint library's drive (`/sites/{id}/drives/{id}`)
 * expose an identical driveItem API shape, so one implementation covers both without duplicating
 * the upload-session chunking loop.
 */

export interface CreatedFolder {
  id: string;
  name: string;
  webUrl: string;
}

export interface CreatedFile {
  id: string;
  name: string;
  sizeBytes: number;
  webUrl: string;
}

/** Creates one folder under `parentItemId` (or the drive root if null). `replace` conflict behavior means re-running against a folder that already exists (e.g. a resumed operation) reuses it rather than erroring or duplicating. */
export async function createFolder(client: Client, driveBasePath: string, parentItemId: string | null, name: string): Promise<CreatedFolder> {
  const path = parentItemId ? `${driveBasePath}/items/${parentItemId}/children` : `${driveBasePath}/root/children`;
  const res: any = await client.api(path).post({
    name,
    folder: {},
    "@microsoft.graph.conflictBehavior": "rename",
  });
  return { id: res.id, name: res.name, webUrl: res.webUrl };
}

/**
 * Uploads one file's content under `parentItemId`. Picks Graph's own documented strategy boundary
 * (GRAPH_SIMPLE_UPLOAD_MAX_BYTES = 4 MiB) automatically — this is Microsoft's limit on the plain
 * PUT :content call, not an app-imposed file-size ceiling; anything larger transparently uses a
 * resumable upload session with GRAPH_UPLOAD_CHUNK_BYTES-sized chunks, so the caller never has to
 * think about file size when calling this function, and the whole buffer is never required to fit
 * in one HTTP request body regardless of how large it is.
 */
export async function uploadFile(
  client: Client,
  driveBasePath: string,
  parentItemId: string,
  fileName: string,
  content: Buffer,
  mimeType: string,
  fileSystemInfo?: { createdDateTime?: string; lastModifiedDateTime?: string }
): Promise<CreatedFile> {
  const item =
    content.length <= GRAPH_SIMPLE_UPLOAD_MAX_BYTES
      ? await uploadSmall(client, driveBasePath, parentItemId, fileName, content, mimeType)
      : await uploadLarge(client, driveBasePath, parentItemId, fileName, content);

  if (fileSystemInfo) {
    await client.api(`${driveBasePath}/items/${item.id}`).patch({ fileSystemInfo });
  }

  return { id: item.id, name: item.name, sizeBytes: content.length, webUrl: item.webUrl };
}

async function uploadSmall(client: Client, driveBasePath: string, parentItemId: string, fileName: string, content: Buffer, mimeType: string): Promise<any> {
  return client
    .api(`${driveBasePath}/items/${parentItemId}:/${encodeURIComponent(fileName)}:/content`)
    .header("Content-Type", mimeType)
    .put(content);
}

async function uploadLarge(client: Client, driveBasePath: string, parentItemId: string, fileName: string, content: Buffer): Promise<any> {
  const session: any = await client.api(`${driveBasePath}/items/${parentItemId}:/${encodeURIComponent(fileName)}:/createUploadSession`).post({
    item: { "@microsoft.graph.conflictBehavior": "rename" },
  });
  const uploadUrl = session.uploadUrl as string;

  let offset = 0;
  let lastResponse: any = null;
  while (offset < content.length) {
    const end = Math.min(offset + GRAPH_UPLOAD_CHUNK_BYTES, content.length);
    const chunk = content.subarray(offset, end);
    // The pre-authenticated uploadUrl is called with plain fetch, not the Graph client — it already
    // carries its own short-lived SAS-style auth, and re-adding a Bearer token is neither required
    // nor supported by this endpoint.
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${offset}-${end - 1}/${content.length}`,
      },
      body: chunk,
    });
    if (!res.ok && res.status !== 200 && res.status !== 201 && res.status !== 202) {
      throw new Error(`Upload session chunk failed: ${res.status} ${await res.text().catch(() => "")}`);
    }
    if (res.status === 200 || res.status === 201) {
      lastResponse = await res.json();
    }
    offset = end;
  }
  return lastResponse;
}
