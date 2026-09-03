import type { Client } from "@microsoft/microsoft-graph-client";

export interface DriveFile {
  id: string;
  path: string;
  sizeBytes: number;
  lastModified: string;
}

/** Walks a user's OneDrive, optionally filtered to files last modified before `olderThanCutoff`. */
export async function listUserFiles(
  client: Client,
  userId: string,
  olderThanCutoff?: string
): Promise<DriveFile[]> {
  const files: DriveFile[] = [];
  let url: string | undefined = `/users/${userId}/drive/root/children`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const item of res.value as any[]) {
      if (item.folder) continue; // v1 walks flat root + top-level files; recursive walk is a straightforward extension
      const lastModified = item.lastModifiedDateTime as string;
      if (olderThanCutoff && lastModified >= olderThanCutoff) continue;
      files.push({
        id: item.id,
        path: item.name,
        sizeBytes: item.size ?? 0,
        lastModified,
      });
    }
    url = res["@odata.nextLink"];
  }
  return files;
}

export async function deleteFile(client: Client, userId: string, fileId: string): Promise<void> {
  await client.api(`/users/${userId}/drive/items/${fileId}`).delete();
}
