import type { Client } from "@microsoft/microsoft-graph-client";
import { createFolder, uploadFile, type CreatedFile, type CreatedFolder } from "./uploadHelper.js";

/**
 * OneDrive Data Dump writes — uses the already-granted Files.ReadWrite.All application permission
 * (docs/azure-ad-app-registration.md), the same permission the existing Cleaning module already
 * uses to enumerate/delete OneDrive content. No new Azure AD consent is required for this workload.
 */

function driveBasePath(userId: string): string {
  return `/users/${userId}/drive`;
}

export async function createOneDriveFolder(client: Client, userId: string, parentItemId: string | null, name: string): Promise<CreatedFolder> {
  return createFolder(client, driveBasePath(userId), parentItemId, name);
}

export async function uploadOneDriveFile(
  client: Client,
  userId: string,
  parentItemId: string,
  fileName: string,
  content: Buffer,
  mimeType: string,
  fileSystemInfo?: { createdDateTime?: string; lastModifiedDateTime?: string }
): Promise<CreatedFile> {
  return uploadFile(client, driveBasePath(userId), parentItemId, fileName, content, mimeType, fileSystemInfo);
}
