import type { Client } from "@microsoft/microsoft-graph-client";
import { createFolder, uploadFile, type CreatedFile, type CreatedFolder } from "./uploadHelper.js";

/**
 * SharePoint Data Dump writes — folder/file/library creation uses the already-granted
 * Sites.ReadWrite.All application permission (docs/azure-ad-app-registration.md), the same
 * permission the existing Cleaning module already uses to enumerate/delete SharePoint content. No
 * new Azure AD consent is required for those.
 *
 * New SITE creation (createSharePointSite below) is different: Microsoft Graph only exposes
 * `POST /sites` on the BETA endpoint (not v1.0) and requires the separate `Sites.Create.All`
 * permission — this app's Azure AD app registration does not currently request it (see
 * docs/azure-ad-app-registration.md's new "Data Dump: new SharePoint sites" section). The function
 * is implemented for real (so it works the moment that permission is granted and consented) but
 * will 403 in any environment that hasn't granted it — this is a genuine, reported Graph
 * permission/authorization error surfaced to the caller, never simulated as success.
 */

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export interface CreatedSite {
  siteId: string;
  webUrl: string;
  displayName: string;
  template: "teamSiteWithoutMicrosoft365Group" | "communicationSite";
}

/** The tenant's root SharePoint URL (e.g. https://contoso.sharepoint.com) — the base a new site's webUrl is built from. */
async function getTenantRootSiteUrl(client: Client): Promise<string> {
  const root: any = await client.api("/sites/root").get();
  return (root.webUrl as string).replace(/\/$/, "");
}

/**
 * Creates a brand-new SharePoint site (spec §6/§8/§9). `template` maps to Graph's own
 * `siteTemplateType` values, confirmed against Microsoft's beta documentation: "communicationSite"
 * → `sitepagepublishing`, "teamSiteWithoutMicrosoft365Group" → `sts` — never a value Graph doesn't
 * document. Site creation is asynchronous on Microsoft's side (returns 202 + an operation to poll);
 * this polls `GET /sites/{id}` a few times with a delay, the same defensive shape as
 * teamsCreation.ts's createTeamFromGroup retry, since a just-created site can briefly 404 before
 * provisioning finishes.
 */
export async function createSharePointSite(
  client: Client,
  options: { displayName: string; urlSlug: string; template: "teamSiteWithoutMicrosoft365Group" | "communicationSite"; description?: string; ownerEmail: string }
): Promise<CreatedSite> {
  const rootUrl = await getTenantRootSiteUrl(client);
  const webUrl = `${rootUrl}/sites/${options.urlSlug}`;
  const graphTemplate = options.template === "communicationSite" ? "sitepagepublishing" : "sts";

  const created: any = await client.api("/sites").version("beta").post({
    name: options.displayName,
    webUrl,
    template: graphTemplate,
    description: options.description,
    ownerIdentityToResolve: { email: options.ownerEmail },
  });

  const siteId = created.id as string;
  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await client.api(`/sites/${siteId}`).get();
      break;
    } catch (err) {
      const status = (err as { statusCode?: number })?.statusCode;
      if (status === 404 && attempt < maxAttempts) {
        await sleep(10_000);
        continue;
      }
      throw err;
    }
  }

  return { siteId, webUrl: created.webUrl ?? webUrl, displayName: options.displayName, template: options.template };
}

export interface CreatedLibrary {
  listId: string;
  driveId: string;
  displayName: string;
}

/** Creates a new document library (a SharePoint list with the documentLibrary template) under an existing site, then resolves its backing drive id for subsequent folder/file creation. */
export async function createDocumentLibrary(client: Client, siteId: string, displayName: string): Promise<CreatedLibrary> {
  const list: any = await client.api(`/sites/${siteId}/lists`).post({
    displayName,
    list: { template: "documentLibrary" },
  });
  const drive: any = await client.api(`/sites/${siteId}/lists/${list.id}/drive`).get();
  return { listId: list.id, driveId: drive.id, displayName };
}

function driveBasePath(siteId: string, driveId: string): string {
  return `/sites/${siteId}/drives/${driveId}`;
}

export async function createSharePointFolder(client: Client, siteId: string, driveId: string, parentItemId: string | null, name: string): Promise<CreatedFolder> {
  return createFolder(client, driveBasePath(siteId, driveId), parentItemId, name);
}

export async function uploadSharePointFile(
  client: Client,
  siteId: string,
  driveId: string,
  parentItemId: string,
  fileName: string,
  content: Buffer,
  mimeType: string,
  fileSystemInfo?: { createdDateTime?: string; lastModifiedDateTime?: string }
): Promise<CreatedFile> {
  return uploadFile(client, driveBasePath(siteId, driveId), parentItemId, fileName, content, mimeType, fileSystemInfo);
}

export interface PermissionGrantResult {
  upn: string;
  role: "read" | "write" | "owner";
  ok: boolean;
  errorMessage?: string;
}

/**
 * Best-effort site-level permission grant for an existing tenant user (spec §13: "do not create
 * external users... if permissions cannot be created through the available application permissions,
 * clearly report that limitation"). A failure here is reported per-user, never thrown — one failed
 * grant must not fail the whole SharePoint workload.
 */
export async function grantSitePermission(client: Client, siteId: string, upn: string, role: "read" | "write" | "owner"): Promise<PermissionGrantResult> {
  const roles = role === "owner" ? ["owner"] : role === "write" ? ["write"] : ["read"];
  try {
    await client.api(`/sites/${siteId}/permissions`).post({
      roles,
      grantedToIdentities: [{ user: { userPrincipalName: upn } }],
    });
    return { upn, role, ok: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { upn, role, ok: false, errorMessage: message };
  }
}
