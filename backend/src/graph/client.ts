import { readFileSync } from "node:fs";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

// One MSAL confidential-client app for the whole service — tenant identity is supplied per
// acquireToken call, not per app instance. See docs/azure-ad-app-registration.md §5: application
// permissions use client-credentials, so there is no per-tenant refresh token to persist.
//
// Built lazily: MSAL throws synchronously if neither a client secret nor certificate is present,
// and this whole feature is optional until the Azure AD app registration exists — an environment
// with no AZURE_* configured should still start the server and serve password login/cleanup UI,
// not crash on import.
let msalApp: ConfidentialClientApplication | null = null;

function getMsalApp(): ConfidentialClientApplication {
  const hasCert = Boolean(config.azure.certThumbprint && config.azure.certPrivateKeyPath);
  if (!config.azure.clientId || (!hasCert && !config.azure.clientSecret)) {
    throw new ApiError(503, "M365_TENANT_ACCESS_NOT_CONFIGURED", "M365 tenant access isn't configured on this deployment yet");
  }
  if (!msalApp) {
    msalApp = new ConfidentialClientApplication({
      auth: {
        clientId: config.azure.clientId,
        ...(hasCert
          ? {
              clientCertificate: {
                thumbprint: config.azure.certThumbprint!,
                privateKey: readPrivateKey(config.azure.certPrivateKeyPath!),
              },
            }
          : { clientSecret: config.azure.clientSecret }),
      },
    });
  }
  return msalApp;
}

function readPrivateKey(path: string): string {
  return readFileSync(path, "utf8");
}

const tokenCache = new Map<string, { token: string; expiresAt: number }>();

async function getAccessTokenForTenant(m365TenantId: string): Promise<string> {
  const cached = tokenCache.get(m365TenantId);
  if (cached && cached.expiresAt > Date.now() + 60_000) return cached.token;

  const result = await getMsalApp().acquireTokenByClientCredential({
    scopes: [GRAPH_SCOPE],
    authority: `https://login.microsoftonline.com/${m365TenantId}`,
  });

  if (!result?.accessToken) {
    throw new Error(`Failed to acquire Graph token for tenant ${m365TenantId}`);
  }

  tokenCache.set(m365TenantId, {
    token: result.accessToken,
    expiresAt: result.expiresOn?.getTime() ?? Date.now() + 55 * 60_000,
  });
  return result.accessToken;
}

/** Returns a Graph client scoped to one customer tenant's application-permission token. */
export async function graphClientForTenant(m365TenantId: string): Promise<Client> {
  const token = await getAccessTokenForTenant(m365TenantId);
  return Client.init({
    authProvider: (done) => done(null, token),
  });
}

export function invalidateTenantTokenCache(m365TenantId: string): void {
  tokenCache.delete(m365TenantId);
}
