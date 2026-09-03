import { readFileSync } from "node:fs";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

/**
 * The Add Clouds connect flow (docs/azure-ad-app-registration.md §4a) — a delegated
 * authorization-code exchange used ONLY to identify the connecting admin and trigger tenant-wide
 * admin consent. Reuses the SAME Azure AD app registration as graph/client.ts's app-only
 * client-credentials client (same clientId/secret), just a different registered redirect URI and
 * a different MSAL request shape (auth-code vs. client-credentials). Never share tokens between
 * the two — this module's tokens are delegated (admin-scoped), graph/client.ts's are app-only.
 *
 * Built lazily, same reasoning as graph/client.ts and microsoftUserAuth.ts: an unconfigured
 * deployment should start and serve everything else, not crash on import.
 */
let msalApp: ConfidentialClientApplication | null = null;

function getMsalApp(): ConfidentialClientApplication {
  const hasCert = Boolean(config.azure.certThumbprint && config.azure.certPrivateKeyPath);
  if (!config.azure.clientId || (!hasCert && !config.azure.clientSecret)) {
    throw new ApiError(503, "M365_CONNECT_NOT_CONFIGURED", "M365 cloud connections aren't configured on this deployment yet");
  }
  if (!config.azure.connectRedirectUri) {
    throw new ApiError(503, "M365_CONNECT_NOT_CONFIGURED", "M365_CONNECT_REDIRECT_URI is not set");
  }
  if (!msalApp) {
    msalApp = new ConfidentialClientApplication({
      auth: {
        clientId: config.azure.clientId,
        ...(hasCert
          ? {
              clientCertificate: {
                thumbprint: config.azure.certThumbprint!,
                privateKey: readFileSync(config.azure.certPrivateKeyPath!, "utf8"),
              },
            }
          : { clientSecret: config.azure.clientSecret }),
      },
    });
  }
  return msalApp;
}

const CONNECT_SCOPES = ["https://graph.microsoft.com/.default", "offline_access", "openid", "profile"];

export async function getM365ConnectAuthorizeUrl(opts: {
  state: string;
  codeChallenge: string;
  codeChallengeMethod: "S256";
}): Promise<string> {
  return getMsalApp().getAuthCodeUrl({
    scopes: CONNECT_SCOPES,
    redirectUri: config.azure.connectRedirectUri,
    state: opts.state,
    codeChallenge: opts.codeChallenge,
    codeChallengeMethod: opts.codeChallengeMethod,
    // Triggers Microsoft's native admin-consent screen if this tenant hasn't granted the app's
    // configured permissions yet — we never build our own consent UI (see reference flow, step 3).
    prompt: "admin_consent",
  });
}

export interface M365DelegatedIdentity {
  m365TenantId: string;
  tenantDomain: string;
  adminUpn: string;
  adminDisplayName: string;
  refreshToken: string;
  accessTokenExpiresOn: Date | null;
}

/** Exchanges the authorization code and resolves the connecting admin's identity + tenant. */
export async function exchangeM365ConnectCode(code: string, codeVerifier: string): Promise<M365DelegatedIdentity> {
  const result = await getMsalApp().acquireTokenByCode({
    code,
    codeVerifier,
    scopes: CONNECT_SCOPES,
    redirectUri: config.azure.connectRedirectUri,
  });

  if (!result?.accessToken) {
    throw new Error("M365 connect code exchange did not return an access token");
  }

  // organization + me — identity capture only (docs/azure-ad-app-registration.md §5), never used
  // for enumeration.
  const graphFetch = async (path: string): Promise<any> => {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${result.accessToken}` },
    });
    if (!res.ok) throw new Error(`Graph identity call ${path} failed: ${res.status}`);
    return res.json();
  };

  const [org, me] = await Promise.all([
    graphFetch("/organization").then((r) => r.value?.[0]),
    graphFetch("/me"),
  ]);

  if (!org?.id) throw new Error("Graph /organization response missing tenant id");

  // MSAL Node's public AuthenticationResult type doesn't expose the raw refresh token (it's
  // managed internally in the token cache) unless a custom cache plugin is wired up. Surfacing it
  // here for encryption/storage requires reading the in-memory token cache directly.
  const tokenCache = getMsalApp().getTokenCache();
  const cacheJson = JSON.parse(await tokenCache.serialize());
  const refreshTokenEntry = Object.values(cacheJson.RefreshToken ?? {})[0] as { secret?: string } | undefined;
  if (!refreshTokenEntry?.secret) {
    throw new Error("No refresh token present in MSAL cache after code exchange — is offline_access scope granted?");
  }

  return {
    m365TenantId: org.id,
    tenantDomain: org.verifiedDomains?.find((d: { isDefault?: boolean }) => d.isDefault)?.name ?? org.displayName ?? org.id,
    adminUpn: me.userPrincipalName ?? me.mail,
    adminDisplayName: me.displayName ?? me.userPrincipalName ?? me.mail,
    refreshToken: refreshTokenEntry.secret,
    accessTokenExpiresOn: result.expiresOn,
  };
}
