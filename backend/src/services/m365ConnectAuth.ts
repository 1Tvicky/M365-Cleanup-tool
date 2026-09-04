import { readFileSync } from "node:fs";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

/**
 * The Add Clouds connect flow (docs/azure-ad-app-registration.md §4a) — a delegated
 * authorization-code exchange used ONLY to identify the connecting admin and trigger tenant-wide
 * admin consent. Reuses the SAME dedicated Microsoft Entra app registration as graph/client.ts's
 * app-only client-credentials client (same clientId/secret — our own registration, never a
 * customer's and never another CloudFuze product's), just a different registered redirect URI and
 * a different MSAL request shape (auth-code vs. client-credentials). Never share tokens between
 * the two — this module's tokens are delegated (admin-scoped), graph/client.ts's are app-only.
 *
 * Built lazily, same reasoning as graph/client.ts and microsoftUserAuth.ts: an unconfigured
 * deployment should start and serve everything else, not crash on import.
 */
let msalApp: ConfidentialClientApplication | null = null;

function getMsalApp(): ConfidentialClientApplication {
  const hasCert = Boolean(config.microsoft.certThumbprint && config.microsoft.certPrivateKeyPath);
  if (!config.microsoft.clientId || (!hasCert && !config.microsoft.clientSecret)) {
    throw new ApiError(503, "M365_CONNECT_NOT_CONFIGURED", "M365 cloud connections aren't configured on this deployment yet");
  }
  if (!config.microsoft.connectRedirectUri) {
    throw new ApiError(503, "M365_CONNECT_NOT_CONFIGURED", "MICROSOFT_REDIRECT_URI is not set");
  }
  if (!msalApp) {
    msalApp = new ConfidentialClientApplication({
      auth: {
        clientId: config.microsoft.clientId,
        authority: config.microsoft.authority,
        ...(hasCert
          ? {
              clientCertificate: {
                thumbprint: config.microsoft.certThumbprint!,
                privateKey: readFileSync(config.microsoft.certPrivateKeyPath!, "utf8"),
              },
            }
          : { clientSecret: config.microsoft.clientSecret }),
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
    redirectUri: config.microsoft.connectRedirectUri,
    state: opts.state,
    codeChallenge: opts.codeChallenge,
    codeChallengeMethod: opts.codeChallengeMethod,
    // "admin_consent" is NOT a valid prompt value on the v2.0 /authorize endpoint (MSAL throws
    // ClientConfigurationError: invalid_prompt_value) — that value only exists on the older,
    // separate /adminconsent endpoint (the legacy flow in routes/auth.ts). "consent" is the
    // correct v2.0 value: combined with the .default scope, it shows Microsoft's native consent
    // screen, and a Global Admin additionally sees a "Consent on behalf of your organization"
    // option there — checking it is what actually grants the app's configured Application
    // permissions tenant-wide. We never build our own consent UI (see reference flow, step 3).
    prompt: "consent",
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
    redirectUri: config.microsoft.connectRedirectUri,
  });

  if (!result?.accessToken) {
    throw new Error("M365 connect code exchange did not return an access token");
  }

  // /me — identity capture only (docs/azure-ad-app-registration.md §5), never used for enumeration.
  // Backed by the default "User.Read" delegated permission every app registration has, so no extra
  // consent is needed for it.
  const graphFetch = async (path: string): Promise<any> => {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${result.accessToken}` },
    });
    if (!res.ok) throw new Error(`Graph identity call ${path} failed: ${res.status}`);
    return res.json();
  };

  const me = await graphFetch("/me");

  // The tenant ID comes straight off the ID token's `tid` claim (MSAL surfaces it as
  // `result.tenantId`) rather than GET /organization, which needs Organization.Read.All /
  // Directory.Read.All — permissions this app deliberately doesn't request (see the permission
  // table in docs/azure-ad-app-registration.md §3).
  if (!result.tenantId) throw new Error("Token response missing tenant id (tid claim)");

  // Prefer the admin's own UPN domain over /organization's "default" verified domain: a tenant's
  // isDefault flag marks whichever domain is set for new-UPN assignment, which frequently stays
  // pinned to the original *.onmicrosoft.com domain even after a custom domain (e.g. cloudfuze.co)
  // is verified and actually used for every real UPN — showing that placeholder instead of the
  // org's recognizable domain confused a real customer during testing. The admin's UPN domain is
  // always present (no extra Graph call/permission needed) and is what they actually sign in with.
  const upnDomain = (me.userPrincipalName ?? me.mail)?.split("@")[1];
  const tenantDomain =
    upnDomain ??
    (await graphFetch("/organization")
      .then((r) => r.value?.[0])
      .then((org: { verifiedDomains?: { isDefault?: boolean; name?: string }[]; displayName?: string }) => org?.verifiedDomains?.find((d) => d.isDefault)?.name ?? org?.displayName)
      .catch(() => undefined));

  // MSAL Node's public AuthenticationResult type doesn't expose the raw refresh token (it's
  // managed internally in the token cache) unless a custom cache plugin is wired up. Surfacing it
  // here for encryption/storage requires reading the in-memory token cache directly.
  const tokenCache = getMsalApp().getTokenCache();
  const cacheJson = JSON.parse(await tokenCache.serialize());
  const refreshTokenEntry = Object.values(cacheJson.RefreshToken ?? {})[0] as { secret?: string } | undefined;
  if (!refreshTokenEntry?.secret) {
    throw new Error("No refresh token present in MSAL cache after code exchange — is offline_access scope granted?");
  }

  const adminUpn: string = me.userPrincipalName ?? me.mail;

  return {
    m365TenantId: result.tenantId,
    tenantDomain: tenantDomain ?? adminUpn?.split("@")[1] ?? result.tenantId,
    adminUpn,
    adminDisplayName: me.displayName ?? me.userPrincipalName ?? me.mail,
    refreshToken: refreshTokenEntry.secret,
    accessTokenExpiresOn: result.expiresOn,
  };
}
