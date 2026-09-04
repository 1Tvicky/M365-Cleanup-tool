function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required env var: ${name}`);
  return value;
}

export const config = {
  port: Number(process.env.PORT ?? 4000),
  nodeEnv: process.env.NODE_ENV ?? "development",

  // Our own dedicated Microsoft Entra app registration for this cleanup tool — one multi-tenant
  // app, owned by us, never a customer's and never shared with any other CloudFuze product (e.g.
  // the migration tool has its own separate registration). See docs/azure-ad-app-registration.md.
  microsoft: {
    clientId: process.env.MICROSOFT_CLIENT_ID ?? "",
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    certThumbprint: process.env.MICROSOFT_CLIENT_CERT_THUMBPRINT,
    certPrivateKeyPath: process.env.MICROSOFT_CLIENT_CERT_PRIVATE_KEY_PATH,
    // Multi-tenant authority — "common" accepts a sign-in from any organizational directory,
    // never a single hardcoded customer tenant.
    authority: process.env.MICROSOFT_AUTHORITY ?? "https://login.microsoftonline.com/common",
    redirectUri: process.env.MICROSOFT_LEGACY_REDIRECT_URI ?? "", // legacy /adminconsent flow — docs/azure-ad-app-registration.md §4
    // Add Clouds per-cloud-type connect flow — docs/azure-ad-app-registration.md §4a. Same app
    // registration/credentials as above, different registered redirect URI.
    connectRedirectUri: process.env.MICROSOFT_REDIRECT_URI ?? "",
  },

  tokenEncryption: {
    // Base64-encoded 32-byte AES-256 key. Dev-only env value — production must fetch this from a
    // KMS at startup instead. See services/tokenEncryption.ts.
    key: process.env.TOKEN_ENCRYPTION_KEY ?? "",
  },

  database: {
    url: process.env.DATABASE_URL ?? "postgres://cleanup:cleanup@localhost:5432/m365_cleanup",
  },

  redis: {
    url: process.env.REDIS_URL ?? "redis://localhost:6379",
  },

  blobStorage: {
    connectionString: process.env.BLOB_STORAGE_CONNECTION_STRING ?? "",
    container: process.env.BLOB_STORAGE_CONTAINER ?? "cleanup-exports",
    exportRetentionDays: Number(process.env.EXPORT_RETENTION_DAYS ?? 90),
  },

  session: {
    jwtSecret: process.env.NODE_ENV === "production" ? required("SESSION_JWT_SECRET") : (process.env.SESSION_JWT_SECRET ?? "dev-only-secret"),
    issuer: process.env.SESSION_JWT_ISSUER ?? "m365-cleanup-utility",
    cookieName: process.env.SESSION_COOKIE_NAME ?? "cf_session",
    cookieDomain: process.env.SESSION_COOKIE_DOMAIN || undefined,
    ttlMinutes: Number(process.env.SESSION_TTL_MINUTES ?? 480),
  },

  password: {
    minLength: Number(process.env.PASSWORD_MIN_LENGTH ?? 8),
    maxLength: Number(process.env.PASSWORD_MAX_LENGTH ?? 128),
  },

  login: {
    maxFailedAttempts: Number(process.env.LOGIN_MAX_FAILED_ATTEMPTS ?? 5),
    lockoutMinutes: Number(process.env.LOGIN_LOCKOUT_MINUTES ?? 15),
  },

  passwordReset: {
    tokenTtlMinutes: Number(process.env.PASSWORD_RESET_TOKEN_TTL_MINUTES ?? 30),
  },

  oauth: {
    google: {
      clientId: process.env.GOOGLE_CLIENT_ID ?? "",
      clientSecret: process.env.GOOGLE_CLIENT_SECRET ?? "",
      redirectUri: process.env.GOOGLE_REDIRECT_URI ?? "",
    },
    operatorAzure: {
      clientId: process.env.OPERATOR_AZURE_CLIENT_ID ?? "",
      clientSecret: process.env.OPERATOR_AZURE_CLIENT_SECRET ?? "",
      redirectUri: process.env.OPERATOR_AZURE_REDIRECT_URI ?? "",
      tenant: process.env.OPERATOR_AZURE_TENANT ?? "common",
    },
    allowedEmailDomains: (process.env.ALLOWED_LOGIN_EMAIL_DOMAINS ?? "")
      .split(",")
      .map((d) => d.trim().toLowerCase())
      .filter(Boolean),
  },

  // Preview results older than this must be re-resolved against live Graph data before confirm.
  previewTtlMinutes: 30,

  // Graph batch execution tuning — see services/rateLimiter.ts
  graph: {
    batchSize: 20,
    maxRetries: 5,
    baseBackoffMs: 500,
    maxBackoffMs: 30_000,
    // Neither the Graph SDK nor Node's fetch has a default timeout — a single hung connection
    // (observed on a large real tenant's chat enumeration) can otherwise stall an entire batch,
    // and therefore the whole job, forever. See services/rateLimiter.ts.
    callTimeoutMs: 30_000,
  },
} as const;
