-- M365 Data Cleanup Utility — Postgres schema
-- Multi-tenant (hundreds+ customer tenants): every job/audit/rbac row is scoped by tenant_id.

CREATE TABLE tenants (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  m365_tenant_id        TEXT NOT NULL UNIQUE,       -- Azure AD directory (tenant) ID
  display_name          TEXT NOT NULL,
  status                TEXT NOT NULL DEFAULT 'consent_pending'
                          CHECK (status IN ('connected', 'consent_pending', 'token_error', 'disconnected')),
  connected_by_admin_upn TEXT,
  connected_at          TIMESTAMPTZ,
  disconnected_at       TIMESTAMPTZ,
  permissions_granted   TEXT[] NOT NULL DEFAULT '{}',
  last_token_refresh_at TIMESTAMPTZ,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- CloudFuze operators (internal users of this tool), not M365 end users.
-- email is stored lowercased (app normalizes on write) so lookups are naturally case-insensitive.
CREATE TABLE operators (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  email                 TEXT NOT NULL UNIQUE,
  display_name          TEXT NOT NULL,
  -- NULL for operators who only ever sign in via Google/Office 365 SSO (no local password set).
  password_hash         TEXT,
  status                TEXT NOT NULL DEFAULT 'unverified'
                          CHECK (status IN ('active', 'unverified', 'disabled', 'deleted')),
  email_verified_at     TIMESTAMPTZ,
  failed_login_count    INTEGER NOT NULL DEFAULT 0,
  locked_until          TIMESTAMPTZ,
  -- Bumped on password change / forced logout; a JWT issued before this instant is rejected
  -- even if not otherwise expired (LOGIN-FP-013, LOGIN-SEC-011 session fixation on reset).
  credentials_valid_after TIMESTAMPTZ NOT NULL DEFAULT now(),
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Single-use, expiring password-reset tokens. Only the SHA-256 hash is stored — the raw token
-- exists only in the emailed link, so a DB read alone can never be used to reset a password.
CREATE TABLE password_reset_tokens (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  operator_id   UUID NOT NULL REFERENCES operators(id) ON DELETE CASCADE,
  token_hash    TEXT NOT NULL UNIQUE,
  expires_at    TIMESTAMPTZ NOT NULL,
  used_at       TIMESTAMPTZ,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX password_reset_tokens_operator_idx ON password_reset_tokens (operator_id);

-- Login/reset events for audit + brute-force review — separate from cleanup job audit_entries.
-- High-volume table; consider partitioning by month in production.
CREATE TABLE auth_events (
  id            BIGSERIAL PRIMARY KEY,
  operator_id   UUID REFERENCES operators(id) ON DELETE SET NULL,
  email         TEXT NOT NULL,
  event         TEXT NOT NULL CHECK (event IN (
                  'login_success', 'login_failed', 'login_locked', 'logout',
                  'password_reset_requested', 'password_reset_completed',
                  'oauth_login_success', 'oauth_login_failed'
                )),
  ip_address    TEXT,
  user_agent    TEXT,
  occurred_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX auth_events_email_time_idx ON auth_events (email, occurred_at DESC);

-- Per-tenant role grants — an operator can be viewer on one tenant, cleanup_admin on another.
CREATE TABLE tenant_roles (
  tenant_id     UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  operator_id   UUID NOT NULL REFERENCES operators(id) ON DELETE CASCADE,
  role          TEXT NOT NULL CHECK (role IN ('viewer', 'cleanup_admin')),
  granted_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
  granted_by    UUID REFERENCES operators(id),
  PRIMARY KEY (tenant_id, operator_id)
);

-- Cached discovery snapshots (teams/sites/users) so the scope picker and preview don't hammer
-- Graph on every keystroke. Refreshed on a schedule + on-demand; previews older than 15 min
-- trigger a live re-check per docs/api-spec.md.
CREATE TABLE discovery_cache (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id     UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  resource_type TEXT NOT NULL CHECK (resource_type IN ('user', 'team', 'site')),
  resource_id   TEXT NOT NULL,
  data          JSONB NOT NULL,
  cached_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (tenant_id, resource_type, resource_id)
);

-- A preview is a resolved, priced-out scope pending confirmation. Expires in 30 min (app-enforced).
CREATE TABLE previews (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id       UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  requested_by    UUID NOT NULL REFERENCES operators(id),
  scope           JSONB NOT NULL,          -- the PreviewRequest body
  result          JSONB NOT NULL,          -- the computed PreviewResult (counts/sizes/breakdown)
  generated_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
  expires_at      TIMESTAMPTZ NOT NULL
);

CREATE TABLE cleanup_jobs (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id             UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  preview_id            UUID NOT NULL REFERENCES previews(id),
  confirmed_by          UUID NOT NULL REFERENCES operators(id),
  confirmed_at          TIMESTAMPTZ NOT NULL DEFAULT now(),
  status                TEXT NOT NULL DEFAULT 'export_in_progress'
                          CHECK (status IN ('export_in_progress', 'queued', 'running',
                                             'completed', 'completed_with_errors', 'failed', 'cancelled')),
  export_manifest_only  BOOLEAN NOT NULL DEFAULT false,
  export_manifest_url   TEXT,
  total_items           INTEGER NOT NULL DEFAULT 0,
  completed_items       INTEGER NOT NULL DEFAULT 0,
  failed_items          INTEGER NOT NULL DEFAULT 0,
  started_at            TIMESTAMPTZ,
  finished_at           TIMESTAMPTZ,
  cancel_requested_at   TIMESTAMPTZ
);

CREATE INDEX cleanup_jobs_tenant_status_idx ON cleanup_jobs (tenant_id, status);

-- Per-item audit trail — one row per Graph delete attempt. Never deleted, only appended.
CREATE TABLE audit_entries (
  id            BIGSERIAL PRIMARY KEY,
  job_id        UUID NOT NULL REFERENCES cleanup_jobs(id) ON DELETE CASCADE,
  tenant_id     UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  item_type     TEXT NOT NULL CHECK (item_type IN ('channel', 'file', 'site_library', 'group')),
  item_id       TEXT NOT NULL,
  item_path     TEXT NOT NULL,
  size_bytes    BIGINT NOT NULL DEFAULT 0,
  result        TEXT NOT NULL CHECK (result IN ('deleted', 'failed', 'skipped')),
  error_code    TEXT,
  occurred_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX audit_entries_job_idx ON audit_entries (job_id);
CREATE INDEX audit_entries_tenant_time_idx ON audit_entries (tenant_id, occurred_at DESC);

-- Top-level "who ran what, when, on what tenant" record — the confirm-step audit log entry
-- referenced by ConfirmResult.auditLogId in docs/api-spec.md, independent of per-item entries.
CREATE TABLE job_audit_log (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  job_id          UUID NOT NULL REFERENCES cleanup_jobs(id) ON DELETE CASCADE,
  tenant_id       UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  operator_id     UUID NOT NULL REFERENCES operators(id),
  scope_summary   JSONB NOT NULL,
  typed_confirmation_matched BOOLEAN NOT NULL,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);
