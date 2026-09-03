-- Add Clouds / Manage Clouds connection layer.
-- Adds a per-(tenant, cloud_type) connection model beneath the existing `tenants` table, which
-- stays the parent registry (one row per Azure AD directory). See docs/cloud-connections-api.md
-- and docs/azure-ad-app-registration.md §4a for the flow this backs.

-- Global (not per-tenant) flag: only operators who can initiate/tear down a tenant connection.
-- Deliberately separate from tenant_roles.role, which governs cleanup actions *within* an
-- already-connected tenant — connecting a cloud is a broader trust decision than running a job
-- inside one that's already connected.
ALTER TABLE operators ADD COLUMN is_internal_admin BOOLEAN NOT NULL DEFAULT false;

CREATE TABLE connections (
  id                       UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id                UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  cloud_type               TEXT NOT NULL CHECK (cloud_type IN ('onedrive', 'sharepoint', 'teams')),
  admin_upn                TEXT NOT NULL,
  display_name             TEXT NOT NULL,
  status                   TEXT NOT NULL DEFAULT 'connecting'
                              CHECK (status IN ('connecting', 'active', 'error', 'needs_reauth', 'disconnected')),
  connected_at             TIMESTAMPTZ,
  disconnected_at          TIMESTAMPTZ,
  last_synced_at           TIMESTAMPTZ,
  -- Delegated admin refresh token from the connect flow — identity confirmation only, never used
  -- for data enumeration (that uses application/client-credentials tokens). See
  -- docs/azure-ad-app-registration.md §5. NULL once disconnected (local revoke).
  encrypted_refresh_token  TEXT,
  token_expiry             TIMESTAMPTZ,
  connected_by_operator_id UUID REFERENCES operators(id),
  last_error               TEXT,
  created_at               TIMESTAMPTZ NOT NULL DEFAULT now(),
  -- One connection per cloud type per tenant — reconnecting re-activates the same row rather than
  -- creating a duplicate (see routes/cloudConnections.ts upsert logic).
  UNIQUE (tenant_id, cloud_type)
);

CREATE INDEX connections_tenant_idx ON connections (tenant_id);
CREATE INDEX connections_status_idx ON connections (status);

-- Shared across all three cloud types, but the columns mean different things per type:
--   onedrive / teams: one row per M365 user (graph_user_id = user id, upn = user's UPN).
--   sharepoint: one row per SITE (graph_user_id = site id, upn = site's webUrl) — SharePoint has
--     no "users" to enumerate the same way. See docs/graph-api-limitations.md.
CREATE TABLE connection_users (
  id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id      UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  graph_user_id      TEXT NOT NULL,
  upn                TEXT NOT NULL,
  display_name       TEXT,
  storage_used_bytes BIGINT NOT NULL DEFAULT 0,
  item_count         INTEGER NOT NULL DEFAULT 0,
  sync_status        TEXT NOT NULL DEFAULT 'pending' CHECK (sync_status IN ('pending', 'synced', 'failed')),
  last_synced_at     TIMESTAMPTZ,
  error_message      TEXT,
  UNIQUE (connection_id, graph_user_id)
);

CREATE INDEX connection_users_connection_idx ON connection_users (connection_id);

CREATE TABLE sync_jobs (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id       UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  status              TEXT NOT NULL DEFAULT 'queued'
                         CHECK (status IN ('queued', 'running', 'completed', 'completed_with_errors', 'failed', 'cancelled')),
  total_users         INTEGER NOT NULL DEFAULT 0,
  processed_users     INTEGER NOT NULL DEFAULT 0,
  started_at          TIMESTAMPTZ,
  finished_at         TIMESTAMPTZ,
  error_log           JSONB NOT NULL DEFAULT '[]',
  cancel_requested_at TIMESTAMPTZ,
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Manage Clouds reads the latest job per connection: ORDER BY created_at DESC LIMIT 1.
CREATE INDEX sync_jobs_connection_idx ON sync_jobs (connection_id, created_at DESC);

-- Audit trail for the connection layer specifically (token exchanges, job lifecycle,
-- resync/disconnect actions) — this is the trust root for the later Cleanup module, so every
-- state-changing action here is logged, not just cleanup executions.
CREATE TABLE connection_events (
  id            BIGSERIAL PRIMARY KEY,
  connection_id UUID REFERENCES connections(id) ON DELETE SET NULL,
  tenant_id     UUID REFERENCES tenants(id) ON DELETE SET NULL,
  event         TEXT NOT NULL CHECK (event IN (
                  'token_exchange', 'connected', 'connect_failed', 'job_started', 'job_finished',
                  'reauth_required', 'resync_requested', 'disconnected'
                )),
  operator_id   UUID REFERENCES operators(id),
  detail        JSONB NOT NULL DEFAULT '{}',
  occurred_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX connection_events_connection_idx ON connection_events (connection_id, occurred_at DESC);
