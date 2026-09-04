import { Router } from "express";
import { config } from "../config/index.js";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";

export const authRouter = Router();

/** See docs/azure-ad-app-registration.md §4 — per-tenant admin consent URL. */
authRouter.get("/consent-url", requireSession, (req, res) => {
  const tenantHint = typeof req.query.tenantHint === "string" ? req.query.tenantHint : "common";
  const consentUrl =
    `https://login.microsoftonline.com/${encodeURIComponent(tenantHint)}/adminconsent` +
    `?client_id=${encodeURIComponent(config.microsoft.clientId)}` +
    `&redirect_uri=${encodeURIComponent(config.microsoft.redirectUri)}`;
  res.json({ consentUrl });
});

/**
 * OAuth redirect target for the M365 TENANT admin-consent grant (not operator login — that's
 * routes/session.ts's /google and /office365 callbacks). Only records a tenant as connected when
 * admin_consent=True is present.
 */
authRouter.get(
  "/callback",
  asyncHandler(async (req, res) => {
    const { tenant, admin_consent } = req.query;

    if (admin_consent !== "True" || typeof tenant !== "string") {
      res.redirect(`/clouds?connected=false&reason=consent_declined`);
      return;
    }

    await query(
      `INSERT INTO tenants (m365_tenant_id, display_name, status, connected_at)
       VALUES ($1, $1, 'connected', now())
       ON CONFLICT (m365_tenant_id)
       DO UPDATE SET status = 'connected', connected_at = now()`,
      [tenant]
    );

    // Discovery snapshot (teams/sites/users) is kicked off as a background job, not inline here —
    // see jobs/cleanupWorker.ts sibling queue for discovery jobs.

    res.redirect(`/clouds?connected=${encodeURIComponent(tenant)}`);
  })
);
