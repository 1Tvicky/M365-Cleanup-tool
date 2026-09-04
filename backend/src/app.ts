import cookieParser from "cookie-parser";
import express, { type ErrorRequestHandler } from "express";
import helmet from "helmet";
import { ZodError } from "zod";
import { config } from "./config/index.js";
import { authRouter } from "./routes/auth.js";
import { sessionRouter } from "./routes/session.js";
import { tenantsRouter } from "./routes/tenants.js";
import { discoveryRouter } from "./routes/discovery.js";
import { cleanupRouter } from "./routes/cleanup.js";
import { jobsRouter } from "./routes/jobs.js";
import { cloudConnectionsRouter, m365ConnectCallbackRouter } from "./routes/cloudConnections.js";
import { cleaningRouter } from "./routes/cleaning.js";
import { ApiError } from "./types/index.js";

export const app = express();

app.set("trust proxy", 1); // req.ip / req.secure reflect the real client behind a load balancer

app.use(helmet());
app.use(cookieParser());
app.use(express.json({ limit: "1mb" }));

/** LOGIN-SEC-009 — reject plaintext HTTP in production; local dev and health checks are exempt. */
app.use((req, res, next) => {
  if (config.nodeEnv === "production" && !req.secure && req.path !== "/healthz") {
    res.status(403).json({ error: { code: "HTTPS_REQUIRED", message: "This API requires HTTPS" } });
    return;
  }
  next();
});

app.get("/healthz", (_req, res) => res.json({ ok: true }));

app.use("/api/v1/auth", sessionRouter); // operator login/logout/forgot-password/SSO — see routes/session.ts
app.use("/api/v1/auth", authRouter); // M365 tenant admin-consent flow — distinct from operator login above
app.use("/api/v1/tenants", tenantsRouter);
app.use("/api/v1", discoveryRouter); // mounts /tenants/:tenantId/{users,teams,sites}
app.use("/api/v1", cleanupRouter); // mounts /tenants/:tenantId/cleanup/*
app.use("/api/v1", jobsRouter); // mounts /jobs/* and /tenants/:tenantId/{jobs,reports}

// Add Clouds / Manage Clouds connection layer — deliberately NOT under /api/v1: /api/clouds/* is
// this feature's own contract (docs/cloud-connections-api.md), and the callback path below must
// match the Azure AD app registration's redirect URI exactly (docs/azure-ad-app-registration.md §4a).
app.use("/api/clouds", cloudConnectionsRouter);
app.use("/api/auth/m365/callback", m365ConnectCallbackRouter);

// Cleaning module (discovery phase) — read-only, reuses connections/tenant_roles from the layer
// above. Same reasoning for living outside /api/v1: its own contract, its own namespace.
app.use("/api/cleaning", cleaningRouter);

const errorHandler: ErrorRequestHandler = (err, _req, res, _next) => {
  if (err instanceof ApiError) {
    res.status(err.status).json({ error: { code: err.code, message: err.message, details: err.details } });
    return;
  }
  if (err instanceof ZodError) {
    res.status(400).json({ error: { code: "VALIDATION_ERROR", message: "Invalid request", details: err.flatten() } });
    return;
  }
  console.error(err);
  res.status(500).json({ error: { code: "INTERNAL_ERROR", message: "Unexpected server error" } });
};
app.use(errorHandler);
