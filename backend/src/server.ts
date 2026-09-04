import { app } from "./app.js";
import "./jobs/cleanupWorker.js"; // starts the BullMQ worker as a side effect of import
import "./jobs/cloudSyncWorker.js"; // ditto, for Add Clouds enumeration jobs
import "./jobs/cleaningScanWorker.js"; // ditto, for Cleaning module discovery scans
import "./jobs/cleanupExecutionWorker.js"; // ditto, for Cleaning module cleanup/deletion execution
import { config } from "./config/index.js";
import { recoverOrphanedJobs } from "./jobs/recoverOrphanedJobs.js";

app.listen(config.port, () => {
  console.log(`M365 Data Cleanup Utility API listening on :${config.port}`);
});

// Best-effort — a failure here shouldn't prevent the server from starting.
recoverOrphanedJobs().catch((err) => console.error("[recovery] failed to reconcile orphaned jobs on startup:", err));
