import { app } from "./app.js";
import "./jobs/cleanupWorker.js"; // starts the BullMQ worker as a side effect of import
import "./jobs/cloudSyncWorker.js"; // ditto, for Add Clouds enumeration jobs
import "./jobs/cleaningScanWorker.js"; // ditto, for Cleaning module discovery scans
import { config } from "./config/index.js";

app.listen(config.port, () => {
  console.log(`M365 Data Cleanup Utility API listening on :${config.port}`);
});
