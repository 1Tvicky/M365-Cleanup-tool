import { app } from "./app.js";
import "./jobs/cleanupWorker.js"; // starts the BullMQ worker as a side effect of import
import { config } from "./config/index.js";

app.listen(config.port, () => {
  console.log(`M365 Data Cleanup Utility API listening on :${config.port}`);
});
