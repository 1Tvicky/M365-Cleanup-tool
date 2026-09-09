import { describe, expect, it } from "vitest";
import { ApiError } from "../types/index.js";
import { getDirectoryClientAs } from "./googleWorkspaceAuth.js";

describe("getDirectoryClientAs", () => {
  it("throws a clear, typed error when no service account key is configured, rather than crashing on import", async () => {
    // Neither GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY_PATH nor _JSON is set in the test environment —
    // this mirrors a deployment that hasn't set up Google Workspace support yet, which must not
    // crash the whole server (see the module's own lazy-loading doc comment).
    await expect(getDirectoryClientAs("admin@example.com")).rejects.toMatchObject({
      status: 503,
      code: "GOOGLE_WORKSPACE_NOT_CONFIGURED",
    });
  });

  it("throws an ApiError instance specifically, not a generic Error", async () => {
    await expect(getDirectoryClientAs("admin@example.com")).rejects.toBeInstanceOf(ApiError);
  });
});
