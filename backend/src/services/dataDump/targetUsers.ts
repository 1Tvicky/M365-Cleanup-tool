import type { BasicUser } from "../../graph/cloudEnumeration.js";

/**
 * Confirmed live against a real tenant (filefuze.co): graph/cloudEnumeration.ts's listAllUsers is
 * deliberately unfiltered for Cleaning's own purposes (matching exactly what /users returns, per
 * that file's own comment) — which for Cleaning is correct, but for Data Dump means the first N
 * "users" picked to generate content for can be guest accounts (`#EXT#` in the UPN, per Microsoft's
 * own convention) or soft-deleted/orphaned directory objects, neither of which has a real OneDrive/
 * mailbox to write into — every generation attempt against one is a guaranteed, wasted failure. Data
 * Dump filters to real internal member accounts before ever picking a target, rather than silently
 * eating a 100%-failure-rate workload when a tenant's directory happens to list guests/leftovers
 * first.
 */
export function filterRealMemberUsers(users: BasicUser[]): BasicUser[] {
  return users.filter((u) => u.upn && !u.upn.includes("#EXT#") && !u.upn.toUpperCase().includes("_DELETED_"));
}
