import { describe, expect, it } from "vitest";
import { filterRealMemberUsers } from "./targetUsers.js";

describe("filterRealMemberUsers", () => {
  it("excludes guest (#EXT#) accounts, confirmed live to have no OneDrive/mailbox to generate into", () => {
    const users = [
      { id: "1", upn: "real.user@contoso.com", displayName: "Real User" },
      { id: "2", upn: "guest_partner.com#EXT#@contoso.onmicrosoft.com", displayName: "Guest" },
    ];
    expect(filterRealMemberUsers(users)).toEqual([users[0]]);
  });

  it("excludes soft-deleted/orphaned placeholder accounts", () => {
    const users = [
      { id: "1", upn: "real.user@contoso.com", displayName: "Real User" },
      { id: "2", upn: "12345__U_DELETED___old.co#EXT#@contoso.onmicrosoft.com", displayName: null },
    ];
    expect(filterRealMemberUsers(users)).toEqual([users[0]]);
  });

  it("keeps every real internal member account untouched", () => {
    const users = [
      { id: "1", upn: "a@contoso.com", displayName: "A" },
      { id: "2", upn: "b@contoso.com", displayName: "B" },
    ];
    expect(filterRealMemberUsers(users)).toEqual(users);
  });
});
