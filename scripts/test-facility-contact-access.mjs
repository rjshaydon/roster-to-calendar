import assert from "node:assert/strict";

import { issueFacilityContactAccessToken, verifyFacilityContactAccessToken } from "../functions/_lib/facility-contact-access.js";

const secret = "contact-access-test-secret-that-is-long-enough";
const now = Date.parse("2026-09-24T09:00:00.000Z");
const token = await issueFacilityContactAccessToken(secret, {
  facilityKey: "DDH", expiresAt: "2026-09-24T09:15:00.000Z", now,
});
assert.ok(token);
assert.deepEqual(await verifyFacilityContactAccessToken(secret, token, { now: now + 14 * 60 * 1000 }), {
  facilityKey: "DDH", expiresAt: "2026-09-24T09:15:00.000Z",
});
assert.equal(await verifyFacilityContactAccessToken(secret, token, { now: now + 15 * 60 * 1000 }), null);
assert.equal(await verifyFacilityContactAccessToken(`${secret}x`, token, { now }), null);
assert.equal(await verifyFacilityContactAccessToken(secret, `${token.slice(0, -1)}x`, { now }), null);
const clamped = await issueFacilityContactAccessToken(secret, {
  facilityKey: "MMC", expiresAt: "2026-09-24T10:00:00.000Z", now,
});
assert.equal((await verifyFacilityContactAccessToken(secret, clamped, { now }))?.expiresAt, "2026-09-24T09:15:00.000Z");
assert.equal(await issueFacilityContactAccessToken("short", { facilityKey: "DDH", now }), "");

console.log("Facility contact access tokens passed: facility-scoped, tamper-proof and limited to 15 minutes.");
