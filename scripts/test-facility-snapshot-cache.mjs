import assert from "node:assert/strict";
import {
  FACILITY_SNAPSHOT_SCHEMA_VERSION,
  facilitySnapshotKey,
  facilitySnapshotRecordIsUsable,
  loadFacilitySnapshot,
} from "../public/static/facility-snapshot-cache.js";

const future = new Date(Date.now() + 60_000).toISOString();
const context = { ownerKey: "doctor@example.com", scopeKey: "site:MMC", accessExpiresAt: future };
const record = { schemaVersion: FACILITY_SNAPSHOT_SCHEMA_VERSION, ownerKey: context.ownerKey, scopeKey: context.scopeKey };
assert.equal(facilitySnapshotRecordIsUsable(record, context), true);
assert.equal(facilitySnapshotRecordIsUsable(record, { ...context, ownerKey: "other@example.com" }), false, "another account must never receive the record");
assert.equal(facilitySnapshotRecordIsUsable(record, { ...context, scopeKey: "site:DDH" }), false, "another access scope must never receive the record");
assert.equal(facilitySnapshotRecordIsUsable(record, { ...context, accessExpiresAt: new Date(Date.now() - 1).toISOString() }), false, "expired authorisation must disable persisted data");
assert.notEqual(
  facilitySnapshotKey({ ...context, kind: "on-shift", query: { date: "2026-09-06", facilityKey: "MMC" } }),
  facilitySnapshotKey({ ...context, kind: "on-shift", query: { date: "2026-09-06", facilityKey: "DDH" } }),
  "different query scopes must have different keys",
);
assert.equal(await loadFacilitySnapshot(context, "on-shift", {}), null, "a runtime without IndexedDB must fail closed");
console.log("Facility snapshot cache isolation, schema and expiry checks passed.");
