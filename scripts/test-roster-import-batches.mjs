import assert from "node:assert/strict";
import { planRosterImportBatches, ROSTER_BATCH_FACT_LIMIT, ROSTER_BATCH_BYTE_LIMIT } from "../functions/_lib/roster-import-batches.js";

const doctors = Array.from({ length: 200 }, (_, index) => ({ key: `DOCTOR ${String(index).padStart(3, "0")}`, displayName: `Doctor ${index}`, seniority: "HMO" }));
const eventsByDoctor = Object.fromEntries(doctors.map((doctor) => [doctor.key, Array.from({ length: 90 }, (_, day) => {
  const date = new Date(Date.UTC(2026, 10, 2 + day)).toISOString().slice(0, 10);
  return { id: `shift-${day}`, source: "mmc", start: `${date}T08:00:00`, end: `${date}T16:00:00`, title: "Day" };
})]));
for (const sourceId of ["monash-adults", "monash-paeds", "dandenong-findmyshift", "vhh-active-medical-roster"]) {
  const payload = { file: { id: "queued-new-term", sourceId, contentHash: "source-content" }, doctors, eventsByDoctor, issuesByDoctor: {} };
  const plan = await planRosterImportBatches(payload);
  assert.equal(plan.eventCount, 18000);
  assert.ok(plan.batches.length > 1);
  assert.equal(plan.batches.reduce((count, batch) => count + batch.eventCount, 0), 18000);
  assert.equal(plan.batches.at(-1).indexedEventCount, 18000);
  for (const batch of plan.batches) {
    assert.ok(batch.facts + 16 <= ROSTER_BATCH_FACT_LIMIT);
    assert.ok(batch.bytes + 4096 <= ROSTER_BATCH_BYTE_LIMIT);
    assert.equal(batch.presenceRows, batch.eventCount);
  }
  const reordered = { ...payload, doctors: [...doctors].reverse(), eventsByDoctor: Object.fromEntries(Object.entries(eventsByDoctor).reverse().map(([key, events]) => [key, [...events].reverse()])) };
  assert.equal((await planRosterImportBatches(reordered)).revision, plan.revision, "retry/reordered input must have an identical revision and batch order");
  const corrected = structuredClone(payload);
  corrected.eventsByDoctor[doctors[0].key][0].title = "Sick leave";
  assert.notEqual((await planRosterImportBatches(corrected)).revision, plan.revision, "a changed plan must not reuse previous batch receipts");
}
const small = { file: { id: "leave", sourceId: "monash-adults" }, doctors: [doctors[0]], eventsByDoctor: { [doctors[0].key]: [{ id: "leave", start: "2026-11-02", end: "2026-12-02", allDay: true }] } };
assert.equal((await planRosterImportBatches(small)).batches[0].presenceRows, 31);
await assert.rejects(planRosterImportBatches({ ...small, eventsByDoctor: { [doctors[0].key]: [small.eventsByDoctor[doctors[0].key][0], small.eventsByDoctor[doctors[0].key][0]] } }), /duplicate roster occurrence/);
await assert.rejects(planRosterImportBatches({ ...small, eventsByDoctor: { UNKNOWN: [] } }), /unknown doctor/);
await assert.rejects(planRosterImportBatches(small, { maximumFacts: 20 }), /exceeds the batch safety budget/);
console.log("Bounded import planning passed 18,000-shift terms for all four sources, deterministic retries, correction fencing, and leave expansion budgets.");

const vhhSpanning = structuredClone(small);
vhhSpanning.file.sourceId = "vhh-active-medical-roster";
const key = vhhSpanning.doctors[0].key;
vhhSpanning.eventsByDoctor = { [key]: [
  { id: "start", start: "2026-09-21T08:00:00", end: "2026-09-21T18:00:00" },
  { id: "end", start: "2027-01-31T22:00:00", end: "2027-02-01T08:00:00" },
] };
const vhhSpanningPlan = await planRosterImportBatches(vhhSpanning);
assert.equal(vhhSpanningPlan.manifest.rosterEndDate, "2027-01-31");
assert.equal(vhhSpanningPlan.manifest.endDate, "2027-02-01");
vhhSpanning.eventsByDoctor[key][1] = { id: "too-far", start: "2027-04-01T08:00:00", end: "2027-04-01T18:00:00" };
await assert.rejects(planRosterImportBatches(vhhSpanning), /180-day/);
