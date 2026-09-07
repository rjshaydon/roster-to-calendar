import assert from "node:assert/strict";

import { D1_BILLING_UNAVAILABLE_REASON, d1AnalyticsQuery, evaluateD1Budget, summarizeAnalyticsPayload, utcDayInterval } from "./d1-quota-budget-lib.mjs";

const productionId = "production-db";
const previewId = "preview-db";
const completeInventory = { complete: true, databases: [{ id: productionId }, { id: previewId }] };
const now = new Date("2026-09-07T03:00:00.000Z");
const payload = analyticsPayload([
  usage(productionId, 100_000, 2_000, 100, 10),
  usage(previewId, 10_000, 100, 5, 1),
], [
  query(productionId, "SELECT id FROM compact WHERE id = ?", 1_000, 100_000, 100, 2_000),
  query(previewId, "SELECT id FROM compact WHERE id = ?", 100, 10_000, 100, 100),
]);
const analytics = summarizeAnalyticsPayload(payload, completeInventory);
assert.equal(analytics.complete, true);
assert.equal(analytics.totals.rowsRead, 110_000);
assert.equal(analytics.totals.rowsWritten, 2_100);
assert.deepEqual(utcDayInterval(now), {
  start: "2026-09-07T00:00:00.000Z",
  observedUntil: "2026-09-07T03:00:00.000Z",
  end: "2026-09-08T00:00:00.000Z",
  date: "2026-09-07",
});
assert.match(d1AnalyticsQuery(), /d1AnalyticsAdaptiveGroups/);
assert.match(d1AnalyticsQuery(), /d1QueriesAdaptiveGroups/);
assert.match(d1AnalyticsQuery(), /datetimeFiveMinutes/);
assert.doesNotMatch(d1AnalyticsQuery(), /databaseId:\s*\$/);

const previousReport = { generatedAt: "2026-09-07T02:40:00.000Z", sampleValid: true, interval: { date: "2026-09-07" }, effectiveUsage: { rowsRead: 100_000, rowsWritten: 2_000 } };
const currentBilling = { rowsRead: 110_000, rowsWritten: 2_100, observedAt: "2026-09-07T02:59:00.000Z" };
const go = evaluateD1Budget({ analytics, inventory: completeInventory, billing: currentBilling, previousReport, now, optionalEstimate: { rowsRead: 27_021, rowsWritten: 2_252 } });
assert.equal(go.decision, "GO");
const graphOnlyGo = evaluateD1Budget({ analytics, inventory: completeInventory, billing: { unavailableReason: D1_BILLING_UNAVAILABLE_REASON }, previousReport, now, optionalEstimate: { rowsRead: 27_021, rowsWritten: 2_252 } });
assert.equal(graphOnlyGo.decision, "GO");
assert.equal(graphOnlyGo.billingMode, "free-plan-analytics-only");

for (const [label, overrides, expectedReason] of [
  ["missing billing", { billing: {} }, "billing-usage-or-documented-unavailability-required"],
  ["unknown billing absence", { billing: { unavailableReason: "unknown" } }, "billing-usage-or-documented-unavailability-required"],
  ["incomplete inventory", { inventory: { ...completeInventory, complete: false } }, "database-inventory-incomplete"],
  ["missing previous sample", { previousReport: null }, "second-sample-required"],
  ["short sample gap", { previousReport: { ...previousReport, generatedAt: "2026-09-07T02:55:00.000Z" } }, "sample-gap-too-short"],
  ["stale billing", { billing: { ...currentBilling, observedAt: "2026-09-07T02:30:00.000Z" } }, "billing-observation-stale"],
  ["incomplete previous sample", { previousReport: { ...previousReport, sampleValid: false } }, "previous-sample-incomplete"],
  ["billing mismatch", { billing: { ...currentBilling, rowsRead: 400_000 } }, "billing-read-reconciliation-failed"],
  ["start threshold", { billing: { ...currentBilling, rowsRead: 500_000 }, analytics: summarizeAnalyticsPayload(analyticsPayload([usage(productionId, 500_000, 2_100, 1, 1)], [query(productionId, "SELECT bounded", 5_000, 500_000, 100, 2_100)]), { complete: true, databases: [{ id: productionId }] }), inventory: { complete: true, databases: [{ id: productionId }] } }, "read-start-threshold-exceeded"],
  ["expensive query", { analytics: summarizeAnalyticsPayload(analyticsPayload([usage(productionId, 110_000, 2_100, 1, 1)], [query(productionId, "SELECT expensive", 2, 110_000, 55_000, 2_100)]), { complete: true, databases: [{ id: productionId }] }), inventory: { complete: true, databases: [{ id: productionId }] } }, "query-per-invocation-limit-exceeded"],
  ["unattributed usage", { analytics: summarizeAnalyticsPayload(analyticsPayload([usage(productionId, 110_000, 2_100, 1, 1)], [query(productionId, "SELECT partial", 10, 1_000, 100, 100)]), { complete: true, databases: [{ id: productionId }] }), inventory: { complete: true, databases: [{ id: productionId }] } }, "query-attribution-incomplete"],
  ["timeline mismatch", { analytics: summarizeAnalyticsPayload(analyticsPayload([usage(productionId, 110_000, 2_100, 1, 1)], [query(productionId, "SELECT bounded", 1_100, 110_000, 100, 2_100)], [timeline(productionId, 80_000, 2_100)]), { complete: true, databases: [{ id: productionId }] }), inventory: { complete: true, databases: [{ id: productionId }] } }, "analytics-time-buckets-do-not-reconcile"],
]) {
  const result = evaluateD1Budget({ analytics, inventory: completeInventory, billing: currentBilling, previousReport, now, ...overrides });
  assert.equal(result.decision, "STOP", label);
  assert.ok(result.reasons.includes(expectedReason), `${label} should include ${expectedReason}`);
}

const unknown = summarizeAnalyticsPayload(analyticsPayload([usage("unknown-db", 1, 0, 1, 0)], []), completeInventory);
assert.equal(unknown.complete, false);
assert.ok(unknown.reasons.includes("unknown-database:unknown-db"));
assert.equal(summarizeAnalyticsPayload({ errors: [{ message: "no" }] }, completeInventory).complete, false);

console.log("Account-wide D1 budget gate passed fail-closed aggregation, reconciliation, burn-rate and threshold checks.");

function analyticsPayload(usageGroups, queryGroups, timelineGroups = usageGroups.map((group) => timeline(group.dimensions.databaseId, group.sum.rowsRead, group.sum.rowsWritten, group.sum.readQueries, group.sum.writeQueries))) {
  return { data: { viewer: { accounts: [{ usage: usageGroups, timeline: timelineGroups, queries: queryGroups }] } } };
}

function usage(databaseId, rowsRead, rowsWritten, readQueries, writeQueries) {
  return { dimensions: { databaseId, date: "2026-09-07" }, sum: { rowsRead, rowsWritten, readQueries, writeQueries } };
}

function timeline(databaseId, rowsRead, rowsWritten, readQueries = 1, writeQueries = 1) {
  return { dimensions: { databaseId, datetimeFiveMinutes: "2026-09-07T02:45:00Z" }, sum: { rowsRead, rowsWritten, readQueries, writeQueries } };
}

function query(databaseId, sql, count, totalRowsRead, averageRowsRead, totalRowsWritten = 0) {
  return {
    count,
    dimensions: { databaseId, query: sql, error: "" },
    avg: { rowsRead: averageRowsRead, rowsReturned: 1, rowsWritten: 0 },
    sum: { rowsRead: totalRowsRead, rowsReturned: count, rowsWritten: totalRowsWritten },
  };
}
