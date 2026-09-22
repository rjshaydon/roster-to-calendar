import assert from "node:assert/strict";

import { D1_BILLING_UNAVAILABLE_REASON, d1AnalyticsQuery, evaluateD1Budget, settledUtcDayInterval, summarizeAnalyticsPayload, utcDayInterval } from "./d1-quota-budget-lib.mjs";

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
assert.deepEqual(settledUtcDayInterval(new Date("2026-09-08T00:10:00.000Z")), {
  start: "2026-09-08T00:00:00.000Z",
  observedUntil: "2026-09-07T23:55:00.000Z",
  end: "2026-09-09T00:00:00.000Z",
  date: "2026-09-08",
  settled: false,
});
assert.deepEqual(settledUtcDayInterval(new Date("2026-09-08T00:20:00.000Z")), {
  start: "2026-09-08T00:00:00.000Z",
  observedUntil: "2026-09-08T00:05:00.000Z",
  end: "2026-09-09T00:00:00.000Z",
  date: "2026-09-08",
  settled: true,
});
assert.match(d1AnalyticsQuery(), /d1AnalyticsAdaptiveGroups/);
assert.match(d1AnalyticsQuery(), /d1QueriesAdaptiveGroups/);
assert.match(d1AnalyticsQuery(), /datetimeFiveMinutes/);
assert.doesNotMatch(d1AnalyticsQuery(), /databaseId:\s*\$/);

const independentlySampledFingerprints = summarizeAnalyticsPayload(
  analyticsPayload(
    [usage(productionId, 16_295, 20, 396, 8)],
    [query(productionId, "SELECT bounded", 394, 14_012, 36, 20)],
  ),
  { complete: true, databases: [{ id: productionId }] },
);
assert.equal(independentlySampledFingerprints.complete, true, "a sub-5,000-row read-only adaptive sampling gap must not invalidate an otherwise exact account timeline");
assert.equal(independentlySampledFingerprints.fingerprintSamplingDifference.rowsRead, 2_283);

const widelySampledFingerprints = summarizeAnalyticsPayload(
  analyticsPayload(
    [usage(productionId, 35_373, 184, 860, 62)],
    [query(productionId, "SELECT sampled", 924, 26_812, 29, 165)],
    [timeline(productionId, 30_770, 154, 800, 52, "2026-09-07T02:40:00Z"), timeline(productionId, 4_603, 30, 60, 10, "2026-09-07T02:45:00Z")],
  ),
  { complete: true, databases: [{ id: productionId }] },
);
assert.equal(widelySampledFingerprints.complete, true, "independently sampled fingerprint totals must not invalidate an exact quota ledger");
assert.deepEqual(widelySampledFingerprints.fingerprintSamplingDifference, { rowsRead: 8_561, rowsWritten: 19, queries: -2 });
assert.deepEqual(widelySampledFingerprints.maximumFiveMinute, { rowsRead: 30_770, rowsWritten: 154 });
const overSampledFingerprints = summarizeAnalyticsPayload(
  analyticsPayload([usage(productionId, 100, 0, 1, 0)], [query(productionId, "SELECT oversampled", 2, 200, 100, 0)]),
  { complete: true, databases: [{ id: productionId }] },
);
assert.equal(overSampledFingerprints.complete, true);
assert.equal(overSampledFingerprints.fingerprintSamplingDifference.rowsRead, -100, "fingerprint over-sampling must remain diagnostic");

const previousReport = { generatedAt: "2026-09-07T02:40:00.000Z", sampleValid: true, interval: { date: "2026-09-07" }, effectiveUsage: { rowsRead: 100_000, rowsWritten: 2_000 } };
const currentBilling = { rowsRead: 110_000, rowsWritten: 2_100, observedAt: "2026-09-07T02:59:00.000Z" };
const go = evaluateD1Budget({ analytics, inventory: completeInventory, billing: currentBilling, previousReport, now, optionalEstimate: { rowsRead: 27_021, rowsWritten: 2_252 } });
assert.equal(go.decision, "GO");
const graphOnlyGo = evaluateD1Budget({ analytics, inventory: completeInventory, billing: { unavailableReason: D1_BILLING_UNAVAILABLE_REASON }, previousReport, now, optionalEstimate: { rowsRead: 27_021, rowsWritten: 2_252 } });
assert.equal(graphOnlyGo.decision, "GO");
assert.equal(graphOnlyGo.billingMode, "free-plan-analytics-only");
const earlyPassiveSample = evaluateD1Budget({
  analytics,
  inventory: completeInventory,
  billing: { unavailableReason: D1_BILLING_UNAVAILABLE_REASON },
  previousReport: null,
  now: new Date("2026-09-07T00:20:00.000Z"),
});
assert.equal(earlyPassiveSample.decision, "STOP");
assert.equal(earlyPassiveSample.sampleValid, true, "a reconciled early sample is valid evidence even though rollout remains blocked");
assert.ok(earlyPassiveSample.reasons.includes("two-hour-passive-baseline-required"));
assert.ok(earlyPassiveSample.reasons.includes("second-sample-required"));

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
const truncated = summarizeAnalyticsPayload(
  analyticsPayload([usage(productionId, 1, 0, 1, 0)], Array.from({ length: 10_000 }, () => query(productionId, "SELECT truncated", 1, 1, 1, 0))),
  { complete: true, databases: [{ id: productionId }] },
);
assert.equal(truncated.complete, false);
assert.ok(truncated.reasons.includes("analytics-pagination-or-limit-truncation"));

const canaryPrevious = {
  ...previousReport,
  interval: { date: "2026-09-07", observedUntil: "2026-09-07T02:40:00.000Z" },
  analytics: {
    maximumFiveMinute: { rowsRead: 5_000, rowsWritten: 10 },
    queryFingerprints: [{ query: "SELECT ordinary" }],
  },
};
const boundedCanaryAnalytics = summarizeAnalyticsPayload(
  analyticsPayload(
    [usage(productionId, 127_000, 2_200, 120, 12)],
    [query(productionId, "SELECT ordinary", 100, 100_000, 1_000, 2_000), query(productionId, "SELECT compact bootstrap", 20, 27_000, 1_350, 200)],
    [timeline(productionId, 100_000, 2_000, 100, 10, "2026-09-07T02:40:00Z"), timeline(productionId, 27_000, 200, 20, 2, "2026-09-07T02:45:00Z")],
  ),
  { complete: true, databases: [{ id: productionId }] },
);
const boundedCanary = evaluateD1Budget({
  analytics: boundedCanaryAnalytics,
  inventory: { complete: true, databases: [{ id: productionId }] },
  billing: { unavailableReason: D1_BILLING_UNAVAILABLE_REASON },
  previousReport: canaryPrevious,
  now,
  optionalEstimate: { rowsRead: 27_021, rowsWritten: 2_252 },
  canary: {
    readCeiling: 27_021,
    requestId: "bootstrap-1",
    requestAttribution: { requestId: "bootstrap-1", d1Statements: 18, d1RowsRead: 27_000, d1RowsWritten: 200, d1Limit: 768, d1MetadataComplete: true },
    reviewedFingerprints: ["SELECT compact bootstrap"],
  },
});
assert.equal(boundedCanary.decision, "GO", boundedCanary.reasons.join(", "));
assert.equal(boundedCanary.canary.envelopeRowsRead, 37_021);

for (const [label, canary, canaryAnalytics, expectedReason] of [
  ["missing attribution", { readCeiling: 27_021, requestId: "bootstrap-1", reviewedFingerprints: ["SELECT compact bootstrap"] }, boundedCanaryAnalytics, "canary-request-attribution-required"],
  ["unreviewed query", { readCeiling: 27_021, requestId: "bootstrap-1", requestAttribution: { requestId: "bootstrap-1", d1Statements: 18, d1RowsRead: 27_000, d1RowsWritten: 200, d1Limit: 768 } }, boundedCanaryAnalytics, "canary-unreviewed-query-fingerprint"],
  ["request mismatch", { readCeiling: 27_021, requestId: "bootstrap-1", requestAttribution: { requestId: "other", d1Statements: 18, d1RowsRead: 27_000, d1RowsWritten: 200, d1Limit: 768 }, reviewedFingerprints: ["SELECT compact bootstrap"] }, boundedCanaryAnalytics, "canary-request-attribution-mismatch"],
  ["metadata incomplete", { readCeiling: 27_021, requestId: "bootstrap-1", requestAttribution: { requestId: "bootstrap-1", d1Statements: 18, d1RowsRead: 27_000, d1RowsWritten: 200, d1Limit: 768, d1MetadataComplete: false }, reviewedFingerprints: ["SELECT compact bootstrap"] }, boundedCanaryAnalytics, "canary-request-metadata-incomplete"],
  ["envelope exceeded", { readCeiling: 27_021, requestId: "bootstrap-1", requestAttribution: { requestId: "bootstrap-1", d1Statements: 18, d1RowsRead: 27_000, d1RowsWritten: 200, d1Limit: 768 }, reviewedFingerprints: ["SELECT compact bootstrap"] }, summarizeAnalyticsPayload(analyticsPayload([usage(productionId, 150_000, 200, 120, 2)], [query(productionId, "SELECT ordinary", 100, 100_000, 1_000, 0), query(productionId, "SELECT compact bootstrap", 20, 50_000, 2_500, 200)], [timeline(productionId, 100_000, 0, 100, 0, "2026-09-07T02:40:00Z"), timeline(productionId, 50_000, 200, 20, 2, "2026-09-07T02:45:00Z")]), { complete: true, databases: [{ id: productionId }] }), "canary-five-minute-envelope-exceeded"],
  ["million-row spike", { readCeiling: 27_021, requestId: "bootstrap-1", requestAttribution: { requestId: "bootstrap-1", d1Statements: 18, d1RowsRead: 27_000, d1RowsWritten: 200, d1Limit: 768 }, reviewedFingerprints: ["SELECT compact bootstrap"] }, summarizeAnalyticsPayload(analyticsPayload([usage(productionId, 1_000_000, 200, 20, 2)], [query(productionId, "SELECT compact bootstrap", 20, 27_000, 1_350, 200)]), { complete: true, databases: [{ id: productionId }] }), "at-a-glance-five-minute-hard-stop"],
]) {
  const result = evaluateD1Budget({
    analytics: canaryAnalytics,
    inventory: { complete: true, databases: [{ id: productionId }] },
    billing: { unavailableReason: D1_BILLING_UNAVAILABLE_REASON },
    previousReport: canaryPrevious,
    now,
    optionalEstimate: { rowsRead: 27_021, rowsWritten: 2_252 },
    canary,
  });
  assert.equal(result.decision, "STOP", label);
  assert.ok(result.reasons.includes(expectedReason), `${label} should include ${expectedReason}`);
}

console.log("Account-wide D1 budget gate passed quota-ledger, sampled-fingerprint, canary-attribution and fail-closed threshold checks.");

function analyticsPayload(usageGroups, queryGroups, timelineGroups = usageGroups.map((group) => timeline(group.dimensions.databaseId, group.sum.rowsRead, group.sum.rowsWritten, group.sum.readQueries, group.sum.writeQueries))) {
  return { data: { viewer: { accounts: [{ usage: usageGroups, timeline: timelineGroups, queries: queryGroups }] } } };
}

function usage(databaseId, rowsRead, rowsWritten, readQueries, writeQueries) {
  return { dimensions: { databaseId, date: "2026-09-07" }, sum: { rowsRead, rowsWritten, readQueries, writeQueries } };
}

function timeline(databaseId, rowsRead, rowsWritten, readQueries = 1, writeQueries = 1, observedAt = "2026-09-07T02:45:00Z") {
  return { dimensions: { databaseId, datetimeFiveMinutes: observedAt }, sum: { rowsRead, rowsWritten, readQueries, writeQueries } };
}

function query(databaseId, sql, count, totalRowsRead, averageRowsRead, totalRowsWritten = 0) {
  return {
    count,
    dimensions: { databaseId, query: sql, error: "" },
    avg: { rowsRead: averageRowsRead, rowsReturned: 1, rowsWritten: 0 },
    sum: { rowsRead: totalRowsRead, rowsReturned: count, rowsWritten: totalRowsWritten },
  };
}
