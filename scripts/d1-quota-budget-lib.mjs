export const D1_BUDGET_LIMITS = Object.freeze({
  dailyReads: 5_000_000,
  dailyWrites: 100_000,
  ordinaryReadReserve: 4_000_000,
  ordinaryWriteReserve: 80_000,
  optionalReadStartMaximum: 500_000,
  optionalWriteStartMaximum: 10_000,
  optionalReadStop: 750_000,
  optionalWriteStop: 15_000,
  maximumOrdinaryRowsPerQuery: 10_000,
  minimumPassiveMs: 2 * 60 * 60 * 1000,
  minimumSampleGapMs: 10 * 60 * 1000,
  maximumReconciliationFraction: 0.05,
  maximumReconciliationFloor: 5_000,
  maximumBillingObservationAgeMs: 15 * 60 * 1000,
  analyticsSettlementMs: 15 * 60 * 1000,
  maximumAnalyticsReconciliationFraction: 0.01,
  maximumAnalyticsReconciliationFloor: 100,
});

export const D1_BILLING_UNAVAILABLE_REASON = "free-plan-dashboard-omits-d1";

export function utcDayInterval(nowValue = new Date()) {
  const now = new Date(nowValue);
  if (!Number.isFinite(now.getTime())) throw new Error("A valid current time is required.");
  const start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  const end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  return { start: start.toISOString(), observedUntil: now.toISOString(), end: end.toISOString(), date: start.toISOString().slice(0, 10) };
}

export function settledUtcDayInterval(nowValue = new Date()) {
  const now = new Date(nowValue);
  const interval = utcDayInterval(now);
  const observedUntil = new Date(now.getTime() - D1_BUDGET_LIMITS.analyticsSettlementMs);
  return {
    ...interval,
    observedUntil: observedUntil.toISOString(),
    settled: observedUntil.getTime() >= Date.parse(interval.start),
  };
}

export function d1AnalyticsQuery() {
  return `query D1AccountQuota($accountTag: string!, $start: Time!, $end: Time!) {
    viewer {
      accounts(filter: { accountTag: $accountTag }) {
        usage: d1AnalyticsAdaptiveGroups(
          limit: 10000
          filter: { datetime_geq: $start, datetime_leq: $end }
        ) {
          dimensions { databaseId }
          sum { readQueries writeQueries rowsRead rowsWritten }
        }
        timeline: d1AnalyticsAdaptiveGroups(
          limit: 10000
          filter: { datetime_geq: $start, datetime_leq: $end }
          orderBy: [datetimeFiveMinutes_ASC]
        ) {
          dimensions { datetimeFiveMinutes databaseId }
          sum { readQueries writeQueries rowsRead rowsWritten }
        }
        queries: d1QueriesAdaptiveGroups(
          limit: 10000
          filter: { datetime_geq: $start, datetime_leq: $end }
        ) {
          count
          dimensions { databaseId query error }
          avg { rowsRead rowsReturned rowsWritten }
          sum { rowsRead rowsReturned rowsWritten }
        }
      }
    }
  }`;
}

export function summarizeAnalyticsPayload(payload, inventory) {
  const reasons = [];
  if (!payload || typeof payload !== "object") return stoppedAnalytics("analytics-response-missing");
  if (Array.isArray(payload.errors) && payload.errors.length) return stoppedAnalytics("analytics-response-errors");
  const accounts = payload?.data?.viewer?.accounts;
  if (!Array.isArray(accounts) || accounts.length !== 1) return stoppedAnalytics("analytics-account-result-invalid");
  const usageGroups = accounts[0]?.usage;
  const timelineGroups = accounts[0]?.timeline;
  const queryGroups = accounts[0]?.queries;
  if (!Array.isArray(usageGroups) || !Array.isArray(timelineGroups) || !Array.isArray(queryGroups)) return stoppedAnalytics("analytics-groups-missing");
  if (usageGroups.length >= 10_000 || timelineGroups.length >= 10_000 || queryGroups.length >= 10_000) reasons.push("analytics-pagination-or-limit-truncation");

  const inventoryIds = new Set((inventory?.databases || []).map((database) => String(database?.id || "")).filter(Boolean));
  const perDatabase = new Map();
  for (const group of usageGroups) {
    const databaseId = String(group?.dimensions?.databaseId || "");
    if (!databaseId) {
      reasons.push("analytics-database-id-missing");
      continue;
    }
    const current = perDatabase.get(databaseId) || emptyUsage(databaseId);
    current.rowsRead += nonNegative(group?.sum?.rowsRead);
    current.rowsWritten += nonNegative(group?.sum?.rowsWritten);
    current.readQueries += nonNegative(group?.sum?.readQueries);
    current.writeQueries += nonNegative(group?.sum?.writeQueries);
    perDatabase.set(databaseId, current);
    if (!inventoryIds.has(databaseId)) reasons.push(`unknown-database:${databaseId}`);
  }

  const queryFingerprints = queryGroups.map((group) => ({
    databaseId: String(group?.dimensions?.databaseId || ""),
    query: String(group?.dimensions?.query || ""),
    error: String(group?.dimensions?.error || ""),
    count: nonNegative(group?.count),
    averageRowsRead: Math.max(
      nonNegative(group?.avg?.rowsRead),
      nonNegative(group?.count) > 0 ? nonNegative(group?.sum?.rowsRead) / nonNegative(group?.count) : 0,
    ),
    averageRowsReturned: nonNegative(group?.avg?.rowsReturned),
    averageRowsWritten: nonNegative(group?.avg?.rowsWritten),
    totalRowsRead: nonNegative(group?.sum?.rowsRead),
    totalRowsReturned: nonNegative(group?.sum?.rowsReturned),
    totalRowsWritten: nonNegative(group?.sum?.rowsWritten),
  })).sort((left, right) => right.totalRowsRead - left.totalRowsRead);
  for (const fingerprint of queryFingerprints) {
    if (!fingerprint.databaseId || !inventoryIds.has(fingerprint.databaseId)) reasons.push(`unknown-query-database:${fingerprint.databaseId || "missing"}`);
  }

  const databases = [...perDatabase.values()].sort((left, right) => left.databaseId.localeCompare(right.databaseId));
  const totals = databases.reduce((total, database) => addUsage(total, database), emptyUsage("account"));
  const timeline = timelineGroups.map((group) => ({
    observedAt: String(group?.dimensions?.datetimeFiveMinutes || ""),
    databaseId: String(group?.dimensions?.databaseId || ""),
    rowsRead: nonNegative(group?.sum?.rowsRead),
    rowsWritten: nonNegative(group?.sum?.rowsWritten),
    readQueries: nonNegative(group?.sum?.readQueries),
    writeQueries: nonNegative(group?.sum?.writeQueries),
  }));
  for (const bucket of timeline) {
    if (!bucket.observedAt || !bucket.databaseId) reasons.push("analytics-timeline-dimension-missing");
    if (bucket.databaseId && !inventoryIds.has(bucket.databaseId)) reasons.push(`unknown-timeline-database:${bucket.databaseId}`);
  }
  const timelineTotals = timeline.reduce((total, bucket) => addUsage(total, bucket), emptyUsage("timeline"));
  const fingerprintTotals = queryFingerprints.reduce((total, fingerprint) => {
    total.rowsRead += fingerprint.totalRowsRead;
    total.rowsWritten += fingerprint.totalRowsWritten;
    total.queries += fingerprint.count;
    return total;
  }, { rowsRead: 0, rowsWritten: 0, queries: 0 });
  if (strictlyDifferent(totals.rowsRead, timelineTotals.rowsRead) || strictlyDifferent(totals.rowsWritten, timelineTotals.rowsWritten)) {
    reasons.push("analytics-time-buckets-do-not-reconcile");
  }
  // Cloudflare's usage and query-fingerprint adaptive groups are sampled
  // independently. Keep timeline reconciliation strict, but tolerate a small
  // absolute read-only sampling gap between those two datasets. Writes retain
  // the strict threshold because even a small unattributed mutation matters.
  if (materiallyDifferent(totals.rowsRead, fingerprintTotals.rowsRead) || strictlyDifferent(totals.rowsWritten, fingerprintTotals.rowsWritten)) {
    reasons.push("query-attribution-incomplete");
  }
  return {
    complete: reasons.length === 0,
    reasons: [...new Set(reasons)],
    databases,
    totals,
    timeline,
    timelineTotals,
    fingerprintTotals,
    unattributed: {
      rowsRead: totals.rowsRead - fingerprintTotals.rowsRead,
      rowsWritten: totals.rowsWritten - fingerprintTotals.rowsWritten,
    },
    queryFingerprints,
  };
}

export function evaluateD1Budget({ analytics, inventory, billing, previousReport, now = new Date(), optionalEstimate = {} } = {}) {
  const interval = analytics?.interval || utcDayInterval(now);
  const reasons = [...(analytics?.reasons || [])];
  if (!inventory?.complete) reasons.push("database-inventory-incomplete");
  if (!analytics?.complete) reasons.push("analytics-incomplete");
  const totals = analytics?.totals || emptyUsage("account");
  const analyticsReads = finiteNonNegative(totals.rowsRead) ?? 0;
  const analyticsWrites = finiteNonNegative(totals.rowsWritten) ?? 0;
  const billingReads = finiteNonNegative(billing?.rowsRead);
  const billingWrites = finiteNonNegative(billing?.rowsWritten);
  const billingUnavailableReason = String(billing?.unavailableReason || "");
  const billingAvailable = billingReads !== null && billingWrites !== null;
  const billingExplicitlyUnavailable = billingReads === null && billingWrites === null && billingUnavailableReason === D1_BILLING_UNAVAILABLE_REASON;
  if (!billingAvailable && !billingExplicitlyUnavailable) reasons.push("billing-usage-or-documented-unavailability-required");
  if ((billingReads === null) !== (billingWrites === null)) reasons.push("billing-usage-incomplete");
  if (billingAvailable && billingUnavailableReason) reasons.push("billing-evidence-conflicting");
  const billingObservedAt = Date.parse(billing?.observedAt || "");
  const nowMs = new Date(now).getTime();
  if (billingAvailable && !Number.isFinite(billingObservedAt)) reasons.push("billing-observation-time-required");
  else if (billingAvailable && (billingObservedAt > nowMs + 30_000 || nowMs - billingObservedAt > D1_BUDGET_LIMITS.maximumBillingObservationAgeMs)) reasons.push("billing-observation-stale");

  const effectiveReads = Math.max(analyticsReads, billingReads ?? 0);
  const effectiveWrites = Math.max(analyticsWrites, billingWrites ?? 0);
  if (billingAvailable && materiallyDifferent(analyticsReads, billingReads)) reasons.push("billing-read-reconciliation-failed");
  if (billingAvailable && materiallyDifferent(analyticsWrites, billingWrites)) reasons.push("billing-write-reconciliation-failed");
  // A sample can be internally trustworthy before it is old enough or cheap
  // enough to authorize work. Preserve that distinction so an early passive
  // sample can be used later to calculate burn rate while still returning STOP.
  const evidenceReasons = [...reasons];
  if (effectiveReads >= D1_BUDGET_LIMITS.optionalReadStartMaximum) reasons.push("read-start-threshold-exceeded");
  if (effectiveWrites >= D1_BUDGET_LIMITS.optionalWriteStartMaximum) reasons.push("write-start-threshold-exceeded");
  if (effectiveReads >= D1_BUDGET_LIMITS.optionalReadStop) reasons.push("read-stop-threshold-exceeded");
  if (effectiveWrites >= D1_BUDGET_LIMITS.optionalWriteStop) reasons.push("write-stop-threshold-exceeded");
  if (new Date(now).getTime() - Date.parse(interval.start) < D1_BUDGET_LIMITS.minimumPassiveMs) reasons.push("two-hour-passive-baseline-required");

  const expensiveFingerprints = (analytics?.queryFingerprints || []).filter((item) => item.averageRowsRead > D1_BUDGET_LIMITS.maximumOrdinaryRowsPerQuery);
  if (expensiveFingerprints.length) reasons.push("query-per-invocation-limit-exceeded");

  let burnRateRowsPerHour = null;
  let projectedReads = null;
  const previousGeneratedAt = Date.parse(previousReport?.generatedAt || "");
  if (!Number.isFinite(previousGeneratedAt)) {
    reasons.push("second-sample-required");
  } else {
    if (previousReport?.sampleValid !== true) reasons.push("previous-sample-incomplete");
    if (previousReport?.interval?.date !== interval.date) reasons.push("previous-sample-not-current-utc-day");
    const gap = nowMs - previousGeneratedAt;
    if (gap < D1_BUDGET_LIMITS.minimumSampleGapMs) reasons.push("sample-gap-too-short");
    const previousReads = finiteNonNegative(previousReport?.effectiveUsage?.rowsRead);
    const previousWrites = finiteNonNegative(previousReport?.effectiveUsage?.rowsWritten);
    if (previousReads === null || previousWrites === null) {
      reasons.push("previous-sample-invalid");
    } else if (effectiveReads < previousReads || effectiveWrites < previousWrites) {
      reasons.push("usage-counter-decreased");
    } else if (gap > 0) {
      burnRateRowsPerHour = (effectiveReads - previousReads) * (60 * 60 * 1000 / gap);
      projectedReads = effectiveReads + burnRateRowsPerHour * ((Date.parse(interval.end) - nowMs) / (60 * 60 * 1000));
      if (projectedReads > D1_BUDGET_LIMITS.ordinaryReadReserve) reasons.push("projected-ordinary-read-reserve-exceeded");
    }
  }

  const estimatedReads = finiteNonNegative(optionalEstimate.rowsRead) ?? 0;
  const estimatedWrites = finiteNonNegative(optionalEstimate.rowsWritten) ?? 0;
  if (effectiveReads + estimatedReads * 2 > D1_BUDGET_LIMITS.optionalReadStop) reasons.push("optional-read-estimate-exceeds-budget");
  if (effectiveWrites + estimatedWrites * 2 > D1_BUDGET_LIMITS.optionalWriteStop) reasons.push("optional-write-estimate-exceeds-budget");

  return {
    decision: reasons.length ? "STOP" : "GO",
    sampleValid: evidenceReasons.length === 0,
    reasons: [...new Set(reasons)],
    interval,
    effectiveUsage: { rowsRead: effectiveReads, rowsWritten: effectiveWrites },
    billingMode: billingAvailable ? "billing-reconciled" : billingExplicitlyUnavailable ? "free-plan-analytics-only" : "unverified",
    burnRateRowsPerHour,
    projectedReads,
    expensiveFingerprints: expensiveFingerprints.slice(0, 20),
    optionalEstimate: { rowsRead: estimatedReads, rowsWritten: estimatedWrites, contingencyMultiplier: 2 },
  };
}

function stoppedAnalytics(reason) {
  return { complete: false, reasons: [reason], databases: [], totals: emptyUsage("account"), timeline: [], timelineTotals: emptyUsage("timeline"), fingerprintTotals: { rowsRead: 0, rowsWritten: 0, queries: 0 }, unattributed: { rowsRead: 0, rowsWritten: 0 }, queryFingerprints: [] };
}

function emptyUsage(databaseId) {
  return { databaseId, rowsRead: 0, rowsWritten: 0, readQueries: 0, writeQueries: 0 };
}

function addUsage(total, usage) {
  total.rowsRead += usage.rowsRead;
  total.rowsWritten += usage.rowsWritten;
  total.readQueries += usage.readQueries;
  total.writeQueries += usage.writeQueries;
  return total;
}

function nonNegative(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function finiteNonNegative(value) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function materiallyDifferent(left, right) {
  const difference = Math.abs(left - right);
  return difference > Math.max(D1_BUDGET_LIMITS.maximumReconciliationFloor, Math.max(left, right) * D1_BUDGET_LIMITS.maximumReconciliationFraction);
}

function strictlyDifferent(left, right) {
  const difference = Math.abs(left - right);
  return difference > Math.max(
    D1_BUDGET_LIMITS.maximumAnalyticsReconciliationFloor,
    Math.max(left, right) * D1_BUDGET_LIMITS.maximumAnalyticsReconciliationFraction,
  );
}
