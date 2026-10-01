export const ACCOUNT_MAINTENANCE_POLICY = Object.freeze({ reads: 4000000, writes: 80000, validityMs: 10 * 60 * 1000 });

// Receipts finished before the settled cutoff are already in account analytics.
// Unfinished/uninstrumented work retains its full reservation. Never subtract
// an estimated maintenance cost from Cloudflare's measured account total.
export function outstandingMaintenance(receipts, cutoff) {
  return receipts.reduce((total, row) => {
    const complete = Number(row.metadata_complete) === 1 && row.finished_at;
    if (!complete) {
      total.reads += Number(row.reserved_reads);
      total.writes += Number(row.reserved_writes);
    } else if (row.finished_at > cutoff) {
      total.reads += Number(row.actual_reads);
      total.writes += Number(row.actual_writes);
    }
    return total;
  }, { reads: 0, writes: 0 });
}

export function accountMaintenanceHeadroom(analytics, outstanding, now = new Date()) {
  const nowMs = new Date(now).getTime();
  const cutoff = Date.parse(analytics.interval.observedUntil);
  const recent = (analytics.fiveMinuteBuckets || []).filter(row => Date.parse(row.observedAt) >= cutoff - 60 * 60 * 1000);
  const elapsed = Math.min(60 * 60 * 1000, cutoff - Date.parse(analytics.interval.start));
  const remaining = Date.parse(analytics.interval.end) - cutoff;
  const totals = recent.reduce((sum, row) => ({ reads: sum.reads + row.rowsRead, writes: sum.writes + row.rowsWritten }), { reads: 0, writes: 0 });
  // Account-wide burn includes other apps and interactive requests. Reserving
  // its projection is conservative even when the last hour contained imports.
  const forecastReads = Math.ceil(totals.reads * remaining / Math.max(elapsed, 1));
  const forecastWrites = Math.ceil(totals.writes * remaining / Math.max(elapsed, 1));
  if (!analytics.complete || !Number.isFinite(cutoff) || cutoff > nowMs || nowMs - cutoff > 25 * 60 * 1000 || elapsed <= 0) throw new Error("Account analytics are unavailable or stale.");
  const reads = Math.max(0, Math.floor(ACCOUNT_MAINTENANCE_POLICY.reads - analytics.totals.rowsRead - outstanding.reads - forecastReads));
  const writes = Math.max(0, Math.floor(ACCOUNT_MAINTENANCE_POLICY.writes - analytics.totals.rowsWritten - outstanding.writes - forecastWrites));
  return { reads, writes, forecastReads, forecastWrites, validUntil: new Date(Math.min(nowMs + ACCOUNT_MAINTENANCE_POLICY.validityMs, Date.parse(analytics.interval.end))).toISOString() };
}
