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

export function accountMaintenanceHeadroom(analytics, outstanding, now = new Date(), settledMaintenance = { reads: 0, writes: 0 }) {
  const nowMs = new Date(now).getTime();
  const cutoff = Date.parse(analytics.interval.observedUntil);
  const elapsed = cutoff - Date.parse(analytics.interval.start);
  const remaining = Date.parse(analytics.interval.end) - cutoff;
  // Project ordinary account traffic, excluding only confirmed measured
  // maintenance that has already settled. A one-off import must not be
  // projected as though it repeats continuously through the rest of the day.
  // Unknown/legacy work stays in the forecast and in its reservation.
  const ordinaryReads = Math.max(0, analytics.totals.rowsRead - settledMaintenance.reads);
  const ordinaryWrites = Math.max(0, analytics.totals.rowsWritten - settledMaintenance.writes);
  const forecastReads = Math.ceil(ordinaryReads * remaining / Math.max(elapsed, 1));
  const forecastWrites = Math.ceil(ordinaryWrites * remaining / Math.max(elapsed, 1));
  if (!analytics.complete || !Number.isFinite(cutoff) || cutoff > nowMs || nowMs - cutoff > 25 * 60 * 1000 || elapsed <= 0) throw new Error("Account analytics are unavailable or stale.");
  const reads = Math.max(0, Math.floor(ACCOUNT_MAINTENANCE_POLICY.reads - analytics.totals.rowsRead - outstanding.reads - forecastReads));
  const writes = Math.max(0, Math.floor(ACCOUNT_MAINTENANCE_POLICY.writes - analytics.totals.rowsWritten - outstanding.writes - forecastWrites));
  return { reads, writes, forecastReads, forecastWrites, validUntil: new Date(Math.min(nowMs + ACCOUNT_MAINTENANCE_POLICY.validityMs, Date.parse(analytics.interval.end))).toISOString() };
}
