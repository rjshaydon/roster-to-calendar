const sessionLimits = new WeakMap();

// A workflow may tighten the daily allowance for one admitted pass. It can
// never expand the global limits. The SQL reservation enforces both atomically.
export function configureRosterMaintenanceBudget(db, session) {
  if (!session) return;
  const writes = Number(session.baseWrites);
  const reads = Number(session.baseReads);
  if (!Number.isFinite(writes) || !Number.isFinite(reads) || writes < 0 || reads < 0) throw new Error("Invalid maintenance session.");
  const sameDay = session.utcDay === new Date().toISOString().slice(0, 10);
  sessionLimits.set(db, { writes: Math.min(10000, (sameDay ? writes : 0) + 5000), reads: Math.min(500000, (sameDay ? reads : 0) + 250000) });
}

// A single shared allowance for automated imports and shared-view rebuilding.
// Reservations include index writes and survive failures; no blind retry gets
// a fresh allowance. Account-wide analytics remains the rollout admission gate.
export async function reserveRosterMaintenanceBudget(db, writes = 0, reads = 0) {
  if (!Number.isFinite(writes) || !Number.isFinite(reads) || writes < 0 || reads < 0) throw new Error("Invalid maintenance cost estimate.");
  const day = new Date().toISOString().slice(0, 10);
  const limits = sessionLimits.get(db) || { writes: 10000, reads: 500000 };
  const results = await db.batch([
    db.prepare("INSERT OR IGNORE INTO roster_import_daily_budget (utc_day, reserved_writes, reserved_reads) VALUES (?, 0, 0)").bind(day),
    db.prepare(`UPDATE roster_import_daily_budget SET reserved_writes = reserved_writes + ?, reserved_reads = reserved_reads + ?
      WHERE utc_day = ? AND reserved_writes + ? <= ? AND reserved_reads + ? <= ?`)
      .bind(Math.ceil(writes) + 4, Math.ceil(reads) + 4, day, Math.ceil(writes) + 4, limits.writes, Math.ceil(reads) + 4, limits.reads),
  ]);
  return Number(results[1]?.meta?.changes || 0) === 1;
}

export function maintenanceBudgetDeferredError() {
  const error = new Error("Automated maintenance allowance reached; progress retained for a later bounded pass.");
  error.code = "ROSTER_MAINTENANCE_DEFERRED";
  return error;
}
