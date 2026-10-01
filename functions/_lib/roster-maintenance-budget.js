// A single shared allowance for automated imports and shared-view rebuilding.
// Reservations include index writes and survive failures; no blind retry gets
// a fresh allowance. Account-wide analytics remains the rollout admission gate.
export async function reserveRosterMaintenanceBudget(db, writes = 0, reads = 0) {
  if (!Number.isFinite(writes) || !Number.isFinite(reads) || writes < 0 || reads < 0) throw new Error("Invalid maintenance cost estimate.");
  const day = new Date().toISOString().slice(0, 10);
  const results = await db.batch([
    db.prepare("INSERT OR IGNORE INTO roster_import_daily_budget (utc_day, reserved_writes, reserved_reads) VALUES (?, 0, 0)").bind(day),
    db.prepare(`UPDATE roster_import_daily_budget SET reserved_writes = reserved_writes + ?, reserved_reads = reserved_reads + ?
      WHERE utc_day = ? AND reserved_writes + ? <= 10000 AND reserved_reads + ? <= 500000`)
      .bind(Math.ceil(writes) + 4, Math.ceil(reads) + 4, day, Math.ceil(writes) + 4, Math.ceil(reads) + 4),
  ]);
  return Number(results[1]?.meta?.changes || 0) === 1;
}

export function maintenanceBudgetDeferredError() {
  const error = new Error("Automated maintenance allowance exhausted; progress retained until the next UTC day.");
  error.code = "ROSTER_MAINTENANCE_DEFERRED";
  return error;
}
