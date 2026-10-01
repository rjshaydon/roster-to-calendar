const sessionLimits = new WeakMap();
const requests = new WeakMap();

export function beginMaintenanceAccounting(db, requestId, enabled) {
  if (enabled) requests.set(db, { id: requestId, reserved: false });
}

export async function finishMaintenanceAccounting(db, meter, bookkeepingDb = db) {
  const request = requests.get(db);
  if (!request?.reserved) return;
  const reserved = await bookkeepingDb.prepare("SELECT reserved_reads,reserved_writes FROM roster_maintenance_receipts WHERE request_id=?").bind(request.id).first();
  if (meter.rowsRead + 24 > Number(reserved?.reserved_reads || 0) || meter.rowsWritten + 24 > Number(reserved?.reserved_writes || 0)) {
    await stopAccountMaintenance(bookkeepingDb, "cost-overrun:request");
  }
  // Work without complete billing metadata keeps its whole reservation.
  // Complete metadata includes all route work, not just the publication core.
  await bookkeepingDb.prepare(`UPDATE roster_maintenance_receipts SET finished_at=?, metadata_complete=?, actual_reads=?, actual_writes=? WHERE request_id=?`)
    .bind(new Date().toISOString(), meter.metadataComplete ? 1 : 0,
      Math.ceil(meter.rowsRead) + 24, Math.ceil(meter.rowsWritten) + 24, request.id).run();
}

export async function stopAccountMaintenance(db, reason) {
  await db.prepare("UPDATE roster_account_budget SET valid_until='', stop_reason=CASE WHEN stop_reason LIKE 'cost-overrun:%' THEN stop_reason ELSE ? END WHERE utc_day=?")
    .bind(String(reason).slice(0, 200), new Date().toISOString().slice(0, 10)).run();
}

// Legacy deployments retain their admitted pass ceilings. Account-aware
// deployments use the durable shared grant instead of caller-supplied limits.
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
  const request = requests.get(db);
  if (request) {
    const estimatedWrites = Math.ceil(writes) + 24;
    const estimatedReads = Math.ceil(reads) + 24;
    const results = await db.batch([
      db.prepare(`UPDATE roster_account_budget SET allocated_writes=allocated_writes+?, allocated_reads=allocated_reads+?
        WHERE utc_day=? AND valid_until>? AND stop_reason=''
        AND allocated_writes+?<=maximum_writes AND allocated_reads+?<=maximum_reads`)
        .bind(estimatedWrites, estimatedReads, day, new Date().toISOString(), estimatedWrites, estimatedReads),
      db.prepare(`INSERT INTO roster_import_daily_budget (utc_day,reserved_writes,reserved_reads)
        SELECT ?,?,? WHERE changes()=1 ON CONFLICT(utc_day) DO UPDATE SET
        reserved_writes=reserved_writes+excluded.reserved_writes,reserved_reads=reserved_reads+excluded.reserved_reads`)
        .bind(day, estimatedWrites, estimatedReads),
      db.prepare(`INSERT INTO roster_maintenance_receipts (request_id,utc_day,reserved_writes,reserved_reads)
        SELECT ?,?,?,? WHERE changes()=1 ON CONFLICT(request_id) DO UPDATE SET
        reserved_writes=reserved_writes+excluded.reserved_writes,reserved_reads=reserved_reads+excluded.reserved_reads`)
        .bind(request.id, day, estimatedWrites, estimatedReads),
    ]);
    const admitted = Number(results[0]?.meta?.changes || 0) === 1;
    if (admitted) request.reserved = true;
    return admitted;
  }
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
