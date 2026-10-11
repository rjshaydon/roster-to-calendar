const sessionLimits = new WeakMap();
const requests = new WeakMap();

export function beginMaintenanceAccounting(db, requestId, enabled, options = {}) {
  if (enabled) requests.set(db, { id: requestId, reserved: false, ...options });
}

export async function finishMaintenanceAccounting(db, meter, bookkeepingDb = db) {
  const request = requests.get(db);
  if (!request?.reserved) return;
  const reserved = await bookkeepingDb.prepare("SELECT utc_day,reserved_reads,reserved_writes,finished_at,reconciled_at FROM roster_maintenance_receipts WHERE request_id=?").bind(request.id).first();
  if (!reserved || reserved.finished_at || reserved.reconciled_at) return;
  // Route totals already include reservation-table mutations. Allow another
  // 24 units for this bounded settlement, including its indexes and refund.
  const actualReads = Math.ceil(meter.rowsRead) + 24;
  const actualWrites = Math.ceil(meter.rowsWritten) + 24;
  if (actualReads > Number(reserved.reserved_reads) || actualWrites > Number(reserved.reserved_writes)) {
    await stopAccountMaintenance(bookkeepingDb, "cost-overrun:request");
  }
  const refundReads = meter.metadataComplete ? Math.max(0, Number(reserved.reserved_reads) - actualReads) : 0;
  const refundWrites = meter.metadataComplete ? Math.max(0, Number(reserved.reserved_writes) - actualWrites) : 0;
  // Settlement and unused-grant release commit together. Repeated settlement
  // cannot refund twice; unknown/lost responses retain the whole reservation.
  await bookkeepingDb.batch([
    bookkeepingDb.prepare(`UPDATE roster_maintenance_receipts SET finished_at=?, metadata_complete=?, actual_reads=?, actual_writes=? WHERE request_id=? AND finished_at='' AND reconciled_at=''`)
      .bind(new Date().toISOString(), meter.metadataComplete ? 1 : 0, actualReads, actualWrites, request.id),
    bookkeepingDb.prepare(`UPDATE roster_account_budget SET allocated_reads=MAX(0,allocated_reads-?),allocated_writes=MAX(0,allocated_writes-?) WHERE utc_day=? AND changes()=1`)
      .bind(refundReads, refundWrites, reserved.utc_day),
  ]);
}

// A deadline is enforced before every metered D1 call. Allow the maximum
// admitted statement count * D1's 30-second query limit, plus settlement calls,
// to drain before a later, validated analytics cutoff can account for this work.
// Legacy receipts have no enforced deadline and are deliberately ineligible.
export async function recoverAbandonedMaintenance(db, day, cutoff) {
  const rows=(await db.prepare(`SELECT request_id,reserved_reads,reserved_writes FROM roster_maintenance_receipts
    WHERE utc_day=? AND reconciled_at='' AND recover_after<>'' AND recover_after<=?
    AND (finished_at='' OR metadata_complete<>1) LIMIT 10`).bind(day,cutoff).all()).results||[];
  let recovered=0;
  for(const row of rows) {
    const result=await db.batch([
      db.prepare(`UPDATE roster_maintenance_receipts SET reconciled_at=? WHERE request_id=? AND utc_day=?
        AND reconciled_at='' AND recover_after<>'' AND recover_after<=? AND (finished_at='' OR metadata_complete<>1)`)
        .bind(cutoff,row.request_id,day,cutoff),
      db.prepare(`UPDATE roster_account_budget SET allocated_reads=MAX(0,allocated_reads-?),allocated_writes=MAX(0,allocated_writes-?)
        WHERE utc_day=? AND changes()=1`).bind(row.reserved_reads,row.reserved_writes,day),
    ]);
    recovered+=Number(result[0]?.meta?.changes||0);
  }
  return recovered;
}

export async function optionalMaintenanceAvailable(db) {
  const day=new Date().toISOString().slice(0,10);
  const unresolved=(await db.prepare(`SELECT request_id FROM roster_maintenance_receipts WHERE utc_day=?
    AND reconciled_at='' AND (finished_at='' OR metadata_complete<>1) LIMIT 3`).bind(day).all()).results||[];
  return unresolved.length<3;
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
    const estimatedWrites = Math.ceil(writes) + 48;
    const estimatedReads = Math.ceil(reads) + 48;
    const optional=request.purpose==='identity';
    if(optional && !await optionalMaintenanceAvailable(db)) return false;
    const deadline=request.deadlineAt||0;
    const recoverAfter=deadline?new Date(deadline+(Number(request.statementLimit||512)+8)*30000+120000).toISOString():'';
    const results = await db.batch([
      db.prepare(`UPDATE roster_account_budget SET allocated_writes=allocated_writes+?, allocated_reads=allocated_reads+?
        WHERE utc_day=? AND valid_until>? AND stop_reason=''
        AND allocated_writes+?<=maximum_writes AND allocated_reads+?<=maximum_reads`)
        .bind(estimatedWrites, estimatedReads, day, new Date().toISOString(), estimatedWrites+(optional?20000:0), estimatedReads+(optional?200000:0)),
      db.prepare(`INSERT INTO roster_import_daily_budget (utc_day,reserved_writes,reserved_reads)
        SELECT ?,?,? WHERE changes()=1 ON CONFLICT(utc_day) DO UPDATE SET
        reserved_writes=reserved_writes+excluded.reserved_writes,reserved_reads=reserved_reads+excluded.reserved_reads`)
        .bind(day, estimatedWrites, estimatedReads),
      db.prepare(`INSERT INTO roster_maintenance_receipts (request_id,utc_day,reserved_writes,reserved_reads,purpose,deadline_at,recover_after)
        SELECT ?,?,?,?,?,?,? WHERE changes()=1 ON CONFLICT(request_id) DO UPDATE SET
        reserved_writes=reserved_writes+excluded.reserved_writes,reserved_reads=reserved_reads+excluded.reserved_reads`)
        .bind(request.id, day, estimatedWrites, estimatedReads,request.purpose||'',deadline?new Date(deadline).toISOString():'',recoverAfter),
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
