// D1 returns one summary, not a growing daily list of receipt objects. The
// original inspection ceilings and reservation fences are intentionally kept.
export async function maintenanceReceiptSummary(db, day, previousDay, dayStart, cutoff) {
  const row = await db.prepare(`WITH today AS (
    SELECT * FROM roster_maintenance_receipts WHERE utc_day=? LIMIT 10001
  ), carry AS (
    SELECT * FROM roster_maintenance_receipts
    WHERE utc_day=? AND recover_after>? AND reconciled_at='' LIMIT 10001
  ), pending AS (SELECT *,1 AS current_day FROM today UNION ALL SELECT *,0 AS current_day FROM carry)
  SELECT
    COALESCE(SUM(current_day),0) AS today_count,
    COALESCE(SUM(1-current_day),0) AS carry_count,
    COALESCE(SUM(CASE WHEN current_day=1 THEN reserved_reads ELSE 0 END),0) AS receipt_reads,
    COALESCE(SUM(CASE WHEN current_day=1 THEN reserved_writes ELSE 0 END),0) AS receipt_writes,
    COALESCE(SUM(CASE WHEN current_day=1 AND reconciled_at='' AND metadata_complete=1
      AND finished_at<>'' AND finished_at<=? THEN MAX(0,actual_reads-24) ELSE 0 END),0) AS settled_reads,
    COALESCE(SUM(CASE WHEN current_day=1 AND reconciled_at='' AND metadata_complete=1
      AND finished_at<>'' AND finished_at<=? THEN MAX(0,actual_writes-24) ELSE 0 END),0) AS settled_writes,
    COALESCE(SUM(CASE WHEN reconciled_at<>'' AND reconciled_at<=? THEN 0
      WHEN metadata_complete<>1 OR finished_at='' THEN reserved_reads
      WHEN finished_at>? THEN actual_reads ELSE 0 END),0) AS outstanding_reads,
    COALESCE(SUM(CASE WHEN reconciled_at<>'' AND reconciled_at<=? THEN 0
      WHEN metadata_complete<>1 OR finished_at='' THEN reserved_writes
      WHEN finished_at>? THEN actual_writes ELSE 0 END),0) AS outstanding_writes
  FROM pending`).bind(day,previousDay,dayStart,cutoff,cutoff,cutoff,cutoff,cutoff,cutoff).first();
  const values = Object.values(row || {});
  if (values.length!==8 || values.some(v=>!Number.isSafeInteger(Number(v)) || Number(v)<0)) throw Error('Maintenance summary unavailable or invalid.');
  if (Number(row.today_count)>10000) throw Error('Maintenance receipt inspection bound reached.');
  if (Number(row.carry_count)>10000) throw Error('Previous-day maintenance receipt inspection bound reached.');
  return {outstanding:{reads:Number(row.outstanding_reads),writes:Number(row.outstanding_writes)},
    settled:{reads:Number(row.settled_reads),writes:Number(row.settled_writes)},
    receiptReads:Number(row.receipt_reads),receiptWrites:Number(row.receipt_writes)};
}
