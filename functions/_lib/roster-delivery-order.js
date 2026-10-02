import { reserveRosterMaintenanceBudget, maintenanceBudgetDeferredError } from './roster-maintenance-budget.js';
const RECENT_RUNS = `SELECT id,status,started_at,source_file_id,file_id FROM roster_sync_runs
  INDEXED BY idx_roster_sync_runs_source_started_id WHERE source_id=?
  ORDER BY started_at DESC,id DESC LIMIT 65`;
const eligible = row => ['queued','processing','success'].includes(row.status);
// Immutable raw-file timestamps prevent old retries replacing newer workbooks.
export async function rosterDeliveryOrder(db, run) {
  if (!run.startedAt) return {};
  const incoming = await db.prepare('SELECT name,last_modified FROM raw_roster_files WHERE file_id=?').bind(run.sourceFileId || run.fileId).first();
  if (!incoming?.name) throw new Error('Queued provider revision metadata is unavailable.');
  const rows = (await db.prepare(`SELECT r.*,f.name,f.last_modified FROM (${RECENT_RUNS}) r LEFT JOIN raw_roster_files f ON f.file_id=COALESCE(NULLIF(r.source_file_id,''),r.file_id)`).bind(run.sourceId).all()).results;
  const isNewer = row => row.id !== run.id && eligible(row) && String(row.name).toLowerCase() === incoming.name.toLowerCase()
    && (Number(row.last_modified)>0 && Number(incoming.last_modified)>0
      ? Number(row.last_modified)>Number(incoming.last_modified)
      : row.started_at>run.startedAt || (row.started_at===run.startedAt && row.id>run.id));
  if (rows.some(isNewer)) return { superseded: true };
  const active = (await db.prepare(`SELECT source_id,name,last_modified FROM roster_files INDEXED BY idx_roster_files_source_active
    WHERE source_type=(SELECT source_type FROM roster_sources WHERE id=?) AND active=1 LIMIT 33`).bind(run.sourceId).all()).results;
  if (active.length>32) throw new Error('Provider revision ordering exceeds its 32-active-file safety window.');
  if (active.some(row=>row.source_id===run.sourceId && String(row.name).toLowerCase()===incoming.name.toLowerCase() && Number(row.last_modified)>Number(incoming.last_modified) && Number(incoming.last_modified)>0)) return {superseded:true};
  if (rows.length>64 && !rows.slice(0,64).some(row=>row.id===run.id)) throw new Error('Provider revision ordering exceeds its 64-run safety window.');
  const modified=Number(incoming.last_modified || 0);
  const guard=()=>db.prepare(`WITH recent AS (${RECENT_RUNS}) SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM recent r
    JOIN raw_roster_files f ON f.file_id=COALESCE(NULLIF(r.source_file_id,''),r.file_id)
    WHERE r.id<>? AND r.status IN ('queued','processing','success') AND LOWER(f.name)=LOWER(?)
    AND CASE WHEN f.last_modified>0 AND ?>0 THEN f.last_modified>?
      ELSE r.started_at>? OR (r.started_at=? AND r.id>?) END) AND NOT EXISTS(SELECT 1 FROM roster_files INDEXED BY idx_roster_files_source_active
      WHERE source_type=(SELECT source_type FROM roster_sources WHERE id=?) AND active=1 AND source_id=?
      AND LOWER(name)=LOWER(?) AND ?>0 AND last_modified>?) THEN 1 ELSE json('stale-roster-delivery') END`)
    .bind(run.sourceId,run.id,incoming.name,modified,modified,run.startedAt,run.startedAt,run.id,run.sourceId,run.sourceId,incoming.name,modified,modified);
  return { guard };
}
export async function skipSupersededRosterDelivery(db, run) {
  if (!await reserveRosterMaintenanceBudget(db,16,160)) throw maintenanceBudgetDeferredError();
  await db.prepare("UPDATE roster_sync_runs SET status='superseded',message=?,completed_at=? WHERE id=? AND status IN ('queued','processing')")
    .bind('Older provider revision skipped; a newer delivery is queued or imported.',new Date().toISOString(),run.id).run();
  return { ok:true,completed:true,superseded:true,doctorCount:0,eventCount:0,fileId:run.fileId };
}
