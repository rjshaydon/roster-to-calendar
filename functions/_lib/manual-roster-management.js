import { validateRosterImportManifest, rosterImportDigest } from './roster-import-batches.js';
import { pinManualAuthority } from './roster-authority.js';
import { beginBoundedRosterImport, stageBoundedRosterBatch, prepareBoundedRosterPresence, prepareBoundedRosterMetadata, activateBoundedRosterTerm } from './roster-import-staging.js';
import { loadRosterFileStatusSummary } from './d1-calendar.js';
import { reserveRosterMaintenanceBudget } from './roster-maintenance-budget.js';
import { automaticFacilityPublicationEnabled, facilityRefreshStatements, facilityTermDates } from './facility-refresh-queue.js';

export const MANUAL_ROSTER_SOURCES = Object.freeze({ mmc: 'manual-mmc', mch: 'manual-mch', ddh: 'manual-ddh', vhh: 'manual-vhh' });
// Omitting the legacy index is safe only while every user reads published
// artifacts. Do not infer this from a client flag or from one cohort's rollout.
export function manualCachedPresenceEnabled(env,sourceType) {
  return Boolean(env.ROSTER_FILES?.get && env.ROSTER_FILES?.put
    && automaticFacilityPublicationEnabled(env,sourceType)
    && env.FACILITY_SHARED_ROLLOUT_ACTIVE === 'true'
    && env.FACILITY_SHARED_EMERGENCY_PAUSED !== 'true'
    && env.FACILITY_SHARED_READER_COHORT === 'all'
    && env.FACILITY_SHARED_READER_SOURCE_ALLOWLIST?.split(',').map(s=>s.trim()).includes(sourceType)
    && env.FACILITY_SHARED_METADATA_ENABLED === 'true'
    && env.FACILITY_SHARED_DAYS_ENABLED === 'true'
    && env.FACILITY_LEGACY_READS_PAUSED === 'true'
    && env.ROSTER_INSIGHT_READS_ENABLED === 'true'
    && env.SUBSCRIPTION_ARTIFACT_MODE === 'ready');
}

export async function handleManualRosterImport(context, body, email) {
  const sourceType = String(body.file?.sourceType || '').toLowerCase();
  const sourceId = MANUAL_ROSTER_SOURCES[sourceType];
  const id = String(body.file?.id || '');
  if (!sourceId || !/^manual:[a-f0-9-]{36}$/.test(id) || body.file.sourceId !== sourceId) throw new Error('A bounded manual roster identity and supported site are required.');
  const db = context.env.ROSTER_DB, runId = `manual:${id}`, revision = String(body.revision || '');
  const file = { ...body.file, sourceId, sourceType, active: false, uploadedBy: email };
  const cachedAllowed=manualCachedPresenceEnabled(context.env,sourceType);
  if(body.manifest?.presenceMode==='cached' && !cachedAllowed)throw new Error('Published manual-import readers are unavailable; keep the existing roster and retry when restored.');
  if(!['bounded-begin','bounded-initialize'].includes(body.phase) && !cachedAllowed){
    const job=await db.prepare('SELECT manifest_json FROM roster_import_jobs WHERE run_id=?').bind(runId).first();
    if(job && JSON.parse(job.manifest_json).presenceMode==='cached')throw new Error('Published manual-import readers are unavailable; the import is retained for retry.');
  }
  let result;
  switch (body.phase) {
    case 'bounded-begin':
    case 'bounded-initialize': {
      validateRosterImportManifest(body.manifest);
      if(body.manifest.fileId!==id || body.manifest.sourceId!==sourceId || await rosterImportDigest(body.manifest)!==revision)throw new Error('Manual import plan does not match its file and revision.');
      if(body.phase==='bounded-begin' && body.deferStaging===true && body.supportsCachedPresence===true) {
        const existing=await db.prepare('SELECT manifest_json,activated FROM roster_import_jobs WHERE run_id=?').bind(runId).first();
        // Existing jobs keep their original mode/hash/cursors. Never turn a
        // partially indexed legacy plan into a differently hashed cached job.
        return {ok:true,mode:'staging-ready',presenceMode:existing?JSON.parse(existing.manifest_json).presenceMode:(cachedAllowed?'cached':undefined)};
      }
      await pinManualAuthority(db,runId,sourceType);
      result=await beginBoundedRosterImport(db,runId,file,{manifest:body.manifest,revision},{allowReplacement:true});
      break;
    }
    case 'bounded-events': result = await stageBoundedRosterBatch(db, runId, file, revision, body.batch); break;
    case 'bounded-presence': result = await prepareBoundedRosterPresence(db, runId, file, revision, body.batch); break;
    case 'bounded-metadata': result = await prepareBoundedRosterMetadata(db, runId, revision, body.batch); break;
    case 'bounded-activate': result = await activateBoundedRosterTerm(db, runId, revision, { publishFacility: automaticFacilityPublicationEnabled(context.env, sourceType) }); break;
    default: throw new Error('A bounded manual import phase is required.');
  }
  if (result.activated || result.completed || body.phase === 'bounded-activate' && !result.deferred) {
    const summary = await loadRosterFileStatusSummary(db, id);
    return { ok:true, indexing:'complete', completed:true, ...result, fileStatus:{ ...summary, status:'populated' } };
  }
  return { ok:true, indexing:'in-progress', ...result };
}

export async function deactivateManualRosterFiles(context, ids) {
  if (!Array.isArray(ids) || !ids.length || ids.length > 4 || ids.some(id => typeof id !== 'string' || !id || id.length > 500)) throw new Error('Remove between one and four roster files per request.');
  const db = context.env.ROSTER_DB;
  if (!await reserveRosterMaintenanceBudget(db, 512, 4096)) return { ok:true, deferred:true };
  const files = (await db.prepare(`SELECT f.id,f.source_type,f.active,f.parsed_at,c.coverage_start,c.coverage_end FROM roster_files f LEFT JOIN roster_file_coverage c ON c.file_id=f.id WHERE f.id IN (${ids.map(() => '?').join(',')})`).bind(...ids).all()).results;
  const statements = [], now = new Date().toISOString();
  for (const file of files) {
    if (!MANUAL_ROSTER_SOURCES[file.source_type] || !file.coverage_start || !file.coverage_end) throw new Error('Roster removal requires prepared coverage for a supported site.');
    if (!file.active) continue;
    statements.push(db.prepare('UPDATE roster_files SET active=0 WHERE id=? AND active=1 AND parsed_at=?').bind(file.id,file.parsed_at));
    statements.push(db.prepare("UPDATE roster_sources SET active_file_id='' WHERE active_file_id=? AND EXISTS(SELECT 1 FROM roster_files WHERE id=? AND active=0 AND parsed_at=?)").bind(file.id,file.id,file.parsed_at));
    statements.push(db.prepare("UPDATE roster_file_status_summaries SET active=0,status_revision=?,updated_at=? WHERE file_id=? AND EXISTS(SELECT 1 FROM roster_files f WHERE f.id=file_id AND f.active=0 AND f.parsed_at=?)").bind(crypto.randomUUID(),now,file.id,file.parsed_at));
    if (automaticFacilityPublicationEnabled(context.env,file.source_type)) statements.push(...facilityRefreshStatements(db,file.source_type,facilityTermDates(file.coverage_start,file.coverage_end),`removed:${file.id}:${file.parsed_at}`));
  }
  if (statements.length) await db.batch(statements);
  const remaining = (await db.prepare(`SELECT id,active FROM roster_files WHERE id IN (${ids.map(() => '?').join(',')})`).bind(...ids).all()).results;
  if (remaining.some(file => file.active)) throw new Error('A concurrent roster update prevented removal. Review the file and retry.');
  return { ok:true, removedFileIds:ids, removedImportIds:ids, allDeactivated:true, verification:ids.map(fileId => ({ fileId,deactivated:true })), sourceTypes:[...new Set(files.map(file => file.source_type))] };
}
