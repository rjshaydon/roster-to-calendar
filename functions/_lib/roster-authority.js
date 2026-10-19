import { reserveRosterMaintenanceBudget, maintenanceBudgetDeferredError } from './roster-maintenance-budget.js';
import { rosterDeliveryOrder } from './roster-delivery-order.js';
import { rosterTermOwnership } from './roster-term-ownership.js';
import { facilityRefreshStatements, facilityTermDates } from './facility-refresh-queue.js';
const all = async (db,sql,...args) => (await db.prepare(sql).bind(...args).all()).results || [];
const sourceFor = sourceId => ({'monash-adults':'mmc','monash-paeds':'mch','vhh-active-medical-roster':'vhh','dandenong-findmyshift':'ddh'})[sourceId] || '';
const snapshotSql = `SELECT COALESCE(json_group_object(term_start,generation),'{}') AS generations FROM (SELECT term_start,generation FROM roster_term_authority WHERE source_type=? ORDER BY term_start LIMIT 65)`;
export async function authoritySnapshot(db,sourceType) {
 const row=await db.prepare(snapshotSql).bind(sourceType).first();
 if(Object.keys(JSON.parse(row.generations)).length>64)throw Error('Roster authority exceeds its term inspection limit.');
 return row.generations;
}
export function authorityFenceStatement(db,runId,sourceType,generations) {
 return db.prepare('INSERT OR IGNORE INTO roster_authority_fences(run_id,source_type,generations_json) VALUES(?,?,?)').bind(runId,sourceType,generations);
}
export async function pinManualAuthority(db,runId,sourceType) {
 if(await db.prepare('SELECT run_id FROM roster_authority_fences WHERE run_id=?').bind(runId).first())return;
 if(!await reserveRosterMaintenanceBudget(db,16,128))throw maintenanceBudgetDeferredError();
 // Snapshot and pin are one SQL statement; another activation cannot slip between them.
 await db.prepare(`INSERT OR IGNORE INTO roster_authority_fences(run_id,source_type,generations_json) SELECT ?,?,generations FROM (${snapshotSql})`).bind(runId,sourceType,sourceType).run();
}
export async function observeProviderAuthority(db,sourceId,fileName,providerVersion) {
 const source=sourceFor(sourceId);if(!source)return;
 const generations=await authoritySnapshot(db,source);
 if(generations==='{}')return;
 const previous=await db.prepare('SELECT provider_version,generations_json FROM roster_provider_observations WHERE source_id=? AND file_name=?').bind(sourceId,fileName.toLowerCase()).first();
 if(previous?.provider_version===providerVersion&&previous.generations_json===generations)return;
 if(!await reserveRosterMaintenanceBudget(db,16,128))throw maintenanceBudgetDeferredError();
 await db.prepare(`INSERT INTO roster_provider_observations(source_id,file_name,provider_version,generations_json) VALUES(?,?,?,?)
 ON CONFLICT(source_id,file_name) DO UPDATE SET provider_version=excluded.provider_version,generations_json=excluded.generations_json`).bind(sourceId,fileName.toLowerCase(),providerVersion,generations).run();
}
export async function providerAuthorityFence(db,sourceId,fileName,providerVersion) {
 const observation=await db.prepare('SELECT generations_json FROM roster_provider_observations WHERE source_id=? AND file_name=? AND provider_version=?').bind(sourceId,fileName.toLowerCase(),providerVersion).first();
 // Direct legacy deliveries have no fresh provider evidence for a manual fallback.
 // They can still update untouched terms, whose generation is zero.
 return observation?.generations_json || '{}';
}
export async function authorityTerms(db,fileId) {
 const rows=await all(db,'SELECT DISTINCT term_start FROM facility_stream_catalog_contributions WHERE file_id=? LIMIT 4',fileId);
 if(!rows.length||rows.length>3)throw Error('Roster authority requires prepared bounded term ownership.');
 return rows.map(r=>r.term_start);
}
export function authorityGuard(db,runId,sourceType,terms) {
 return db.prepare(`SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM roster_term_authority a
 WHERE a.source_type=? AND a.term_start IN (SELECT value FROM json_each(?))
 AND (a.generation<>COALESCE((SELECT CAST(json_extract(f.generations_json,'$."'||a.term_start||'"') AS INTEGER) FROM roster_authority_fences f WHERE f.run_id=?),0) OR (?=0 AND a.source_id LIKE 'manual-%' AND COALESCE((SELECT enabled FROM roster_authority_policy WHERE source_type=a.source_type),0)<>1)))
 THEN 1 ELSE json('stale-roster-authority') END`).bind(sourceType,JSON.stringify(terms),runId,runId.startsWith('manual:')?1:0);
}
export function authorityActivationStatements(db,{runId,sourceType,sourceId,fileId,revision,terms,now,gate='1',gateArgs=[]}) {
 return terms.map(term=>db.prepare(`INSERT INTO roster_term_authority(source_type,term_start,generation,file_id,source_id,content_revision,activated_at)
 SELECT ?,?,1,?,?,?,? WHERE ${gate} AND (? LIKE 'manual-%' OR EXISTS(SELECT 1 FROM roster_term_authority WHERE source_type=? AND term_start=?))
 ON CONFLICT(source_type,term_start) DO UPDATE SET generation=roster_term_authority.generation+1,file_id=excluded.file_id,source_id=excluded.source_id,content_revision=excluded.content_revision,activated_at=excluded.activated_at`)
 .bind(sourceType,term,fileId,sourceId,revision,now,...gateArgs,sourceId,sourceType,term));
}
// A successful retained receipt is not proof that its parsed file still contains
// those bytes: later patches may reuse its file ID. Require its exact raw revision
// and completed prepared snapshot, then restore control pointers only.
export async function restoreRetainedAuthority(db,run,{publishFacility=false,expectedGenerations=null}={}) {
 if(!run?.fileId||!sourceFor(run.sourceId))return null;
 const source=sourceFor(run.sourceId);
 const manual=await all(db,"SELECT term_start,generation FROM roster_term_authority WHERE source_type=? AND source_id LIKE 'manual-%' LIMIT 65",source);
 if(!manual.length)return null;
 let terms;try{terms=await authorityTerms(db,run.fileId);}catch{return {status:'repair-required',download:false,reason:'retained-snapshot-unavailable'};}
 if(!manual.some(row=>terms.includes(row.term_start)))return null;
 if(!Number((await db.prepare('SELECT enabled FROM roster_authority_policy WHERE source_type=?').bind(source).first())?.enabled))return {status:'handoff-paused',download:false};
 const retained=await db.prepare(`SELECT f.*,s.derived_state,s.content_revision,c.coverage_start,c.coverage_end,
 r.name AS raw_name,r.last_modified AS raw_modified FROM roster_files f
 JOIN roster_file_status_summaries s ON s.file_id=f.id JOIN roster_file_coverage c ON c.file_id=f.id
 JOIN raw_roster_files r ON r.file_id=? WHERE f.id=? AND f.source_id=?`).bind(run.sourceFileId||run.fileId,run.fileId,run.sourceId).first();
 if(!retained||retained.derived_state!=='ready'||!retained.content_revision||retained.active||retained.parsed_at>run.completedAt||Number(retained.last_modified)!==Number(retained.raw_modified))return {status:'repair-required',download:false};
 const active=await all(db,'SELECT id,parsed_at FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type=? AND active=1 ORDER BY id LIMIT 33',source);
 if(active.length>32)throw Error('Roster handoff exceeds its active-file inspection limit.');
 const incoming=await rosterTermOwnership(db,run.fileId),retire=[];
 let start=retained.coverage_start,end=retained.coverage_end;
 for(const old of active){const own=await rosterTermOwnership(db,old.id);
  if(own.firstDate>incoming.lastDate||own.lastDate<incoming.firstDate)continue;
  if(own.firstTerm!==incoming.firstTerm||own.lastTerm!==incoming.lastTerm||own.firstDate<incoming.firstDate||own.lastDate>incoming.lastDate)return {status:'repair-required',download:false};
  retire.push(old.id);
  const c=await db.prepare('SELECT coverage_start,coverage_end FROM roster_file_coverage WHERE file_id=?').bind(old.id).first();
  start=start<c.coverage_start?start:c.coverage_start;end=end>c.coverage_end?end:c.coverage_end;
 }
 const ordering=await rosterDeliveryOrder(db,run,[],false);
 if(ordering.superseded)return {status:'newer-revision-pending',download:false};
 if(!retire.length)return {status:'repair-required',download:false};
 if(!await reserveRosterMaintenanceBudget(db,1024,40000))return {status:'deferred',download:false,reason:'account-budget'};
 const generation=expectedGenerations||await authoritySnapshot(db,source),id=`handoff:${crypto.randomUUID()}`,now=new Date().toISOString();
 const fence=JSON.stringify(active);
 const statements=[...(ordering.guard?[ordering.guard()]:[]),authorityFenceStatement(db,id,source,generation),authorityGuard(db,id,source,terms),
  db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type=? AND active=1)=json_array_length(?)
   AND NOT EXISTS(SELECT 1 FROM roster_files f INDEXED BY idx_roster_files_source_active WHERE f.source_type=? AND f.active=1 AND NOT EXISTS(SELECT 1 FROM json_each(?) p WHERE json_extract(p.value,'$.id')=f.id AND json_extract(p.value,'$.parsed_at')=f.parsed_at))
   AND EXISTS(SELECT 1 FROM roster_files WHERE id=? AND active=0 AND parsed_at=?) THEN 1 ELSE json('roster-handoff-conflict') END`).bind(source,fence,source,fence,run.fileId,retained.parsed_at),
  db.prepare('UPDATE roster_files SET active=1 WHERE id=?').bind(run.fileId),
  db.prepare("UPDATE roster_file_status_summaries SET active=1,status_revision=?,updated_at=? WHERE file_id=?").bind(crypto.randomUUID(),now,run.fileId),
  ...retire.flatMap(fileId=>[db.prepare('UPDATE roster_files SET active=0 WHERE id=?').bind(fileId),db.prepare('UPDATE roster_file_status_summaries SET active=0,status_revision=?,updated_at=? WHERE file_id=?').bind(crypto.randomUUID(),now,fileId),
   db.prepare('UPDATE roster_sources SET active_file_id=? WHERE active_file_id=?').bind(run.fileId,fileId)]),
  ...authorityActivationStatements(db,{runId:id,sourceType:source,sourceId:run.sourceId,fileId:run.fileId,revision:retained.content_revision,terms,now}),
  ...(publishFacility?facilityRefreshStatements(db,source,facilityTermDates(start,end),`handoff:${run.fileId}:${retained.content_revision}:${id}`):[]),
  db.prepare('DELETE FROM roster_authority_fences WHERE run_id=?').bind(id),
 ];
 await db.batch(statements);
 return {status:'authority-restored',download:false,fileId:run.fileId};
}

// Fresh metadata may resume retained bytes after an intervening manual upload.
// Use a distinct deterministic run/staging identity; never repin an old job.
export async function reconcileQueuedAuthority(db,run,fileName) {
 const source=sourceFor(run.sourceId);if(!source)return null;
 const current=await authoritySnapshot(db,source);
 const previous=await db.prepare('SELECT generations_json FROM roster_authority_fences WHERE run_id=?').bind(run.id).first();
 if(current===(previous?.generations_json||'{}'))return null;
 const manual=await db.prepare("SELECT 1 AS found FROM roster_term_authority WHERE source_type=? AND source_id LIKE 'manual-%' LIMIT 1").bind(source).first();
 if(!manual)return null;
 const raw=await db.prepare('SELECT file_id,object_key,name FROM raw_roster_files WHERE file_id=?').bind(run.sourceFileId||run.fileId).first();
 if(!raw?.object_key||raw.name.toLowerCase()!==fileName.toLowerCase())return {status:'repair-required',download:false};
 const bytes=await crypto.subtle.digest('SHA-256',new TextEncoder().encode([run.sourceId,run.providerVersion,raw.file_id,current].join('|')));
 const key=[...new Uint8Array(bytes)].map(b=>b.toString(16).padStart(2,'0')).join('');
 const id=`authority:${key}`,fileId=`${raw.file_id}:authority:${key.slice(0,16)}`;
 const priorRecovery=await db.prepare('SELECT status FROM roster_sync_runs WHERE id=?').bind(id).first();
 if(priorRecovery&&['failed','superseded'].includes(priorRecovery.status))return {status:'repair-required',download:false};
 if(!await reserveRosterMaintenanceBudget(db,64,512))return {status:'deferred',download:false};
 await db.batch([
  db.prepare(`INSERT OR IGNORE INTO roster_sync_runs(id,source_id,trigger_type,provider_version,content_hash,file_id,source_file_id,status,message,started_at)
   VALUES(?,?,'authority-reconcile',?,?,?,?,'queued','Fresh provider confirmation after manual activation; reusing retained bytes.',?)`).bind(id,run.sourceId,run.providerVersion,run.contentHash,fileId,raw.file_id,new Date().toISOString()),
  authorityFenceStatement(db,id,source,current),
  db.prepare("UPDATE roster_sync_runs SET status='superseded',message='Fresh provider reconciliation owns this retained input.' WHERE id=? AND id<>? AND status IN ('queued','processing')").bind(run.id,id),
 ]);
 return {status:'resuming',download:false};
}
export async function activeAuthorityFile(db,fileId) {
 return Boolean((await db.prepare('SELECT active FROM roster_files WHERE id=?').bind(fileId).first())?.active);
}
