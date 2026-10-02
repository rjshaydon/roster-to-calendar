import { rosterDeliveryOrder } from '../functions/_lib/roster-delivery-order.js';
import { listQueuedRosterSyncRuns, claimRosterDispatch } from '../functions/_lib/d1-calendar.js';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { buildAutomatedDerivedRosterPayload } from '../functions/_lib/automation-import.js';
import { planRosterImportBatches, rosterImportDigest } from '../public/static/roster-import-batches.js';
import { beginBoundedRosterImport, stageBoundedRosterBatch, prepareBoundedRosterPresence, prepareBoundedRosterMetadata, activateBoundedRosterTerm } from '../functions/_lib/roster-import-staging.js';
import { onRequestPost as derivedHandler } from '../functions/api/automation/derived.js';
import { beginMaintenanceAccounting } from '../functions/_lib/roster-maintenance-budget.js';
class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; this.sql = []; this.failRunIncludes = ""; }
  prepare(sql) {
    const owner = this;
    owner.sql.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() {
        if (owner.failRunIncludes && sql.includes(owner.failRunIncludes)) {
          owner.failRunIncludes = "";
          throw new Error("Injected D1 statement failure");
        }
        const result = owner.sqlite.prepare(sql).run(...this.args);
        owner.rowsWritten += Number(result.changes || 0);
        return { success: true, meta: { changes: Number(result.changes || 0), rows_read: Number(result.changes || 0) + 1, rows_written: Number(result.changes || 0) } };
      },
      // Metadata is simulated here to test reservation/refund logic, not to
      // claim SQLite result counts measure Cloudflare index billing.
      async all() { const results = owner.sqlite.prepare(sql).all(...this.args); return { success: true, results, meta: { rows_read: results.length + 1, rows_written: 0 } }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) {
    if (this.beforeBatch) { const fn = this.beforeBatch; this.beforeBatch = null; fn(statements); }
    const results = [];
    const ownsTransaction = !this.sqlite.isTransaction;
    if (ownsTransaction) this.sqlite.exec("BEGIN");
    try {
      for (const statement of statements) results.push(await statement.run());
      if (ownsTransaction) this.sqlite.exec("COMMIT");
      return results;
    } catch (error) {
      if (ownsTransaction) this.sqlite.exec("ROLLBACK");
      throw error;
    }
  }
}

const wire = value => JSON.parse(JSON.stringify(value));
for(const value of [{optional:undefined,keep:null},[undefined,,NaN],{nested:{missing:undefined},list:[undefined]}, {when:new Date('2026-10-02T00:00:00Z')}, {missing:{toJSON(){return undefined}},keep:null}]) {
  assert.equal(await rosterImportDigest(value),await rosterImportDigest(wire(value)),'hash must match the transmitted JSON');
}
function legacyCanonical(value) {
  if(Array.isArray(value))return `[${value.map(legacyCanonical).join(',')}]`;
  if(value&&typeof value==='object')return `{${Object.keys(value).sort().map(key=>`${JSON.stringify(key)}:${legacyCanonical(value[key])}`).join(',')}}`;
  return JSON.stringify(value);
}
async function legacyDigest(value) {
  return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(legacyCanonical(value))))].map(byte=>byte.toString(16).padStart(2,'0')).join('');
}
const sqlite=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(name=>name.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL(`../migrations/${name}`,import.meta.url),'utf8'));
// Production raw-file metadata predates explicit migrations; model its verified columns.
for(const [column,definition] of [['name',"TEXT NOT NULL DEFAULT ''"],['source_type',"TEXT NOT NULL DEFAULT ''"],['size','INTEGER NOT NULL DEFAULT 0'],['last_modified','INTEGER NOT NULL DEFAULT 0']])sqlite.exec(`ALTER TABLE raw_roster_files ADD COLUMN ${column} ${definition}`);
const db=new LocalD1(sqlite);
const day=new Date().toISOString().slice(0,10);
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)').run(day,50000000,5000000,new Date(Date.now()+3600000).toISOString());
let request=0;
async function call(fn){beginMaintenanceAccounting(db,'wire-fixture-'+(++request),true);return fn();}
for(const [sourceId,path] of [['monash-adults','fixtures/AdultTerm1.2026.xlsx'],['monash-paeds','fixtures/Paeds_Term_2_2026.xlsx']]) {
  const fileBytes=await readFile(new URL('../'+path,import.meta.url));
  const payload=await buildAutomatedDerivedRosterPayload({file:new File([fileBytes],path.split('/').at(-1),{lastModified:1}),sourceId,contentHash:'fixture',fileId:'wire:'+sourceId});
  const plan=await planRosterImportBatches(payload);
  assert.equal(plan.revision,await rosterImportDigest(wire(plan.manifest)));
  for(const batch of plan.batches)assert.equal(plan.manifest.batches[batch.index].hash,await rosterImportDigest(wire(batch)));
  const run='wire-run:'+sourceId, file=payload.file;
  const legacyManifest={...plan.manifest,batches:await Promise.all(plan.batches.map(async batch=>({...plan.manifest.batches[batch.index],hash:await legacyDigest(batch)})))};
  const legacyRevision=await legacyDigest(legacyManifest);
  await call(()=>beginBoundedRosterImport(db,run,file,{manifest:wire(legacyManifest),revision:legacyRevision},{allowReplacement:true}));
  const initialFence=sqlite.prepare('SELECT promotion_fence_json FROM roster_import_jobs WHERE run_id=?').get(run).promotion_fence_json;
  const before=db.rowsWritten;
  await assert.rejects(call(()=>stageBoundedRosterBatch(db,run,file,legacyRevision,wire(plan.batches[0]))),/differs from the pinned plan/);
  assert.equal(db.rowsWritten,before,'wire hash rejection must precede writes');
  assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(file.id).n,0);
  await assert.rejects(call(()=>beginBoundedRosterImport(db,run,file,plan,{allowReplacement:true})),/new run is required/,'recovery is not available to ordinary manual requests');
  const alteredManifest={...plan.manifest,doctors:plan.manifest.doctors.map((doctor,index)=>index?doctor:{...doctor,displayName:'Changed Identity'})};
  const alteredRevision=await rosterImportDigest(alteredManifest);
  await assert.rejects(call(()=>beginBoundedRosterImport(db,run,file,{manifest:alteredManifest,revision:alteredRevision},{allowReplacement:true,allowEmptyPlanRecovery:true})),/new run is required/);
  sqlite.prepare('UPDATE roster_import_jobs SET next_batch=1 WHERE run_id=?').run(run);
  await assert.rejects(call(()=>beginBoundedRosterImport(db,run,file,plan,{allowReplacement:true,allowEmptyPlanRecovery:true})),/new run is required/);
  sqlite.prepare('UPDATE roster_import_jobs SET next_batch=0 WHERE run_id=?').run(run);
  const recovered=await call(()=>beginBoundedRosterImport(db,run,file,wire(plan),{allowReplacement:true,allowEmptyPlanRecovery:true}));
  assert.equal(recovered.nextBatch,0);
  assert.equal(sqlite.prepare('SELECT promotion_fence_json FROM roster_import_jobs WHERE run_id=?').get(run).promotion_fence_json,initialFence);
  // A plan/cursor change after preflight must roll back the entire fact batch.
  db.beforeBatch=()=>{db.beforeBatch=()=>sqlite.prepare("UPDATE roster_import_jobs SET plan_revision='concurrent' WHERE run_id=?").run(run)};
  await assert.rejects(call(()=>stageBoundedRosterBatch(db,run,file,plan.revision,wire(plan.batches[0]))),/malformed JSON/);
  assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(file.id).n,0);
  sqlite.prepare('UPDATE roster_import_jobs SET plan_revision=? WHERE run_id=?').run(plan.revision,run);
  for(const batch of plan.batches) await call(()=>stageBoundedRosterBatch(db,run,file,plan.revision,wire(batch)));
  const staged=db.rowsWritten;
  assert.equal((await call(()=>stageBoundedRosterBatch(db,run,file,plan.revision,wire(plan.batches[0])))).duplicate,true);
  assert.equal(db.rowsWritten,staged);
  for(const batch of plan.batches)await call(()=>prepareBoundedRosterPresence(db,run,file,plan.revision,wire(batch)));
  await call(()=>prepareBoundedRosterMetadata(db,run,plan.revision));
  await call(()=>activateBoundedRosterTerm(db,run,plan.revision));
  assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(file.id).active,1);
  assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(file.id).n,payload.eventCount);
  sqlite.prepare('INSERT INTO roster_sources(id,source_type,enabled,active_file_id) VALUES(?,?,1,?)').run(sourceId,file.sourceType,file.id);
  const env={ROSTER_DB:db,ROSTER_AUTOMATION_TOKEN:'fixture-token',ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_QUEUE_ENABLED:'true',ROSTER_AUTOMATION_SOURCE_ALLOWLIST:sourceId,ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT:'1250',ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED:'true',ROSTER_BOUNDED_REPLACEMENT_ENABLED:'true'};
  async function api(body){
    beginMaintenanceAccounting(db,'wire-api-'+(++request),true);
    const response=await derivedHandler({env,waitUntil(){},request:new Request('https://fixture.test/api/automation/derived',{method:'POST',headers:{'content-type':'application/json',authorization:'Bearer fixture-token'},body:JSON.stringify(body)})});
    return {status:response.status,data:await response.json()};
  }
  async function correction(name,changed){
    const incoming={...file,id:'incoming:'+sourceId+':'+name};
    const events=wire(payload.eventsByDoctor);
    let remaining=changed;
    for(const rows of Object.values(events))for(const row of rows)if(remaining-->0)row.title+=' corrected';
    const desired={...payload,file:incoming,eventsByDoctor:events};
    const next=await planRosterImportBatches(desired);
    const runId='correction:'+sourceId+':'+name;
    sqlite.prepare('INSERT INTO roster_sync_runs(id,source_id,file_id,content_hash,status) VALUES(?,?,?,?,?)').run(runId,sourceId,incoming.id,'fixture','queued');
    const common={sourceId,runId,file:incoming,revision:next.revision};
    const probe=await api({...common,phase:'bounded-begin',manifest:next.manifest});
    assert.equal(probe.status,200,JSON.stringify(probe));
    assert.equal(probe.data.mode,'complete','same-term corrections try the bounded diff first');
    assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_import_jobs WHERE run_id=?').get(runId).n,0);
    const result=await api({...common,phase:'complete',doctors:desired.doctors,eventsByDoctor:desired.eventsByDoctor,issuesByDoctor:desired.issuesByDoctor});
    return {incoming,next,runId,common,result};
  }
  const small=await correction('small',1);
  assert.equal(small.result.status,200,JSON.stringify(small.result));
  assert.equal(small.result.data.fileId,file.id,'small corrections preserve the stable active file');
  assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(file.id).n,payload.eventCount);
  const beforeLarge=sqlite.prepare('SELECT id,event_json FROM roster_events WHERE file_id=? ORDER BY id').all(file.id);
  const large=await correction('large',1300);
  assert.equal(large.result.status,422,JSON.stringify(large.result));
  assert.equal(large.result.data.code,'ROSTER_INCREMENTAL_BUDGET');
  assert.equal(large.result.data.boundedReplacementRequired,true);
  assert.equal(sqlite.prepare('SELECT status FROM roster_sync_runs WHERE id=?').get(large.runId).status,'queued');
  assert.deepEqual(sqlite.prepare('SELECT id,event_json FROM roster_events WHERE file_id=? ORDER BY id').all(file.id),beforeLarge,'oversized diff cannot mutate active facts');
  const fallback=await api({...large.common,phase:'bounded-begin',forceStaging:true,manifest:large.next.manifest});
  assert.equal(fallback.status,200,JSON.stringify(fallback));
  assert.equal(fallback.data.mode,'bounded');
  assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(file.id).active,1,'forced staging retains working roster');
  assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(large.incoming.id).active,0);
  const orderedOld={...file,id:'ordered-old:'+sourceId};
  const orderedNew={...file,id:'ordered-new:'+sourceId};
  function deliveryRow(id,raw,modified,started,status='queued') {
    sqlite.prepare('INSERT INTO raw_roster_files(file_id,name,last_modified) VALUES(?,?,?)').run(raw.id,raw.name,modified);
    sqlite.prepare('INSERT INTO roster_sync_runs(id,source_id,file_id,source_file_id,content_hash,status,started_at) VALUES(?,?,?,?,?,?,?)').run(id,sourceId,raw.id,raw.id,'fixture',status,started);
    return {id,sourceId,fileId:raw.id,sourceFileId:raw.id,status,startedAt:started};
  }
  const oldDelivery=deliveryRow('ordered-old-run:'+sourceId,orderedOld,100,'2026-10-01T00:00:00Z');
  const newDelivery=deliveryRow('ordered-new-run:'+sourceId,orderedNew,200,'2026-10-02T00:00:00Z','success');
  const activeBefore=sqlite.prepare('SELECT id,event_json FROM roster_events WHERE file_id=? ORDER BY id').all(file.id);
  assert.equal((await rosterDeliveryOrder(db,oldDelivery)).superseded,true);
  assert.equal((await listQueuedRosterSyncRuns(db,sourceId,1)).length,0,'new successful delivery excludes older queued versions');
  assert.equal((await claimRosterDispatch(db,{sourceId})).reason,'queue-empty','obsolete backlog cannot trigger repeated GitHub dispatch');
  const oldPlan=await planRosterImportBatches({...payload,file:orderedOld});
  const skipped=await api({sourceId,runId:oldDelivery.id,file:orderedOld,phase:'bounded-begin',manifest:oldPlan.manifest,revision:oldPlan.revision});
  assert.equal(skipped.status,200,JSON.stringify(skipped));
  assert.equal(skipped.data.superseded,true);
  assert.equal(sqlite.prepare('SELECT status FROM roster_sync_runs WHERE id=?').get(oldDelivery.id).status,'superseded');
  assert.deepEqual(sqlite.prepare('SELECT id,event_json FROM roster_events WHERE file_id=? ORDER BY id').all(file.id),activeBefore);
  const late=deliveryRow('late-old:'+sourceId,{...orderedOld,id:'late-old-raw:'+sourceId},50,'2026-10-03T00:00:00Z');
  assert.equal((await rosterDeliveryOrder(db,late)).superseded,true,'late arrival cannot bypass provider timestamps');
  const differentTerm=deliveryRow('different-term:'+sourceId,{...orderedOld,id:'different-term-raw:'+sourceId,name:'Other term.xlsx'},10,'2026-10-03T00:00:01Z');
  assert.ok((await rosterDeliveryOrder(db,differentTerm)).guard,'a distinct workbook term remains independent');
  const fence=await rosterDeliveryOrder(db,newDelivery);
  const concurrentNewest=deliveryRow('concurrent-new:'+sourceId,{...orderedNew,id:'concurrent-new-raw:'+sourceId},300,'2026-10-04T00:00:00Z');
  await assert.rejects(db.batch([fence.guard(),db.prepare('UPDATE roster_files SET active=0 WHERE id=?').bind(file.id)]),/malformed JSON/);
  assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(file.id).active,1,'transaction guard prevents stale writes after preflight');
  sqlite.prepare(`WITH RECURSIVE n(value) AS (SELECT 1 UNION ALL SELECT value+1 FROM n WHERE value<100001)
    INSERT INTO roster_sync_runs(id,source_id,status,started_at) SELECT ?||':'||value,?,'failed','1900-01-01T00:00:00Z' FROM n`).run('ordering-history:'+sourceId,sourceId);
  assert.ok((await rosterDeliveryOrder(db,concurrentNewest)).guard,'large retained run history does not prevent the latest delivery');
  const unprovable=deliveryRow('unprovable-old:'+sourceId,{...orderedOld,id:'unprovable-raw:'+sourceId,name:'Ancient isolated term.xlsx'},1,'1800-01-01T00:00:00Z');
  await assert.rejects(rosterDeliveryOrder(db,unprovable),/64-run safety window/,'old work cannot fall back to scanning history');
  const orderingSql=db.sql.find(sql=>sql.includes('SELECT r.*,f.name,f.last_modified'));
  const orderingPlan=sqlite.prepare('EXPLAIN QUERY PLAN '+orderingSql).all(sourceId).map(row=>row.detail).join('\n');
  assert.ok(orderingPlan.includes('idx_roster_sync_runs_source_started_id'));
  assert.ok(!/SCAN roster_sync_runs|SCAN f/.test(orderingPlan),orderingPlan);
  console.log('Delivery ordering: successor/late-arrival exclusion, independent terms, dispatch, atomic races and 100,001 historical runs passed.');
  console.log(`${sourceId}: ${payload.eventCount} real workbook events survived JSON transport, guarded empty-plan recovery, staging and activation.`);
}
console.log('Wire protocol: optional/sparse/null/date values, hash rejection before writes, plan recovery bounds, stale-transaction rollback and replay passed.');
