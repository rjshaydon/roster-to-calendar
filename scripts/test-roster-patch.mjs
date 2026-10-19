import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile,readdir } from 'node:fs/promises';
import { buildAutomatedDerivedRosterPayload } from '../functions/_lib/automation-import.js';
import { replaceDerivedRosterFile,refreshFacilityOverviewMaterializationForFile } from '../functions/_lib/d1-calendar.js';
import { beginRosterPatch,readRosterPatchPage,pinRosterPatchPlan,stageRosterPatchBatch,activateRosterPatch } from '../functions/_lib/roster-patch.js';
import { planRosterPatch,executeRosterPatch } from './roster-patch-plan.mjs';
import { beginMaintenanceAccounting } from '../functions/_lib/roster-maintenance-budget.js';

class D1 {
 constructor(sqlite){this.sqlite=sqlite;this.sql=[];this.fail='';}
 prepare(sql){const owner=this;owner.sql.push(sql);return {args:[],bind(...args){this.args=args;return this;},
  async first(){return owner.sqlite.prepare(sql).get(...this.args)||null;},
  async all(){return {results:owner.sqlite.prepare(sql).all(...this.args)};},
  async run(){if(owner.fail && sql.includes(owner.fail)){owner.fail='';throw Error('Injected patch failure');}const r=owner.sqlite.prepare(sql).run(...this.args);return {meta:{changes:Number(r.changes)}};}};}
 async batch(statements){this.sqlite.exec('SAVEPOINT patch_test');try{const r=[];for(const s of statements)r.push(await s.run());this.sqlite.exec('RELEASE patch_test');return r;}catch(e){this.sqlite.exec('ROLLBACK TO patch_test');this.sqlite.exec('RELEASE patch_test');throw e;}}
}
for(const [sourceId,path] of [['monash-adults','AdultTerm1.2026.xlsx'],['monash-paeds','Paeds_Term_2_2026.xlsx']]) {
 const sqlite=new DatabaseSync(':memory:');
 for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
 const db=new D1(sqlite),bytes=await readFile(new URL('../fixtures/'+path,import.meta.url));
 const payload=await buildAutomatedDerivedRosterPayload({file:new File([bytes],path,{lastModified:1}),sourceId,contentHash:'fixture',fileId:'retained:'+sourceId});
 await replaceDerivedRosterFile(db,{...payload.file,active:true},payload.doctors,payload.eventsByDoctor,payload.issuesByDoctor);
 const day=new Date().toISOString().slice(0,10);
 sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)').run(day,50000000,5000000,new Date(Date.now()+3600000).toISOString());
 let callId=0;const call=fn=>{beginMaintenanceAccounting(db,'patch-'+(++callId),true);return fn();};
 const run={id:'patch-run',sourceId}, baseline=await call(()=>beginRosterPatch(db,run,payload.file.id));
 const stored={};
 for(const section of ['doctors','events','issues','staff','catalog']) {
  stored[section]=[];let after='';
  do {const page=await call(()=>readRosterPatchPage(db,run,{section,after}));assert.ok(page.rows.length<=32);stored[section].push(...page.rows);after=page.next;}while(after);
 }
 assert.equal(stored.events.length,payload.eventCount);
 const eventQuery=db.sql.find(sql=>sql.includes('AS patch_cursor FROM roster_events'));
 const readPlan=sqlite.prepare('EXPLAIN QUERY PLAN '+eventQuery).all(payload.file.id,0).map(r=>r.detail).join('\n');
 assert.match(readPlan,/idx_roster_events_file/);assert.doesNotMatch(readPlan,/SCAN roster_events|TEMP B-TREE/);
 const noChange=await planRosterPatch(payload,baseline,stored);
 assert.equal(noChange.manifest.changedFacts,0);assert.equal(noChange.batches.length,0);
 const desired=structuredClone(payload),key=Object.keys(desired.eventsByDoctor).find(k=>desired.eventsByDoctor[k].length>=3);
 desired.eventsByDoctor[key].shift();desired.eventsByDoctor[key][0].title+=' corrected';
 desired.eventsByDoctor[key].push({...desired.eventsByDoctor[key][1],id:'new-occurrence'});
 desired.doctors.find(d=>d.key===key).seniority='Senior Registrar';
 const plan=await planRosterPatch(desired,baseline,stored),retry=await planRosterPatch(desired,baseline,stored);
 const cachedPlan=await planRosterPatch(desired,baseline,stored,{presenceMode:'cached'});
 assert.equal(cachedPlan.manifest.changedFacts,plan.manifest.changedFacts);
 assert.ok(cachedPlan.batches.every(b=>b.entries.every(e=>e.kind!=='presence')));
 assert.deepEqual(cachedPlan.batches.flatMap(b=>b.entries),plan.batches.flatMap(b=>b.entries).filter(e=>e.kind!=='presence'),'cached mode preserves every clinical and compact fact');
 assert.equal(plan.patchRevision,retry.patchRevision,'a resumed planner must produce identical hashes');
 assert.ok(plan.manifest.changedFacts<10);
 await call(()=>pinRosterPatchPlan(db,run,{manifest:plan.manifest,patchRevision:plan.patchRevision}));
 const bad=structuredClone(plan.batches[0]);bad.entries[0].row[0]='other-file';
 await assert.rejects(call(()=>stageRosterPatchBatch(db,run,{batch:bad,patchRevision:plan.patchRevision})),/pinned plan/);
 for(const batch of plan.batches) await call(()=>stageRosterPatchBatch(db,run,{batch,patchRevision:plan.patchRevision}));
 assert.equal((await call(()=>stageRosterPatchBatch(db,run,{batch:plan.batches[0],patchRevision:plan.patchRevision}))).duplicate,true);
 const previous=sqlite.prepare('SELECT id,event_json FROM roster_events ORDER BY id').all();
 const activate=()=>call(()=>activateRosterPatch(db,run,{patchRevision:plan.patchRevision},{file:desired.file,publishFacility:true}));
 db.fail='UPDATE roster_file_coverage';
 await assert.rejects(activate(),/Injected patch failure/);
 assert.deepEqual(sqlite.prepare('SELECT id,event_json FROM roster_events ORDER BY id').all(),previous,'activation failure preserves all previous events');
 assert.equal(sqlite.prepare('SELECT activated FROM roster_patch_jobs').get().activated,0);
 const activated=await activate();assert.equal(activated.fileId,payload.file.id);
 assert.equal((await activate()).activated,true,'lost activation response is safe to replay');
 assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_events').get().n,payload.eventCount);
 const comparison=await refreshFacilityOverviewMaterializationForFile(db,payload.file.id,{dryRun:true,contentRevision:plan.manifest.contentRevision});
 assert.equal(comparison.proposedWrites,0,'patch compact facts must match the original historical-grade implementation');
 const after=sqlite.prepare('SELECT id,event_json FROM roster_events ORDER BY id').all();
 assert.equal(after.filter(r=>previous.some(p=>p.id===r.id && p.event_json!==r.event_json)).length,1,'unchanged roster events must not be rewritten');
 const removed=previous.find(p=>!after.some(r=>r.id===p.id));
 assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_daily_presence WHERE event_id=?').get(removed.id).n,0);
 const restarted=await call(()=>beginRosterPatch(db,run,payload.file.id));assert.equal(restarted.activated,true);
 const resume=await executeRosterPatch(desired,restarted,body=>call(()=>activateRosterPatch(db,run,body,{file:desired.file,publishFacility:true})));
 assert.equal(resume.activated,true,'a retry after activation must not inspect the stale snapshot');
 const race={id:'race-run',sourceId};await call(()=>beginRosterPatch(db,race,payload.file.id));
 sqlite.prepare("UPDATE roster_files SET parsed_at='changed' WHERE id=?").run(payload.file.id);
 await assert.rejects(call(()=>readRosterPatchPage(db,race,{section:'events'})),/malformed JSON/,'a changed baseline cannot be inspected or applied');
 console.log(`${sourceId}: ${payload.eventCount} events; indexed pages, bounded diff, stable resume, atomic rollback, presence cleanup and compact-grade parity passed.`);
}
