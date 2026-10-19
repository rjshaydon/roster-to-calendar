import assert from 'node:assert/strict';
import {beginMaintenanceAccounting} from '../functions/_lib/roster-maintenance-budget.js';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {planRosterImportBatches} from '../functions/_lib/roster-import-batches.js';
import {executeBoundedRosterImport} from './roster-import-driver.mjs';
import {handleManualRosterImport} from '../functions/_lib/manual-roster-management.js';
import {beginBoundedRosterImport,stageBoundedRosterBatch,prepareBoundedRosterPresence,prepareBoundedRosterMetadata,activateBoundedRosterTerm} from '../functions/_lib/roster-import-staging.js';
import {authorityFenceStatement,authoritySnapshot,observeProviderAuthority,providerAuthorityFence,restoreRetainedAuthority,reconcileQueuedAuthority} from '../functions/_lib/roster-authority.js';
import {createRosterSyncRun,upsertRawRosterFile,upsertRosterSource} from '../functions/_lib/d1-calendar.js';
import {onRequestPost as metadataCheck} from '../functions/api/automation/roster-check.js';
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
        let result;try{result=owner.sqlite.prepare(sql).run(...this.args);}catch(error){error.message += ` SQL: ${sql} Args: ${JSON.stringify(this.args)}`;throw error;}
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

const sqlite=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
for(const [column,definition] of [['name',"TEXT NOT NULL DEFAULT ''"],['source_type',"TEXT NOT NULL DEFAULT ''"],['size','INTEGER NOT NULL DEFAULT 0'],['last_modified','INTEGER NOT NULL DEFAULT 0']])sqlite.exec(`ALTER TABLE raw_roster_files ADD COLUMN ${column} ${definition}`);
const db=new LocalD1(sqlite);
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)').run(new Date().toISOString().slice(0,10),50000000,5000000,new Date(Date.now()+3600000).toISOString());
beginMaintenanceAccounting(db,'authority-fixture',true);
const env={ROSTER_DB:db,ROSTER_AUTOMATION_TOKEN:'fixture',ROSTER_METADATA_CHECK_ENABLED:'true',ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_SOURCE_ALLOWLIST:'monash-adults',ROSTER_AUTOMATION_ENABLED_SOURCES:'monash-adults',ROSTER_GITHUB_AUTO_PROCESS_ENABLED:'false'};
await upsertRosterSource(db,{id:'monash-adults',sourceType:'mmc',provider:'sharepoint',enabled:true});
const active=id=>sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(id)?.active;
let generation=0;
async function fixture({manual=false,from='2026-08-03',to='2026-11-01',label='Automatic',version='1'}={}){
 const file={id:(manual?'manual:':'automation:')+crypto.randomUUID(),sourceId:manual?'manual-mmc':'monash-adults',sourceType:'mmc',name:from==='2026-11-02'?'AdultTerm4.2026.xlsx':'AdultTerm3.2026.xlsx',lastModified:1000,size:10,parserVersion:'fixture'};
 const doctor={key:'EXAMPLE',displayName:'Example Doctor',seniority:'Intern'};
 const events=[from,to].map((day,i)=>({id:file.id+':'+i,source:'MMC',title:label,rawValue:label,seniority:'Intern',start:day+'T08:00:00',end:day+'T16:00:00'}));
 const plan=await planRosterImportBatches({file,contentHash:'a'.repeat(64),doctors:[doctor],eventsByDoctor:{EXAMPLE:events},issuesByDoctor:{}});
 const runId=manual?'manual:'+file.id:'run:'+crypto.randomUUID();
 if(!manual){
  await upsertRawRosterFile(db,file,{objectKey:'retained/'+file.id,uploadedAt:new Date().toISOString()});
  await createRosterSyncRun(db,{id:runId,sourceId:file.sourceId,triggerType:'sharepoint',providerVersion:version,contentHash:'a'.repeat(64),fileId:file.id,status:'processing',startedAt:new Date().toISOString()});
  await db.batch([authorityFenceStatement(db,runId,'mmc',await authoritySnapshot(db,'mmc'))]);
 }
 const run={id:runId,fileId:file.id,sourceFileId:file.id,sourceId:file.sourceId,contentHash:'a'.repeat(64),providerVersion:version,startedAt:new Date().toISOString(),completedAt:'2099-01-01T00:00:00Z',status:'success'};
 const call=body=>manual?handleManualRosterImport({env}, {file,...body},'owner@example.test'):(async()=>{
  switch(body.phase){
   case 'bounded-begin':return beginBoundedRosterImport(db,runId,file,{manifest:body.manifest,revision:body.revision},{allowReplacement:true});
   case 'bounded-events':return stageBoundedRosterBatch(db,runId,file,body.revision,body.batch);
   case 'bounded-presence':return prepareBoundedRosterPresence(db,runId,file,body.revision,body.batch);
   case 'bounded-metadata':return prepareBoundedRosterMetadata(db,runId,body.revision);
   case 'bounded-activate':return activateBoundedRosterTerm(db,runId,body.revision);
  }
 })();
 return {file,plan,run,call};
}
async function importAll(f){return executeBoundedRosterImport(f.plan,f.call);}
const original=await fixture();await importAll(original);
sqlite.prepare("UPDATE roster_sync_runs SET status='success',completed_at=? WHERE id=?").run(original.run.completedAt,original.run.id);
const manual=await fixture({manual:true,label:'Manual correction'});await importAll(manual);
assert.equal(active(original.file.id),0);assert.equal(active(manual.file.id),1);
assert.equal(sqlite.prepare('SELECT generation FROM roster_term_authority').get().generation,1);
const beforeEvents=sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events').get().n;
const restored=await restoreRetainedAuthority(db,original.run);
assert.equal(restored.status,'authority-restored');assert.equal(active(original.file.id),1);assert.equal(active(manual.file.id),0);
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events').get().n,beforeEvents,'handoff never rewrites clinical facts');
const writes=db.rowsWritten;
assert.equal(await restoreRetainedAuthority(db,original.run),null);assert.equal(db.rowsWritten,writes,'already-authoritative path is zero-write');
// A next-term manual draft is not displaced by current-term reconciliation.
const future=await fixture({manual:true,from:'2026-11-02',to:'2027-01-31'});await importAll(future);
assert.equal(await restoreRetainedAuthority(db,original.run),null);assert.equal(active(future.file.id),1);
// Old automatic processing cannot overwrite a manual activation after it began.
const delayed=await fixture({version:'2'});
await executeBoundedRosterImport(delayed.plan,async body=>body.phase==='bounded-activate'?{deferred:true}:delayed.call(body));
const newerManual=await fixture({manual:true,label:'New manual'});await importAll(newerManual);
await assert.rejects(delayed.call({phase:'bounded-activate',revision:delayed.plan.revision}),/stale-roster-authority|Active rosters changed/);
assert.equal(active(newerManual.file.id),1);
const resumed=await reconcileQueuedAuthority(db,delayed.run,delayed.file.name);
assert.equal(resumed.status,'resuming');
const count=sqlite.prepare("SELECT COUNT(*) AS n FROM roster_sync_runs WHERE trigger_type='authority-reconcile'").get().n;
await reconcileQueuedAuthority(db,delayed.run,delayed.file.name);
assert.equal(sqlite.prepare("SELECT COUNT(*) AS n FROM roster_sync_runs WHERE trigger_type='authority-reconcile'").get().n,count,'fresh confirmation coalesces by manual generation');
const changed=await fixture({version:'3',label:'Official correction'});await importAll(changed);
assert.equal(active(changed.file.id),1);assert.equal(active(newerManual.file.id),0);assert.equal(active(future.file.id),1);
// Confirmation is pinned before retrieval, not to whichever authority exists later.
await observeProviderAuthority(db,'monash-adults',changed.file.name,'4');
const observed=await providerAuthorityFence(db,'monash-adults',changed.file.name,'4');
const intervening=await fixture({manual:true,label:'Later manual'});await importAll(intervening);
assert.notEqual(observed,await authoritySnapshot(db,'mmc'));
// Exact raw revision guard: a later in-place patch is not the earlier receipt.
const patched={...changed.run,completedAt:'2000-01-01T00:00:00Z'};
assert.equal((await restoreRetainedAuthority(db,patched)).status,'repair-required');assert.equal(active(intervening.file.id),1);
// A transaction failure must preserve the manual file and its generation.
const currentGeneration=await authoritySnapshot(db,'mmc');
const candidate=await fixture({version:'5'});await importAll(candidate);
const latest=await fixture({manual:true,label:'Latest manual'});await importAll(latest);
sqlite.prepare("UPDATE roster_sync_runs SET status='success',completed_at=? WHERE id=?").run(candidate.run.completedAt,candidate.run.id);
db.failRunIncludes='UPDATE roster_files SET active=1 WHERE id=?';
await assert.rejects(restoreRetainedAuthority(db,candidate.run),/Injected/);assert.equal(active(latest.file.id),1);assert.equal(active(candidate.file.id),0);
// Missing/unavailable inventory does not restore anything or download.
const response=await metadataCheck({env,request:new Request('https://fixture/api/automation/roster-check',{method:'POST',headers:{authorization:'Bearer fixture','content-type':'application/json'},body:JSON.stringify({mode:'reconcile',libraries:[{site:'monash',unavailable:true,files:[]},{site:'vhh',files:[]}]})})},new Date('2026-10-19T12:00:00Z'));
assert.equal(response.status,200);assert.deepEqual((await response.json()).downloads,[]);assert.equal(active(latest.file.id),1);
// Exercise the real unchanged-provider HTTP path: restore once, then zero writes.
const metadataRequest=()=>({env,request:new Request('https://fixture/api/automation/roster-check',{method:'POST',headers:{authorization:'Bearer fixture','content-type':'application/json'},body:JSON.stringify({sourceId:'monash-adults',fileName:candidate.file.name,providerVersion:'5'})})});
const restoredHttp=await metadataCheck(metadataRequest(),new Date('2026-10-19T12:00:00Z'));
assert.equal(restoredHttp.status,200);assert.equal((await restoredHttp.json()).status,'authority-restored');
const unchangedWrites=db.rowsWritten;
const unchangedHttp=await metadataCheck(metadataRequest(),new Date('2026-10-19T12:00:00Z'));
assert.equal((await unchangedHttp.json()).status,'unchanged');assert.equal(db.rowsWritten,unchangedWrites);
console.log('Authority: automatic/manual handoff, retained zero-event-write restore, repeat no-op, next-term isolation, stale-run rejection, fresh confirmation/coalescing, exact snapshot validation and transactional rollback passed.');
