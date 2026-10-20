import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {buildVhhDerivedRosterPayload,VHH_ROSTER_SOURCE_ID} from '../functions/_lib/vhh-roster.js';
import {planVhhHistoryWindows} from './vhh-history-plan.mjs';
import {extractVhhHistoryWindow} from './vhh-history-extract.mjs';
import {prepareVhhHistoryPromotion,requirePreservedPrefix,historyJobId,historyFileName} from '../functions/_lib/vhh-history-preservation.js';
import {loadPublishedFacilityDays,rosterCoverageContains} from '../functions/_lib/facility-overview-cache.js';
import {beginBoundedRosterImport,stageBoundedRosterBatch,prepareBoundedRosterPresence,prepareBoundedRosterMetadata,activateBoundedRosterTerm} from '../functions/_lib/roster-import-staging.js';
import {beginRosterPatch,readRosterPatchPage,pinRosterPatchPlan,stageRosterPatchBatch,activateRosterPatch} from '../functions/_lib/roster-patch.js';
import {planRosterImportBatches,rosterImportDigest} from '../functions/_lib/roster-import-batches.js';
import {attachRosterCompactPlan} from './roster-compact-plan.mjs';
import {executeRosterPatch} from './roster-patch-plan.mjs';
import {beginMaintenanceAccounting} from '../functions/_lib/roster-maintenance-budget.js';
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
        const prepared = owner.sqlite.prepare(sql);
        if (/^\s*SELECT\b/i.test(sql) || /^\s*WITH\b/i.test(sql) && !/\b(INSERT|UPDATE|DELETE)\b/i.test(sql)) {
          const results = prepared.all(...this.args);
          return {success:true,results,meta:{changes:0,rows_read:results.length+1,rows_written:0}};
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

const sqlite=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(name=>name.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL(`../migrations/${name}`,import.meta.url),'utf8'));
// Production raw-file metadata predates explicit migrations; model its verified columns.
for(const [column,definition] of [['name',"TEXT NOT NULL DEFAULT ''"],['source_type',"TEXT NOT NULL DEFAULT ''"],['size','INTEGER NOT NULL DEFAULT 0'],['last_modified','INTEGER NOT NULL DEFAULT 0']])sqlite.exec(`ALTER TABLE raw_roster_files ADD COLUMN ${column} ${definition}`);
const db=new LocalD1(sqlite);
const day=new Date().toISOString().slice(0,10);
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)').run(day,50000000,5000000,new Date(Date.now()+3600000).toISOString());
let request=0;
async function call(fn){beginMaintenanceAccounting(db,'wire-fixture-'+(++request),true);return fn();}
const dates=(from,to)=>{const result=[];for(let d=from;d<=to;d=new Date(Date.parse(d+'T12:00:00Z')+86400000).toISOString().slice(0,10))result.push(d);return result;};
const block=(from,to,visible=true)=>({sheetName:from,blockIndex:1,visible,dates:dates(from,to).map(date=>({date,sourceColumn:'B'})),rows:[{sourceRow:2,shiftLabel:'AM REG',sourceShiftLabel:'AM REG',assignments:dates(from,to).map((date,i)=>({date,namesText:'Example, Doctor',sourceCell:String.fromCharCode(66+i)+'2'}))}]});
const extract={sourceId:VHH_ROSTER_SOURCE_ID,fileName:'Active Medical Roster.xlsx',blocks:[block('2026-09-21','2026-10-04',false),block('2026-10-05','2026-10-18',false),block('2026-10-19','2026-11-01')]};
const active=[{id:'old',firstDate:'2026-09-21',lastDate:'2026-11-01'}];
assert.deepEqual(planVhhHistoryWindows(extract,active,'2026-10-20'),[{from:'2026-09-21',to:'2026-10-18',overlap:'old'}]);
assert.deepEqual(planVhhHistoryWindows(extract,[{...active[0],firstDate:'2026-10-19'}],'2026-10-20'),[{from:'2026-09-21',to:'2026-10-18',overlap:''}]);
assert.deepEqual(planVhhHistoryWindows({...extract,blocks:[...extract.blocks,block('2026-11-02','2026-11-15',false)]},[{firstDate:'2026-09-21',lastDate:'2026-10-18'},active[0]],'2026-10-20'),[{from:'2026-09-21',to:'2026-10-18',overlap:'old'}]);
assert.throws(()=>requirePreservedPrefix('2026-09-21','2026-10-19',[]),/without a prepared archive/);
assert.throws(()=>requirePreservedPrefix('2026-09-21','2026-10-19',[{from:'2026-09-22',to:'2026-10-18'}]),/without a prepared archive/);
assert.equal(rosterCoverageContains({startDate:'2026-08-24',endDate:'2026-09-21',rosterEndDate:'2026-09-20'},'2026-09-21'),false);
const input=structuredClone(extract);for(const b of input.blocks)b.visible=true;
const old=buildVhhDerivedRosterPayload({extract:input,fileId:'old',contentHash:'a'.repeat(64),lastModified:1});
const incoming=buildVhhDerivedRosterPayload({extract,fileId:'new',contentHash:'b'.repeat(64),lastModified:2});
async function prepare(runId,payload,activate=false,options={}) {
 const plan=await attachRosterCompactPlan(await planRosterImportBatches(payload),payload);plan.manifest.presenceMode='cached';plan.revision=await rosterImportDigest(plan.manifest);
 await call(()=>beginBoundedRosterImport(db,runId,payload.file,plan,options));
 for(const batch of plan.batches)await call(()=>stageBoundedRosterBatch(db,runId,payload.file,plan.revision,batch));
 await call(()=>prepareBoundedRosterPresence(db,runId,payload.file,plan.revision,plan.batches[0]));
 for(const batch of plan.compactBatches)await call(()=>prepareBoundedRosterMetadata(db,runId,plan.revision,batch));
 await call(()=>prepareBoundedRosterMetadata(db,runId,plan.revision));
 if(activate)await call(()=>activateBoundedRosterTerm(db,runId,plan.revision));
 return plan;
}
await prepare('old-run',old,true);
const parent='test-rollover', window={from:'2026-09-21',to:'2026-10-18'};
const archive=buildVhhDerivedRosterPayload({extract:extractVhhHistoryWindow(extract,{...window,fileName:historyFileName(window.from,window.to)}),fileId:'archive',contentHash:'c'.repeat(64),lastModified:2});
await prepare(historyJobId(parent,window.from,window.to),archive);
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='archive'").get().active,0);
const run={id:parent,sourceId:VHH_ROSTER_SOURCE_ID};
async function step(body){return call(async()=>{
 switch(body.phase){
 case 'bounded-patch-start':return beginRosterPatch(db,run,body.targetFileId,true);
 case 'bounded-patch-page':return readRosterPatchPage(db,run,body);
 case 'bounded-patch-plan':return pinRosterPatchPlan(db,run,body);
 case 'bounded-patch-stage':return stageRosterPatchBatch(db,run,body);
 case 'bounded-patch-activate':return activateRosterPatch(db,run,body,{file:incoming.file,publishFacility:true,historyWindows:[window]});
 case 'bounded-finalize':return {completed:true};
 }
});}
const oldCount=sqlite.prepare("SELECT COUNT(*) n FROM roster_events WHERE file_id='old'").get().n;
db.failRunIncludes='UPDATE roster_import_jobs SET activated=1 WHERE run_id=? AND plan_revision=?';
await assert.rejects(executeRosterPatch(incoming,{fileId:'old'},step),/Injected/);
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_events WHERE file_id='old'").get().n,oldCount,'failed promotion retains every old shift');
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='archive'").get().active,0,'failed promotion cannot expose archive duplicates');
await executeRosterPatch(incoming,{fileId:'old'},step);
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='archive'").get().active,1);
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1").get().n,oldCount);
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_events WHERE file_id='old' AND start_date<'2026-10-19'").get().n,0);
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_term_authority").get().n,0,'archives do not assume whole-term authority');
// The fallback full replacement must promote its history in the same transaction.
const replacementExtract={sourceId:VHH_ROSTER_SOURCE_ID,fileName:'Active Medical Roster.xlsx',blocks:[block('2026-10-19','2026-10-25',false),block('2026-10-26','2026-11-01')]};
const replacement=buildVhhDerivedRosterPayload({extract:replacementExtract,fileId:'replacement',contentHash:'e'.repeat(64),lastModified:3});
const archiveWindow={from:'2026-10-19',to:'2026-10-25'};
const replacementArchive=buildVhhDerivedRosterPayload({extract:extractVhhHistoryWindow(replacementExtract,{...archiveWindow,fileName:historyFileName(archiveWindow.from,archiveWindow.to)}),fileId:'replacement-archive',contentHash:'f'.repeat(64),lastModified:3});
await prepare(historyJobId('replacement-run',archiveWindow.from,archiveWindow.to),replacementArchive);
const replacementPlan=await prepare('replacement-run',replacement,false,{allowReplacement:true});
const activate=()=>call(()=>activateBoundedRosterTerm(db,'replacement-run',replacementPlan.revision,{publishFacility:true,historyWindows:[archiveWindow]}));
db.failRunIncludes='UPDATE roster_import_jobs SET activated=1 WHERE run_id=? AND plan_revision=?';
await assert.rejects(activate(),/Injected/);
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='old'").get().active,1);
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='replacement'").get().active,0);
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='replacement-archive'").get().active,0);
await activate();
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='old'").get().active,0);
assert.equal(sqlite.prepare("SELECT active FROM roster_files WHERE id='replacement-archive'").get().active,1);
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1").get().n,oldCount);
const manifest={terms:[{termStart:'2026-08-03',termEnd:'2026-11-01',visibleFrom:'2026-07-01',coverage:[{startDate:'2026-10-19',endDate:'2026-11-02',rosterEndDate:'2026-11-01'}]}],days:{}};
const bucket={get:async key=>key.endsWith('manifest.json')?{arrayBuffer:async()=>new TextEncoder().encode(JSON.stringify(manifest)).buffer}:null};
const absent=await loadPublishedFacilityDays(bucket,['vhh'],'2026-09-22','2026-10-20');
assert.equal(absent.preparing,false);assert.equal(absent.missing[0].reason,'coverage-unavailable');
const broken=await loadPublishedFacilityDays(bucket,['vhh'],'2026-10-19','2026-10-20');
assert.equal(broken.preparing,true,'missing artifact inside real coverage remains a preparation error');
console.log('VHH hidden history: bounded gap planning, atomic rollover rollback/retry, no duplicate shifts, grade-preserving archive, overnight boundary and missing-data responses passed.');
