import assert from 'node:assert/strict';
import {readFile,readdir} from 'node:fs/promises';
import {DatabaseSync} from 'node:sqlite';
import {staffHistoryResponse} from '../functions/api/state.js';
import {buildStaffHistoryCatalogue,matchesStaffHistory,coalesceStaffHistory,createStaffHistoryClient} from '../public/static/staff-history.js';
import {prepareStaffHistory} from './process-staff-history.mjs';
import {readStaffHistoryView,staffHistoryAccessAllowed,STAFF_HISTORY_POINTER,staffHistorySourceTerms,staffHistoryWatermark,staffHistoryEnabledFor} from '../functions/_lib/staff-history-cache.js';
import {onRequestPost} from '../functions/api/automation/staff-history.js';
import {storeCachedSnapshot} from '../functions/_lib/d1-calendar.js';
class R2 {
 objects=new Map();gets=0;puts=0;heads=0;
 async get(key){this.gets++;const o=this.objects.get(key);return o?{etag:o.etag,arrayBuffer:async()=>o.bytes.buffer.slice(o.bytes.byteOffset,o.bytes.byteOffset+o.bytes.byteLength),json:async()=>JSON.parse(new TextDecoder().decode(o.bytes)),customMetadata:o.options?.customMetadata}:null;}
 async head(key){this.heads++;const o=this.objects.get(key);return o?{etag:o.etag}:null;}
 async put(key,value,options={}){const old=this.objects.get(key);if(options.onlyIf?.etagMatches&&old?.etag!==options.onlyIf.etagMatches||options.onlyIf?.etagDoesNotMatch==='*'&&old)return null;const bytes=typeof value==='string'?new TextEncoder().encode(value):new Uint8Array(value);this.puts++;this.objects.set(key,{bytes,etag:String(this.puts),options});return {etag:String(this.puts)};}
}
class D1 {
 sql=[];constructor(sqlite){this.sqlite=sqlite;}
 prepare(sql){this.sql.push(sql);const db=this;return {args:[],bind(...args){this.args=args;return this;},async first(){return db.sqlite.prepare(sql).get(...this.args)||null;},async all(){return {results:db.sqlite.prepare(sql).all(...this.args)};},async run(){const r=db.sqlite.prepare(sql).run(...this.args);return {meta:{changes:Number(r.changes),rows_read:1,rows_written:Number(r.changes)}};},execute(){return /^\s*SELECT/i.test(sql)?this.all():this.run();}};}
 async batch(stmts){this.sqlite.exec('BEGIN');try{const results=[];for(const stmt of stmts)results.push(await stmt.execute());this.sqlite.exec('COMMIT');return results;}catch(e){this.sqlite.exec('ROLLBACK');throw e;}}
}
const sqlite=new DatabaseSync(':memory:');for(const name of(await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
for(const definition of ['name TEXT','last_modified INTEGER'])sqlite.exec('ALTER TABLE raw_roster_files ADD COLUMN '+definition);
const db=new D1(sqlite),r2=new R2(),env={ROSTER_DB:db,ROSTER_FILES:r2,ROSTER_AUTOMATION_TOKEN:'synthetic-token',STAFF_HISTORY_BUILD_ENABLED:'true'};
const today=new Intl.DateTimeFormat('en-CA',{timeZone:'Australia/Melbourne'}).format(new Date());
// Stable historical terms; future visibility tested independently below.
const member=(key,name,grade='HMO',extra={})=>({doctorKey:key,displayName:name,seniority:grade,contributionCount:1,membershipSource:'roster',...extra});
const terms=[['2026-02-02','2026-05-03'],['2026-05-04','2026-08-02'],['2026-08-03','2026-11-01']];
async function source(sourceType,lists){const pointers=[];for(const [i,members]of lists.entries()){const [termStart,termEnd]=terms[i];const staffKey=`test/${sourceType}/${termStart}.json.gz`;await storeCachedSnapshot(r2,staffKey,{members,seniorityOverrides:i===1?[{termStart,doctorKey:'HAYDON R',seniority:'Junior Registrar'}]:[]});pointers.push({termStart,termEnd,staffKey,staffRevision:'revision'+i});}await storeCachedSnapshot(r2,`facility-overview/v1/${sourceType}/manifest.json`,{terms:pointers});}
await source('mmc',[[member('HAYDON R','Richard HAYDON'),member('OLD','Historical Only'),member('OTHER','Sam Lee'),member('CONT','Continuing SMS','SMS',{contributionCount:0,membershipSource:'sms-continuity'}),member('DIR','Directory Only','SMS',{contributionCount:0,membershipSource:'provider-only'})],[member('HAYDON R','HAYDON, Richard')],[member('HAYDON R','Richard HAYDON','SMS')]]);
await source('ddh',[[member('RICHARD','Rich HAYDON','HMO'),member('OTHER','Sam Lee')]]);
const now=new Date().toISOString();sqlite.prepare("INSERT INTO roster_people(person_id,preferred_display_name,provenance,review_state,created_at,updated_at)VALUES(?,?,'test','approved',?,?)").run('person:haydon-richard','Richard Haydon',now,now);
for(const [sourceType,key,name]of [['mmc','HAYDON R','Richard HAYDON'],['ddh','RICHARD','Rich HAYDON'],['vhh','ALIAS','Haydon, Richie']])sqlite.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id,provenance,confidence,review_state,created_at,updated_at,approved_by)VALUES(?,?,?,'person:haydon-richard','test','test','approved',?,?,'test')").run(sourceType,key,name,now,now);
async function request(body){const response=await onRequestPost({env,request:new Request('https://test.invalid/api/automation/staff-history',{method:'POST',headers:{Authorization:'Bearer synthetic-token','Content-Type':'application/json'},body:JSON.stringify(body)})});const value=await response.json();assert.equal(response.status,200,JSON.stringify(value));return value;}
const result=await prepareStaffHistory({request});assert.equal(result.status,'complete');assert.equal(result.count,4);
// A queued import cannot make a read-only approved-alias export unsafe: source
// objects are immutable and final revision fences reject mixed generations.
sqlite.prepare("INSERT INTO roster_sync_runs(id,source_id,status,started_at)VALUES('synthetic-staging','monash-adults','queued',?)").run(now);
assert.ok((await request({mode:'aliases'})).aliases.length>0);
sqlite.prepare("UPDATE roster_sync_runs SET status='success' WHERE id='synthetic-staging'").run();

const streamed=await readStaffHistoryView(r2,{mode:'names'},today,{streamNames:true});assert.equal(streamed.headers.get('content-encoding'),'gzip');const decoded=await new Response(streamed.body.pipeThrough(new DecompressionStream('gzip'))).json();assert.equal(decoded.people.length,4);assert.equal(decoded.total,4);assert.equal(decoded.next,null);assert.ok(decoded.revision);
const names=await readStaffHistoryView(r2,{mode:'names'},today);assert.equal(names.people.length,4);assert.equal(names.next,null);
const richard=names.people.find(p=>p.displayName==='Richard Haydon');assert.ok(matchesStaffHistory(richard,'haydon'));assert.ok(matchesStaffHistory(richard,'HAYDON, Richie'));assert.ok(matchesStaffHistory(richard,'Richard Hay'));assert.ok(names.people.some(p=>p.displayName==='Historical Only'));assert.equal(names.people.filter(p=>p.displayName==='Sam Lee').length,2,'unapproved identical names stay separate');assert.ok(!JSON.stringify(names).includes('Continuing SMS'));assert.ok(!JSON.stringify(names).includes('Directory Only'));
const sqlBefore=db.sql.length,details=await readStaffHistoryView(r2,{mode:'histories',revision:names.revision,slots:[richard.slot]},today);assert.equal(db.sql.length,sqlBefore,'interactive history never queries D1');const history=details.people[0].history;assert.equal(history.length,4);assert.deepEqual(history.filter(h=>h.sourceType==='mmc').map(h=>h.grade),['HMO','Junior Registrar','SMS']);
const puts=r2.puts,sql=db.sql.length;assert.equal((await prepareStaffHistory({request})).status,'unchanged');assert.equal(r2.puts,puts);assert.equal(db.sql.length,sql,'unchanged preparation performs no D1 reads or writes');
assert.equal((await readStaffHistoryView(r2,{mode:'names',cachedRevision:names.revision},today)).unchanged,true);
// A changed day-manifest tag with identical staff evidence acknowledges the
// new source watermark without any identity queries or catalogue rewrites.
const sourceObject=r2.objects.get('facility-overview/v1/mmc/manifest.json');sourceObject.etag='synthetic-day-only-update';
const beforeAcknowledge=db.sql.length,putsBeforeAcknowledge=r2.puts;
assert.equal((await prepareStaffHistory({request})).status,'unchanged-history');assert.equal(db.sql.length,beforeAcknowledge);assert.equal(r2.puts,putsBeforeAcknowledge+1);assert.equal((await readStaffHistoryView(r2,{mode:'names',cachedRevision:names.revision},today)).unchanged,true);

assert.equal(staffHistoryAccessAllowed({role:'user',facilityOverviewEnabled:true},{mode:'all',isSms:true}),true);assert.equal(staffHistoryAccessAllowed({role:'user',facilityOverviewEnabled:true},{mode:'sites',isSms:false}),false);assert.equal(staffHistoryAccessAllowed({role:'user',facilityOverviewEnabled:false},{mode:'all',isSms:true}),false);assert.equal(staffHistoryAccessAllowed({role:'user',nonClinical:true,facilityOverviewEnabled:true},{mode:'all',isSms:true}),false);assert.equal(staffHistoryAccessAllowed({role:'creator'},{mode:'all'}),true);
// A changed identity marker invalidates the coherent catalogue before new data
// can be combined with old memberships; a missing shard is preparation, not no match.
await r2.put('identity/revision.json','{"revision":"changed"}');assert.equal((await readStaffHistoryView(r2,{mode:'names'},today)).preparing,true);
assert.equal((await prepareStaffHistory({request})).status,'complete');
const pointer=JSON.parse(new TextDecoder().decode(r2.objects.get(STAFF_HISTORY_POINTER).bytes));r2.objects.delete(`staff-history/v1/${pointer.revision}/histories/0.json.gz`);assert.equal((await readStaffHistoryView(r2,{mode:'histories',slots:[richard.slot],revision:pointer.revision},today)).preparing,true);
// Coalescing keeps grade transitions, breaks, simultaneous hospitals and future
// allocations distinct; it never turns roster listing into attendance evidence.
const row=(start,end,grade='HMO',sourceType='mmc',upcoming=false)=>({termStart:start,termEnd:end,grade,sourceType,upcoming,rosterNames:['Example']});const merged=coalesceStaffHistory([row(...terms[0]),row(...terms[1]),row(...terms[2],'SMS'),row('2027-02-01','2027-05-02','HMO'),row(...terms[0],'HMO','ddh'),row('2026-11-02','2027-01-31','SMS','mmc',true)]);assert.equal(merged.length,5);assert.equal(merged.find(r=>r.termStart==='2026-02-02'&&r.sourceType==='mmc').termEnd,'2026-08-02');
// Browser cache and late logout: typing does not fetch, histories are reused,
// concurrent catalogue loads coalesce, and a late response cannot refill logout.
let calls=0,resolve;const client=createStaffHistoryClient(async input=>{calls++;if(input.mode==='names')return {revision:'one',people:[richard],next:null};return {revision:'one',people:[{slot:richard.slot,history}]};});await Promise.all([client.load(),client.load()]);assert.equal(calls,1);for(const q of ['r','ri','rich','haydon'])client.matches(q);assert.equal(calls,1);await client.details([richard]);await client.details([richard]);assert.equal(calls,2);
const late=createStaffHistoryClient(()=>new Promise(r=>{resolve=r;}));const loading=late.load();late.clear();resolve({revision:'old',people:[richard],next:null});await loading;assert.equal(late.matches('Richard').length,0);
// A failed later catalogue page must not commit its revision: a retry must
// fetch the first page again rather than accept an empty 'unchanged' cache.
let failPage=true;const pagedCalls=[];const paged=createStaffHistoryClient(async input=>{pagedCalls.push(input);if(input.page===0)return {revision:'new',people:[richard],next:1};if(failPage)throw Error('synthetic lost page');return {revision:'new',people:[{...richard,slot:99,displayName:'Historical Later Page',names:['Historical Later Page']}],next:null};});
await assert.rejects(()=>paged.load(),/lost page/);assert.equal(paged.revision,'');failPage=false;await paged.load();assert.equal(pagedCalls[2].cachedRevision,'');assert.equal(paged.matches('Later Page').length,1);
await assert.rejects(()=>readStaffHistoryView(r2,{mode:'histories',slots:Array.from({length:13},(_,i)=>i)},today),/twelve/);
// Execute the actual shared response gate; denial precedes any R2 history
// read in both the fast authenticated route and legacy direct-handler fallback.
let reads=0;
const denied=await staffHistoryResponse({env:{STAFF_HISTORY_ENABLED:'true',ROSTER_FILES:{async get(){reads++;throw Error('Unexpected history read');}}}},{mode:'names'},{role:'user',facilityOverviewEnabled:true},{mode:'site',isSms:false});
assert.equal(denied.status,403);assert.equal(reads,0);
// Old cached staff objects without contribution counts are not silently lost.
// The background endpoint uses the indexed contribution table, excluding an
// inactive file and carried-forward directory entries without reading events.
await storeCachedSnapshot(r2,'legacy-staff',{members:[{doctorKey:'RECOVER',displayName:'Recovered Historical Doctor',seniority:'SMS'},{doctorKey:'CONTINUITY',displayName:'Continuing senior',membershipSource:'sms-continuity'}]});
await storeCachedSnapshot(r2,'facility-overview/v1/vhh/manifest.json',{terms:[{...Object.fromEntries(['termStart','termEnd'].map((k,i)=>[k,terms[0][i]])),staffKey:'legacy-staff',staffRevision:'legacy-evidence'}]});
sqlite.prepare("INSERT INTO roster_files(id,name,source_type,active)VALUES('legacy-active','Synthetic.xlsx','vhh',1),('legacy-inactive','Synthetic old.xlsx','vhh',0)").run();
sqlite.prepare("INSERT INTO facility_term_staff_contributions(source_type,term_start,doctor_key,file_id,display_name,seniority)VALUES('vhh','2026-02-02','RECOVER','legacy-active','Recovered Historical Doctor','HMO'),('vhh','2026-02-02','RETIRED','legacy-inactive','Retired File Doctor','SMS')").run();
assert.equal(staffHistoryEnabledFor({STAFF_HISTORY_ENABLED:'creator'},{role:'creator'}),true);assert.equal(staffHistoryEnabledFor({STAFF_HISTORY_ENABLED:'creator'},{role:'user'}),false);
const beforeLegacy=db.sql.length,watermark=await staffHistoryWatermark(r2,today),legacy=await request({mode:'fragment',sourceType:'vhh',termStart:'2026-02-02',inputRevision:watermark.inputRevision});assert.deepEqual(legacy.fragment.members.map(m=>m.displayName),['Recovered Historical Doctor']);assert.equal(legacy.fragment.members[0].grade,'HMO','dated contribution grade wins over a newer directory grade');assert.ok(db.sql.slice(beforeLegacy).some(sql=>sql.includes('facility_term_staff_contributions')));assert.ok(!db.sql.some(sql=>/FROM roster_events/i.test(sql)));
await storeCachedSnapshot(r2,'facility-overview/v1/casey/manifest.json',{terms:[{termStart:'2026-11-02',termEnd:'2027-01-31',staffKey:'future',staffRevision:'future'}]});
assert.equal((await staffHistorySourceTerms(r2,'casey','2026-09-30')).terms.length,0);assert.equal((await staffHistorySourceTerms(r2,'casey','2026-10-01')).terms.length,1,'future term respects its normal release date');
console.log('Staff history: bounded runner + real API/cache, historical-only and approved aliases, distinct names, grades, continuity exclusions, zero-D1 reads, unchanged no-op, invalidation, paging cache and permission gates passed.');
