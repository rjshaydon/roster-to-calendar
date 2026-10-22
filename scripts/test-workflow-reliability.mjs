import assert from 'node:assert/strict';
import {createMaintenanceAdmission,isReadOnlyPatchPage,retryablePatchReadFailure,retryPatchRead} from './automation-checkpoint-client.mjs';
import {createStaffHistoryRequest,runStaffHistory} from './process-staff-history.mjs';

let now=Date.parse('2026-10-22T01:00:00Z'),calls=[];
const ensure=createMaintenanceAdmission({now:()=>now,request:async body=>{
 calls.push(body);return {reused:true,validUntil:new Date(now+90000).toISOString()};
}});
await ensure();now+=29000;await ensure();assert.equal(calls.length,1);
now+=1000;await ensure();assert.equal(calls.length,2,'renew a reused grant before its real expiry, not five minutes after receipt');
await ensure({force:true});assert.equal(calls.at(-1).reuseCurrentGrant,false,'a failed page can request one fresh analytics admission');
let legacy=0;const older=createMaintenanceAdmission({now:()=>now,request:async()=>{legacy++;return {ok:true};}});
await older();now+=30000;await older();assert.equal(legacy,2,'mixed deployments without expiry are checked conservatively');
const midnight=createMaintenanceAdmission({now:()=>now,request:async()=>({budget:{validUntil:'2026-10-23T00:00:00Z'}})});
await midnight();now=Date.parse('2026-10-23T00:00:00Z');assert.equal((await midnight()).budget.validUntil,'2026-10-23T00:00:00Z','UTC rollover never reuses a local admission indefinitely');

for(const phase of ['bounded-patch-stage','bounded-patch-activate','bounded-finalize','failed'])assert.equal(isReadOnlyPatchPage('/api/automation/derived',{phase}),false);
assert.equal(isReadOnlyPatchPage('/api/automation/derived',{phase:'bounded-patch-page'}),true);
assert.equal(retryablePatchReadFailure(422,{error:'D1_ERROR: Network connection lost.'}),true);
assert.equal(retryablePatchReadFailure(422,{error:'D1_ERROR: malformed JSON: SQLITE_ERROR'}),false,'baseline fences must fail immediately');
for(const status of [400,401,403,409])assert.equal(retryablePatchReadFailure(status),false);
let attempts=0,delays=[];const cursor='same-read-cursor';
assert.equal(await retryPatchRead(async()=>{attempts++;if(attempts<3)throw Object.assign(Error('connection lost'),{retryablePatchRead:true});return cursor;},{sleep:async ms=>delays.push(ms)}),cursor);
assert.equal(attempts,3);assert.deepEqual(delays,[2000,4000]);
attempts=0;await assert.rejects(retryPatchRead(async()=>{attempts++;throw Object.assign(Error('persistent failure'),{retryablePatchRead:true});},{sleep:async()=>{}}),/persistent failure/);assert.equal(attempts,3);
attempts=0;await assert.rejects(retryPatchRead(async()=>{attempts++;throw Error('stale baseline');},{sleep:async()=>assert.fail('stale reads must not retry')}),/stale baseline/);assert.equal(attempts,1);

const clientFor=(value,status=200)=>createStaffHistoryRequest({base:'https://test.invalid',token:'synthetic',fetch:async()=>Response.json(value,{status})});
for(const [value,status] of [[{status:'deferred'},200],[{status:'deferred'},503],[{status:'stale'},409]]){
 const outcome=await runStaffHistory({request:clientFor(value,status)});assert.equal(outcome.status,value.status,'expected postponement is a normal workflow result');
}
for(const [value,status] of [[{error:'unexpected failure'},500],[{error:'unauthorized'},401],[{status:'deferred'},401]])await assert.rejects(runStaffHistory({request:clientFor(value,status)}),/checkpoint failed/,'unexpected failures must remain red');
await assert.rejects(runStaffHistory({request:createStaffHistoryRequest({base:'https://test.invalid',token:'synthetic',fetch:async()=>{throw Error('network unavailable');}})}),/network unavailable/);
let requests=[];const client=createStaffHistoryRequest({base:'https://test.invalid',token:'synthetic',now:()=>now,fetch:async(_env,url,options)=>{
 const body=JSON.parse(options.body);requests.push([url,body]);return Response.json(url.endsWith('account-budget')?{budget:{validUntil:new Date(now+120000).toISOString()}}:{status:'complete'});
}});
await client({mode:'fragment'});now+=61000;await client({mode:'aliases'});
assert.equal(requests.filter(([url])=>url.endsWith('account-budget')).length,2,'staff history also renews the actual expiry');
console.log('Workflow reliability: expected deferrals, real failure alerts, allowance expiry/rollover, bounded read retries and stale/write exclusions passed.');

// Exercise the actual queue executable and its transport, not only classifiers.
// Every fetch is intercepted; no credential or network access is involved.
const {spawnSync}=await import('node:child_process');
const preload=`
 import {readFile} from 'node:fs/promises';
 let raw=0,pages=[],writes=0,admissions=0;
 process.on('exit',()=>console.log('CHECKPOINT_TEST:'+JSON.stringify({raw,pages,writes,admissions})));
 globalThis.fetch=async(url,options={})=>{
  const path=new URL(url).pathname,body=JSON.parse(options.body||'{}');
  if(path.endsWith('/account-budget')){admissions++;return Response.json({ok:true,validUntil:new Date(Date.now()+600000).toISOString()});}
  if(path.endsWith('/pending'))return Response.json({boundedImportEnabled:true,patchImportEnabled:true,runs:[{id:'synthetic-run',sourceId:'monash-adults',fileId:'synthetic-file',fileName:'AdultTerm1.2026.xlsx',lastModified:1}]});
  if(path.endsWith('/facility-refresh'))return Response.json({idle:true});
  if(path.endsWith('/parser-config'))return Response.json({});
  if(path.endsWith('/raw')){raw++;return new Response(await readFile('fixtures/AdultTerm1.2026.xlsx'));}
  if(path.endsWith('/derived')){
   if(body.phase==='bounded-begin'){
    writes++;if(process.env.CHECKPOINT_SCENARIO==='write-lost')throw Error('uncertain write');
    return Response.json({mode:'patch',fileId:'synthetic-file',baselineRevision:'baseline',baselineParsedAt:'2026-10-22T00:00:00Z'});
   }
   if(body.phase==='bounded-patch-page'){
    pages.push(body.after||'');
    if(process.env.CHECKPOINT_SCENARIO==='stale')return Response.json({error:'D1_ERROR: malformed JSON: SQLITE_ERROR'},{status:422});
    if(pages.length===1)return Response.json({error:'D1_ERROR: Network connection lost.'},{status:422});
    if(!body.after)return Response.json({rows:[],next:'cursor-1'});
    return Response.json({deferred:true});
   }
  }
  throw Error('Unexpected synthetic checkpoint '+path+' '+body.phase);
 };
`;
for(const scenario of ['read-recovered','stale','write-lost']){
 const child=spawnSync(process.execPath,['--input-type=module','--eval',preload+"\ntry{await import('./scripts/process-roster-queue.mjs');process.exit(0);}catch(error){console.error(error.message);process.exit(1);}"],{
  encoding:'utf8',timeout:30000,env:{PATH:process.env.PATH,ROSTER_AUTOMATION_TOKEN:'synthetic',ROSTER_AUTOMATION_SOURCE_ID:'monash-adults',ROSTER_AUTOMATION_BASE_URL:'https://test.invalid',CHECKPOINT_SCENARIO:scenario}
 });
 if(child.error)throw Error(scenario+": "+child.error.code+"\n"+child.stdout+"\n"+child.stderr);
 const line=child.stdout.split('\n').find(s=>s.startsWith('CHECKPOINT_TEST:'));assert.ok(line,child.stderr);
 const state=JSON.parse(line.slice('CHECKPOINT_TEST:'.length));assert.equal(state.raw,1,'a same-run retry does not download or parse the workbook again');
 if(scenario==='read-recovered'){
  assert.equal(child.status,0,child.stderr);assert.deepEqual(state.pages,['','','cursor-1','cursor-1']);assert.equal(state.admissions,2,'read deferral refreshes admission once then retains progress');
 }else{
  assert.notEqual(child.status,0,'real failures retain a nonzero job exit');
  if(scenario==='stale')assert.equal(state.pages.length,1);else assert.equal(state.writes,1,'an uncertain write is never retried blindly');
 }
}
console.log('Actual queue transport recovers transient reads at the same cursor without re-downloading, stops stale baselines and preserves uncertain writes.');
