import assert from 'node:assert/strict';
import {continueIdentityMaintenance,createIdentityAdmission} from './process-identity-maintenance.mjs';
let calls=0;
const options={mode:'register',report:()=>{}};
const done=await continueIdentityMaintenance({...options,request:async()=>({status:++calls===3?'complete':'running'})});
assert.equal(done.status,'complete');assert.equal(calls,3);
for(const status of ['import-active','registration-active','usage-or-service-deferred','outside-audit-window']) {
 calls=0;
 const stopped=await continueIdentityMaintenance({...options,request:async()=>{calls++;return {status};}});
 assert.equal(stopped.status,'deferred');assert.equal(calls,1,`${status} must not spin or bypass admission`);
}
calls=0;let waits=0;
const collision=await continueIdentityMaintenance({...options,wait:async ms=>{assert.equal(ms,30000);waits++;},request:async()=>({status:++calls===1?'busy':'complete'})});
assert.equal(collision.status,'complete');assert.equal(calls,2);assert.equal(waits,1);
calls=0;waits=0;
const busy=await continueIdentityMaintenance({...options,wait:async()=>{waits++;},request:async()=>{calls++;return {status:'busy'};}});
assert.equal(busy.reason,'busy');assert.equal(calls,4);assert.equal(waits,3,'persistent lease contention has a fixed retry bound');
calls=0;
await continueIdentityMaintenance({...options,maxMs:30000,wait:async()=>{throw Error('Must not exceed time cap');},request:async()=>{calls++;return {status:'busy'};}});
assert.equal(calls,1);
calls=0;
const capped=await continueIdentityMaintenance({...options,maxSteps:2,request:async()=>{calls++;return {status:'running'};}});
assert.equal(capped.reason,'checkpoint-or-time-limit');assert.equal(calls,2);
let instant=0;calls=0;
await continueIdentityMaintenance({...options,maxMs:10,now:()=>instant,request:async()=>{calls++;instant+=10;return {status:'running'};}});
assert.equal(calls,1,'time limit stops before issuing another checkpoint');
calls=0;
await assert.rejects(()=>continueIdentityMaintenance({...options,request:async()=>{calls++;throw Error('Uncertain response');}}),/Uncertain response/);
assert.equal(calls,1,'network failures never trigger blind retries');
await assert.rejects(()=>continueIdentityMaintenance({...options,maxSteps:601}),/Invalid checkpoint/);
console.log('Identity continuation passed: serial checkpoints, completion, admission/priority stops, time/step caps and no uncertain retries.');

let clock=0,admissions=0,deferred=false;
const admit=createIdentityAdmission({now:()=>clock,request:async()=>{admissions++;return {ok:true,deferred};}});
assert(await admit());assert(await admit());assert.equal(admissions,1,'one separate initial admission, including at clock zero');
clock=300000;deferred=true;assert.equal(await admit(),false);assert.equal(admissions,2,'expired admission cannot bypass a quota deferral');
deferred=false;assert(await admit());assert.equal(admissions,3);
const uncertainAdmission=createIdentityAdmission({request:async()=>{throw Error('Uncertain admission');}});await assert.rejects(uncertainAdmission(),/Uncertain admission/);
console.log('Separate identity admission preserves cache expiry, fail-closed deferral and uncertain-error handling.');
