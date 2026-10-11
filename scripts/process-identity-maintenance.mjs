import {pathToFileURL} from 'node:url';
import {guardedFetch} from '../functions/_lib/outbound-network.js';

export async function continueIdentityMaintenance({mode,request,maxSteps=600,maxMs=15*60*1000,now=Date.now,report=console.log,wait=ms=>new Promise(resolve=>setTimeout(resolve,ms))}) {
 if(!['register','audit'].includes(mode))throw Error('Choose register or audit.');
 if(!Number.isInteger(maxSteps)||maxSteps<1||maxSteps>600)throw Error('Invalid checkpoint limit.');
 const started=now();
 let busyResponses=0;
 for(let step=0;step<maxSteps && now()-started<maxMs;step++) {
  const result=await request(mode);
  if(step%25===0 || result.status!=='running')report(JSON.stringify({mode,checkpoint:step+1,status:result.status,examined:result.examined,candidates:result.candidates}));
  if(['complete','already-complete','unchanged'].includes(result.status))return {status:'complete',checkpoints:step+1};
  // A confirmed lease collision performed no identity mutation. Give the
  // watchdog time to release it, within the same request/time bounds. Never
  // retry errors, uncertain responses or a quota/source-priority deferral.
  if(result.status==='busy' && ++busyResponses<=3 && now()-started+30000<maxMs) {
   await wait(30000);
   continue;
  }
  if(result.status!=='running')return {status:'deferred',reason:result.status,checkpoints:step+1};
  busyResponses=0;
 }
 return {status:'deferred',reason:'checkpoint-or-time-limit'};
}

if(process.argv[1] && import.meta.url===pathToFileURL(process.argv[1]).href) {
 const base=String(process.env.ROSTER_AUTOMATION_BASE_URL||'https://roster-to-calendar.pages.dev').replace(/\/$/,'');
 const token=String(process.env.ROSTER_AUTOMATION_TOKEN||'');
 if(!token)throw Error('ROSTER_AUTOMATION_TOKEN required.');
 const result=await continueIdentityMaintenance({mode:process.env.IDENTITY_MAINTENANCE_MODE,
  request:async mode=>{
   // Each call has its own metered reservation, account admission, import
   // priority guard and one-name lease. Never retry an uncertain response.
   const response=await guardedFetch(process.env,`${base}/api/automation/identity-maintenance`,{
    method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({mode}),signal:AbortSignal.timeout(120000)}, {label:'Bounded identity catch-up'});
   if([429,503].includes(response.status))return {status:'usage-or-service-deferred'};
   if(!response.ok)throw Error(`Identity checkpoint failed: HTTP ${response.status}`);
   return await response.json();
  }});
 console.log(JSON.stringify(result));
}
