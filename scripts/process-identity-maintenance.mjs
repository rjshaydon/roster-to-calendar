import {continueExternalIdentityAudit} from './process-external-identity-audit.mjs';
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
 let useLegacy=true;
 if(process.env.IDENTITY_MAINTENANCE_MODE==='audit' && process.env.IDENTITY_EXTERNAL_AUDIT_ENABLED==='true') {
  let lastAdmission=0;
  const request=async body=>{
   if(Date.now()-lastAdmission>5*60*1000){
    const response=await guardedFetch(process.env,`${base}/api/automation/account-budget`,{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({reuseCurrentGrant:true})},{label:'Identity audit account admission'});
    const admission=await response.json();if(!response.ok||admission.deferred)return {status:'deferred'};lastAdmission=Date.now();
   }
   const response=await guardedFetch(process.env,`${base}/api/automation/identity-audit`,{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify(body),signal:AbortSignal.timeout(120000)},{label:'External identity audit checkpoint'});
   if([429,503].includes(response.status))return {status:'deferred'};
   if(response.status===409){const result=await response.json();if(result.status==='lease-expired')return result;throw Object.assign(Error('Identity evidence changed'),{staleEvidence:result.status==='stale'});}
   if(!response.ok)throw Error(`External audit callback failed: HTTP ${response.status}`);
   return await response.json();
  };
  const external=await continueExternalIdentityAudit({request,dispatchId:process.env.IDENTITY_AUDIT_DISPATCH_ID||''});
  useLegacy=external.status==='external-disabled';
  if(!useLegacy)console.log(JSON.stringify(external));
 }
 if(useLegacy) {
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
}
