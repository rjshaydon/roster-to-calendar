import {guardedFetch} from '../../_lib/outbound-network.js';
import {DOCTOR_SOURCE_VERSION,doctorSourceKey,doctorSourceCurrentKey} from '../../_lib/doctor-source-artifacts.js';
import {automatedRosterWritesEnabled,rosterWritePausedResponse} from '../../_lib/roster-automation-guard.js';
import {loadCachedSnapshot} from '../../_lib/d1-calendar.js';
import {doctorMonthIndexKey,doctorMonthRowsKey,validDoctorMonth,DOCTOR_CALENDAR_MAX_BYTES,DOCTOR_CALENDAR_VERSION} from '../../_lib/doctor-calendar-artifacts.js';
import {facilityMetadataManifestKey,facilityMonthObjectKey} from '../../_lib/facility-overview-cache.js';
export async function onRequestPost({request,env}){
 const auth=request.headers.get('authorization'),runner=Boolean(env.ROSTER_AUTOMATION_TOKEN)&&auth==='Bearer '+env.ROSTER_AUTOMATION_TOKEN,watchdog=Boolean(env.ROSTER_WATCHDOG_TOKEN)&&auth==='Bearer '+env.ROSTER_WATCHDOG_TOKEN;
 if(!runner&&!watchdog)return Response.json({error:'Unauthorized.'},{status:401});
 if(!runner){const input=await request.json().catch(()=>null);return input?.mode==='check-source'?checkDoctorSourceBuild(env):Response.json({error:'Unauthorized.'},{status:401});}
 if(!automatedRosterWritesEnabled(env))return rosterWritePausedResponse();
 const r2=env.ROSTER_FILES;if(!r2)return Response.json({error:'Artifact store unavailable.'},{status:503});
 try{
  if(request.headers.get('content-type')==='application/octet-stream'){
   const header=request.headers.get('x-doctor-artifact')||'';if(header.length>4096)return bad();
   const m=JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(header),c=>c.charCodeAt(0))));
   const sourceKind=['source-person','source-catalogue'].includes(m.kind);
   if((sourceKind?!validManifestKey(m.sourceType,m.manifestKey):!validDoctorMonth(m.sourceType,m.month,m.parentRevision))||!/^[a-f0-9]{64}$/.test(m.digest)||!Number.isSafeInteger(m.bytes)||m.bytes<2||m.bytes>DOCTOR_CALENDAR_MAX_BYTES||Number(request.headers.get('content-length'))!==m.bytes||!request.body||!['index','rows','source-person','source-catalogue'].includes(m.kind))return bad();
   // Only a retained immutable monthly publication can authorize a derivative.
   if(!await r2.head(sourceKind?m.manifestKey:facilityMonthObjectKey(m.sourceType,m.month,m.parentRevision)))return Response.json({status:'gone'},{status:409});
   const key=sourceKind?doctorSourceKey(m.sourceType,m.kind==='source-person'?'people':'catalogues',m.digest):m.kind==='index'?doctorMonthIndexKey(m.sourceType,m.month,m.parentRevision):doctorMonthRowsKey(m.sourceType,m.month,m.digest);
   if(await r2.head(key)){if(m.kind==='source-person')await sourceReceipt(r2,m,key);return Response.json({status:'unchanged',key});}
   const checksum=Uint8Array.from(m.digest.match(/../g),s=>parseInt(s,16)).buffer;
   await r2.put(key,request.body,{onlyIf:{etagDoesNotMatch:'*'},sha256:checksum,httpMetadata:{contentType:'application/json'},customMetadata:{version:sourceKind?DOCTOR_SOURCE_VERSION:DOCTOR_CALENDAR_VERSION,parentRevision:String(m.parentRevision||''),digest:m.digest}});
   if(m.kind==='source-person')await sourceReceipt(r2,m,key);
   if(m.kind==='index')await r2.put(`doctor-calendars/v1/${m.sourceType}/${m.month}/latest.json`,JSON.stringify({indexKey:key}),{onlyIf:m.latestEtag?{etagMatches:m.latestEtag}:{etagDoesNotMatch:'*'}});
   return Response.json({status:'complete',key});
  }
  const input=await request.json();
  if(input.mode==='check-source')return checkDoctorSourceBuild(env);
  if(input.mode==='finish-source'){const lease=await r2.get('doctor-calendars/v2/build-lease.json'),value=lease?await new Response(lease.body).json():null;if(value?.id===input.dispatchId)await r2.put('doctor-calendars/v2/build-lease.json',JSON.stringify({...value,until:0}),{onlyIf:{etagMatches:lease.etag}});return Response.json({status:'complete'});}
  if(!['mmc','mch','ddh','vhh','casey'].includes(input.sourceType))return bad();
  if(input.mode==='object'){
   if(!validExportKey(input.sourceType,input.key))return bad();
   const object=await r2.get(input.key);return object?new Response(object.body,{headers:{'Content-Type':'application/octet-stream','x-object-etag':object.etag||''}}):Response.json({status:'gone'},{status:409});
  }
  if(input.mode==='source-receipt'){
   if(!/^[a-f0-9]{64}$/.test(input.inputSignature))return bad();
   return Response.json({receipt:await loadCachedSnapshot(r2,doctorSourceKey(input.sourceType,'prepared-people',input.inputSignature))});
  }
  if(input.mode==='source-current'){
   const object=await r2.get(doctorSourceCurrentKey(input.sourceType));
   return Response.json({current:object?await new Response(object.body).json():null,etag:object?.etag||''});
  }
  if(['prepare-source','activate-source'].includes(input.mode)){
   if(!/^[a-f0-9]{64}$/.test(input.digest)||input.catalogueKey!==doctorSourceKey(input.sourceType,'catalogues',input.digest)||!await r2.head(input.catalogueKey))return bad();
   const value={version:DOCTOR_SOURCE_VERSION,catalogueKey:input.catalogueKey,digest:input.digest,manifestRevision:input.manifestRevision,baseEtag:input.baseEtag||''};
   if(input.mode==='prepare-source'){
    if(!/^[a-f0-9]{64}$/.test(input.operationRevision)||!/^[a-f0-9]{64}$/.test(input.manifestRevision))return bad();
    const key=doctorSourceKey(input.sourceType,'generations',input.operationRevision),old=await r2.get(key),prior=old?await new Response(old.body).json():null;
    if(prior&&prior.manifestRevision!==input.manifestRevision)return Response.json({status:'stale'},{status:409});
    if(!await r2.put(key,JSON.stringify(value),{onlyIf:old?{etagMatches:old.etag}:{etagDoesNotMatch:'*'}}))return Response.json({status:'stale'},{status:409});
   }else{
    const head=await r2.head(facilityMetadataManifestKey(input.sourceType));if(!head||head.etag!==input.manifestEtag)return Response.json({status:'stale'},{status:409});
    const saved=await r2.put(doctorSourceCurrentKey(input.sourceType),JSON.stringify({...value,parentEtag:head.etag}),{onlyIf:input.baseEtag?{etagMatches:input.baseEtag}:{etagDoesNotMatch:'*'}});
    if(!saved)return Response.json({status:'stale'},{status:409});
   }
   return Response.json({status:'complete'});
  }
  if(input.mode==='manifest'){
   const manifest=await loadCachedSnapshot(r2,facilityMetadataManifestKey(input.sourceType));
   return Response.json({status:manifest?'running':'gone',months:manifest?.months||{}});
  }
  if(!validDoctorMonth(input.sourceType,input.month,input.parentRevision))return bad();
  const indexKey=doctorMonthIndexKey(input.sourceType,input.month,input.parentRevision);
  if(input.mode==='previous'){
   const object=await r2.get(`doctor-calendars/v1/${input.sourceType}/${input.month}/latest.json`);
   const latest=object?await new Response(object.body).json():null;
   const index=latest?.indexKey?.startsWith(`doctor-calendars/v1/${input.sourceType}/${input.month}/indexes/`)?await loadCachedSnapshot(r2,latest.indexKey):null;
   return Response.json({doctors:index?.doctors||{},etag:object?.etag||''});
  }
  if(input.mode==='check')return Response.json({status:await r2.head(indexKey)?'unchanged':'missing'});
  if(input.mode==='month'){
   const object=await r2.get(facilityMonthObjectKey(input.sourceType,input.month,input.parentRevision));
   return object?new Response(object.body,{headers:{'Content-Type':'application/octet-stream'}}):Response.json({status:'gone'},{status:409});
  }
  return bad();
 }catch(error){return Response.json({error:/checksum|digest/i.test(error.message)?'Artifact checksum rejected.':'Doctor calendar operation failed.'},{status:400});}
}
function bad(){return Response.json({error:'Invalid doctor calendar request.'},{status:400});}

function validManifestKey(source,key){return ['mmc','mch','ddh','vhh','casey'].includes(source)&&(key===facilityMetadataManifestKey(source)||new RegExp('^facility-overview/v1/'+source+'/manifests/[a-f0-9]{64}-[a-f0-9]{64}\\.json\\.gz$').test(key||''));}
function validExportKey(source,key){return validManifestKey(source,key)||new RegExp('^facility-overview/v1/'+source+'/staff/\\d{4}-\\d{2}-\\d{2}/[a-f0-9]{64}\\.json\\.gz$').test(key||'')||new RegExp('^doctor-calendars/v1/'+source+'/\\d{4}-\\d{2}/(rows|indexes)/[a-f0-9]{64}\\.json$').test(key||'')||new RegExp('^doctor-calendars/v2/'+source+'/catalogues/[a-f0-9]{64}\\.json$').test(key||'');}

async function sourceReceipt(r2,m,key){
 if(!/^[a-f0-9]{64}$/.test(m.inputSignature||'')||typeof m.doctorKey!=='string'||m.doctorKey.length>512)return;
 await r2.put(doctorSourceKey(m.sourceType,'prepared-people',m.inputSignature),JSON.stringify({doctorKey:m.doctorKey,pointer:{key,digest:m.digest,inputSignature:m.inputSignature}}),{onlyIf:{etagDoesNotMatch:'*'}});
}

async function checkDoctorSourceBuild(env){
 if(!automatedRosterWritesEnabled(env)||!['build','ready'].includes(env.SUBSCRIPTION_ARTIFACT_MODE))return Response.json({status:'disabled'});
 const r2=env.ROSTER_FILES,key='doctor-calendars/v2/build-lease.json',object=await r2.get(key),lease=object?await new Response(object.body).json():{};
 if(lease.until>Date.now())return Response.json({status:'busy'});
 const states=await Promise.all(['mmc','mch','ddh','vhh','casey'].map(async source=>{const [head,current]=await Promise.all([r2.head(facilityMetadataManifestKey(source)),loadCachedSnapshot(r2,doctorSourceCurrentKey(source))]);return Boolean(head)&&(!current||current.version!==DOCTOR_SOURCE_VERSION||current.parentEtag!==head.etag);}));
 if(!states.some(Boolean))return Response.json({status:'unchanged'});
 if(!env.GITHUB_ACTIONS_TOKEN)return Response.json({status:'dispatch-unconfigured'});
 const next={id:crypto.randomUUID(),until:Date.now()+20*60000};if(!await r2.put(key,JSON.stringify(next),{onlyIf:object?{etagMatches:object.etag}:{etagDoesNotMatch:'*'}}))return Response.json({status:'busy'});
 const response=await guardedFetch(env,'https://api.github.com/repos/rjshaydon/roster-to-calendar/actions/workflows/doctor-calendars.yml/dispatches',{method:'POST',headers:{Authorization:'Bearer '+env.GITHUB_ACTIONS_TOKEN,Accept:'application/vnd.github+json','Content-Type':'application/json','User-Agent':'roster-to-calendar','X-GitHub-Api-Version':'2022-11-28'},body:JSON.stringify({ref:env.ROSTER_GITHUB_WORKFLOW_REF||'main',inputs:{dispatch_id:next.id}}),signal:AbortSignal.timeout(15000)},{label:'Doctor calendar dispatch'});
 return Response.json({status:response.ok?'dispatched':'dispatch-uncertain'});
}
