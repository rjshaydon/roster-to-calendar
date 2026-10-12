import {subscriptionEvidence,prepareSubscriptionExport,subscriptionEventPage,subscriptionDigest,queueSubscriptionArtifact,SUBSCRIPTION_ARTIFACT_VERSION} from '../../_lib/subscription-artifacts.js';
import {identityMaintenancePreflight} from '../../_lib/identity-maintenance-preflight.js';
import {reserveRosterMaintenanceBudget} from '../../_lib/roster-maintenance-budget.js';
import {guardedFetch,localFeatureDisabledResponse} from '../../_lib/outbound-network.js';
const LEASE='subscriptions/build-lease.json',PREFIX='subscriptions/requests/',TTL=20*60000;
export async function onRequestPost(context){
 const input=context.data?.subscriptionBuildInput??await context.request.json().catch(()=>({}));
 const auth=context.request.headers.get('authorization'),env=context.env,r2=env.ROSTER_FILES,db=env.ROSTER_DB;
 const runner=Boolean(env.ROSTER_AUTOMATION_TOKEN)&&auth==='Bearer '+env.ROSTER_AUTOMATION_TOKEN;
 const watchdog=input.mode==='check'&&Boolean(env.ROSTER_WATCHDOG_TOKEN)&&auth==='Bearer '+env.ROSTER_WATCHDOG_TOKEN;
 if(!runner&&!watchdog)return Response.json({error:'Unauthorized.'},{status:401});
 if(!['build','ready'].includes(env.SUBSCRIPTION_ARTIFACT_MODE)||env.IDENTITY_REVIEW_ENABLED!=='true')return Response.json({status:'disabled'});
 const local=localFeatureDisabledResponse(env,'Subscription calendar assembly');if(local)return local;
 if(!['prime','check','start','next','snapshot','prepare','events','publish','finish'].includes(input.mode))return Response.json({error:'Invalid operation.'},{status:400});
 try{
  if(input.mode==='prime'){
   const deferred=await subscriptionAdmission(env);if(deferred)return Response.json(deferred);
   const cursor=String(input.cursor||'');if(cursor.length>512)return Response.json({error:'Invalid prime cursor.'},{status:400});
   const result=await db.prepare(`SELECT token FROM subscription_tokens WHERE token>? UNION SELECT subscription_token AS token FROM account_profiles WHERE subscription_token<>'' AND subscription_token>? ORDER BY token LIMIT 3`).bind(cursor,cursor).all();
   const rows=result.results||[],page=rows.slice(0,2);
   for(const {token} of page)for(const view of ['full','range'])await queueSubscriptionArtifact(r2,token,view,{revision:'0'.repeat(64)});
   return Response.json({status:'running',accounts:page.length,cursor:rows.length>2?page.at(-1).token:null});
  }
  const previous=await r2.get(LEASE),lease=previous?await previous.json():{},now=Date.now();
  if(['check','start'].includes(input.mode)){
   if(lease.until>now&&!(input.mode==='start'&&input.dispatchId===lease.id))return Response.json({status:'busy'});
   if(!(await r2.list({prefix:PREFIX,limit:1})).objects.length)return Response.json({status:'unchanged'});
   if(input.mode==='check'&&!env.GITHUB_ACTIONS_TOKEN)return Response.json({status:'dispatch-unconfigured'});
   const next={id:input.mode==='start'&&lease.until>now?lease.id:crypto.randomUUID(),until:now+TTL};
   if(!await r2.put(LEASE,JSON.stringify(next),{onlyIf:previous?{etagMatches:previous.etag}:{etagDoesNotMatch:'*'}}))return Response.json({status:'busy'});
   if(input.mode==='start')return Response.json({status:'running',leaseId:next.id});
   const response=await guardedFetch(env,'https://api.github.com/repos/rjshaydon/roster-to-calendar/actions/workflows/subscription-build.yml/dispatches',{
    method:'POST',headers:{Accept:'application/vnd.github+json',Authorization:'Bearer '+env.GITHUB_ACTIONS_TOKEN,'Content-Type':'application/json','X-GitHub-Api-Version':'2022-11-28','User-Agent':'roster-to-calendar'},
    body:JSON.stringify({ref:env.ROSTER_GITHUB_WORKFLOW_REF||'main',inputs:{dispatch_id:next.id}}),signal:AbortSignal.timeout(15000)
   },{label:'Subscription workflow dispatch'});
   return Response.json({status:response.ok?'dispatched':'dispatch-uncertain'});
  }
  if(!input.leaseId||lease.id!==input.leaseId||lease.until<=now)return Response.json({status:'lease-expired'},{status:409});
  if(input.mode==='finish'){
   await r2.put(LEASE,JSON.stringify({id:lease.id,until:0}),{onlyIf:{etagMatches:previous.etag}});
   return Response.json({status:'complete'});
  }
  if(input.mode==='next'){
   const list=await r2.list({prefix:PREFIX,limit:1});
   return Response.json({status:list.objects.length?'running':'complete',queueKey:list.objects[0]?.key||null});
  }
  const queueKey=String(input.queueKey||'');
  if(!/^subscriptions\/requests\/[a-f0-9]{64}\/(full|range)\/[a-f0-9]{64}\.json$/.test(queueKey))return Response.json({error:'Invalid request key.'},{status:400});
  const queued=await r2.get(queueKey);if(!queued)return Response.json({status:'gone'});
  const request=await queued.json(),identitiesEnabled=env.IDENTITY_REVIEW_ENABLED==='true';
  // Optional exports never compete with source imports or bypass the shared
  // reservation gate. Unknown callback costs remain held by middleware.
  const deferred=await subscriptionAdmission(env);if(deferred)return Response.json(deferred);
  const exportKey=`subscriptions/exports/${await subscriptionDigest(queueKey)}.json`;
  if(input.mode==='snapshot'){
   const evidence=await subscriptionEvidence(db,request.token,request.view,identitiesEnabled);
   if(!evidence){await r2.delete(queueKey);return Response.json({status:'gone'});}
   if(await r2.head(evidence.key)){await r2.delete(queueKey);await r2.delete(exportKey);return Response.json({status:'unchanged'});}
   // Pin the small evidence record separately from alias preparation, keeping
   // each complete metered callback within the Free CPU allowance.
   await r2.put(exportKey,JSON.stringify({evidence,revision:evidence.revision,until:now+TTL,phase:'pending'}));
   return Response.json({status:'preparing',revision:evidence.revision});
  }
  const exported=await r2.get(exportKey),snapshot=exported?await exported.json():null;
  if(!snapshot||snapshot.until<=now||snapshot.revision!==input.revision)return Response.json({status:'stale'},{status:409});
  if(input.mode==='prepare'){
   if(snapshot.phase==='pending'){
    // Provisional alias preparation cannot authorize an event export. Every
    // later page and promotion rechecks current ownership and all input facts.
    const prepared=await prepareSubscriptionExport(db,request.token,request.view,identitiesEnabled,snapshot.evidence,{deferFinalEvidenceCheck:true});
    await r2.put(exportKey,JSON.stringify({...prepared,until:snapshot.until,phase:'ready'}));
    return Response.json({status:'running',revision:prepared.revision,record:prepared.record,claims:prepared.claims,aliases:prepared.aliases,inputs:prepared.inputs});
   }
   return Response.json({status:'running',revision:snapshot.revision,record:snapshot.record,claims:snapshot.claims,aliases:snapshot.aliases,inputs:snapshot.inputs});
  }
  const evidence=await subscriptionEvidence(db,request.token,request.view,identitiesEnabled);
  if(!evidence){await r2.delete(queueKey);await r2.delete(exportKey);return Response.json({status:'gone'});}
  if(snapshot.phase!=='ready'||evidence.revision!==snapshot.revision)return Response.json({status:'stale'},{status:409});
  if(input.mode==='events')return Response.json({status:'running',...await subscriptionEventPage(db,snapshot.aliases,snapshot.start,snapshot.end,input.cursor)});
  if(input.mode==='publish'){
   const ics=String(input.ics||''),bytes=new TextEncoder().encode(ics);
   if(bytes.length>2*1024*1024||!ics.startsWith('BEGIN:VCALENDAR\r\n')||!ics.endsWith('END:VCALENDAR\r\n')||ics.includes('\0'))return Response.json({error:'Invalid calendar artifact.'},{status:400});
   if(await subscriptionDigest(ics)!==input.digest)return Response.json({error:'Calendar digest mismatch.'},{status:400});
   // The content-addressed key is the promotion fence. A later ownership/input
   // change yields a different key on GET, even if this upload finishes late.
   await r2.put(evidence.key,bytes,{onlyIf:{etagDoesNotMatch:'*'},httpMetadata:{contentType:'text/calendar; charset=utf-8'},customMetadata:{revision:evidence.revision,version:SUBSCRIPTION_ARTIFACT_VERSION,displayName:String(input.displayName||'roster').replace(/[\r\n]/g,'').slice(0,200)}});
   await r2.delete(queueKey);await r2.delete(exportKey);
   return Response.json({status:'complete'});
  }
  return Response.json({error:'Invalid phase.'},{status:400});
 }catch(error){return Response.json({error:String(error.message||'Subscription assembly failed.')},{status:500});}
}

async function subscriptionAdmission(env){
 const db=env.ROSTER_DB,preflight=await identityMaintenancePreflight(db,{accountAware:env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'});
 if(preflight.active_import)return {status:'deferred',reason:'import-active'};
 if(env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'&&(!Number(preflight.current_grant)||Number(preflight.unfinished_count)>=3))return {status:'deferred',reason:'account-admission'};
 if(!await reserveRosterMaintenanceBudget(db,0,131072))return {status:'deferred',reason:'account-headroom'};
 return null;
}
