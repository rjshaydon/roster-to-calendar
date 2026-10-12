import {assembleSubscriptionCalendar} from '../functions/api/feed.js';
import {normalizeSubscriptionEvents} from '../functions/_lib/d1-calendar.js';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
export async function processSubscriptionBuild({call,admit,dispatchId='',prime=false,maximumCalendars=20,deadlineMs=12*60000,now=()=>Date.now()}){
 const started=now();
 if(prime){
  let cursor=null;
  do{
   if(now()-started>=deadlineMs)return {status:'yielded',built:0};
   if(!await admit())return {status:'deferred',built:0};
   const result=await call({mode:'prime',cursor});
   if(result.status!=='running')return {status:result.status,built:0};
   cursor=result.cursor;
  }while(cursor);
 }
 const begin=await call({mode:'start',dispatchId});
 if(begin.status!=='running')return {status:begin.status,built:0};
 const leaseId=begin.leaseId;let built=0;
 while(built<maximumCalendars&&now()-started<deadlineMs){
  const next=await call({mode:'next',leaseId});
  if(next.status==='complete'){await call({mode:'finish',leaseId});return {status:'complete',built};}
  if(next.status!=='running')return {status:next.status,built};
  if(!await admit()){await call({mode:'finish',leaseId});return {status:'deferred',built};}
  const base={leaseId,queueKey:next.queueKey};
  let snapshot=await call({...base,mode:'snapshot'});
  if(snapshot.status==='preparing')snapshot=await call({...base,mode:'prepare',revision:snapshot.revision});
  if(['unchanged','gone'].includes(snapshot.status))continue;
  if(snapshot.status==='deferred'){await call({mode:'finish',leaseId});return {status:'deferred',built};}
  if(snapshot.status!=='running')return {status:snapshot.status,built};
  const events=[...snapshot.inputs.designations];let cursor=null,pages=0;
  do{
   if(now()-started>=deadlineMs){await call({mode:'finish',leaseId});return {status:'yielded',built};}
   const page=await call({...base,mode:'events',revision:snapshot.revision,cursor});
   if(page.status==='deferred'){await call({mode:'finish',leaseId});return {status:'deferred',built};}
   if(page.status!=='running'){if(['stale','gone'].includes(page.status))await call({mode:'finish',leaseId});return {status:page.status,built};}
   events.push(...page.events);cursor=page.cursor;
   // Explicit refusal rather than silent truncation; normal multi-year calendars
   // remain complete. This cap protects runner memory, never publishes a prefix.
   if(++pages>1024)throw Error('Subscription exceeds the complete-export safety bound.');
  }while(cursor);
  const calendar=assembleSubscriptionCalendar(snapshot.record,next.queueKey.split('/')[3]==='range'?'range':'full',snapshot.claims,
   {...snapshot.inputs,events:normalizeSubscriptionEvents(events,snapshot.aliases)},true);
  const digest=createHash('sha256').update(calendar.ics).digest('hex');
  const result=await call({...base,mode:'publish',revision:snapshot.revision,ics:calendar.ics,displayName:calendar.displayName,digest});
  if(result.status==='deferred'){await call({mode:'finish',leaseId});return {status:'deferred',built};}
  if(!['complete','unchanged','gone'].includes(result.status)){if(result.status==='stale')await call({mode:'finish',leaseId});return {status:result.status,built};}
  built++;
 }
 await call({mode:'finish',leaseId});return {status:'yielded',built};
}
if(import.meta.url===pathToFileURL(process.argv[1]||'').href){
 const base=String(process.env.ROSTER_AUTOMATION_BASE_URL||'').replace(/\/$/,''),token=process.env.ROSTER_AUTOMATION_TOKEN;
 if(!base||!token)throw Error('Subscription runner credentials unavailable.');
 async function post(path,input){
  const r=await fetch(base+path,{method:'POST',headers:{authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify(input),signal:AbortSignal.timeout(120000)});
  const value=await r.json().catch(()=>null);
  if(r.status===409&&value?.status)return value;
  if(!r.ok||!value)throw Error('Subscription callback failed; progress and unknown costs retained.');
  return value;
 }
 let until=0;
 const result=await processSubscriptionBuild({prime:process.env.SUBSCRIPTION_BUILD_PRIME==='true',dispatchId:process.env.SUBSCRIPTION_BUILD_DISPATCH_ID||'',call:input=>post('/api/automation/subscription-build',input),
  admit:async()=>{if(Date.now()<until)return true;const budget=await post('/api/automation/account-budget',{reuseCurrentGrant:true});if(budget.deferred)return false;until=Date.now()+4*60000;return true;}});
 console.log(JSON.stringify(result));
}
