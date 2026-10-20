import {describeStaffHistory,staffHistorySourceTerms,staffHistoryWatermark,readStaffHistoryFragment,historyObjectKey,STAFF_HISTORY_POINTER,requestStaffHistoryBuild} from '../../_lib/staff-history-cache.js';
import {storeCachedSnapshot} from '../../_lib/d1-calendar.js';
import {reserveRosterMaintenanceBudget} from '../../_lib/roster-maintenance-budget.js';
import {STAFF_HISTORY_SHARD} from '../../../public/static/staff-history.js';
import {identityMaintenancePreflight} from '../../_lib/identity-maintenance-preflight.js';
export async function onRequestPost(context){
 const {env,request}=context;
 if(!env.ROSTER_AUTOMATION_TOKEN||request.headers.get('Authorization')!=='Bearer '+env.ROSTER_AUTOMATION_TOKEN)return Response.json({error:'Unauthorized.'},{status:401});
 if(env.STAFF_HISTORY_BUILD_ENABLED!=='true')return Response.json({status:'disabled'});
 const input=await request.json(),r2=env.ROSTER_FILES,db=env.ROSTER_DB;
 try{
  if(input.mode==='enqueue'){await requestStaffHistoryBuild(context);return Response.json({status:'queued'});}
  if(input.mode==='describe'){
   const watermark=await staffHistoryWatermark(r2),object=await r2.get(STAFF_HISTORY_POINTER),pointer=object?await object.json():null;
   if(pointer?.inputRevision===watermark.inputRevision)return Response.json({status:'unchanged'});
   const descriptor=await describeStaffHistory(r2);
   return Response.json({status:'ready',descriptor,projectionInput:pointer?.projectionInput||''});
  }
  if(['terms','fragment'].includes(input.mode)){
   if((await staffHistoryWatermark(r2)).inputRevision!==input.inputRevision)return Response.json({status:'stale'},{status:409});
   const source=await staffHistorySourceTerms(r2,input.sourceType);
   if(input.mode==='terms')return Response.json(source);
   const term=source.terms.find(t=>t.termStart===input.termStart);
   if(!term)throw Error('Unknown published term.');
   return Response.json({fragment:await readStaffHistoryFragment(r2,term,{loadEvidence:async()=>{
    const preflight=await identityMaintenancePreflight(db,{accountAware:env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'});
    if(preflight.active_import||Number(preflight.unfinished_count)>=3||!Number(preflight.current_grant)||!await reserveRosterMaintenanceBudget(db,0,65536))throw Error('Older staff history evidence deferred.');
    const rows=await db.prepare(`SELECT c.doctor_key,c.display_name,c.seniority,c.membership_source FROM facility_term_staff_contributions c
     JOIN roster_files f ON f.id=c.file_id AND f.active=1 WHERE c.source_type=? AND c.term_start=? LIMIT 513`).bind(term.sourceType,term.termStart).all();
    if(rows.results.length>512)throw Error('Older staff history evidence exceeds its bounded page.');
    return rows.results.map(row=>({doctorKey:row.doctor_key,displayName:row.display_name,seniority:row.seniority,membershipSource:row.membership_source,contributionCount:1}));
   }})});
  }
  if(input.mode==='acknowledge'){
   const object=await r2.get(STAFF_HISTORY_POINTER),pointer=object?await object.json():null;
   if(!pointer||!input.projectionInput||pointer.projectionInput!==input.projectionInput||input.inputRevision!==(await staffHistoryWatermark(r2)).inputRevision)return Response.json({status:'stale'},{status:409});
   const result=await r2.put(STAFF_HISTORY_POINTER,JSON.stringify({...pointer,inputRevision:input.inputRevision}),{httpMetadata:{contentType:'application/json'},onlyIf:{etagMatches:object.etag}});
   return Response.json({status:result===null?'stale':'unchanged-history',count:pointer.count});
  }
  if(input.mode==='aliases'){
   // Keyset pagination on the existing alias primary key, with one coherent
   // identity revision. Never inspect accounts, events or the candidate graph.
   const preflight=await identityMaintenancePreflight(db,{accountAware:env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'});
   if(preflight.active_import||Number(preflight.unfinished_count)>=3||!Number(preflight.current_grant)||!await reserveRosterMaintenanceBudget(db,0,2048))return Response.json({status:'deferred'},{status:503});
   const source=String(input.afterSource||''),key=String(input.afterKey||'');
   if(source.length>20||key.length>200)throw Error('Invalid alias cursor.');
   const results=await db.batch([
    db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1'),
    db.prepare(`SELECT a.source_type,a.doctor_key,a.display_name,a.person_id,p.preferred_display_name FROM roster_person_aliases a JOIN roster_people p ON p.person_id=a.person_id
     WHERE (a.source_type,a.doctor_key)>(?,?) AND a.review_state='approved' AND p.status='active' ORDER BY a.source_type,a.doctor_key LIMIT 129`).bind(source,key)
   ]);
   const revision=Number(results[0].results[0].revision);
   if(input.identityRevision!=null&&input.identityRevision!==revision)return Response.json({status:'stale'},{status:409});
   const rows=results[1].results,page=rows.slice(0,128);
   return Response.json({identityRevision:revision,aliases:page.map(a=>({sourceType:a.source_type,doctorKey:a.doctor_key,displayName:a.display_name,personId:a.person_id,preferredName:a.preferred_display_name})),next:rows.length>128?{afterSource:page.at(-1).source_type,afterKey:page.at(-1).doctor_key}:null});
  }
  if(input.mode==='store'){
   const key=historyObjectKey(input.revision,input.kind,input.index),people=input.people;
   if(!Array.isArray(people)||people.length>(input.kind==='names'?128:STAFF_HISTORY_SHARD)||JSON.stringify(people).length>500000)throw Error('Staff history object limit exceeded.');
   if(!await r2.head(key))await storeCachedSnapshot(r2,key,{people},{revision:input.revision});
   return Response.json({ok:true});
  }
  if(input.mode==='finish'){
   const descriptor=input.descriptor;
   if(!descriptor||descriptor.inputRevision!==(await staffHistoryWatermark(r2)).inputRevision)return Response.json({status:'stale'},{status:409});
   if(!Number.isInteger(input.count)||input.count<0||input.count>4096)throw Error('Staff history catalogue limit exceeded.');
   const pages=Math.max(1,Math.ceil(input.count/128)),shards=Math.ceil(input.count/STAFF_HISTORY_SHARD);
   // Validate completion in bounded checkpoints before exposing the pointer.
   if(!await r2.head(historyObjectKey(input.revision,'names',pages-1))||shards&&!await r2.head(historyObjectKey(input.revision,'histories',shards-1)))throw Error('Staff history assembly incomplete.');
   if(!await reserveRosterMaintenanceBudget(db,0,8))return Response.json({status:'deferred'},{status:503});
   const identity=await db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1').first();
   if(Number(identity.revision)!==input.identityRevision)return Response.json({status:'stale'},{status:409});
   const pointer={revision:input.revision,inputRevision:descriptor.inputRevision,count:input.count,pages,projectionInput:String(input.projectionInput||''),missing:input.missing||[],preparedAt:new Date().toISOString()};
   const old=await r2.get(STAFF_HISTORY_POINTER);
   const result=await r2.put(STAFF_HISTORY_POINTER,JSON.stringify(pointer),{httpMetadata:{contentType:'application/json'},...(old?.etag?{onlyIf:{etagMatches:old.etag}}:{onlyIf:{etagDoesNotMatch:'*'}})});
   return Response.json({status:result===null?'stale':'complete',count:input.count});
  }
  throw Error('Invalid staff history operation.');
 }catch{
  console.warn(JSON.stringify({event:'staff-history',mode:input.mode,status:'interrupted'}));
  return Response.json({error:'Staff history preparation interrupted.'},{status:500});
 }
}
