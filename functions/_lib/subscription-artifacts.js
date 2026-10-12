import {accountIdentityAliases,expandApprovedIdentityAliases} from './doctor-identity.js';
import {hospitalLocationsFromSession,normalizeHospitalLocationMap,customEventForAccount,designationLeaveEventFromRow} from './d1-calendar.js';

export const SUBSCRIPTION_ARTIFACT_VERSION='subscription-v1';
const sources=['mmc','mch','ddh','vhh','casey'];
export async function subscriptionDigest(value){
 const bytes=new TextEncoder().encode(String(value));
 return [...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('');
}
// One indexed, read-only binding call. Evidence includes ownership, every input
// that the renderer uses, approved identity links, and atomic roster generations.
// Sentinels refuse oversized evidence instead of publishing a partial calendar.
export async function subscriptionEvidence(db,token,view='full',identitiesEnabled=true){
 const row=await db.prepare(`WITH RECURSIVE owner AS (
  SELECT p.email,p.real_name,p.role,CASE WHEN json_valid(s.session_json) THEN s.session_json ELSE '{}' END session_json
  FROM account_profiles p LEFT JOIN account_states s ON s.email=p.email
  WHERE p.email=COALESCE((SELECT email FROM subscription_tokens WHERE token=?),
   (SELECT email FROM account_profiles WHERE subscription_token=? LIMIT 1))
 ), claims AS (SELECT c.* FROM account_claims c JOIN owner o ON o.email=c.email),
 seeds AS (
  SELECT a.person_id FROM account_people a JOIN owner o ON a.email=o.email WHERE o.role NOT IN ('creator','owner')
  UNION SELECT a.person_id FROM roster_person_aliases a JOIN claims c ON a.source_type=c.source_type AND a.doctor_key=c.doctor_key
  UNION SELECT a.person_id FROM owner o,json_each('["mmc","mch","ddh","vhh","casey"]') s JOIN roster_person_aliases a ON a.source_type=s.value AND a.doctor_key=json_extract(o.session_json,'$.doctorKey') WHERE o.role IN ('creator','owner')
 ), people(person_id,depth) AS (
  SELECT person_id,0 FROM seeds UNION SELECT r.person_id,p.depth+1 FROM people p
   JOIN roster_person_redirects r ON r.old_person_id=p.person_id AND r.active=1 WHERE p.depth<16
 ), aliases AS (SELECT a.* FROM roster_person_aliases a WHERE a.person_id IN (SELECT person_id FROM people) AND a.review_state='approved'),
 wanted AS (
  SELECT source_type,doctor_key FROM claims UNION SELECT source_type,doctor_key FROM aliases
  UNION SELECT s.value,json_extract(o.session_json,'$.doctorKey') FROM owner o,json_each('["mmc","mch","ddh","vhh","casey"]') s WHERE o.role IN ('creator','owner')
 ) SELECT o.email,o.real_name,o.role,
 json_object('doctorKey',json_extract(o.session_json,'$.doctorKey'),'settings',json_extract(o.session_json,'$.settings'),
  'overrides',json_extract(o.session_json,'$.overrides'),'exportRange',json_extract(o.session_json,'$.exportRange'),
  'customEvents',json_extract(o.session_json,'$.customEvents')) session_json,
 (SELECT json_group_array(json_array(source_type,doctor_key,display_name)) FROM (SELECT * FROM claims ORDER BY source_type,doctor_key LIMIT 17)) claims_json,
 (SELECT json_group_array(json_array(source_type,doctor_key,display_name,person_id,updated_at)) FROM (SELECT * FROM aliases ORDER BY source_type,doctor_key LIMIT 65)) aliases_json,
 (SELECT json_group_array(json_array(person_id,status,preferred_display_name,merged_into_person_id,version)) FROM (SELECT * FROM roster_people WHERE person_id IN (SELECT person_id FROM people) ORDER BY person_id LIMIT 65)) people_json,
 (SELECT json_group_array(json_array(person_id,updated_at)) FROM account_people WHERE email=o.email) account_people_json,
 (SELECT revision FROM roster_identity_revision WHERE id=1) identity_revision,
 (SELECT json_group_array(json_array(id,source_type,parsed_at,content_revision)) FROM (
  SELECT f.id,f.source_type,f.parsed_at,c.content_revision FROM roster_files f LEFT JOIN roster_file_coverage c ON c.file_id=f.id
  WHERE f.active=1 AND f.source_type IN (SELECT source_type FROM wanted) ORDER BY f.source_type,f.id LIMIT 1025)) files_json,
 (SELECT json_group_array(json_array(id,source_type,doctor_key,designation,term_start,term_end,seniority,updated_at,display_name)) FROM (
  SELECT d.* FROM facility_staff_designations d JOIN wanted w ON w.source_type=d.source_type AND w.doctor_key=d.doctor_key
  WHERE d.active=1 AND d.designation<>'previous_staff' ORDER BY d.id LIMIT 257)) designations_json,
 (SELECT json_group_array(json_array(source_type,location)) FROM (SELECT * FROM account_hospital_locations WHERE email=o.email ORDER BY source_type LIMIT 6)) locations_json,
 (SELECT json_group_array(json_array(id,title,start_date,end_date,all_day,start_time,end_time,location,include)) FROM (
  SELECT * FROM custom_events WHERE owner_email=o.email ORDER BY id LIMIT 201)) custom_json
 FROM owner o`).bind(token,token).first();
 if(!row)return null;
 for(const [key,maximum] of [['claims_json',16],['aliases_json',64],['people_json',64],['files_json',1024],['designations_json',256],['locations_json',5],['custom_json',200]]){
  if(JSON.parse(row[key]).length>maximum)throw Object.assign(Error('Subscription evidence exceeds its safe bound.'),{code:'SUBSCRIPTION_ARTIFACT_UNSUPPORTED'});
 }
 const serialized=JSON.stringify([SUBSCRIPTION_ARTIFACT_VERSION,view,identitiesEnabled,row]);
 if(new TextEncoder().encode(serialized).length>256*1024)throw Object.assign(Error('Subscription evidence exceeds its safe byte bound.'),{code:'SUBSCRIPTION_ARTIFACT_UNSUPPORTED'});
 const revision=await subscriptionDigest(serialized),tokenKey=await subscriptionDigest(token);
 return {revision,key:`subscriptions/v1/${tokenKey}/${view}/${revision}.ics`,row};
}
export async function prepareSubscriptionExport(db,token,view,identitiesEnabled,evidence=null){
 evidence ||= await subscriptionEvidence(db,token,view,identitiesEnabled);
 if(!evidence)return null;
 const row=evidence.row;
 const record={email:row.email,realName:row.real_name,role:row.role,
  claims:JSON.parse(row.claims_json).map(([sourceType,key,displayName])=>({sourceType,key,displayName})),
  state:{session:JSON.parse(row.session_json)}};
 let claims=(record.claims||[]).map(c=>({...c,key:String(c.key||'').trim()})).filter(c=>c.key);
 if(identitiesEnabled){
  claims=['creator','owner'].includes(record.role)
   ?await expandApprovedIdentityAliases(db,sources.map(sourceType=>({sourceType,key:String(record.state?.session?.doctorKey||'').trim()})).filter(a=>a.key))
   :await accountIdentityAliases(db,record.email,claims);
 }else if(['creator','owner'].includes(record.role)){
  claims=sources.map(sourceType=>({sourceType,key:String(record.state?.session?.doctorKey||'').trim()})).filter(a=>a.key);
 }
 const aliases=[...new Map(claims.map(a=>[a.sourceType+'|'+a.key,{sourceType:a.sourceType,key:a.key}])).values()];
 if(aliases.length>16)throw Error('Calendar identity limit exceeded.');
 const session=record.state?.session||{};
 const range=view==='range'?session.exportRange||{}:{};
 const start=String(range.startDate||'0000-01-01'),end=range.allFuture===false?String(range.endDate||start):'9999-12-31';
 const locationRows=JSON.parse(row.locations_json);
 const locations=locationRows.length?normalizeHospitalLocationMap(Object.fromEntries(locationRows)):hospitalLocationsFromSession(session);
 const custom=JSON.parse(row.custom_json).map(([id,title,start_date,end_date,all_day,start_time,end_time,location,include])=>
  customEventForAccount({id,title,start_date,end_date,all_day,start_time,end_time,location,include},record.email));
 const allowed=new Set(aliases.map(a=>a.sourceType+'|'+a.key));
 const designations=JSON.parse(row.designations_json).map(([id,source_type,doctor_key,designation,term_start,term_end,seniority,updated_at,display_name])=>
  ({id,source_type,doctor_key,designation,term_start,term_end,seniority,updated_at,display_name,active:1}))
  .filter(d=>allowed.has(d.source_type+'|'+d.doctor_key)&&d.term_start<=end&&d.term_end>=start)
  .sort((a,b)=>a.term_start.localeCompare(b.term_start)||a.source_type.localeCompare(b.source_type)||a.display_name.localeCompare(b.display_name));
 const current=await subscriptionEvidence(db,token,view,identitiesEnabled);
 if(current?.revision!==evidence.revision)throw Error('Subscription inputs changed during export.');
 return {revision:evidence.revision,record,claims,aliases,start,end,inputs:{locations,custom,designations:designations.map(designationLeaveEventFromRow).filter(Boolean)}};
}
export async function subscriptionEventPage(db,aliases,start,end,cursor=null){
 if(!Array.isArray(aliases)||aliases.length>16||aliases.some(a=>!sources.includes(a.sourceType)||!a.key))throw Error('Invalid subscription aliases.');
 if(cursor&&(!Array.isArray(cursor)||cursor.length!==3||cursor.some(v=>typeof v!=='string')))throw Error('Invalid subscription cursor.');
 const after=cursor||['','',''];
 const result=await db.prepare(`WITH wanted AS (SELECT json_extract(value,'$.sourceType') source_type,json_extract(value,'$.key') doctor_key FROM json_each(?))
 SELECT e.file_id,e.doctor_key,e.id,e.event_json FROM wanted w
 CROSS JOIN roster_files f INDEXED BY idx_roster_files_source_active
 CROSS JOIN roster_events e INDEXED BY idx_roster_events_file_doctor
 WHERE f.source_type=w.source_type AND f.active=1 AND e.file_id=f.id AND e.doctor_key=w.doctor_key
 AND e.start_date<=? AND e.end_date>=? AND (e.file_id,e.doctor_key,e.id)>(?,?,?)
 ORDER BY e.file_id,e.doctor_key,e.id LIMIT 65`).bind(JSON.stringify(aliases),end,start,...after).all();
 const rows=result.results||[],page=rows.slice(0,64),last=page.at(-1);
 return {events:page.map(r=>JSON.parse(r.event_json)),cursor:rows.length>64?[last.file_id,last.doctor_key,last.id]:null};
}
export async function readSubscriptionArtifact(r2,evidence){
 const object=await r2.get(evidence.key);
 if(!object)return null;
 if(object.customMetadata?.revision!==evidence.revision||object.customMetadata?.version!==SUBSCRIPTION_ARTIFACT_VERSION)throw Error('Subscription artifact metadata is inconsistent.');
 return {body:object.body,displayName:object.customMetadata.displayName||'roster'};
}
export async function queueSubscriptionArtifact(r2,token,view,evidence){
 const key=`subscriptions/requests/${await subscriptionDigest(token)}/${view}/${evidence.revision}.json`;
 await r2.put(key,JSON.stringify({token,view,revision:evidence.revision}),{onlyIf:{etagDoesNotMatch:'*'},httpMetadata:{contentType:'application/json'}});
 return key;
}
