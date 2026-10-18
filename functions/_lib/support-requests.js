import {AuthenticationError} from './supabase-auth.js';
import {normalizeEmail} from './email.js';
import {resolvePersonId,queryIdentityPerson,previewIdentityOperation,prepareIdentityCommit,previewIdentityReversal,prepareIdentityReversal} from './doctor-identity.js';
import {previewOwnershipTransfer,prepareOwnershipTransfer,previewOwnershipReversal,prepareOwnershipReversal} from './request-ownership.js';
export {requestSummary} from './request-summary.js';

export const SUPPORT_ADMIN='rhaydon@gmail.com';
export const requestAdmin=session=>session.email===SUPPORT_ADMIN&&['owner','creator'].includes(session.role);
const uuid=/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const sources=new Set(['mmc','mch','ddh','vhh','casey']);
const fail=(message='This request changed. Open it again before continuing.',status=409)=>new AuthenticationError('requests-conflict',message,status);
const all=async(db,sql,args=[])=>((await db.prepare(sql).bind(...args).all()).results||[]);
const assertion=(db,condition,args=[])=>db.prepare(`SELECT CASE WHEN ${condition} THEN 1 ELSE json('request-conflict') END AS request_guard`).bind(...args);
const digest=async body=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(body)))),b=>b.toString(16).padStart(2,'0')).join('');
function text(value,max,required=true){if(typeof value!=='string'||value.length>max||required&&!value.trim()||/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value))throw fail('Enter a shorter, valid message.',400);return value.trim();}
function id(value){if(!uuid.test(value||''))throw fail('Invalid request reference.',400);return value;}
function admin(session){if(!requestAdmin(session))throw new AuthenticationError('requests-forbidden','Admin access is required.',403);}
export function requestAuthority(db,session,requireAdmin=false){
 const now=Date.now();return assertion(db,`EXISTS(SELECT 1 FROM account_sessions s JOIN account_profiles a ON a.email=s.email WHERE s.token_hash=? AND s.email=? AND s.idle_expires_at>? AND s.absolute_expires_at>?${requireAdmin?" AND a.email='rhaydon@gmail.com' AND a.role IN ('creator','owner')":''})`,[session.token_hash,session.email,now,now]);
}
async function mutation(db,session,input){
 id(input.mutationId);const fingerprint=await digest(input),old=await db.prepare('SELECT digest,result_json FROM support_mutations WHERE actor_email=? AND mutation_id=?').bind(session.email,input.mutationId).first();
 if(old){if(old.digest!==fingerprint)throw fail('That action has already been used. Reopen the request to check its result.');return {replay:JSON.parse(old.result_json)};}
 return {fingerprint};
}
async function commit(db,session,input,prepared,statements,result,requireAdmin=false){
 try{await db.batch([requestAuthority(db,session,requireAdmin||requestAdmin(session)),db.prepare('INSERT INTO support_mutations(actor_email,mutation_id,digest,result_json,created_at) VALUES(?,?,?,?,?)')
  .bind(session.email,input.mutationId,prepared.fingerprint,JSON.stringify(result),new Date().toISOString()),...statements]);}
 catch(error){
  const retry=await mutation(db,session,input);if(retry.replay)return retry.replay;
  if(/malformed JSON|request-conflict|UNIQUE constraint/i.test(error.message||''))throw fail();throw error;
 }
 return result;
}
function event(db,caseId,session,kind,note='',operationId=''){return db.prepare('INSERT INTO support_case_events(case_id,event_id,kind,actor,note,operation_id,created_at) VALUES(?,?,?,?,?,?,?)')
 .bind(caseId,crypto.randomUUID(),kind,session.email,note,operationId,new Date().toISOString());}
async function loadCase(db,session,caseId){
 const row=await db.prepare('SELECT * FROM support_cases WHERE case_id=?').bind(id(caseId)).first();
 if(!row)throw new AuthenticationError('requests-not-found','Request not found.',404);
 if(!requestAdmin(session)&&!await db.prepare('SELECT channel_id FROM support_channels WHERE case_id=? AND participant_email=?').bind(caseId,session.email).first())throw new AuthenticationError('requests-not-found','Request not found.',404);
 return row;
}
async function channel(db,session,caseId,channelId){
 const row=await db.prepare(`SELECT * FROM support_channels WHERE case_id=? AND ${requestAdmin(session)&&channelId?'channel_id':'participant_email'}=?`)
  .bind(caseId,requestAdmin(session)&&channelId?id(channelId):session.email).first();
 if(!row)throw new AuthenticationError('requests-not-found','Conversation not found.',404);return row;
}
async function context(db,input){
 let personId='',alias=null;
 if(input.alias){
  if(!sources.has(input.alias.sourceType))throw fail('Unknown roster site.',400);
  const key=text(input.alias.key,200);const row=await db.prepare('SELECT source_type,doctor_key,display_name,person_id FROM roster_person_aliases WHERE source_type=? AND doctor_key=?').bind(input.alias.sourceType,key).first();
  if(!row)throw fail('That roster name is not in the identity directory. Describe the problem without selecting it.',400);
  personId=(await resolvePersonId(db,row.person_id)).person_id;alias={sourceType:row.source_type,key:row.doctor_key,displayName:row.display_name};
 }else if(input.personId)personId=(await resolvePersonId(db,text(input.personId,160))).person_id;
 return {personId,alias};
}
export async function inspectSupportCreation(db,session,input,{legacy}={}){
 const prepared=await mutation(db,session,input);if(prepared.replay)return prepared.replay;
 const kind=legacy?'identity':input.kind;if(!['identity','support'].includes(kind))throw fail('Choose a request type.',400);
 const subject=text(legacy?.subject||input.subject,120),body=text(legacy?.body||input.body,4000),details=legacy?{personId:'',alias:null}:await context(db,input);
 const participant=legacy?.email||session.email;if(legacy)admin(session);
 if(kind==='identity'&&details.personId){
  const existing=await db.prepare("SELECT c.case_id,h.channel_id FROM support_cases c INDEXED BY idx_support_identity_open JOIN support_channels h ON h.case_id=c.case_id AND h.participant_email=c.requester_email WHERE c.requester_email=? AND c.person_id=? AND c.person_id<>'' AND c.kind='identity' AND c.status='open'").bind(participant,details.personId).first();
  if(existing)return {ok:true,caseId:existing.case_id,channelId:existing.channel_id,alreadyOpen:true};
 }
 return {prepared,kind,subject,body,details,participant};
}
export async function createSupportRequest(db,session,input,{legacy,review}={}){
 const creation=review||await inspectSupportCreation(db,session,input,{legacy});if(creation.ok)return creation;
 const {prepared,kind,subject,body,details,participant}=creation;
 const caseId=crypto.randomUUID(),channelId=crypto.randomUUID(),now=new Date().toISOString(),isAdmin=requestAdmin(session)&&!legacy;
 const result={ok:true,caseId,channelId};
 const statements=[db.prepare(`INSERT INTO support_cases(case_id,requester_email,kind,subject,person_id,context_json,created_at,updated_at,legacy_reference) VALUES(?,?,?,?,?,?,?,?,?)`)
  .bind(caseId,participant,kind,subject,details.personId,JSON.stringify(details),now,now,legacy?.reference||null),
  db.prepare(`INSERT INTO support_channels(channel_id,case_id,participant_email,last_seq,participant_read_seq,admin_read_seq,waiting,updated_at) VALUES(?,?,?,1,?,?,?,?)`)
   .bind(channelId,caseId,participant,isAdmin||legacy?0:1,isAdmin?1:0,isAdmin?'participant':'admin',now),
  db.prepare('INSERT INTO support_messages(channel_id,sequence,author_type,body,created_at) VALUES(?,1,?,?,?)').bind(channelId,legacy?'system':isAdmin?'admin':'participant',body,now),event(db,caseId,session,legacy?'legacy-import':'created')];
 return commit(db,session,input,prepared,statements,result,Boolean(legacy));
}
export async function supportMutationOutcome(db,session,input){
 const row=await db.prepare('SELECT result_json FROM support_mutations WHERE actor_email=? AND mutation_id=?').bind(session.email,id(input.mutationId)).first();
 if(!row)return {ok:true,found:false};
 const result=JSON.parse(row.result_json);if(result.operationId)admin(session);
 return {ok:true,found:true,result};
}
function cursor(value){if(!value)return null;try{const pair=JSON.parse(value);if(!Array.isArray(pair)||pair.length!==2||!uuid.test(pair[1])||typeof pair[0]!=='string'||pair[0].length>40)throw Error();return pair;}catch{throw fail('Invalid page cursor.',400);}}
export async function listSupportRequests(db,session,input={}){
 const status=input.status==='resolved'?'resolved':'open',after=cursor(input.cursor),filter=['unread','waiting'].includes(input.filter)?input.filter:'';
 let rows;
 if(requestAdmin(session))rows=await all(db,`SELECT case_id,subject,kind,status,version,updated_at,requester_email,has_admin_unread AS unread,waiting_admin
  FROM support_cases WHERE status=? ${filter==='unread'?'AND has_admin_unread=1':filter==='waiting'?'AND waiting_admin=1':''}
  ${after?'AND (updated_at,case_id)<(?,?)':''} ORDER BY updated_at DESC,case_id DESC LIMIT 21`,[status,...(after||[])]);
 else rows=await all(db,`SELECT c.case_id,CASE WHEN c.requester_email=? THEN c.subject ELSE 'Roster identity review' END AS subject,c.kind,c.status,c.version,h.updated_at,h.channel_id,h.participant_unread AS unread,h.waiting
  FROM support_channels h JOIN support_cases c ON c.case_id=h.case_id WHERE h.participant_email=? AND h.case_status=? ${filter==='unread'?'AND h.participant_unread=1':filter==='waiting'?"AND h.waiting='admin'":''}
  ${after?'AND (h.updated_at,h.channel_id)<(?,?)':''} ORDER BY h.updated_at DESC,h.channel_id DESC LIMIT 21`,[session.email,session.email,status,...(after||[])]);
 const items=rows.slice(0,20),last=items.at(-1);return {ok:true,items,nextCursor:rows.length>20?JSON.stringify([last.updated_at,requestAdmin(session)?last.case_id:last.channel_id]):''};
}
export async function readSupportRequest(db,session,input){
 const c=await loadCase(db,session,input.caseId);
 const channels=requestAdmin(session)?await all(db,`SELECT channel_id,participant_email,case_status,waiting,last_seq,participant_read_seq,admin_read_seq,version,updated_at FROM support_channels WHERE case_id=? ORDER BY channel_id LIMIT 9`,[c.case_id]):[];
 if(channels.length>8)throw fail('Too many private channels; narrow this case.');
 const selected=await channel(db,session,c.case_id,input.channelId||channels[0]?.channel_id);
 const revision=c.version+':'+selected.version;
 if(input.cachedRevision===revision&&!input.before)return {ok:true,unchanged:true,revision};
 const before=input.before===undefined?Number.MAX_SAFE_INTEGER:Number(input.before);if(!Number.isSafeInteger(before)||before<1)throw fail('Invalid message page.',400);
 const messages=await all(db,'SELECT sequence,author_type,body,created_at FROM support_messages WHERE channel_id=? AND sequence<? ORDER BY sequence DESC LIMIT 21',[selected.channel_id,before]);
 const result={ok:true,revision,case:{caseId:c.case_id,kind:c.kind,subject:requestAdmin(session)||c.requester_email===session.email?c.subject:'Roster identity review',status:c.status,version:c.version,personId:c.person_id,context:JSON.parse(c.context_json),operationId:c.resolution_operation_id},
  channel:{channelId:selected.channel_id,lastSeq:selected.last_seq,waiting:selected.waiting,readSeq:requestAdmin(session)?selected.admin_read_seq:selected.participant_read_seq},
  messages:messages.slice(0,20).reverse(),before:messages.length>20?messages[19].sequence:0};
 if(requestAdmin(session)){
  result.case.requesterEmail=c.requester_email;result.channels=channels;
  result.history=await all(db,'SELECT kind,note,operation_id,created_at FROM support_case_events WHERE case_id=? ORDER BY created_at DESC,event_id DESC LIMIT 20',[c.case_id]);
 }
 return result;
}
export async function markSupportRead(db,session,input,{admitWrite}={}){
 const c=await loadCase(db,session,input.caseId),h=await channel(db,session,c.case_id,input.channelId);
 const sequence=Number(input.sequence);if(!Number.isSafeInteger(sequence)||sequence<0||sequence>h.last_seq)throw fail('Invalid read position.',400);
 const field=requestAdmin(session)?'admin_read_seq':'participant_read_seq';if(h[field]>=sequence)return {ok:true,unchanged:true};
 if(admitWrite&&!await admitWrite())return {ok:true,readDeferred:true};
 await db.batch([requestAuthority(db,session,requestAdmin(session)),db.prepare(`UPDATE support_channels SET ${field}=max(${field},?),version=version+1 WHERE channel_id=? AND ${field}<? AND last_seq>=?`).bind(sequence,h.channel_id,sequence,sequence)]);
 return {ok:true};
}
function caseGuard(db,c){return assertion(db,'EXISTS(SELECT 1 FROM support_cases WHERE case_id=? AND version=? AND requester_email=? AND person_id=? AND status=?)',[c.case_id,c.version,c.requester_email,c.person_id,c.status]);}
function appendMessage(db,h,body,author,now){
 const seq=h.last_seq+1;return [assertion(db,'EXISTS(SELECT 1 FROM support_channels WHERE channel_id=? AND last_seq=? AND participant_email=?)',[h.channel_id,h.last_seq,h.participant_email]),
 db.prepare('INSERT INTO support_messages(channel_id,sequence,author_type,body,created_at) VALUES(?,?,?,?,?)').bind(h.channel_id,seq,author,body,now),
 db.prepare(`UPDATE support_channels SET last_seq=?,${author==='admin'?'admin_read_seq':'participant_read_seq'}=?,waiting=?,version=version+1,updated_at=? WHERE channel_id=?`)
 .bind(seq,seq,author==='admin'?'participant':'admin',now,h.channel_id)];
}
export async function replySupportRequest(db,session,input){
 const c=await loadCase(db,session,input.caseId),h=await channel(db,session,c.case_id,input.channelId),prepared=await mutation(db,session,input);if(prepared.replay)return prepared.replay;
 if(c.status!=='open')throw fail('Reopen this request before replying.');
 const body=text(input.body,4000),now=new Date().toISOString();
 return commit(db,session,input,prepared,[caseGuard(db,c),...appendMessage(db,h,body,requestAdmin(session)?'admin':'participant',now),db.prepare('UPDATE support_cases SET version=version+1,updated_at=? WHERE case_id=?').bind(now,c.case_id)],{ok:true,caseId:c.case_id});
}
function decisionStatements(db,c,session,channels,{status,body,operationId='',kind='resolved'}){
 const now=new Date().toISOString(),scope=JSON.stringify(channels.map(h=>({id:h.channel_id,seq:h.last_seq,email:h.participant_email}))),statements=[caseGuard(db,c),
  assertion(db,`(SELECT count(*) FROM support_channels WHERE case_id=?)=? AND NOT EXISTS(
   SELECT 1 FROM json_each(?) x LEFT JOIN support_channels h ON h.channel_id=json_extract(x.value,'$.id')
   WHERE h.channel_id IS NULL OR h.case_id<>? OR h.last_seq<>json_extract(x.value,'$.seq') OR h.participant_email<>json_extract(x.value,'$.email'))`,[c.case_id,channels.length,scope,c.case_id]),
  db.prepare("INSERT INTO support_messages(channel_id,sequence,author_type,body,created_at) SELECT json_extract(value,'$.id'),json_extract(value,'$.seq')+1,'admin',?,? FROM json_each(?)").bind(body,now,scope)];
 statements.push(db.prepare("UPDATE support_channels SET last_seq=last_seq+1,admin_read_seq=last_seq+1,case_status=?,waiting=?,version=version+1,updated_at=? WHERE case_id=?").bind(status,status==='resolved'?'none':'admin',now,c.case_id),
  db.prepare('UPDATE support_cases SET status=?,version=version+1,updated_at=?,resolution_operation_id=? WHERE case_id=?').bind(status,now,operationId||c.resolution_operation_id,c.case_id),event(db,c.case_id,session,kind,body,operationId));
 return statements;
}
export async function setSupportStatus(db,session,input){
 const c=await loadCase(db,session,input.caseId),prepared=await mutation(db,session,input);if(prepared.replay)return prepared.replay;
 const status=input.status;if(!['open','resolved'].includes(status))throw fail('Invalid request status.',400);
 if(status==='resolved')admin(session);
 if(c.version!==input.version)throw fail();
 if(c.status===status)return commit(db,session,input,prepared,[],{ok:true,caseId:c.case_id,unchanged:true},status==='resolved');
 const body=text(input.body|| (status==='open'?'Request reopened for admin review.':''),1000),channels=await all(db,'SELECT * FROM support_channels WHERE case_id=? ORDER BY channel_id LIMIT 9',[c.case_id]);
 if(channels.length>8)throw fail('Too many private channels.');
 const statements=requestAdmin(session)?decisionStatements(db,c,session,channels,{status,body,kind:status==='open'?'reopened':'resolved'}):
 [caseGuard(db,c),...appendMessage(db,channels.find(h=>h.participant_email===session.email),body,'participant',new Date().toISOString()),
 db.prepare("UPDATE support_channels SET case_status='open',waiting='admin',version=version+1 WHERE case_id=?").bind(c.case_id),db.prepare("UPDATE support_cases SET status='open',version=version+1,updated_at=? WHERE case_id=?").bind(new Date().toISOString(),c.case_id),event(db,c.case_id,session,'reopened')];
 return commit(db,session,input,prepared,statements,{ok:true,caseId:c.case_id},status==='resolved');
}
export async function supportIdentityEvidence(db,session,input){
 admin(session);const c=await loadCase(db,session,input.caseId);if(!c.person_id)return {ok:true,identity:null};
 const person=await resolvePersonId(db,c.person_id);return {ok:true,identity:await queryIdentityPerson(db,person.person_id)};
}
export async function attachSupportIdentity(db,session,input){
 admin(session);const c=await loadCase(db,session,input.caseId),prepared=await mutation(db,session,input);if(prepared.replay)return prepared.replay;
 if(c.status!=='open'||c.version!==input.version)throw fail();
 const person=await resolvePersonId(db,text(input.personId,160)),now=new Date().toISOString();
 return commit(db,session,input,prepared,[caseGuard(db,c),assertion(db,"EXISTS(SELECT 1 FROM roster_people WHERE person_id=? AND version=? AND status='active')",[person.person_id,person.version]),
  db.prepare('UPDATE support_cases SET person_id=?,context_json=?,version=version+1,updated_at=? WHERE case_id=?').bind(person.person_id,JSON.stringify({personId:person.person_id,alias:null,reviewedBy:session.email}),now,c.case_id),event(db,c.case_id,session,'identity-context-reviewed',person.preferred_display_name)],{ok:true,caseId:c.case_id},true);
}
async function reviewedOperation(db,c,input){
 if(!c.person_id)throw fail('Attach a reviewed roster identity before applying an identity change.');
 const person=await resolvePersonId(db,c.person_id);
 if(!['merge','name','id','alias-move'].includes(input.operation?.kind)||!input.operation?.personIds?.includes(person.person_id))throw fail('Review this case’s person. Use account reassignment to change account access.');
}
function reviewedRecipient(c,input){
 const target=normalizeEmail(input.targetEmail||c.requester_email);
 if(!c.requester_email||target!==c.requester_email)throw fail('This request can assign access only to its requester. Review a separate request for another receiving account.');return target;
}
async function saveOwnershipPreview(db,session,c,preview){
 const previewId=crypto.randomUUID(),now=Date.now(),payload=JSON.stringify(preview);
 if(new TextEncoder().encode(payload).length>196608)throw fail('This ownership review exceeds its safe evidence bound.');
 await db.batch([requestAuthority(db,session,true),
  db.prepare('INSERT INTO support_ownership_previews(preview_id,actor_email,case_id,case_version,expires_at,payload_json) VALUES(?,?,?,?,?,?)').bind(previewId,session.email,c.case_id,c.version,now+300000,payload),
  db.prepare('DELETE FROM support_ownership_previews WHERE preview_id IN (SELECT preview_id FROM support_ownership_previews WHERE actor_email=? AND expires_at<=? ORDER BY expires_at,preview_id LIMIT 10)').bind(session.email,now)]);
 // The commit snapshot stays private on the server. The review needs names,
 // sites and the affected accounts, not every legacy claim and timestamp.
 return {kind:preview.kind,personId:preview.personId||preview.before.people[0].person_id,targetEmail:preview.targetEmail,
  ...(preview.owners?{owners:preview.owners}:{}),before:{people:preview.before.people.map(p=>({preferred_display_name:p.preferred_display_name})),aliases:preview.before.aliases.map(a=>({source_type:a.source_type,display_name:a.display_name||a.doctor_key}))},previewToken:'prepared:'+previewId};
}
async function loadOwnershipPreview(db,session,c,input,kind){
 const previewId=id(String(input.previewToken||'').replace(/^prepared:/,''));
 const saved=await db.prepare('SELECT payload_json FROM support_ownership_previews WHERE preview_id=? AND actor_email=? AND case_id=? AND case_version=? AND expires_at>?').bind(previewId,session.email,c.case_id,c.version,Date.now()).first();
 if(!saved)throw fail('This review expired or changed. Review the ownership decision again.');
 const preview=JSON.parse(saved.payload_json);
 if(preview.kind!==kind||preview.caseId!==c.case_id||preview.caseVersion!==c.version||kind==='account-transfer'&&preview.targetEmail!==reviewedRecipient(c,input)||kind==='transfer-reversal'&&preview.operationId!==c.resolution_operation_id)throw fail();
 return {preview,previewId};
}
export async function previewSupportResolution(db,session,input){
 admin(session);if(!['identity','transfer','reverse'].includes(input.kind))throw fail('Choose a supported identity decision.',400);const c=await loadCase(db,session,input.caseId);if(c.version!==input.version)throw fail();
 if(input.kind==='transfer')return {ok:true,preview:await saveOwnershipPreview(db,session,c,await previewOwnershipTransfer(db,{personId:c.person_id,targetEmail:reviewedRecipient(c,input),caseId:c.case_id,caseVersion:c.version}))};
 if(input.kind==='reverse'){
  const op=await db.prepare('SELECT kind FROM roster_identity_operations WHERE operation_id=?').bind(c.resolution_operation_id).first();
  if(!op)throw fail('There is no identity operation to reverse.');
  return {ok:true,preview:op.kind==='account-transfer'?await saveOwnershipPreview(db,session,c,await previewOwnershipReversal(db,{operationId:c.resolution_operation_id,caseId:c.case_id,caseVersion:c.version})):await previewIdentityReversal(db,c.resolution_operation_id,{caseId:c.case_id,version:c.version})};
 }
 await reviewedOperation(db,c,input);
 return {ok:true,preview:await previewIdentityOperation(db,{...input.operation,actor:session.email,requestContext:{caseId:c.case_id,version:c.version}})};
}
export async function commitSupportResolution(db,session,input){
 admin(session);if(!['identity','transfer','reverse'].includes(input.kind))throw fail('Choose a supported identity decision.',400);const c=await loadCase(db,session,input.caseId),prepared=await mutation(db,session,input);if(prepared.replay)return prepared.replay;
 if(c.version!==input.version||input.kind!=='reverse'&&c.status!=='open')throw fail();
 const reason=text(input.reason,1000),options={...input.operation,previewToken:input.previewToken,reason,actor:session.email,confirmAccountEmails:input.confirmAccountEmails,requestContext:{caseId:c.case_id,version:c.version}};let identity,ownership;
 if(input.kind==='transfer'){
  ownership=await loadOwnershipPreview(db,session,c,input,'account-transfer');
  identity=await prepareOwnershipTransfer(db,{...options,previewToken:ownership.preview.previewToken,personId:c.person_id,targetEmail:reviewedRecipient(c,input),caseId:c.case_id,caseVersion:c.version},ownership.preview);
 }
 else if(input.kind==='reverse'){
  const op=await db.prepare('SELECT kind FROM roster_identity_operations WHERE operation_id=?').bind(c.resolution_operation_id).first();
  if(op?.kind==='account-transfer'){
   ownership=await loadOwnershipPreview(db,session,c,input,'transfer-reversal');
   identity=await prepareOwnershipReversal(db,{...options,previewToken:ownership.preview.previewToken,operationId:c.resolution_operation_id,caseId:c.case_id,caseVersion:c.version},ownership.preview);
  }else identity=await prepareIdentityReversal(db,{...options,operationId:c.resolution_operation_id});
 }else{await reviewedOperation(db,c,input);identity=await prepareIdentityCommit(db,options);}
 const channels=await all(db,'SELECT * FROM support_channels WHERE case_id=? ORDER BY channel_id LIMIT 9',[c.case_id]);if(channels.length>8)throw fail('Too many private channels.');
 const additions=[];
 // A reviewed reassignment notifies each displaced holder privately even if
 // no earlier question was needed. Never forward the requester's explanation.
 if(input.kind==='transfer')for(const recipient of identity.preview.owners){
  if(channels.some(h=>h.participant_email===recipient))continue;
  if(channels.length>=8)throw fail('Too many private participants; narrow this case before changing access.');
  const h={channel_id:crypto.randomUUID(),case_id:c.case_id,participant_email:recipient,last_seq:0};channels.push(h);
  additions.push(db.prepare("INSERT INTO support_channels(channel_id,case_id,participant_email,waiting,updated_at) VALUES(?,?,?,'none',?)").bind(h.channel_id,c.case_id,recipient,new Date().toISOString()));
 }
 const statements=[...identity.statements,...additions,...decisionStatements(db,c,session,channels,{status:input.kind==='reverse'?'open':'resolved',body:reason,operationId:identity.result.operationId,kind:input.kind==='reverse'?'reversed':'identity-resolved'})];
 if(ownership)statements.unshift(assertion(db,'EXISTS(SELECT 1 FROM support_ownership_previews WHERE preview_id=? AND actor_email=? AND case_id=? AND case_version=? AND expires_at>?)',[ownership.previewId,session.email,c.case_id,c.version,Date.now()]),db.prepare('DELETE FROM support_ownership_previews WHERE preview_id=?').bind(ownership.previewId));
 if(statements.length>440)throw fail('Narrow this identity change before applying it.');
 return commit(db,session,input,prepared,statements,{ok:true,caseId:c.case_id,...identity.result},true);
}
export async function inviteSupportHolder(db,session,input){
 admin(session);const c=await loadCase(db,session,input.caseId),prepared=await mutation(db,session,input);if(prepared.replay)return prepared.replay;
 if(c.version!==input.version||!c.person_id||c.status!=='open')throw fail();
 const targetEmail=normalizeEmail(input.participantEmail),identity=await queryIdentityPerson(db,(await resolvePersonId(db,c.person_id)).person_id);
 if(![...identity.accounts,...identity.claimOwners].some(a=>a.email===targetEmail))throw fail('Choose a currently linked account from the reviewed identity.');
 const channels=await all(db,'SELECT channel_id FROM support_channels WHERE case_id=? LIMIT 9',[c.case_id]);if(channels.length>=8)throw fail('Too many private channels.');
 const existing=await db.prepare('SELECT channel_id FROM support_channels WHERE case_id=? AND participant_email=?').bind(c.case_id,targetEmail).first();
 if(existing)return commit(db,session,input,prepared,[],{ok:true,caseId:c.case_id,channelId:existing.channel_id},true);
 const body=text(input.body,4000),channelId=crypto.randomUUID(),now=new Date().toISOString();
 const statements=[caseGuard(db,c),assertion(db,`EXISTS(SELECT 1 FROM account_people WHERE email=? AND person_id=?) OR EXISTS(SELECT 1 FROM account_claims a JOIN roster_person_aliases r ON r.source_type=a.source_type AND r.doctor_key=a.doctor_key WHERE a.email=? AND r.person_id=?)`,[targetEmail,identity.people[0].person_id,targetEmail,identity.people[0].person_id]),
 db.prepare(`INSERT INTO support_channels(channel_id,case_id,participant_email,last_seq,admin_read_seq,waiting,updated_at) VALUES(?,?,?,1,1,'participant',?)`).bind(channelId,c.case_id,targetEmail,now),
 db.prepare("INSERT INTO support_messages(channel_id,sequence,author_type,body,created_at) VALUES(?,1,'admin',?,?)").bind(channelId,body,now),
 db.prepare('UPDATE support_cases SET version=version+1,updated_at=? WHERE case_id=?').bind(now,c.case_id),event(db,c.case_id,session,'private-review-started')];
 return commit(db,session,input,prepared,statements,{ok:true,caseId:c.case_id,channelId},true);
}
export async function legacySupportReports(db,session,input){
 admin(session);const after=text(input.after||'',254,false),profiles=await all(db,"SELECT email,real_name,admin_issues_json FROM account_profiles WHERE email>? ORDER BY email LIMIT 21",[after]);
 const reports=[];for(const profile of profiles.slice(0,20))for(const issue of JSON.parse(profile.admin_issues_json||'[]'))if(String(issue.id||'').startsWith('identity:'))reports.push({email:profile.email,name:profile.real_name,id:issue.id,message:String(issue.message||'').slice(0,4000),createdAt:issue.firstSeenAt||''});
 return {ok:true,reports:reports.slice(0,100),next:profiles.length>20?profiles[19].email:''};
}
export async function convertLegacyReport(db,session,input){
 admin(session);const profile=await db.prepare('SELECT email,admin_issues_json FROM account_profiles WHERE email=?').bind(normalizeEmail(input.email)).first();
 const issue=profile&&JSON.parse(profile.admin_issues_json||'[]').find(i=>i.id===input.reportId&&String(i.id).startsWith('identity:'));
 if(!issue)throw fail('Earlier report not found.',404);
 const reference=await digest([profile.email,issue.id]),existing=await db.prepare('SELECT case_id FROM support_cases WHERE legacy_reference=?').bind(reference).first();
 if(existing)return {ok:true,caseId:existing.case_id};
 return createSupportRequest(db,session,input,{legacy:{email:profile.email,reference,subject:'Earlier identity report',body:('Earlier report ('+String(issue.firstSeenAt||'date not recorded').slice(0,40)+'):\n'+text(issue.message||'Roster identity needs review.',3900)).slice(0,4000)}});
}
