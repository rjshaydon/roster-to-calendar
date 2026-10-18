// Bounded exact ownership transaction; preparation never mutates a roster.
import {AuthenticationError} from './supabase-auth.js';
import {resolvePersonId} from './doctor-identity.js';
import {normalizeEmail} from './email.js';
const conflict=message=>new AuthenticationError('requests-conflict',message||'The roster links changed. Review this request again.',409);
const rows=async(db,sql,args=[])=>((await db.prepare(sql).bind(...args).all()).results||[]);
const guard=(db,condition,args=[])=>db.prepare(`SELECT CASE WHEN ${condition} THEN 1 ELSE json('request-conflict') END AS request_guard`).bind(...args);
const hash=async value=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value)))),x=>x.toString(16).padStart(2,'0')).join('');
const limited=(items,max)=>{if(items.length>max)throw conflict('This ownership scope is too large. Review fewer linked accounts first.');return items;};
async function ownershipScope(db,personId,targetEmail){
 const person=await resolvePersonId(db,personId);
 const snapshot=await db.batch([
  db.prepare('SELECT * FROM roster_person_aliases WHERE person_id=? ORDER BY source_type,doctor_key LIMIT 17').bind(person.person_id),
  db.prepare('SELECT * FROM account_people WHERE person_id=? ORDER BY email LIMIT 9').bind(person.person_id),
  db.prepare('SELECT * FROM account_people WHERE email=?').bind(targetEmail),
  db.prepare('SELECT * FROM account_claims WHERE email=? ORDER BY source_type,doctor_key LIMIT 17').bind(targetEmail),
  db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1')]);
 const aliases=limited(snapshot[0].results||[],16),accounts=limited(snapshot[1].results||[],8),targetLink=snapshot[2].results?.[0]||null,targetClaims=limited(snapshot[3].results||[],16),revision=snapshot[4].results[0].revision;
 if(!aliases.length||aliases.some(a=>a.review_state!=='approved'))throw conflict('Approve and review the roster names in Identities before assigning this person.');
 const predicate=aliases.map(()=>'(source_type=? AND doctor_key=?)').join(' OR ');
 const claimOwners=limited(await rows(db,`SELECT * FROM account_claims INDEXED BY idx_account_claims_source_doctor_email WHERE ${predicate} ORDER BY email,source_type,doctor_key LIMIT 65`,aliases.flatMap(a=>[a.source_type,a.doctor_key])),64);
 const emails=limited([...new Set([...accounts,...claimOwners,{email:targetEmail}].map(a=>a.email))].sort(),8);
 const placeholders=emails.map(()=>'?').join(',');
 const profiles=await rows(db,`SELECT email,real_name FROM account_profiles WHERE email IN (${placeholders}) ORDER BY email`,emails);
 if(profiles.length!==emails.length)throw conflict('An account no longer exists.');
 if(targetLink&&targetLink.person_id!==person.person_id)throw conflict('The receiving account is linked to another person. Review that identity first.');
 if(targetClaims.some(c=>!aliases.some(a=>a.source_type===c.source_type&&a.doctor_key===c.doctor_key)))throw conflict('The receiving account has other roster names. Review those links before replacing its identity.');
 return {people:[person],aliases,accounts,claimOwners,redirects:[],profiles,targetLink,targetClaims,revision};
}
export async function previewOwnershipTransfer(db,{personId,targetEmail,caseId,caseVersion}){
 targetEmail=normalizeEmail(targetEmail);if(!targetEmail||targetEmail.length>254)throw conflict('Choose an existing receiving account.');
 const before=await ownershipScope(db,personId,targetEmail);
 const owners=[...new Set([...before.accounts,...before.claimOwners].map(a=>a.email))].filter(e=>e!==targetEmail).sort();
 if(!owners.length&&before.targetLink&&before.targetClaims.length===before.aliases.length)throw conflict('This account already has all the reviewed roster links. Resolve the request without changing access.');
 const preview={kind:'account-transfer',personId:before.people[0].person_id,targetEmail,caseId,caseVersion,before,owners};
 return {...preview,previewToken:await hash(preview)};
}
// Pack read-only assertions into <=90-bind statements. All still execute
// before mutations in the same D1 transaction; no guard is deferred or removed.
function packedGuards(db,conditions){
 const result=[];let sql=[],args=[];
 const flush=()=>{if(sql.length)result.push(guard(db,sql.map(s=>'('+s+')').join(' AND '),args));sql=[];args=[];};
 for(const [condition,values]of conditions){if(args.length+values.length>90)flush();sql.push(condition);args.push(...values);}flush();return result;
}
function exactRows(table,items){return items.map(row=>[`EXISTS(SELECT 1 FROM ${table} WHERE ${Object.keys(row).map(k=>`${k} IS ?`).join(' AND ')})`,Object.values(row)]);}
function ownershipGuards(db,before,targetEmail){
 const id=before.people[0].person_id;
 const checks=[['(SELECT revision FROM roster_identity_revision WHERE id=1)=?',[before.revision]],
  ...exactRows('roster_people',before.people),...exactRows('roster_person_aliases',before.aliases),...exactRows('account_people',before.accounts),...exactRows('account_claims',before.claimOwners),
  ['(SELECT count(*) FROM roster_person_aliases WHERE person_id=?)=?',[id,before.aliases.length]],
  ['(SELECT count(*) FROM account_people WHERE person_id=?)=?',[id,before.accounts.length]],
  ['(SELECT count(*) FROM account_claims WHERE email=?)=?',[targetEmail,before.targetClaims.length]],...exactRows('account_claims',before.targetClaims)];
 if(!before.targetLink)checks.push(['NOT EXISTS(SELECT 1 FROM account_people WHERE email=?)',[targetEmail]]);
 for(const profile of before.profiles)checks.push(['EXISTS(SELECT 1 FROM account_profiles WHERE email=?)',[profile.email]]);
 for(const alias of before.aliases)checks.push(['(SELECT count(*) FROM account_claims WHERE source_type=? AND doctor_key=?)=?',
  [alias.source_type,alias.doctor_key,before.claimOwners.filter(c=>c.source_type===alias.source_type&&c.doctor_key===alias.doctor_key).length]]);
 return packedGuards(db,checks);
}
function aliasSelection(){return "(source_type,doctor_key) IN (SELECT json_extract(value,'$.source_type'),json_extract(value,'$.doctor_key') FROM json_each(?))";}
function invalidate(db,emails,operationId,now){
 const json=JSON.stringify(emails),scope="owner_type IN ('user-account','creator-account') AND owner_id IN (SELECT value FROM json_each(?))";
 return [guard(db,`NOT EXISTS(SELECT owner_id FROM snapshot_registry WHERE ${scope} GROUP BY owner_id HAVING count(*)>16)`,[json]),
 db.prepare(`UPDATE snapshot_registry SET status='stale',requested_revision=?,updated_at=? WHERE ${scope}`).bind('identity:'+operationId,now,json)];
}
export async function prepareOwnershipTransfer(db,input){
 const preview=await previewOwnershipTransfer(db,input);
 if(preview.previewToken!==input.previewToken)throw conflict();
 const reason=String(input.reason||'').trim();if(!reason||reason.length>1000)throw conflict('Explain the reviewed ownership decision.');
 const {before,targetEmail,personId}=preview,now=new Date().toISOString(),id=crypto.randomUUID();
 const statements=ownershipGuards(db,before,targetEmail);
 statements.push(db.prepare('DELETE FROM account_people WHERE person_id=? AND email<>?').bind(personId,targetEmail));
 statements.push(db.prepare(`INSERT INTO account_people(email,person_id,created_at,updated_at) VALUES(?,?,?,?)
  ON CONFLICT(email) DO UPDATE SET person_id=excluded.person_id,updated_at=excluded.updated_at`).bind(targetEmail,personId,now,now));
 statements.push(db.prepare(`DELETE FROM account_claims WHERE ${aliasSelection()} AND email<>?`).bind(JSON.stringify(before.aliases),targetEmail),
  db.prepare(`INSERT INTO account_claims(email,source_type,doctor_key,display_name,matched_at,updated_at)
   SELECT ?,json_extract(value,'$.source_type'),json_extract(value,'$.doctor_key'),json_extract(value,'$.display_name'),?,? FROM json_each(?) WHERE 1
   ON CONFLICT(email,source_type,doctor_key) DO NOTHING`).bind(targetEmail,now,now,JSON.stringify(before.aliases)));
 statements.push(db.prepare('UPDATE roster_people SET version=version+1,updated_at=?,approved_by=? WHERE person_id=?').bind(now,input.actor,personId),db.prepare('UPDATE roster_identity_revision SET revision=revision+1 WHERE id=1'));
 const after={...before,people:before.people.map(p=>({...p,version:p.version+1,updated_at:now,approved_by:input.actor})),
  accounts:[before.targetLink||{email:targetEmail,person_id:personId,created_at:now,updated_at:now}],
  claimOwners:before.aliases.map(a=>before.targetClaims.find(c=>c.source_type===a.source_type&&c.doctor_key===a.doctor_key)||{email:targetEmail,source_type:a.source_type,doctor_key:a.doctor_key,display_name:a.display_name,matched_at:now,updated_at:now})};
 // Preserve the actual target-link timestamp for dependency checking.
 after.accounts=after.accounts.map(a=>({...a,person_id:personId,updated_at:now}));
 const affected={personIds:[personId],accountEmails:before.profiles.map(p=>p.email),aliases:before.aliases.map(a=>({sourceType:a.source_type,key:a.doctor_key}))};
 statements.push(...invalidate(db,affected.accountEmails,id,now),
  db.prepare(`INSERT INTO roster_identity_operations(operation_id,kind,status,actor,reason,created_at,before_json,after_json,affected_json)
   VALUES(?,'account-transfer','committed',?,?,?,?,?,?)`).bind(id,input.actor,reason,now,JSON.stringify(before),JSON.stringify(after),JSON.stringify(affected)),
  db.prepare('INSERT INTO roster_identity_jobs(operation_id,affected_json,created_at) VALUES(?,?,?)').bind(id,JSON.stringify(affected),now),
  db.prepare('INSERT INTO roster_identity_operation_people(person_id,operation_id,created_at) VALUES(?,?,?)').bind(personId,id,now));
 return {statements,result:{operationId:id,affected,status:'committed'},preview};
}
export async function previewOwnershipReversal(db,{operationId,caseId,caseVersion}){
 const op=await db.prepare("SELECT * FROM roster_identity_operations WHERE operation_id=? AND kind='account-transfer' AND status='committed'").bind(operationId).first();
 if(!op)throw conflict('This ownership operation cannot be reversed.');
 const expected=JSON.parse(op.after_json),original=JSON.parse(op.before_json),targetEmail=expected.accounts[0].email;
 const current=await ownershipScope(db,expected.people[0].person_id,targetEmail);
 for(const key of ['people','aliases','accounts','claimOwners'])if(JSON.stringify(current[key])!==JSON.stringify(expected[key]))throw conflict('Later identity or claim changes depend on this transfer. Review them first.');
 const preview={kind:'transfer-reversal',operationId,caseId,caseVersion,before:current,after:original,affected:JSON.parse(op.affected_json),targetEmail};
 return {...preview,previewToken:await hash(preview)};
}
export async function prepareOwnershipReversal(db,input){
 const preview=await previewOwnershipReversal(db,input);if(preview.previewToken!==input.previewToken)throw conflict();
 const reason=String(input.reason||'').trim();if(!reason||reason.length>1000)throw conflict('Explain the reversal.');
 const id=crypto.randomUUID(),now=new Date().toISOString(),statements=ownershipGuards(db,preview.before,preview.targetEmail),original=preview.after;
 statements.push(guard(db,"NOT EXISTS(SELECT 1 FROM json_each(?) x LEFT JOIN account_profiles p ON p.email=json_extract(x.value,'$.email') WHERE p.email IS NULL)",[JSON.stringify(original.profiles)]),
  guard(db,"NOT EXISTS(SELECT 1 FROM json_each(?) x JOIN account_people p ON p.email=json_extract(x.value,'$.email') WHERE p.person_id<>?)",[JSON.stringify(original.accounts),original.people[0].person_id]),
  db.prepare(`DELETE FROM account_claims WHERE ${aliasSelection()}`).bind(JSON.stringify(original.aliases)),
  db.prepare('DELETE FROM account_people WHERE person_id=?').bind(original.people[0].person_id),
  db.prepare(`INSERT INTO account_people(email,person_id,created_at,updated_at)
   SELECT json_extract(value,'$.email'),json_extract(value,'$.person_id'),json_extract(value,'$.created_at'),? FROM json_each(?)`).bind(now,JSON.stringify(original.accounts)),
  db.prepare(`INSERT INTO account_claims(email,source_type,doctor_key,display_name,matched_at,updated_at)
   SELECT json_extract(value,'$.email'),json_extract(value,'$.source_type'),json_extract(value,'$.doctor_key'),json_extract(value,'$.display_name'),json_extract(value,'$.matched_at'),json_extract(value,'$.updated_at') FROM json_each(?)`).bind(JSON.stringify(original.claimOwners)));
 const person=original.people[0],restored={...person,version:preview.before.people[0].version+1,updated_at:now,approved_by:input.actor};
 statements.push(db.prepare('UPDATE roster_people SET version=?,updated_at=?,approved_by=? WHERE person_id=?').bind(restored.version,now,input.actor,person.person_id),db.prepare('UPDATE roster_identity_revision SET revision=revision+1 WHERE id=1'),
  ...invalidate(db,preview.affected.accountEmails,id,now),db.prepare("UPDATE roster_identity_operations SET status='reversed',reversed_by=? WHERE operation_id=? AND status='committed'").bind(id,input.operationId),
  db.prepare(`INSERT INTO roster_identity_operations(operation_id,kind,status,actor,reason,created_at,reversed_operation_id,before_json,after_json,affected_json)
   VALUES(?,'reverse','committed',?,?,?,?,?,?,?)`).bind(id,input.actor,reason,now,input.operationId,JSON.stringify(preview.before),JSON.stringify({...original,people:[restored]}),JSON.stringify(preview.affected)),
  db.prepare('INSERT INTO roster_identity_jobs(operation_id,affected_json,created_at) VALUES(?,?,?)').bind(id,JSON.stringify(preview.affected),now),
  db.prepare('INSERT INTO roster_identity_operation_people(person_id,operation_id,created_at) VALUES(?,?,?)').bind(person.person_id,id,now));
 return {statements,result:{operationId:id,affected:preview.affected,status:'committed'}};
}
