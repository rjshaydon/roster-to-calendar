import {identitySuggestionProposalStatement} from './identity-suggestion-proposal.js';
import {reserveRosterMaintenanceBudget} from './roster-maintenance-budget.js';
import {activeRosterImportsQuery} from './roster-delivery-health.js';
import {doctorIdentityPairKey} from '../../public/static/doctor-identity.js';
const invalid=message=>Object.assign(Error(message),{code:'IDENTITY_AUDIT_EVIDENCE'});
const sources=new Set(['mmc','mch','ddh','vhh','casey']);
const hex=v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v);
const text=(v,max)=>typeof v==='string'&&v.length>0&&v.length<=max;
async function digest(value){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value))))].map(v=>v.toString(16).padStart(2,'0')).join('');}
const priority=activeRosterImportsQuery(['monash-adults','monash-paeds','vhh-active-medical-roster','dandenong-findmyshift']);
// The authenticated endpoint owns its lease and meters optional identity work.
// These functions also reserve costs and authenticate pinned evidence, then
// guard source priority, current evidence and the cursor in each transaction.
function identityGuard(db,revision){return db.prepare(`SELECT CASE WHEN (SELECT revision FROM roster_identity_revision WHERE id=1)=?
  AND NOT EXISTS(${priority.sql}) THEN 1 ELSE json('external-audit-stale-or-import-active') END`).bind(revision,...priority.bindings);}
function featureGuard(db,features){
 const json=JSON.stringify(features);
 return db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM json_each(?) j
 JOIN roster_identity_features f ON f.source_type=json_extract(j.value,'$.source') AND f.doctor_key=json_extract(j.value,'$.key')
 JOIN roster_person_aliases a ON a.source_type=f.source_type AND a.doctor_key=f.doctor_key
 JOIN roster_people p ON p.person_id=a.person_id WHERE f.fingerprint=json_extract(j.value,'$.fingerprint') AND p.status='active'
 AND (json_extract(j.value,'$.personId') IS NULL OR a.person_id=json_extract(j.value,'$.personId') AND f.display_name=json_extract(j.value,'$.displayName')))=?
 THEN 1 ELSE json('external-audit-feature-conflict') END`).bind(json,features.length);
}
function validateManifest(m,bytes){
 if(!m || m.schemaVersion!==1 || m.algorithmVersion!=='external-pairs-v1' || typeof m.fullAudit!=='boolean'
  || !Number.isSafeInteger(m.identityRevision) || m.identityRevision<0 || !Number.isInteger(m.nameCount)||m.nameCount<0||m.nameCount>2000
  || !Number.isInteger(m.pendingCount)||m.pendingCount<0||m.pendingCount>m.nameCount
  || m.fullAudit&&m.pendingCount!==m.nameCount
  || !Number.isInteger(m.candidateCount)||m.candidateCount<0||m.candidateCount>10000||!hex(m.snapshotHash)
  || bytes.length>262144)throw invalid('Invalid external identity manifest.');
 for(const [field,total,limit,max]of[['batches',m.candidateCount,5,2000],['featureBatches',m.pendingCount,16,125]]){
  const list=m[field];if(!Array.isArray(list)||list.length>max)throw invalid('Invalid identity descriptors.');
  let count=0;for(const [i,b]of list.entries()){if(b.index!==i||!hex(b.hash)||!Number.isInteger(b.count)||b.count<1||b.count>limit)throw invalid('Invalid identity descriptor.');count+=b.count;}
  if(count!==total)throw invalid('Incomplete identity descriptor counts.');
 }
}
async function jobFor(db,runId,revision,kind='candidates',index=0){
 const path=`$.${kind==='candidates'?'batches':'featureBatches'}[${Number.isInteger(index)&&index>=0?index:0}]`;
 const job=await db.prepare(`SELECT run_id,plan_revision,identity_revision,next_candidate_batch,next_feature_batch,status,
 json_array_length(manifest_json,'$.batches') AS candidate_batches,json_array_length(manifest_json,'$.featureBatches') AS feature_batches,
 json_extract(manifest_json,'$.pendingCount') AS pending_count,json_extract(manifest_json,'$.nameCount') AS name_count,json_extract(manifest_json,?) AS descriptor
 FROM external_identity_audit_jobs WHERE run_id=?`).bind(path,runId).first();
 if(!job||job.plan_revision!==revision)throw invalid('Unpinned or changed identity audit.');return job;
}
export async function pinExternalIdentityAudit(db,{runId,manifest,revision}){
 if(!text(runId,100)||!hex(revision))throw invalid('Invalid identity audit identity.');const serialized=JSON.stringify(manifest),bytes=new TextEncoder().encode(serialized);validateManifest(manifest,bytes);
 const pinnedHash=[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(v=>v.toString(16).padStart(2,'0')).join('');
 if(pinnedHash!==revision)throw invalid('Identity manifest hash mismatch.');
 const old=await db.prepare('SELECT plan_revision,next_candidate_batch,next_feature_batch,status FROM external_identity_audit_jobs WHERE run_id=?').bind(runId).first();
 if(old){if(old.plan_revision!==revision)throw invalid('Audit already pinned to different evidence.');return progress(old);}
 if(!await reserveRosterMaintenanceBudget(db,64,65536))return {deferred:true};
 const now=new Date().toISOString();
 await db.batch([identityGuard(db,manifest.identityRevision),
  db.prepare("SELECT CASE WHEN (SELECT COUNT(*) FROM (SELECT 1 FROM roster_identity_features LIMIT 2001))=? THEN 1 ELSE json('external-audit-cache-count-conflict') END").bind(manifest.nameCount),
  db.prepare('INSERT OR IGNORE INTO external_identity_audit_jobs(run_id,plan_revision,identity_revision,manifest_json,created_at,updated_at)VALUES(?,?,?,?,?,?)').bind(runId,revision,manifest.identityRevision,serialized,now,now),
  db.prepare("SELECT CASE WHEN EXISTS(SELECT 1 FROM external_identity_audit_jobs WHERE run_id=? AND plan_revision=?) THEN 1 ELSE json('external-audit-pin-conflict') END").bind(runId,revision),
 ]);return {nextCandidateBatch:0,nextFeatureBatch:0,completed:false};
}
const progress=job=>({nextCandidateBatch:job.next_candidate_batch,nextFeatureBatch:job.next_feature_batch,completed:job.status==='complete'});
export async function applyExternalIdentityAuditBatch(db,{runId,revision,kind,batch}){
 if(!['candidates','features'].includes(kind))throw invalid('Invalid identity batch kind.');
 const job=await jobFor(db,runId,revision,kind,batch?.index),items=batch?.[kind];
 const descriptor=job.descriptor?JSON.parse(job.descriptor):null;
 if(!descriptor || !Array.isArray(items)||items.length!==descriptor.count || new TextEncoder().encode(JSON.stringify(batch)).length>24576
  || await digest(batch)!==descriptor.hash)throw invalid('Identity batch differs from its pinned evidence.');
 const receipt=await db.prepare('SELECT payload_hash FROM external_identity_audit_receipts WHERE run_id=? AND kind=? AND batch_index=?').bind(runId,kind,batch.index).first();
 if(receipt){if(receipt.payload_hash!==descriptor.hash)throw invalid('Conflicting identity receipt.');return {...progress(job),duplicate:true};}
 const cursor=kind==='candidates'?'next_candidate_batch':'next_feature_batch';
 if(job.status!=='running'||batch.index!==job[cursor] || kind==='features'&&job.next_candidate_batch!==job.candidate_batches)throw invalid('Identity batches are incomplete or out of order.');
 const mutations=[],features=[],seen=new Set(),now=new Date().toISOString();
 if(kind==='candidates')for(const c of items){
  for(const [alias,fp]of [[c.left,c.leftFingerprint],[c.right,c.rightFingerprint]]){
   if(!alias||!sources.has(alias.sourceType)||!text(alias.key,200)||!text(alias.displayName,200)||!text(alias.personId,200)||alias.marker!==alias.sourceType+':'+alias.key||!hex(fp))throw invalid('Invalid proposal alias.');
   features.push({source:alias.sourceType,key:alias.key,fingerprint:fp,personId:alias.personId,displayName:alias.displayName});
  }
  const pair=[c.left,c.right].sort((a,b)=>a.marker<b.marker?-1:1);
  if(c.left.personId===c.right.personId || c.pairKey!==doctorIdentityPairKey(c.left,c.right)||seen.has(c.pairKey)
   || !hex(c.fingerprint)||await digest(pair.map(a=>[a.sourceType,a.key,a.displayName,a.personId]))!==c.fingerprint
   || !text(c.evidence?.reason,500)||!['formatting','likely','check'].includes(c.evidence?.confidence))throw invalid('Invalid proposal evidence.');
  seen.add(c.pairKey);mutations.push(identitySuggestionProposalStatement(db,c,now));
 }else for(const f of items){
  const colon=String(f.marker||'').indexOf(':'),source=String(f.marker||'').slice(0,colon),key=String(f.marker||'').slice(colon+1);
  if(!sources.has(source)||!text(key,200)||!hex(f.fingerprint)||seen.has(f.marker))throw invalid('Invalid feature completion.');
  seen.add(f.marker);features.push({source,key,fingerprint:f.fingerprint});
 }
 if(kind==='features'){
  const json=JSON.stringify(features);
  mutations.push(db.prepare(`INSERT INTO external_identity_audit_features(run_id,source_type,doctor_key,fingerprint)
   SELECT ?,json_extract(value,'$.source'),json_extract(value,'$.key'),json_extract(value,'$.fingerprint') FROM json_each(?)`).bind(runId,json),
   db.prepare(`UPDATE roster_identity_features SET audited_fingerprint=fingerprint WHERE (source_type,doctor_key) IN
    (SELECT json_extract(value,'$.source'),json_extract(value,'$.key') FROM json_each(?)) AND audited_fingerprint<>fingerprint`).bind(json));
 }

 if(!await reserveRosterMaintenanceBudget(db,64+items.length*(kind==='candidates'?12:8),65536))return {deferred:true};
 await db.batch([identityGuard(db,job.identity_revision),
  db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM external_identity_audit_jobs WHERE run_id=? AND plan_revision=? AND status='running' AND ${cursor}=? AND (?='candidates' OR next_candidate_batch=?))
    THEN 1 ELSE json('external-audit-cursor-conflict') END`).bind(runId,revision,batch.index,kind,job.candidate_batches),
  featureGuard(db,features),...mutations,
  db.prepare('INSERT INTO external_identity_audit_receipts(run_id,kind,batch_index,payload_hash)VALUES(?,?,?,?)').bind(runId,kind,batch.index,descriptor.hash),
  db.prepare(`UPDATE external_identity_audit_jobs SET ${cursor}=${cursor}+1,updated_at=? WHERE run_id=? AND ${cursor}=?`).bind(now,runId,batch.index),
 ]);return {...progress(job),[kind==='candidates'?'nextCandidateBatch':'nextFeatureBatch']:batch.index+1};
}
export async function finishExternalIdentityAudit(db,{runId,revision}){
 const job=await jobFor(db,runId,revision);if(job.status==='complete')return progress(job);
 if(job.next_candidate_batch!==job.candidate_batches||job.next_feature_batch!==job.feature_batches)throw invalid('Identity audit has unfinished checkpoints.');
 if(!await reserveRosterMaintenanceBudget(db,32,65536))return {deferred:true};
 await db.batch([identityGuard(db,job.identity_revision),
  db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM external_identity_audit_receipts WHERE run_id=? AND kind='candidates')=?
    AND (SELECT COUNT(*) FROM external_identity_audit_receipts WHERE run_id=? AND kind='features')=?
    AND (SELECT COUNT(*) FROM external_identity_audit_features WHERE run_id=?)=?
    AND EXISTS(SELECT 1 FROM external_identity_audit_jobs WHERE run_id=? AND plan_revision=? AND status='running' AND next_candidate_batch=? AND next_feature_batch=?)
    AND (SELECT COUNT(*) FROM (SELECT 1 FROM roster_identity_features LIMIT 2001))=?
    AND NOT EXISTS(SELECT 1 FROM roster_identity_features INDEXED BY idx_identity_feature_pending WHERE audited_fingerprint<>fingerprint LIMIT 1)
    AND NOT EXISTS(SELECT 1 FROM external_identity_audit_features e LEFT JOIN roster_identity_features f ON f.source_type=e.source_type AND f.doctor_key=e.doctor_key
      WHERE e.run_id=? AND (f.doctor_key IS NULL OR f.fingerprint<>e.fingerprint OR f.audited_fingerprint<>e.fingerprint))
    THEN 1 ELSE json('external-audit-incomplete') END`).bind(runId,job.candidate_batches,runId,job.feature_batches,runId,job.pending_count,runId,revision,job.candidate_batches,job.feature_batches,job.name_count,runId),
  db.prepare("UPDATE external_identity_audit_jobs SET status='complete',updated_at=? WHERE run_id=? AND plan_revision=? AND next_candidate_batch=? AND next_feature_batch=? AND status='running'")
    .bind(new Date().toISOString(),runId,revision,job.candidate_batches,job.feature_batches),
 ]);return {...progress(job),completed:true};
}
