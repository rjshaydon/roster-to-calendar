import {identitySuggestionProposalStatement} from './identity-suggestion-proposal.js';
import {reserveRosterMaintenanceBudget} from './roster-maintenance-budget.js';
import {activeRosterImportsQuery} from './roster-delivery-health.js';
import {doctorIdentityPairKey} from '../../public/static/doctor-identity.js';
const sources=new Set(['mmc','mch','ddh','vhh','casey']);
const hex=v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v);
const text=(v,max)=>typeof v==='string'&&v.length>0&&v.length<=max;
async function digest(value){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value))))].map(v=>v.toString(16).padStart(2,'0')).join('');}
const priority=activeRosterImportsQuery(['monash-adults','monash-paeds','vhh-active-medical-roster','dandenong-findmyshift']);
// Internal primitives, not an enabled endpoint. The future caller authenticates
// and meters optional identity work; these functions also reserve its costs and
// guard source priority, current evidence and the cursor in each transaction.
function identityGuard(db,revision){return db.prepare(`SELECT CASE WHEN (SELECT revision FROM roster_identity_revision WHERE id=1)=?
  AND NOT EXISTS(${priority.sql}) THEN 1 ELSE json('external-audit-stale-or-import-active') END`).bind(revision,...priority.bindings);}
function featureGuard(db,source,key,fingerprint,alias=null){
 const extra=alias?' AND a.person_id=? AND f.display_name=?':'';
 return db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_identity_features f
  JOIN roster_person_aliases a ON a.source_type=f.source_type AND a.doctor_key=f.doctor_key
  JOIN roster_people p ON p.person_id=a.person_id WHERE f.source_type=? AND f.doctor_key=? AND f.fingerprint=? AND p.status='active'${extra})
  THEN 1 ELSE json('external-audit-feature-conflict') END`).bind(source,key,fingerprint,...(alias?[alias.personId,alias.displayName]:[]));
}
function validateManifest(m){
 if(!m || m.schemaVersion!==1 || m.algorithmVersion!=='external-pairs-v1' || typeof m.fullAudit!=='boolean'
  || !Number.isSafeInteger(m.identityRevision) || m.identityRevision<0 || !Number.isInteger(m.nameCount)||m.nameCount<0||m.nameCount>2000
  || !Number.isInteger(m.pendingCount)||m.pendingCount<0||m.pendingCount>m.nameCount
  || m.fullAudit&&m.pendingCount!==m.nameCount
  || !Number.isInteger(m.candidateCount)||m.candidateCount<0||m.candidateCount>10000||!hex(m.snapshotHash)
  || new TextEncoder().encode(JSON.stringify(m)).length>262144)throw Error('Invalid external identity manifest.');
 for(const [field,total,limit,max]of[['batches',m.candidateCount,5,2000],['featureBatches',m.pendingCount,16,125]]){
  const list=m[field];if(!Array.isArray(list)||list.length>max)throw Error('Invalid identity descriptors.');
  let count=0;for(const [i,b]of list.entries()){if(b.index!==i||!hex(b.hash)||!Number.isInteger(b.count)||b.count<1||b.count>limit)throw Error('Invalid identity descriptor.');count+=b.count;}
  if(count!==total)throw Error('Incomplete identity descriptor counts.');
 }
}
async function jobFor(db,runId,revision){
 const job=await db.prepare('SELECT * FROM external_identity_audit_jobs WHERE run_id=?').bind(runId).first();
 if(!job||job.plan_revision!==revision)throw Error('Unpinned or changed identity audit.');return job;
}
export async function pinExternalIdentityAudit(db,{runId,manifest,revision}){
 if(!text(runId,100)||!hex(revision))throw Error('Invalid identity audit identity.');validateManifest(manifest);
 if(await digest(manifest)!==revision)throw Error('Identity manifest hash mismatch.');
 const old=await db.prepare('SELECT * FROM external_identity_audit_jobs WHERE run_id=?').bind(runId).first();
 if(old){if(old.plan_revision!==revision)throw Error('Audit already pinned to different evidence.');return progress(old);}
 if(!await reserveRosterMaintenanceBudget(db,64,65536))return {deferred:true};
 const now=new Date().toISOString();
 await db.batch([identityGuard(db,manifest.identityRevision),
  db.prepare("SELECT CASE WHEN (SELECT COUNT(*) FROM (SELECT 1 FROM roster_identity_features LIMIT 2001))=? THEN 1 ELSE json('external-audit-cache-count-conflict') END").bind(manifest.nameCount),
  db.prepare('INSERT OR IGNORE INTO external_identity_audit_jobs(run_id,plan_revision,identity_revision,manifest_json,created_at,updated_at)VALUES(?,?,?,?,?,?)').bind(runId,revision,manifest.identityRevision,JSON.stringify(manifest),now,now),
  db.prepare("SELECT CASE WHEN EXISTS(SELECT 1 FROM external_identity_audit_jobs WHERE run_id=? AND plan_revision=?) THEN 1 ELSE json('external-audit-pin-conflict') END").bind(runId,revision),
 ]);return {nextCandidateBatch:0,nextFeatureBatch:0,completed:false};
}
const progress=job=>({nextCandidateBatch:job.next_candidate_batch,nextFeatureBatch:job.next_feature_batch,completed:job.status==='complete'});
export async function applyExternalIdentityAuditBatch(db,{runId,revision,kind,batch}){
 if(!['candidates','features'].includes(kind))throw Error('Invalid identity batch kind.');
 const job=await jobFor(db,runId,revision),m=JSON.parse(job.manifest_json),items=batch?.[kind];
 const descriptor=(kind==='candidates'?m.batches:m.featureBatches)[batch?.index];
 if(!descriptor || !Array.isArray(items)||items.length!==descriptor.count || new TextEncoder().encode(JSON.stringify(batch)).length>24576
  || await digest(batch)!==descriptor.hash)throw Error('Identity batch differs from its pinned evidence.');
 const receipt=await db.prepare('SELECT payload_hash FROM external_identity_audit_receipts WHERE run_id=? AND kind=? AND batch_index=?').bind(runId,kind,batch.index).first();
 if(receipt){if(receipt.payload_hash!==descriptor.hash)throw Error('Conflicting identity receipt.');return {...progress(job),duplicate:true};}
 const cursor=kind==='candidates'?'next_candidate_batch':'next_feature_batch';
 if(job.status!=='running'||batch.index!==job[cursor] || kind==='features'&&job.next_candidate_batch!==m.batches.length)throw Error('Identity batches are incomplete or out of order.');
 const mutations=[],seen=new Set(),now=new Date().toISOString();
 if(kind==='candidates')for(const c of items){
  for(const [alias,fp]of [[c.left,c.leftFingerprint],[c.right,c.rightFingerprint]]){
   if(!alias||!sources.has(alias.sourceType)||!text(alias.key,200)||!text(alias.displayName,200)||!text(alias.personId,200)||alias.marker!==alias.sourceType+':'+alias.key||!hex(fp))throw Error('Invalid proposal alias.');
   mutations.push(featureGuard(db,alias.sourceType,alias.key,fp,alias));
  }
  const pair=[c.left,c.right].sort((a,b)=>a.marker<b.marker?-1:1);
  if(c.left.personId===c.right.personId || c.pairKey!==doctorIdentityPairKey(c.left,c.right)||seen.has(c.pairKey)
   || !hex(c.fingerprint)||await digest(pair.map(a=>[a.sourceType,a.key,a.displayName,a.personId]))!==c.fingerprint
   || !text(c.evidence?.reason,500)||!['formatting','likely','check'].includes(c.evidence?.confidence))throw Error('Invalid proposal evidence.');
  seen.add(c.pairKey);mutations.push(identitySuggestionProposalStatement(db,c,now));
 }else for(const f of items){
  const colon=String(f.marker||'').indexOf(':'),source=String(f.marker||'').slice(0,colon),key=String(f.marker||'').slice(colon+1);
  if(!sources.has(source)||!text(key,200)||!hex(f.fingerprint)||seen.has(f.marker))throw Error('Invalid feature completion.');
  seen.add(f.marker);mutations.push(featureGuard(db,source,key,f.fingerprint),
   db.prepare('INSERT INTO external_identity_audit_features(run_id,source_type,doctor_key,fingerprint)VALUES(?,?,?,?)').bind(runId,source,key,f.fingerprint),
   db.prepare('UPDATE roster_identity_features SET audited_fingerprint=? WHERE source_type=? AND doctor_key=? AND fingerprint=? AND audited_fingerprint<>?').bind(f.fingerprint,source,key,f.fingerprint,f.fingerprint));
 }
 if(!await reserveRosterMaintenanceBudget(db,64+items.length*(kind==='candidates'?12:8),65536))return {deferred:true};
 await db.batch([identityGuard(db,job.identity_revision),
  db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM external_identity_audit_jobs WHERE run_id=? AND plan_revision=? AND status='running' AND ${cursor}=? AND (?='candidates' OR next_candidate_batch=?))
    THEN 1 ELSE json('external-audit-cursor-conflict') END`).bind(runId,revision,batch.index,kind,m.batches.length),
  ...mutations,
  db.prepare('INSERT INTO external_identity_audit_receipts(run_id,kind,batch_index,payload_hash)VALUES(?,?,?,?)').bind(runId,kind,batch.index,descriptor.hash),
  db.prepare(`UPDATE external_identity_audit_jobs SET ${cursor}=${cursor}+1,updated_at=? WHERE run_id=? AND ${cursor}=?`).bind(now,runId,batch.index),
 ]);return {...progress(job),[kind==='candidates'?'nextCandidateBatch':'nextFeatureBatch']:batch.index+1};
}
export async function finishExternalIdentityAudit(db,{runId,revision}){
 const job=await jobFor(db,runId,revision);if(job.status==='complete')return progress(job);
 const m=JSON.parse(job.manifest_json);
 if(job.next_candidate_batch!==m.batches.length||job.next_feature_batch!==m.featureBatches.length)throw Error('Identity audit has unfinished checkpoints.');
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
    THEN 1 ELSE json('external-audit-incomplete') END`).bind(runId,m.batches.length,runId,m.featureBatches.length,runId,m.pendingCount,runId,revision,m.batches.length,m.featureBatches.length,m.nameCount,runId),
  db.prepare("UPDATE external_identity_audit_jobs SET status='complete',updated_at=? WHERE run_id=? AND plan_revision=? AND next_candidate_batch=? AND next_feature_batch=? AND status='running'")
    .bind(new Date().toISOString(),runId,revision,m.batches.length,m.featureBatches.length),
 ]);return {...progress(job),completed:true};
}
