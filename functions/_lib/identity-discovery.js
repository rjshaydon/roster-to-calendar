import {flattenDoctorIdentities,normalizeDoctorIdentityName,buildDoctorIdentityCandidates} from '../../public/static/doctor-identity.js';
import {harmlessIdentityKey} from './doctor-identity.js';
const all=async(db,sql,args=[]) => (await db.prepare(sql).bind(...args).all()).results||[];
async function digest(value) { return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value))))].map(b=>b.toString(16).padStart(2,'0')).join(''); }
// Caller supplies authoritative published directory data, never client names.
// Keep production requests small enough for Pages CPU limits. The durable
// cursor resumes after each request; at most 24 neighbours/block, no events.
export const IDENTITY_BATCH_SIZE = 5;
// On reopening the UI, consult the cached names rather than restarting a
// completed registration pass. This reads names only, never roster events.
export async function initializePublishedIdentityBatch(db,doctors,options={}) {
 const cached=await all(db,'SELECT source_type,doctor_key,display_name FROM roster_identity_features LIMIT 2001');
 if(cached.length>2000) throw Error('Identity name cache exceeds the initialization inspection limit.');
 const names=new Map(cached.map(a=>[a.source_type+':'+a.doctor_key,a.display_name]));
 const identities=flattenDoctorIdentities(doctors);
 if(identities.every(a=>names.get(a.marker)===a.displayName)) return {status:'complete',examined:0,candidates:0};
 return auditIdentityBatch(db,doctors,{...options,register:true});
}
export async function auditIdentityBatch(db,doctors,{runId,actor='identity-audit',register=false,weekKey='',sourceTypes=[],cachedDirectory=false}={}) {
 const startedAt=Date.now();
 const scope=[...new Set(sourceTypes)].sort();
 if(scope.some(s=>!['mmc','mch','ddh','vhh','casey'].includes(s))) throw Error('Unsupported audit site.');
 const scopeJson=JSON.stringify(scope);
 const now=new Date().toISOString();
 let run=runId?await db.prepare('SELECT * FROM roster_identity_audit_runs WHERE run_id=?').bind(runId).first():null;
 if(runId && !run) throw Error('Identity audit not found.');
 if(!run) run=await db.prepare("SELECT * FROM roster_identity_audit_runs WHERE status='running' AND mode=? AND scope_json=? ORDER BY created_at LIMIT 1").bind(register?'register':'audit',scopeJson).first();
 if(run && run.scope_json!==scopeJson) throw Error('Resume the same audit sites.');
 if(run && run.mode!==(register?'register':'audit')) throw Error('Resume the same kind of identity audit.');
 if(run?.status==='complete') return {runId:run.run_id,status:'complete',cursor:run.cursor,examined:run.examined,candidates:run.candidates};
 if(!run){run={run_id:crypto.randomUUID(),cursor:'',examined:0,candidates:0,status:'running'};await db.prepare("INSERT INTO roster_identity_audit_runs(run_id,status,mode,scope_json,week_key,created_at,updated_at) VALUES(?,'running',?,?,?,?,?)").bind(run.run_id,register?'register':'audit',scopeJson,weekKey,now,now).run();}
 const lease=crypto.randomUUID();
 const acquired=await db.prepare('UPDATE roster_identity_audit_runs SET lease_token=?,lease_until=?,week_key=CASE WHEN ?<>\'\' THEN ? ELSE week_key END WHERE run_id=? AND lease_until<?').bind(lease,new Date(Date.now()+120000).toISOString(),weekKey,weekKey,run.run_id,now).run();
 if(!(acquired.meta?.changes||acquired.changes)) return {runId:run.run_id,status:'busy'};
 const batchSize=register?(cachedDirectory?1:IDENTITY_BATCH_SIZE):1;
 const scopeClause=scope.length?'AND source_type IN ('+scope.map(()=>'?').join(',')+')':'';
 const page=cachedDirectory&&register?await all(db,`SELECT source_type,doctor_key,display_name FROM roster_doctors
   WHERE source_type IN ('mmc','mch','ddh','vhh','casey') ${scopeClause}
   AND length(doctor_key) BETWEEN 1 AND 200 AND length(display_name) BETWEEN 1 AND 200
   AND (source_type,doctor_key)>(?,?) ORDER BY source_type,doctor_key LIMIT 2`,
   [...scope,run.cursor.split(':')[0]||'',run.cursor.slice(run.cursor.indexOf(':')+1)||'']):null;
 const identities=flattenDoctorIdentities(page?page.map(a=>({sourceType:a.source_type,key:a.doctor_key,displayName:a.display_name})):doctors).filter(a=>(!scope.length || scope.includes(a.sourceType)) && ['mmc','mch','ddh','vhh','casey'].includes(a.sourceType) && a.key.length<=200 && a.displayName.length<=200).sort((a,b)=>a.marker<b.marker?-1:a.marker>b.marker?1:0);
 if(page?.length && !identities.some(a=>a.marker===page[0].source_type+':'+page[0].doctor_key)) throw Error('Cached roster name requires review before automatic identity registration.');
 const pending=register?[]:await all(db,`SELECT source_type,doctor_key,display_name FROM roster_identity_features INDEXED BY idx_identity_feature_pending WHERE audited_fingerprint<>fingerprint ${scope.length?'AND source_type IN ('+scope.map(()=>'?').join(',')+')':''} ORDER BY source_type,doctor_key LIMIT ${batchSize+1}`,scope);
 const selected=register?identities.filter(a=>a.marker>run.cursor).slice(0,batchSize):pending.slice(0,batchSize).map(a=>({sourceType:a.source_type,key:a.doctor_key,displayName:a.display_name,marker:a.source_type+':'+a.doctor_key}));
 let candidateCount=0, skippedLargeBlocks=0; const processed=[];
 for(const alias of selected) {
  if(Date.now()-startedAt>=10000) break;
  processed.push(alias);
  const name=normalizeDoctorIdentityName(alias.displayName), normalized=harmlessIdentityKey(alias.displayName);
  let existing=await db.prepare('SELECT * FROM roster_person_aliases WHERE source_type=? AND doctor_key=?').bind(alias.sourceType,alias.key).first();
  if(!existing && !register) continue;
  if(!existing) {
   const matches=normalized?await all(db,`SELECT DISTINCT a.person_id FROM roster_identity_features f JOIN roster_person_aliases a ON a.source_type=f.source_type AND a.doctor_key=f.doctor_key JOIN roster_people p ON p.person_id=a.person_id WHERE f.normalized_key=? AND a.review_state='approved' AND p.status='active' LIMIT 2`,[normalized]):[];
   const owners=await all(db,'SELECT email FROM account_claims INDEXED BY idx_account_claims_source_doctor_email WHERE source_type=? AND doctor_key=? LIMIT 3',[alias.sourceType,alias.key]);
   let personId=matches.length===1 && owners.length<3?matches[0].person_id:'';
   if(personId) {
    const aliases=await all(db,'SELECT doctor_key FROM roster_person_aliases WHERE person_id=? LIMIT 16',[personId]);
    if(aliases.length>=16) personId='';
    const linked=await all(db,'SELECT email FROM account_people WHERE person_id=? LIMIT 3',[personId]);
    // Never automatically connect a separately claimed account.
    if(owners.some(o=>!linked.some(a=>a.email===o.email))) personId='';
   }
   const autoLinked=Boolean(personId);
   if(!personId) personId='person:auto-'+(await digest([alias.sourceType,alias.key])).slice(0,24);
   const revision=(await db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1').first()).revision;
   const statements=[db.prepare("SELECT CASE WHEN (SELECT revision FROM roster_identity_revision WHERE id=1)=? AND NOT EXISTS(SELECT 1 FROM roster_person_aliases WHERE source_type=? AND doctor_key=?) THEN 1 ELSE json('identity-conflict') END").bind(revision,alias.sourceType,alias.key)];
   if(autoLinked) statements.push(db.prepare("SELECT CASE WHEN (SELECT COUNT(*) FROM (SELECT email FROM account_claims WHERE source_type=? AND doctor_key=? LIMIT 4))=? THEN 1 ELSE json('identity-conflict') END").bind(alias.sourceType,alias.key,owners.length));
   if(autoLinked) for(const owner of owners) statements.push(db.prepare("SELECT CASE WHEN EXISTS(SELECT 1 FROM account_claims WHERE email=? AND source_type=? AND doctor_key=?) THEN 1 ELSE json('identity-conflict') END").bind(owner.email,alias.sourceType,alias.key));
   if(autoLinked) {
    statements.push(db.prepare("SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_people WHERE person_id=? AND status='active') THEN 1 ELSE json('identity-conflict') END").bind(personId));
    for(const owner of owners) statements.push(db.prepare("SELECT CASE WHEN EXISTS(SELECT 1 FROM account_people WHERE email=? AND person_id=?) THEN 1 ELSE json('identity-conflict') END").bind(owner.email,personId));
   }
   if(!autoLinked) statements.push(db.prepare("INSERT INTO roster_people(person_id,preferred_display_name,provenance,review_state,created_at,updated_at) VALUES(?,?,'published-roster','approved',?,?)").bind(personId,alias.displayName,now,now));
   statements.push(db.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id,provenance,confidence,review_state,created_at,updated_at,approved_by) VALUES(?,?,?,?,?,'formatting-only','approved',?,?,?)").bind(alias.sourceType,alias.key,alias.displayName,personId,autoLinked?'harmless-formatting':'published-roster',now,now,actor));
   statements.push(db.prepare('INSERT INTO roster_identity_registrations(source_type,doctor_key,person_id,actor,reason,created_at) VALUES(?,?,?,?,?,?)').bind(alias.sourceType,alias.key,personId,actor,autoLinked?'Harmless formatting only':'New published roster identity',now));
   statements.push(db.prepare('UPDATE roster_identity_revision SET revision=revision+1 WHERE id=1'));
   await db.batch(statements);
   existing={person_id:personId};
  }
  const featureHash=await digest([alias.sourceType,alias.key,alias.displayName,existing.person_id]);
  const previousFeature=await db.prepare('SELECT fingerprint,audited_fingerprint FROM roster_identity_features WHERE source_type=? AND doctor_key=?').bind(alias.sourceType,alias.key).first();
  if(!register && previousFeature?.audited_fingerprint===featureHash) continue;
  await db.prepare(`INSERT INTO roster_identity_features(source_type,doctor_key,display_name,normalized_key,given_block,surname_block,fingerprint,updated_at) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(source_type,doctor_key) DO UPDATE SET display_name=excluded.display_name,normalized_key=excluded.normalized_key,given_block=excluded.given_block,surname_block=excluded.surname_block,fingerprint=excluded.fingerprint,updated_at=excluded.updated_at WHERE roster_identity_features.fingerprint<>excluded.fingerprint`).bind(alias.sourceType,alias.key,alias.displayName,normalized,`${name.given}:${name.surname.slice(0,3)}`,`${name.surname}:${name.givenInitial}`,featureHash,now).run();
  if(register) continue;
  const neighbours=new Map();
  for(const [column,value] of [['given_block',`${name.given}:${name.surname.slice(0,3)}`],['surname_block',`${name.surname}:${name.givenInitial}`]]) {
   const matches=await all(db,`SELECT f.source_type,f.doctor_key,f.display_name,a.person_id FROM roster_identity_features f JOIN roster_person_aliases a ON a.source_type=f.source_type AND a.doctor_key=f.doctor_key WHERE f.${column}=? ORDER BY f.source_type,f.doctor_key LIMIT 25`,[value]);
   if(matches.length>24){skippedLargeBlocks++;continue;}
   for(const item of matches) neighbours.set(`${item.source_type}:${item.doctor_key}`,{sourceType:item.source_type,key:item.doctor_key,displayName:item.display_name,personId:item.person_id});
  }
  const candidates=register?[]:buildDoctorIdentityCandidates([{...alias,personId:existing.person_id},...neighbours.values()],{bucketLimit:50,candidateLimit:20}).candidates.filter(c=>c.left.marker===alias.marker || c.right.marker===alias.marker);
  for(const candidate of candidates.slice(0,10)) {
   const fingerprint=await digest([candidate.left,candidate.right].sort((a,b)=>a.marker<b.marker?-1:1).map(a=>[a.sourceType,a.key,a.displayName,a.personId]));
   const previous=await db.prepare('SELECT fingerprint,status FROM roster_identity_candidates WHERE pair_key=?').bind(candidate.pairKey).first();
   if(previous?.fingerprint===fingerprint) continue;
   await db.prepare(`INSERT INTO roster_identity_candidates(pair_key,left_json,right_json,evidence_json,fingerprint,status,updated_at) VALUES(?,?,?,?,?,'pending',?) ON CONFLICT(pair_key) DO UPDATE SET left_json=excluded.left_json,right_json=excluded.right_json,evidence_json=excluded.evidence_json,fingerprint=excluded.fingerprint,status='pending',reviewer='',updated_at=excluded.updated_at`).bind(candidate.pairKey,JSON.stringify(candidate.left),JSON.stringify(candidate.right),JSON.stringify({reason:candidate.reason,confidence:candidate.confidence,changedEvidence:Boolean(previous)}),fingerprint,now).run(); candidateCount++;
  }
  await db.prepare('UPDATE roster_identity_features SET audited_fingerprint=? WHERE source_type=? AND doctor_key=?').bind(featureHash,alias.sourceType,alias.key).run();
 }
 let cursor=processed.at(-1)?.marker || run.cursor, complete=register?!identities.some(a=>a.marker>cursor):pending.length<=processed.length;
 if(register && complete && cachedDirectory) {
  const size=await db.prepare('SELECT COUNT(*) AS n FROM (SELECT 1 FROM roster_doctors LIMIT 2001)').first();
  if(Number(size?.n)>2000) throw Error('Cached roster names exceed the completion inspection limit.');
  const missing=await db.prepare(`SELECT d.source_type,d.doctor_key FROM roster_doctors d
    LEFT JOIN roster_person_aliases a ON a.source_type=d.source_type AND a.doctor_key=d.doctor_key
    LEFT JOIN roster_identity_features f ON f.source_type=d.source_type AND f.doctor_key=d.doctor_key
    WHERE d.source_type IN ('mmc','mch','ddh','vhh','casey')
    ${scope.length?'AND d.source_type IN ('+scope.map(()=>'?').join(',')+')':''}
    AND length(d.doctor_key) BETWEEN 1 AND 200 AND length(d.display_name) BETWEEN 1 AND 200
    AND (a.person_id IS NULL OR f.display_name IS NULL OR f.display_name<>d.display_name)
    ORDER BY d.source_type,d.doctor_key LIMIT 1`).bind(...scope).first();
  if(missing) {cursor='';complete=false;}
 } else if(register && complete) {
  // The authoritative directory can grow behind a saved cursor (for example
  // when historical names are restored). Do not publish an unchanged receipt
  // until every current name has a registered alias and matching cache entry.
  const cached=await all(db,`SELECT f.source_type,f.doctor_key,f.display_name FROM roster_identity_features f
    JOIN roster_person_aliases a ON a.source_type=f.source_type AND a.doctor_key=f.doctor_key LIMIT 2001`);
  if(cached.length>2000) throw Error('Registered identity name cache exceeds the completion inspection limit.');
  const names=new Map(cached.map(a=>[a.source_type+':'+a.doctor_key,a.display_name]));
  const missed=identities.findIndex(a=>names.get(a.marker)!==a.displayName);
  if(missed>=0) {cursor=identities[missed-1]?.marker||'';complete=false;}
 }
 await db.prepare("UPDATE roster_identity_audit_runs SET cursor=?,status=?,examined=examined+?,candidates=candidates+?,updated_at=?,lease_token='',lease_until='' WHERE run_id=? AND cursor=? AND lease_token=?").bind(cursor,complete?'complete':'running',processed.length,candidateCount,now,run.run_id,run.cursor,lease).run();
 return {runId:run.run_id,status:complete?'complete':'running',cursor,examined:run.examined+processed.length,candidates:run.candidates+candidateCount,skippedLargeBlocks,durationMs:Date.now()-startedAt};
}
export async function queryIdentityCandidates(db,{after='',status='pending'}={}) {
 // Page before checking current links, so a queue of obsolete suggestions
 // cannot turn a Creator page request into a scan of every old candidate.
 if(!['pending','rejected'].includes(status)) throw Error('Unsupported suggestion view.');
 const list=await all(db,`WITH page AS (SELECT * FROM roster_identity_candidates WHERE status=? AND pair_key>? ORDER BY pair_key LIMIT 26)
 SELECT page.*,l.person_id AS current_left,r.person_id AS current_right FROM page
 LEFT JOIN roster_person_aliases l ON l.source_type=json_extract(page.left_json,'$.sourceType') AND l.doctor_key=json_extract(page.left_json,'$.key')
 LEFT JOIN roster_person_aliases r ON r.source_type=json_extract(page.right_json,'$.sourceType') AND r.doctor_key=json_extract(page.right_json,'$.key') ORDER BY page.pair_key`,[status,String(after)]);
 const candidates=list.slice(0,25).map(r=>({...r,left:JSON.parse(r.left_json),right:JSON.parse(r.right_json),evidence:JSON.parse(r.evidence_json)})).filter(r=>r.current_left===r.left.personId && r.current_right===r.right.personId && r.current_left!==r.current_right);
 return {candidates,next:list.length>25?list[24].pair_key:''};
}
export async function rejectIdentityCandidate(db,{pairKey,fingerprint,actor}) {
 const result=await db.prepare("UPDATE roster_identity_candidates SET status='rejected',reviewer=?,updated_at=? WHERE pair_key=? AND fingerprint=? AND status='pending'").bind(actor,new Date().toISOString(),pairKey,fingerprint).run();
 if(!result.meta?.changes && !result.changes) throw Error('Candidate changed; reload it before rejecting.');
 return {status:'rejected'};
}

export async function restoreIdentityCandidate(db,{pairKey,fingerprint}) {
 const result=await db.prepare("UPDATE roster_identity_candidates SET status='pending',reviewer='',updated_at=? WHERE pair_key=? AND fingerprint=? AND status='rejected'").bind(new Date().toISOString(),pairKey,fingerprint).run();
 if(!result.meta?.changes && !result.changes) throw Error('Suggestion changed; reload it before restoring.');
 return {status:'pending'};
}
