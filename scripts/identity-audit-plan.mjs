import {createHash} from 'node:crypto';
import {flattenDoctorIdentities,compareDoctorIdentityPair} from '../public/static/doctor-identity.js';
const sources = new Set(['mmc','mch','ddh','vhh','casey']);
const hash = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');
const wire = value => JSON.parse(JSON.stringify(value));
export const IDENTITY_AUDIT_ALGORITHM_VERSION = 'external-pairs-v1';
export const IDENTITY_AUDIT_MAXIMUM_NAMES = 2000;
export const IDENTITY_PROPOSAL_BATCH_SIZE = 5;
export const IDENTITY_FEATURE_BATCH_SIZE = 16;

export async function loadIdentityAuditSnapshot(requestPage) {
  const rows=[], seen=new Set(); let cursor={}, identityRevision=null;
  do {
    const page=await requestPage({...cursor,identityRevision});
    if (!Number.isSafeInteger(page.identityRevision) || page.identityRevision<0 || identityRevision!=null && identityRevision!==page.identityRevision
      || !Array.isArray(page.rows) || page.rows.length>32) throw Error('Invalid or changed identity snapshot page.');
    identityRevision=page.identityRevision;
    for(const row of page.rows) {
      const marker=JSON.stringify([row.source_type,row.doctor_key]);
      if(seen.has(marker)) throw Error('Identity snapshot cursor repeated a name.');
      seen.add(marker);rows.push(row);
    }
    if(rows.length>IDENTITY_AUDIT_MAXIMUM_NAMES) throw Error('Identity name cache exceeds its reviewed bound; no audit completion is allowed.');
    cursor=page.next;
    if(cursor && (!page.rows.length || cursor.afterSource!==page.rows.at(-1).source_type || cursor.afterKey!==page.rows.at(-1).doctor_key)) throw Error('Invalid identity snapshot continuation.');
  } while(cursor);
  return {identityRevision,rows};
}

// Pure Node preparation: no network, database mutations or account credentials.
export function planIdentityAudit(snapshot,{maximumCandidates=10000,fullAudit=false}={}) {
  if(!Number.isSafeInteger(snapshot?.identityRevision) || snapshot.identityRevision<0 || !Array.isArray(snapshot.rows) || snapshot.rows.length>IDENTITY_AUDIT_MAXIMUM_NAMES
    || typeof fullAudit!=='boolean' || !Number.isSafeInteger(maximumCandidates) || maximumCandidates<1 || maximumCandidates>10000) throw Error('Invalid identity audit snapshot/bound.');
  const rows=[...snapshot.rows].sort((a,b)=>a.source_type<b.source_type?-1:a.source_type>b.source_type?1:a.doctor_key<b.doctor_key?-1:a.doctor_key>b.doctor_key?1:0);
  const pending=new Map(), features=new Map();
  for(const r of rows) {
    if(!sources.has(r.source_type) || !r.doctor_key || r.doctor_key.length>200 || !r.display_name || r.display_name.length>200
      || !r.person_id || r.person_status!=='active' || !/^[a-f0-9]{64}$/.test(r.fingerprint)
      || r.fingerprint!==hash([r.source_type,r.doctor_key,r.display_name,r.person_id])) throw Error('Identity registration/evidence needs refresh before audit.');
    const marker=r.source_type+':'+r.doctor_key;
    if(features.has(marker)) throw Error('Duplicate cached identity.');
    features.set(marker,r);
    if(fullAudit || r.audited_fingerprint!==r.fingerprint) pending.set(marker,r.fingerprint);
  }
  const identities=flattenDoctorIdentities(rows.map(r=>({sourceType:r.source_type,key:r.doctor_key,displayName:r.display_name,personId:r.person_id})));
  if(identities.length!==rows.length || identities.some(a=>!features.has(a.marker))) throw Error('A cached name needs review; no name may be silently omitted.');
  const groups=new Map();
  for(const alias of identities) {const initial=alias.name.givenInitial;if(!groups.has(initial))groups.set(initial,[]);groups.get(initial).push(alias);}
  const candidates=[];let comparisonCount=0;
  // Every evidence rule requires the same first initial. Compare all pairs in
  // those groups, including surname-prefix typos and arbitrarily dense blocks.
  for(const aliases of groups.values()) for(let i=0;i<aliases.length;i++) for(let j=i+1;j<aliases.length;j++) {
    const left=aliases[i],right=aliases[j];
    if(!pending.has(left.marker) && !pending.has(right.marker)) continue;
    comparisonCount++;
    const candidate=compareDoctorIdentityPair(left,right);
    if(!candidate) continue;
    if(candidates.length>=maximumCandidates) throw Error('Suggestion envelope exceeded; retain pending names for review, never truncate an audit.');
    const pair=[candidate.left,candidate.right].sort((a,b)=>a.marker<b.marker?-1:1);
    candidates.push({pairKey:candidate.pairKey,left:wire(candidate.left),right:wire(candidate.right),
      fingerprint:hash(pair.map(a=>[a.sourceType,a.key,a.displayName,a.personId])),
      leftFingerprint:features.get(left.marker).fingerprint,rightFingerprint:features.get(right.marker).fingerprint,
      evidence:{reason:candidate.reason,confidence:candidate.confidence}});
  }
  candidates.sort((a,b)=>a.pairKey<b.pairKey?-1:a.pairKey>b.pairKey?1:0);
  const batches=[];
  for(let i=0;i<candidates.length;i+=IDENTITY_PROPOSAL_BATCH_SIZE) {
    const batch={index:batches.length,candidates:candidates.slice(i,i+IDENTITY_PROPOSAL_BATCH_SIZE)};
    if(Buffer.byteLength(JSON.stringify(batch))>24576) throw Error('Identity proposal exceeds its byte envelope.');
    batches.push(batch);
  }
  const featuresToFinish=[...pending].map(([marker,fingerprint])=>({marker,fingerprint}));
  const featureBatches=[];
  for(let i=0;i<featuresToFinish.length;i+=IDENTITY_FEATURE_BATCH_SIZE) featureBatches.push({index:featureBatches.length,features:featuresToFinish.slice(i,i+IDENTITY_FEATURE_BATCH_SIZE)});
  const manifest={schemaVersion:1,algorithmVersion:IDENTITY_AUDIT_ALGORITHM_VERSION,fullAudit,identityRevision:snapshot.identityRevision,nameCount:rows.length,
    snapshotHash:hash(rows.map(r=>[r.source_type,r.doctor_key,r.fingerprint])),
    pendingCount:featuresToFinish.length,candidateCount:candidates.length,
    featureBatches:featureBatches.map(b=>({index:b.index,count:b.features.length,hash:hash(b)})),
    batches:batches.map(b=>({index:b.index,count:b.candidates.length,hash:hash(b)}))};
  return {manifest,batches,featureBatches,revision:hash(manifest),comparisonCount};
}
