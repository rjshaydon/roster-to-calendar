import {comparisonPresence} from '../public/static/comparison-policy.js';
import {COMPARISON_INDEX_VERSION} from '../functions/_lib/roster-comparison.js';
import {createHash} from 'node:crypto';
import {gunzipSync} from 'node:zlib';
import {DOCTOR_SOURCE_VERSION,DOCTOR_SOURCE_LAYOUT,doctorSourceKey} from '../functions/_lib/doctor-source-artifacts.js';
import {publishDoctorMonth} from './process-doctor-calendars.mjs';
const hash=value=>createHash('sha256').update(value).digest('hex');
const decode=bytes=>JSON.parse(Buffer.from(bytes).subarray(0,2).equals(Buffer.from([31,139]))?gunzipSync(bytes).toString('utf8'):Buffer.from(bytes).toString('utf8'));
export async function publishDoctorSource({sourceType,request,manifestKey,operationRevision}){
 const downloaded=await request({mode:'object',sourceType,key:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{binary:true,metadata:true});
 const manifest=decode(downloaded.bytes),previous=await request({mode:'source-current',sourceType});
 const oldCatalogue=previous.current?.catalogueKey?decode(await request({mode:'object',sourceType,key:previous.current.catalogueKey},{binary:true})):null;
 if(!operationRevision&&previous.current?.version===DOCTOR_SOURCE_VERSION&&previous.current.layout===DOCTOR_SOURCE_LAYOUT&&Array.isArray(oldCatalogue?.months)&&previous.current?.parentEtag===downloaded.etag&&previous.current.comparisonIndexes&&previous.current.comparisonCertificate&&previous.current.comparisonVersion===COMPARISON_INDEX_VERSION)return {status:'unchanged',changed:0};
 const months=Object.entries(manifest.months||{}).sort(([a],[b])=>a<b?-1:a>b?1:0),byDoctor=new Map();
 for(const [month,pointer] of months){
  await publishDoctorMonth({sourceType,month,revision:pointer.revision,request});
  const index=decode(await request({mode:'object',sourceType,key:`doctor-calendars/v1/${sourceType}/${month}/indexes/${pointer.revision}.json`},{binary:true}));
  for(const [doctorKey,entry] of Object.entries(index.doctors)){if(!byDoctor.has(doctorKey))byDoctor.set(doctorKey,[]);byDoctor.get(doctorKey).push({month,...entry});}
 }
 if(byDoctor.size>4096)throw Error('Doctor source identity bound exceeded.');
 const overrides=new Map(),terms=[];
 for(const term of manifest.terms||[]){
  const staff=term.staffKey?decode(await request({mode:'object',sourceType,key:term.staffKey},{binary:true})):null;
  const coverage=Array.isArray(term.coverage)?term.coverage:Array.isArray(staff?.coverage)?staff.coverage:manifest.coverage;
  const empty=Array.isArray(coverage)&&!coverage.length&&(!term.staffKey||Boolean(staff));
  terms.push({termStart:term.termStart,termEnd:term.termEnd,coverage:Array.isArray(coverage)?coverage.map(c=>({startDate:c.startDate,endDate:c.endDate})):null,empty,staffUnavailable:!staff});
  for(const override of staff?.seniorityOverrides||[])if(override.termStart===term.termStart){if(!overrides.has(override.doctorKey))overrides.set(override.doctorKey,[]);overrides.get(override.doctorKey).push(override);}
 }
 const doctors=Object.create(null),entries=[...byDoctor].sort(([a],[b])=>a<b?-1:a>b?1:0);let changed=0;
 for(let offset=0;offset<entries.length;offset+=4){
  const results=await Promise.allSettled(entries.slice(offset,offset+4).map(async([doctorKey,pointers])=>{
   const ownOverrides=(overrides.get(doctorKey)||[]).map(o=>({sourceType:o.sourceType,doctorKey:o.doctorKey,termStart:o.termStart,seniority:o.seniority,useRosterSeniority:o.useRosterSeniority})),inputSignature=hash(JSON.stringify([DOCTOR_SOURCE_VERSION,COMPARISON_INDEX_VERSION,pointers.map(p=>[p.month,p.digest]),ownOverrides]));
   const prior=Object.hasOwn(oldCatalogue?.doctors||{},doctorKey)?oldCatalogue.doctors[doctorKey]:null;
   if(prior?.inputSignature===inputSignature){doctors[doctorKey]=prior;return;}
   const saved=(await request({mode:'source-receipt',sourceType,inputSignature})).receipt;
   if(saved?.doctorKey===doctorKey&&saved.pointer?.inputSignature===inputSignature&&saved.pointer.key===doctorSourceKey(sourceType,'people',saved.pointer.digest)){doctors[doctorKey]=saved.pointer;return;}
   const events=[],seen=new Set();let displayName=doctorKey;
   for(const pointer of pointers){const data=decode(await request({mode:'object',sourceType,key:pointer.key},{binary:true}));
    for(const row of data.rows){displayName=row.displayName||displayName;const date=String(row.event.start||'').slice(0,10),term=terms.find(t=>t.termStart<=date&&t.termEnd>=date),override=ownOverrides.find(o=>o.termStart===term?.termStart&&o.sourceType===sourceType);
     const event=override&&!override.useRosterSeniority?{...row.event,seniority:override.seniority,facilitySeniorityOverride:true}:row.event;
     const marker=event.id||JSON.stringify(event);if(!seen.has(marker)){seen.add(marker);events.push(event);}
    }
   }
   const presence={displayName,months:comparisonPresence(events,sourceType),terms:Object.fromEntries(terms.filter(term=>!term.empty).map(term=>[term.termStart,comparisonPresence(events.filter(event=>String(event.start||'').slice(0,10)>=term.termStart&&String(event.start||'').slice(0,10)<=term.termEnd),sourceType)]))};
   const body=JSON.stringify({version:DOCTOR_SOURCE_VERSION,sourceType,doctorKey,events}),digest=hash(body),key=doctorSourceKey(sourceType,'people',digest);
   if(prior?.digest===digest&&prior.key===key){await request({mode:'source-reuse',sourceType,digest,inputSignature,doctorKey,presence,manifestKey:downloaded.key||manifestKey||`facility-overview/v1/${sourceType}/manifest.json`});doctors[doctorKey]={...prior,inputSignature,presence};return;}
   const result=await request({sourceType,kind:'source-person',digest,inputSignature,doctorKey,presence,manifestKey:downloaded.key||manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body});
   if(!['complete','unchanged'].includes(result.status))throw Error('Doctor source upload did not complete.');if(result.status==='complete')changed++;
   doctors[doctorKey]={key,digest,inputSignature,presence};
  }));for(const result of results)if(result.status==='rejected')throw result.reason;
 }
 const comparisonIndexes=Object.create(null);
 for(const [month] of months){
   const members=Object.entries(doctors).filter(([,pointer])=>pointer.presence?.months?.[month]).map(([key,pointer])=>[key,pointer.presence.displayName||key,pointer.presence.months[month],Object.fromEntries(Object.entries(pointer.presence.terms||{}).filter(([,dates])=>dates[month]).map(([term,dates])=>[term,dates[month]]))]).sort(([a],[b])=>a.localeCompare(b));
   const body=JSON.stringify({version:COMPARISON_INDEX_VERSION,sourceType,month,doctors:members}),digest=hash(body),key=doctorSourceKey(sourceType,'catalogues',digest);
   if(previous.current?.comparisonIndexes?.[month]?.digest!==digest)await request({sourceType,kind:'source-catalogue',digest,manifestKey:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body});
   comparisonIndexes[month]={key,digest};
 }
 const certificateBody=JSON.stringify({version:COMPARISON_INDEX_VERSION,sourceType,indexes:comparisonIndexes}),certificateDigest=hash(certificateBody),certificateKey=doctorSourceKey(sourceType,'catalogues',certificateDigest);
 if(previous.current?.comparisonCertificate?.digest!==certificateDigest)await request({sourceType,kind:'source-catalogue',digest:certificateDigest,manifestKey:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body:certificateBody});
 const comparisonCertificate={key:certificateKey,digest:certificateDigest};
 const ordered=Object.fromEntries(Object.keys(doctors).sort().map(k=>[k,doctors[k]])),body=JSON.stringify({version:DOCTOR_SOURCE_VERSION,sourceType,terms,months:months.map(([month])=>month),doctors:ordered}),digest=hash(body),catalogueKey=doctorSourceKey(sourceType,'catalogues',digest);
 await request({sourceType,kind:'source-catalogue',digest,manifestKey:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body});
 const shards=Object.create(null),groups=Array.from({length:16},()=>Object.create(null));
 for(const [doctorKey,pointer] of Object.entries(ordered))groups[parseInt(hash(doctorKey)[0],16)][doctorKey]={key:pointer.key,digest:pointer.digest};
 for(let offset=0;offset<16;offset+=4){
  const results=await Promise.allSettled(groups.slice(offset,offset+4).map(async(doctors,index)=>{
   const shard=(offset+index).toString(16),body=JSON.stringify({version:DOCTOR_SOURCE_VERSION,sourceType,shard,doctors}),digest=hash(body),key=doctorSourceKey(sourceType,'catalogues',digest);
   if(previous.current?.shards?.[shard]?.digest!==digest)await request({sourceType,kind:'source-catalogue',digest,manifestKey:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body});
   shards[shard]={key,digest};
  }));for(const result of results)if(result.status==='rejected')throw result.reason;
 }
 const result=await request({mode:operationRevision?'prepare-source':'activate-source',sourceType,operationRevision,manifestRevision:manifest.revision,manifestEtag:downloaded.etag,catalogueKey,digest,baseEtag:previous.etag||'',layout:DOCTOR_SOURCE_LAYOUT,metadata:{terms,months:months.map(([month])=>month)},shards,comparisonIndexes,comparisonCertificate});
 if(!['complete','unchanged'].includes(result.status))throw Error('Doctor source activation did not complete.');return {status:'complete',changed,doctors:entries.length};
}
