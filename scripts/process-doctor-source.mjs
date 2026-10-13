import {createHash} from 'node:crypto';
import {gunzipSync} from 'node:zlib';
import {DOCTOR_SOURCE_VERSION,doctorSourceKey} from '../functions/_lib/doctor-source-artifacts.js';
import {publishDoctorMonth} from './process-doctor-calendars.mjs';
const hash=value=>createHash('sha256').update(value).digest('hex');
const decode=bytes=>JSON.parse(Buffer.from(bytes).subarray(0,2).equals(Buffer.from([31,139]))?gunzipSync(bytes).toString('utf8'):Buffer.from(bytes).toString('utf8'));
export async function publishDoctorSource({sourceType,request,manifestKey,operationRevision}){
 const downloaded=await request({mode:'object',sourceType,key:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{binary:true,metadata:true});
 const manifest=decode(downloaded.bytes),previous=await request({mode:'source-current',sourceType});
 const oldCatalogue=previous.current?.catalogueKey?decode(await request({mode:'object',sourceType,key:previous.current.catalogueKey},{binary:true})):null;
 if(!operationRevision&&previous.current?.version===DOCTOR_SOURCE_VERSION&&Array.isArray(oldCatalogue?.months)&&previous.current?.parentEtag===downloaded.etag)return {status:'unchanged',changed:0};
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
   const ownOverrides=(overrides.get(doctorKey)||[]).map(o=>({sourceType:o.sourceType,doctorKey:o.doctorKey,termStart:o.termStart,seniority:o.seniority,useRosterSeniority:o.useRosterSeniority})),inputSignature=hash(JSON.stringify([DOCTOR_SOURCE_VERSION,pointers.map(p=>[p.month,p.digest]),ownOverrides]));
   const prior=Object.hasOwn(oldCatalogue?.doctors||{},doctorKey)?oldCatalogue.doctors[doctorKey]:null;
   if(prior?.inputSignature===inputSignature){doctors[doctorKey]=prior;return;}
   const saved=(await request({mode:'source-receipt',sourceType,inputSignature})).receipt;
   if(saved?.doctorKey===doctorKey&&saved.pointer?.inputSignature===inputSignature&&saved.pointer.key===doctorSourceKey(sourceType,'people',saved.pointer.digest)){doctors[doctorKey]=saved.pointer;return;}
   const events=[],seen=new Set();
   for(const pointer of pointers){const data=decode(await request({mode:'object',sourceType,key:pointer.key},{binary:true}));
    for(const row of data.rows){const date=String(row.event.start||'').slice(0,10),term=terms.find(t=>t.termStart<=date&&t.termEnd>=date),override=ownOverrides.find(o=>o.termStart===term?.termStart&&o.sourceType===sourceType);
     const event=override&&!override.useRosterSeniority?{...row.event,seniority:override.seniority,facilitySeniorityOverride:true}:row.event;
     const marker=event.id||JSON.stringify(event);if(!seen.has(marker)){seen.add(marker);events.push(event);}
    }
   }
   const body=JSON.stringify({version:DOCTOR_SOURCE_VERSION,sourceType,doctorKey,events}),digest=hash(body),key=doctorSourceKey(sourceType,'people',digest);
   if(prior?.digest===digest&&prior.key===key){doctors[doctorKey]={...prior,inputSignature};return;}
   const result=await request({sourceType,kind:'source-person',digest,inputSignature,doctorKey,manifestKey:downloaded.key||manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body});
   if(!['complete','unchanged'].includes(result.status))throw Error('Doctor source upload did not complete.');if(result.status==='complete')changed++;
   doctors[doctorKey]={key,digest,inputSignature};
  }));for(const result of results)if(result.status==='rejected')throw result.reason;
 }
 const ordered=Object.fromEntries(Object.keys(doctors).sort().map(k=>[k,doctors[k]])),body=JSON.stringify({version:DOCTOR_SOURCE_VERSION,sourceType,terms,months:months.map(([month])=>month),doctors:ordered}),digest=hash(body),catalogueKey=doctorSourceKey(sourceType,'catalogues',digest);
 await request({sourceType,kind:'source-catalogue',digest,manifestKey:manifestKey||`facility-overview/v1/${sourceType}/manifest.json`},{body});
 const result=await request({mode:operationRevision?'prepare-source':'activate-source',sourceType,operationRevision,manifestRevision:manifest.revision,manifestEtag:downloaded.etag,catalogueKey,digest,baseEtag:previous.etag||''});
 if(!['complete','unchanged'].includes(result.status))throw Error('Doctor source activation did not complete.');return {status:'complete',changed,doctors:entries.length};
}
