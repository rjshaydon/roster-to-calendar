import {loadCachedSnapshot} from './d1-calendar.js';
import {rosterTermVisible} from '../../public/static/roster-term-policy.js';
export const DOCTOR_SOURCE_VERSION='doctor-source-v2';
export const DOCTOR_SOURCE_LAYOUT='hashed-catalogue-v1';
export async function doctorCatalogueShard(doctorKey){const digest=new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(doctorKey)));return (digest[0]>>4).toString(16);}
const validSource=s=>['mmc','mch','ddh','vhh','casey'].includes(s);
export function doctorSourceKey(source,kind,digest){if(!validSource(source)||!['people','catalogues','generations','prepared-people'].includes(kind)||!/^[a-f0-9]{64}$/.test(digest))throw Error('Invalid doctor source artifact.');return `doctor-calendars/v2/${source}/${kind}/${digest}.json`;}
export function doctorSourceCurrentKey(source){if(!validSource(source))throw Error('Invalid doctor source.');return `doctor-calendars/v2/${source}/current.json`;}
export async function loadPublishedDoctorSourceRange(r2,sourceTypes,startDate,endDate,today,markers){
 const length=(Date.parse(endDate+'T12:00:00Z')-Date.parse(startDate+'T12:00:00Z'))/86400000;
 if(!Number.isFinite(length)||length<0||length>370)return {preparing:true,events:[],missing:[],visibleTerms:[],revision:''};
 const selected=[],missing=[];let pending=false;
 await Promise.all(sourceTypes.map(async sourceType=>{
  const [head,current]=await Promise.all([r2.head(`facility-overview/v1/${sourceType}/manifest.json`),loadCachedSnapshot(r2,doctorSourceCurrentKey(sourceType))]);
  if(!head){missing.push({sourceType,startDate,endDate,reason:'manifest-unavailable'});return;}
  if(!current||current.version!==DOCTOR_SOURCE_VERSION||current.parentEtag!==head.etag){pending=true;return;}
  if(current.catalogueKey!==doctorSourceKey(sourceType,'catalogues',current.digest)){pending=true;return;}
  let catalogue;
  if(current.layout===DOCTOR_SOURCE_LAYOUT&&current.metadata&&current.shards){
   const keys=[...markers].filter(marker=>marker.startsWith(sourceType+'|')).map(marker=>marker.slice(sourceType.length+1));
   const shards=[...new Set(await Promise.all(keys.map(doctorCatalogueShard)))];
   const pieces=await Promise.all(shards.map(async shard=>{
    const pointer=current.shards[shard];
    if(!pointer||pointer.key!==doctorSourceKey(sourceType,'catalogues',pointer.digest))return null;
    const value=await loadCachedSnapshot(r2,pointer.key);
    return value?.version===DOCTOR_SOURCE_VERSION&&value.sourceType===sourceType&&value.shard===shard&&value.doctors&&Object.keys(value.doctors).length<=4096?value:null;
   }));
   if(pieces.some(piece=>!piece)){pending=true;return;}
   catalogue={version:DOCTOR_SOURCE_VERSION,sourceType,terms:current.metadata.terms,months:current.metadata.months,doctors:Object.assign(Object.create(null),...pieces.map(piece=>piece.doctors))};
  }else catalogue=await loadCachedSnapshot(r2,current.catalogueKey);
  if(!catalogue||catalogue.version!==DOCTOR_SOURCE_VERSION||catalogue.sourceType!==sourceType||!Array.isArray(catalogue.terms)||!Array.isArray(catalogue.months)||!catalogue.doctors||Object.keys(catalogue.doctors).length>4096){pending=true;return;}
  const terms=catalogue.terms.filter(term=>rosterTermVisible(term,today)&&term.termEnd>=startDate&&term.termStart<=endDate);
  if(!terms.length){missing.push({sourceType,startDate,endDate,reason:'history-unavailable'});return;}
  if(!terms.every(t=>t.empty)){
   const first=startDate.slice(0,7),last=endDate.slice(0,7);
   for(let month=first;month<=last;){
    if(terms.some(t=>!t.empty&&t.termStart.slice(0,7)<=month&&t.termEnd.slice(0,7)>=month)&&!(catalogue.months||[]).includes(month))missing.push({sourceType,month,startDate,endDate,reason:'month-unavailable'});
    const next=new Date(month+'-01T12:00:00Z');next.setUTCMonth(next.getUTCMonth()+1);month=next.toISOString().slice(0,7);
   }
  }
  for(const term of terms)if(term.staffUnavailable&&!term.empty)missing.push({sourceType,termStart:term.termStart,startDate,endDate,reason:'staff-unavailable'});
  missing.push(...doctorRangeGaps(sourceType,terms,startDate,endDate));
  const doctors=[];
  for(const marker of markers){if(!marker.startsWith(sourceType+'|'))continue;const key=marker.slice(sourceType.length+1),pointer=Object.hasOwn(catalogue.doctors,key)?catalogue.doctors[key]:null;
   if(pointer&&pointer.key!==doctorSourceKey(sourceType,'people',pointer.digest)){pending=true;return;}
   doctors.push({doctorKey:key,pointer});
  }
  selected.push({sourceType,terms,doctors});
 }));
 selected.sort((a,b)=>sourceTypes.indexOf(a.sourceType)-sourceTypes.indexOf(b.sourceType));missing.sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b)));
 const facts=selected.map(s=>[s.sourceType,s.terms,s.doctors.map(d=>[d.doctorKey,d.pointer?.digest||'absent'])]);
 const revision=[...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify([facts,missing]))))].map(v=>v.toString(16).padStart(2,'0')).join('');
 return {preparing:pending||!selected.length,events:[],missing,revision,visibleTerms:selected.flatMap(s=>s.terms.map(t=>({sourceType:s.sourceType,termStart:t.termStart,termEnd:t.termEnd}))),
  async readEvents(){
   const groups=await Promise.all(selected.map(async source=>{
    const docs=await Promise.all(source.doctors.map(async({doctorKey,pointer})=>{
     if(!pointer)return [];
     const data=await loadCachedSnapshot(r2,pointer.key);
     if(!data||data.version!==DOCTOR_SOURCE_VERSION||data.sourceType!==source.sourceType||data.doctorKey!==doctorKey||!Array.isArray(data.events)||data.events.length>50000)throw Error('Complete doctor calendar artifact unavailable.');
     return data.events.filter(event=>{const date=String(event.start||'').slice(0,10);return date>=startDate&&date<=endDate&&source.terms.some(t=>!t.empty&&t.termStart<=date&&t.termEnd>=date);}).map(event=>({sourceType:source.sourceType,doctorKey,event}));
    }));return docs.flat();
   }));return groups.flat();
  }};
}

export function doctorRangeGaps(sourceType,terms,startDate,endDate){
 const next=date=>{const d=new Date(date+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+1);return d.toISOString().slice(0,10);};
 const endBoundary=next(endDate),bounds=new Set([startDate,endBoundary]);
 for(const term of terms){for(const date of [term.termStart,next(term.termEnd)])if(date>startDate&&date<endBoundary)bounds.add(date);
  for(const coverage of term.coverage||[])for(const date of [coverage.startDate,next(coverage.endDate)])if(date>startDate&&date<endBoundary)bounds.add(date);
 }
 const dates=[...bounds].sort(),gaps=[];let gap=null;
 for(let i=0;i<dates.length-1;i++){
  const date=dates[i],term=terms.find(t=>t.termStart<=date&&t.termEnd>=date),covered=term&&(term.empty||!Array.isArray(term.coverage)||term.coverage.some(c=>c.startDate<=date&&c.endDate>=date));
  const reason=!term?'history-unavailable':!covered?'coverage-unavailable':'';
  const last=new Date(dates[i+1]+'T12:00:00Z');last.setUTCDate(last.getUTCDate()-1);const end=last.toISOString().slice(0,10);
  if(reason&&gap?.reason===reason)gap.endDate=end;else{if(gap)gaps.push(gap);gap=reason?{sourceType,startDate:date,endDate:end,reason}:null;}
 }
 if(gap)gaps.push(gap);return gaps;
}
