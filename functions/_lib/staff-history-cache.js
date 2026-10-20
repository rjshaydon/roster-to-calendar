import {loadCachedSnapshot,storeCachedSnapshot} from './d1-calendar.js';
const facilityMetadataManifestKey=source=>`facility-overview/v1/${source}/manifest.json`;
import {rosterTermVisible,melbourneDateKey} from '../../public/static/roster-term-policy.js';
import {STAFF_HISTORY_VERSION,STAFF_HISTORY_PAGE,STAFF_HISTORY_SHARD} from '../../public/static/staff-history.js';
import {guardedFetch} from './outbound-network.js';
export const STAFF_HISTORY_SOURCES=['mmc','mch','ddh','vhh','casey'];
export const STAFF_HISTORY_POINTER='staff-history/v1/current.json';
const ROOT='staff-history/v1/';
export async function historyDigest(value){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value))))].map(v=>v.toString(16).padStart(2,'0')).join('');}
export async function staffHistoryWatermark(r2,today=melbourneDateKey()){
 const heads=await Promise.all([...STAFF_HISTORY_SOURCES.map(facilityMetadataManifestKey),'identity/revision.json'].map(key=>r2.head(key)));
 const identityMarker=heads.at(-1)?.etag||'';
 const sourceEtags=heads.slice(0,-1).map(o=>o?.etag||'');
 return {inputRevision:await historyDigest([STAFF_HISTORY_VERSION,today.slice(0,7),sourceEtags,identityMarker]),identityMarker,sourceEtags,month:today.slice(0,7)};
}
export async function describeStaffHistory(r2,today=melbourneDateKey()){
 const watermark=await staffHistoryWatermark(r2,today);
 return {...watermark,today,sources:STAFF_HISTORY_SOURCES.filter((s,i)=>watermark.sourceEtags[i]),missing:STAFF_HISTORY_SOURCES.filter((s,i)=>!watermark.sourceEtags[i]).map(sourceType=>({sourceType,reason:'manifest-unavailable'}))};
}
export async function staffHistorySourceTerms(r2,sourceType,today=melbourneDateKey()){
 if(!STAFF_HISTORY_SOURCES.includes(sourceType))throw Error('Unknown staff history source.');
 const manifest=await loadCachedSnapshot(r2,facilityMetadataManifestKey(sourceType));
 if(!manifest||!Array.isArray(manifest.terms)||manifest.terms.length>64)throw Error('Staff history source unavailable.');
 const terms=[],missing=[];
 for(const term of manifest.terms)if(rosterTermVisible(term,today)){
  if(term.staffKey&&term.staffRevision)terms.push({sourceType,termStart:term.termStart,termEnd:term.termEnd,staffKey:term.staffKey,staffRevision:term.staffRevision});
  else missing.push({sourceType,termStart:term.termStart,reason:'staff-unavailable'});
 }
 return {terms,missing};
}
export async function readStaffHistoryFragment(r2,term,{loadEvidence}={}){
 const key=ROOT+'fragments/v2/'+term.sourceType+'/'+term.termStart+'/'+term.staffRevision+'.json.gz';
 const cached=await loadCachedSnapshot(r2,key);if(cached)return cached;
 const staff=await loadCachedSnapshot(r2,term.staffKey);
 if(!staff)return {...term,staffKey:undefined,members:[],missing:true};
 if(!Array.isArray(staff.members)||staff.members.length>512||(staff.seniorityOverrides||[]).length>1024)throw Error('Staff history fragment limit exceeded.');
 const overrides=new Map((staff.seniorityOverrides||[]).filter(o=>o.termStart===term.termStart&&!o.useRosterSeniority).map(o=>[o.doctorKey,o.seniority]));
 let members=staff.members;
 const uncertain=members.some(m=>m.membershipSource!=='sms-continuity'&&m.membershipSource!=='provider-only'&&!Number.isFinite(Number(m.contributionCount)));
 // Older publications need real contribution evidence, never a guess based on
 // a senior clinician's continued directory membership. This callback is only
 // available to the metered background runner, not the interactive API.
 if(uncertain){if(!loadEvidence)throw Error('Older staff history requires contribution evidence.');members=await loadEvidence();}
 const fragment={sourceType:term.sourceType,termStart:term.termStart,termEnd:term.termEnd,members:members.filter(m=>Number(m.contributionCount)>0&&m.membershipSource!=='sms-continuity'&&m.membershipSource!=='provider-only').map(m=>({doctorKey:m.doctorKey,displayName:m.displayName,grade:overrides.get(m.doctorKey)||m.seniority||'',contributionCount:m.contributionCount,membershipSource:m.membershipSource}))};
 await storeCachedSnapshot(r2,key,fragment,{revision:term.staffRevision});return fragment;
}
export function historyObjectKey(revision,kind,index){
 if(!/^[a-f0-9]{64}$/.test(revision)||!['names','histories'].includes(kind)||!Number.isInteger(index)||index<0||index>=256)throw Error('Invalid staff history object.');
 return ROOT+revision+'/'+kind+'/'+index+'.json.gz';
}
export async function readStaffHistoryView(r2,input,today=melbourneDateKey(),{streamNames=false}={}){
 const current=await r2.get(STAFF_HISTORY_POINTER);let pointer=current?await current.json():null;
 if(!pointer)return {preparing:true};
 // Published generations remain coherent while a replacement is assembled.
 // Authorization is checked separately on every request. Pin continuation and
 // detail reads to their immutable generation, including during a pointer swap.
 if(input.revision&&input.revision!==pointer.revision){
  if(!/^[a-f0-9]{64}$/.test(input.revision))return {stale:true};
  const pinned=await r2.get(ROOT+input.revision+'/manifest.json');
  if(!pinned)return {stale:true,revision:pointer.revision};
  pointer=await pinned.json();
 }
 if(input.mode==='names'){
  const page=Number(input.page||0);if(!Number.isInteger(page)||page<0||page>=Math.max(1,pointer.pages))throw Error('Invalid staff history page.');
  if(page===0&&input.cachedRevision===pointer.revision)return {unchanged:true,revision:pointer.revision};
  // Offline pages already contain the response envelope. Send their gzip
  // bytes directly; browser decompression/JSON parsing need no Worker CPU.
  if(streamNames){
   const object=await r2.get(historyObjectKey(pointer.revision,'names',page));
   if(!object)return {preparing:true};
   return new Response(object.body||await object.arrayBuffer(),{encodeBody:'manual',headers:{'Content-Type':'application/json','Content-Encoding':'gzip','Cache-Control':'private, no-store'}});
  }
  const data=await loadCachedSnapshot(r2,historyObjectKey(pointer.revision,'names',page));
  if(!data)return {preparing:true};
  return {revision:pointer.revision,people:data.people,missing:pointer.missing,next:page+1<pointer.pages?page+1:null,total:pointer.count};
 }
 if(input.mode!=='histories')throw Error('Invalid staff history view.');
 const slots=[...new Set(input.slots||[])];
 if(!slots.length||slots.length>12||slots.some(n=>!Number.isInteger(n)||n<0||n>=pointer.count))throw Error('Choose up to twelve staff history results.');
 const shards=await Promise.all([...new Set(slots.map(n=>Math.floor(n/STAFF_HISTORY_SHARD)))].map(async index=>[index,await loadCachedSnapshot(r2,historyObjectKey(pointer.revision,'histories',index))]));
 if(shards.some(([,value])=>!value))return {preparing:true};
 const wanted=new Set(slots),people=shards.flatMap(([,value])=>value.people).filter(p=>wanted.has(p.slot));
 return {revision:pointer.revision,people};
}
// Optional preparation never blocks roster publication. A CAS lease prevents
// concurrent source/identity updates from dispatching duplicate runner jobs.
export async function requestStaffHistoryBuild(context){
 const env=context.env,r2=env.ROSTER_FILES;
 if(env.STAFF_HISTORY_BUILD_ENABLED!=='true'||!env.GITHUB_ACTIONS_TOKEN||!r2?.head)return;
 try{
  const watermark=await staffHistoryWatermark(r2),current=await r2.get(STAFF_HISTORY_POINTER);
  if(current&&(await current.json()).inputRevision===watermark.inputRevision)return;
  const key=ROOT+'dispatch.json',old=await r2.get(key),state=old?await old.json():{};
  if(state.until>Date.now())return;
  const saved=await r2.put(key,JSON.stringify({until:Date.now()+20*60000}),{...(old?.etag?{onlyIf:{etagMatches:old.etag}}:{onlyIf:{etagDoesNotMatch:'*'}})});if(saved===null)return;
  await guardedFetch(env,'https://api.github.com/repos/rjshaydon/roster-to-calendar/actions/workflows/staff-history.yml/dispatches',{method:'POST',headers:{Authorization:'Bearer '+env.GITHUB_ACTIONS_TOKEN,Accept:'application/vnd.github+json','Content-Type':'application/json','X-GitHub-Api-Version':'2022-11-28','User-Agent':'roster-to-calendar'},body:JSON.stringify({ref:env.ROSTER_GITHUB_WORKFLOW_REF||'main'}),signal:AbortSignal.timeout(15000)},{label:'Staff history preparation'});
 }catch{console.warn('Staff history preparation deferred.');}
}
export function scheduleStaffHistoryBuild(context){if(context.env.STAFF_HISTORY_BUILD_ENABLED==='true'&&context.waitUntil)context.waitUntil(requestStaffHistoryBuild(context));}

export function staffHistoryAccessAllowed(record,access){
 const creator=['creator','owner'].includes(record?.role);
 return Boolean((creator||record?.facilityOverviewEnabled===true)&&access?.mode==='all'&&(creator||record?.nonClinical!==true&&access?.isSms===true));
}

export function staffHistoryEnabledFor(env,record){
 return env.STAFF_HISTORY_ENABLED==='true'||env.STAFF_HISTORY_ENABLED==='creator'&&['creator','owner'].includes(record?.role);
}
