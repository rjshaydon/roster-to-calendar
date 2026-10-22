import {appendFile} from 'node:fs/promises';
import {createMaintenanceAdmission} from './automation-checkpoint-client.mjs';
import {pathToFileURL} from 'node:url';
import {createHash} from 'node:crypto';
import {buildStaffHistoryCatalogue,STAFF_HISTORY_PAGE,STAFF_HISTORY_SHARD} from '../public/static/staff-history.js';
import {guardedFetch} from '../functions/_lib/outbound-network.js';
export class StaffHistoryDeferred extends Error {
 constructor(status='deferred',mode='',httpStatus=200){super('Staff history preparation postponed.');this.status=status;this.mode=mode;this.httpStatus=httpStatus;}
}
export function createStaffHistoryRequest({base,token,fetch=guardedFetch,env=process.env,now=Date.now}) {
 let lastMode='',lastHttpStatus=0;
 const post=async(path,body)=>{
  lastMode=body.mode||'admission';lastHttpStatus=0;
  const response=await fetch(env,base+path,{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify(body),signal:AbortSignal.timeout(120000)},{label:'Staff history checkpoint'});
  lastHttpStatus=response.status;
  const value=await response.json();
  const expected=value.status==='stale'&&[200,409].includes(response.status)||value.status==='deferred'&&[200,503].includes(response.status);
  if(expected)throw new StaffHistoryDeferred(value.status,lastMode,lastHttpStatus);
  if(!response.ok)throw Error('Staff history checkpoint failed.');
  return value;
 };
 const admission=createMaintenanceAdmission({request:body=>post('/api/automation/account-budget',body),now});
 const request=async body=>{
  if(['repair-grades','fragment','aliases','finish'].includes(body.mode)){
   const value=await admission();
   if(value.deferred||value.paused)throw new StaffHistoryDeferred('deferred','admission',200);
  }
  return post('/api/automation/staff-history',body);
 };
 request.diagnostic=()=>({mode:lastMode,httpStatus:lastHttpStatus});
 return request;
}
export async function runStaffHistory({request,dispatch=false}) {
 try{return dispatch?await request({mode:'enqueue'}):await prepareStaffHistory({request});}
 catch(error){if(error instanceof StaffHistoryDeferred)return {status:error.status,mode:error.mode,httpStatus:error.httpStatus};throw error;}
}
export async function prepareStaffHistory({request,maxAliasPages=64,repairPass=0}){
 const initial=await request({mode:'describe'});if(['unchanged','disabled','deferred'].includes(initial.status))return initial;
 const {descriptor}=initial,fragments=[],missing=[...descriptor.missing],terms=[];
 for(const sourceType of descriptor.sources){const source=await request({mode:'terms',sourceType,inputRevision:descriptor.inputRevision});terms.push(...source.terms);missing.push(...source.missing);}
 if(terms.length>320)throw Error('Published staff history exceeds the term limit.');
 for(const term of terms){
  const {fragment}=await request({mode:'fragment',inputRevision:descriptor.inputRevision,sourceType:term.sourceType,termStart:term.termStart});
  if(!fragment.missing&&fragment.members.some(m=>!m.grade||['unknown','not recorded'].includes(m.grade.toLowerCase()))){
   let repaired;for(let step=0;step<128;step++){repaired=await request({mode:'repair-grades',sourceType:term.sourceType,termStart:term.termStart,staffRevision:term.staffRevision});if(repaired.status!=='continue')break;}
   if(repaired.status==='source-repaired'){if(repairPass>=16)throw Error('Grade repair restart limit.');return prepareStaffHistory({request,maxAliasPages,repairPass:repairPass+1});}
   if(['deferred','stale'].includes(repaired.status))throw new StaffHistoryDeferred(repaired.status,'repair-grades');
   if(repaired.status!=='complete')throw Error('Unexpected grade repair result.');
  }
  if(fragment.missing)missing.push({sourceType:term.sourceType,termStart:term.termStart,reason:'staff-unavailable'});else fragments.push(fragment);
 }
 // Day/month publication tags may change while the staff history stays the
 // same. A semantic checkpoint avoids re-exporting identities or rebuilding
 // catalogue pages when only shift times/unchanged staff metadata changed.
 const projectionInput=createHash('sha256').update(JSON.stringify([1,descriptor.identityMarker,missing,fragments.map(f=>[f.sourceType,f.termStart,f.termEnd,f.members.map(m=>[m.doctorKey,m.displayName,m.grade])])])).digest('hex');
 if(initial.projectionInput===projectionInput)return request({mode:'acknowledge',inputRevision:descriptor.inputRevision,projectionInput});
 const aliases=[];let cursor={},identityRevision;
 for(let page=0;page<maxAliasPages;page++){
  const result=await request({mode:'aliases',...cursor,identityRevision});
  identityRevision=result.identityRevision;aliases.push(...result.aliases);
  if(!result.next)break;
  if(page===maxAliasPages-1)throw Error('Alias page limit exceeded.');
  cursor=result.next;
 }
 const people=buildStaffHistoryCatalogue(fragments,aliases,descriptor.today);
 if(people.length>4096)throw Error('Staff history catalogue limit exceeded.');
 const revision=createHash('sha256').update(JSON.stringify({people,missing})).digest('hex');
 for(let index=0;index<Math.max(1,Math.ceil(people.length/STAFF_HISTORY_PAGE));index++)await request({mode:'store',revision,kind:'names',index,metadata:{revision,missing,next:index+1<Math.ceil(people.length/STAFF_HISTORY_PAGE)?index+1:null,total:people.length},people:people.slice(index*STAFF_HISTORY_PAGE,(index+1)*STAFF_HISTORY_PAGE).map(({history,...p})=>p)});
 for(let index=0;index<Math.ceil(people.length/STAFF_HISTORY_SHARD);index++)await request({mode:'store',revision,kind:'histories',index,people:people.slice(index*STAFF_HISTORY_SHARD,(index+1)*STAFF_HISTORY_SHARD).map(({names,...p})=>p)});
 return request({mode:'finish',descriptor,revision,count:people.length,identityRevision,missing,projectionInput});
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href){
 const base=String(process.env.ROSTER_AUTOMATION_BASE_URL||'https://roster-to-calendar.pages.dev').replace(/\/$/,''),token=process.env.ROSTER_AUTOMATION_TOKEN;
 if(!token)throw Error('ROSTER_AUTOMATION_TOKEN required.');
 const request=createStaffHistoryRequest({base,token});
 try{
  const result=await runStaffHistory({request,dispatch:process.argv.includes('--dispatch')});
  console.log(JSON.stringify({event:'staff-history',status:result.status,count:result.count,mode:result.mode,httpStatus:result.httpStatus}));
  if(process.env.GITHUB_STEP_SUMMARY)await appendFile(process.env.GITHUB_STEP_SUMMARY,`Staff history: ${result.status}. Existing published history remains available.\n`);
 }catch{
  console.error(JSON.stringify({event:'staff-history',status:'failed',...request.diagnostic()}));
  process.exitCode=1;
 }
}
