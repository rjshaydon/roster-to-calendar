import {pathToFileURL} from 'node:url';
import {createHash} from 'node:crypto';
import {buildStaffHistoryCatalogue,STAFF_HISTORY_PAGE,STAFF_HISTORY_SHARD} from '../public/static/staff-history.js';
import {guardedFetch} from '../functions/_lib/outbound-network.js';
export async function prepareStaffHistory({request,maxAliasPages=64}){
 const initial=await request({mode:'describe'});if(['unchanged','disabled','deferred'].includes(initial.status))return initial;
 const {descriptor}=initial,fragments=[],missing=[...descriptor.missing],terms=[];
 for(const sourceType of descriptor.sources){const source=await request({mode:'terms',sourceType,inputRevision:descriptor.inputRevision});terms.push(...source.terms);missing.push(...source.missing);}
 if(terms.length>320)throw Error('Published staff history exceeds the term limit.');
 for(const term of terms){
  const {fragment}=await request({mode:'fragment',inputRevision:descriptor.inputRevision,sourceType:term.sourceType,termStart:term.termStart});
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
 let lastAdmission=0,lastMode='describe',lastHttpStatus=0;
 try{
 const request=async body=>{
  lastMode=body.mode;lastHttpStatus=0;
  if(['fragment','aliases','finish'].includes(body.mode)&&Date.now()-lastAdmission>300000){
   const response=await guardedFetch(process.env,base+'/api/automation/account-budget',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify({reuseCurrentGrant:true}),signal:AbortSignal.timeout(120000)},{label:'Staff history admission'});
   const value=await response.json();if(!response.ok||value.deferred)throw Error('Staff history preparation deferred.');lastAdmission=Date.now();
  }
  const response=await guardedFetch(process.env,base+'/api/automation/staff-history',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify(body),signal:AbortSignal.timeout(120000)},{label:'Staff history checkpoint'});
  lastHttpStatus=response.status;
  if(!response.ok)throw Error('Staff history preparation deferred or superseded.');
  const value=await response.json();if(['stale','deferred'].includes(value.status))throw Error('Staff history preparation deferred or superseded.');return value;
 };
 const result=process.argv.includes('--dispatch')?await request({mode:'enqueue'}):await prepareStaffHistory({request});
 console.log(JSON.stringify({event:'staff-history',status:result.status,count:result.count}));
 }catch{console.log(JSON.stringify({event:'staff-history',status:'deferred',mode:lastMode,httpStatus:lastHttpStatus}));process.exitCode=1;}
}
