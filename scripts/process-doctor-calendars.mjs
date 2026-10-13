import {createHash} from 'node:crypto';
import {gunzipSync} from 'node:zlib';
import {pathToFileURL} from 'node:url';
import {DOCTOR_CALENDAR_VERSION,DOCTOR_CALENDAR_MAX_BYTES,doctorMonthRowsKey} from '../functions/_lib/doctor-calendar-artifacts.js';
export async function publishDoctorMonth({sourceType,month,revision,request}){
 const scope={sourceType,month,parentRevision:revision};
 if((await request({...scope,mode:'check'})).status==='unchanged')return {status:'unchanged',changed:0};
 const bytes=await request({...scope,mode:'month'},{binary:true});
 const snapshot=JSON.parse(Buffer.from(bytes).subarray(0,2).equals(Buffer.from([31,139]))?gunzipSync(bytes).toString('utf8'):Buffer.from(bytes).toString('utf8'));
 if(snapshot.sourceType!==sourceType||snapshot.month!==month||!Array.isArray(snapshot.rows)||snapshot.rows.length>20000)throw Error('Invalid monthly publication.');
 const groups=new Map();
 for(const row of snapshot.rows){if(row.sourceType!==sourceType||typeof row.doctorKey!=='string'||!row.doctorKey||!row.event||typeof row.event!=='object')throw Error('Monthly identity mismatch.');if(!groups.has(row.doctorKey))groups.set(row.doctorKey,[]);groups.get(row.doctorKey).push(row);}
 if(groups.size>1024)throw Error('Doctor month identity bound exceeded.');
 const previous=await request({...scope,mode:'previous'});
 const doctors=Object.create(null);let changed=0;
 const entries=[...groups].sort(([a],[b])=>a.localeCompare(b));
 for(let offset=0;offset<entries.length;offset+=4){
  const results=await Promise.allSettled(entries.slice(offset,offset+4).map(async([doctorKey,rows])=>{
   if(rows.length>2048)throw Error('Doctor month event bound exceeded.');
   const body=JSON.stringify({version:DOCTOR_CALENDAR_VERSION,sourceType,month,doctorKey,rows}),digest=createHash('sha256').update(body).digest('hex'),key=doctorMonthRowsKey(sourceType,month,digest);
   if(Buffer.byteLength(body)>DOCTOR_CALENDAR_MAX_BYTES)throw Error('Doctor month byte bound exceeded.');
   if(Object.hasOwn(previous.doctors||{},doctorKey)&&previous.doctors[doctorKey].digest===digest&&previous.doctors[doctorKey].key===key){doctors[doctorKey]={digest,key};return;}
   const result=await request({...scope,kind:'rows',digest},{body});if(!['complete','unchanged'].includes(result.status))throw Error('Doctor publication did not complete.');
   if(result.status==='complete')changed++;doctors[doctorKey]={digest,key};
  }));
  for(const result of results)if(result.status==='rejected')throw result.reason;
 }
 // Index is the completion certificate. Failed rows never publish a partial index.
 const orderedDoctors=Object.fromEntries(Object.keys(doctors).sort().map(key=>[key,doctors[key]]));
 const body=JSON.stringify({version:DOCTOR_CALENDAR_VERSION,...scope,doctors:orderedDoctors}),digest=createHash('sha256').update(body).digest('hex');
 if(Buffer.byteLength(body)>DOCTOR_CALENDAR_MAX_BYTES)throw Error('Doctor index byte bound exceeded.');
 const result=await request({...scope,kind:'index',digest,latestEtag:previous.etag||''},{body});if(!['complete','unchanged'].includes(result.status))throw Error('Doctor index did not complete.');
 return {status:'complete',doctors:groups.size,changed};
}
export function doctorCalendarRequest({base,token,deadlineAt=Infinity}){return async(input,{binary=false,metadata=false,body}={})=>{
 if(Date.now()>=deadlineAt)throw Object.assign(Error('Doctor build time bound reached.'),{code:'DOCTOR_BUILD_YIELD'});
 const headers={authorization:'Bearer '+token,'Content-Type':'application/json'};let payload=JSON.stringify(input);
 if(body!==undefined){payload=Buffer.from(body);headers['Content-Type']='application/octet-stream';headers['Content-Length']=String(payload.length);headers['x-doctor-artifact']=Buffer.from(JSON.stringify({...input,bytes:payload.length})).toString('base64');}
 const response=await fetch(base+'/api/automation/doctor-calendars',{method:'POST',headers,body:payload,signal:AbortSignal.timeout(120000)});
 if(!response.ok)throw Error('Doctor calendar callback failed: '+response.status);
 if(binary){const bytes=new Uint8Array(await response.arrayBuffer());return metadata?{bytes,etag:response.headers.get('x-object-etag')||'',key:input.key}:bytes;}return response.json();
};}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href){
 if(!process.env.ROSTER_AUTOMATION_BASE_URL||!process.env.ROSTER_AUTOMATION_TOKEN)throw Error('Doctor calendar credentials unavailable.');
 const {publishDoctorSource}=await import('./process-doctor-source.mjs');
 const base=process.env.ROSTER_AUTOMATION_BASE_URL,token=process.env.ROSTER_AUTOMATION_TOKEN;
 const request=doctorCalendarRequest({base,token,deadlineAt:Date.now()+12*60000});let settled=false;
 try{
  for(const sourceType of ['mmc','mch','ddh','vhh','casey']){
   const sourceManifest=await request({mode:'manifest',sourceType});const {months}=sourceManifest;
   if(sourceManifest.status==='gone')continue;
   for(const [month,pointer] of Object.entries(months))console.log(JSON.stringify({sourceType,month,...await publishDoctorMonth({sourceType,month,revision:pointer.revision,request})}));
   console.log(JSON.stringify({sourceType,...await publishDoctorSource({sourceType,request})}));
  }
  settled=true;console.log(JSON.stringify({status:'complete'}));
 }catch(error){
  if(error.code!=='DOCTOR_BUILD_YIELD')throw error;
  settled=true;console.log(JSON.stringify({status:'yielded',reason:'time-bound'}));
 }finally{
  if(settled&&process.env.DOCTOR_CALENDAR_DISPATCH_ID)await doctorCalendarRequest({base,token})({mode:'finish-source',dispatchId:process.env.DOCTOR_CALENDAR_DISPATCH_ID});
 }
}
