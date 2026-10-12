export const SUBSCRIPTION_ARTIFACT_BYTES=2*1024*1024;
const META_LIMIT=8192;
export function subscriptionBuildWire(input){
 if(input.mode==='publish'&&typeof input.ics==='string'){
  const {ics,...metadata}=input,body=new TextEncoder().encode(ics);
  if(body.length>SUBSCRIPTION_ARTIFACT_BYTES)throw Error('Calendar artifact exceeds its safe byte bound.');
  if(!ics.startsWith('BEGIN:VCALENDAR\r\n')||!ics.endsWith('END:VCALENDAR\r\n')||ics.includes('\0'))throw Error('Invalid calendar artifact.');
  const encoded=new TextEncoder().encode(JSON.stringify({...metadata,bytes:body.length}));
  if(encoded.length>META_LIMIT*3/4)throw Error('Calendar metadata exceeds its safe byte bound.');
  const header=btoa(String.fromCharCode(...encoded)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
  return {headers:{'Content-Type':'text/calendar; charset=utf-8','Content-Length':String(body.length),'x-subscription-metadata':header},body};
 }
 return {headers:{'Content-Type':'application/json'},body:JSON.stringify(input)};
}
export async function readBoundedSubscriptionBytes(request,maximum=SUBSCRIPTION_ARTIFACT_BYTES){
 const length=request.headers.get('content-length');
 if(length&&Number(length)>maximum)throw Error('Calendar artifact exceeds its safe byte bound.');
 if(length!==null&&/^\d+$/.test(length)){
  const bytes=new Uint8Array(await request.arrayBuffer());
  if(bytes.length>maximum||bytes.length!==Number(length))throw Error('Calendar artifact length is inconsistent.');
  return bytes;
 }
 const reader=request.body?.getReader();if(!reader)return new Uint8Array();
 const chunks=[];let total=0;
 while(true){const {done,value}=await reader.read();if(done)break;total+=value.length;
  if(total>maximum){await reader.cancel();throw Error('Calendar artifact exceeds its safe byte bound.');}chunks.push(value);
 }
 const bytes=new Uint8Array(total);let offset=0;for(const chunk of chunks){bytes.set(chunk,offset);offset+=chunk.length;}return bytes;
}
export async function readSubscriptionBuildInput(request){
 try{
  if(request.headers.get('content-type')?.split(';')[0].trim().toLowerCase()==='text/calendar'){
   const encoded=request.headers.get('x-subscription-metadata')||'';
   if(!encoded||encoded.length>META_LIMIT||!/^[A-Za-z0-9_-]+$/.test(encoded))return null;
   const base64=encoded.replace(/-/g,'+').replace(/_/g,'/'),binary=atob(base64+'='.repeat((4-base64.length%4)%4));
   const input=JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(Uint8Array.from(binary,c=>c.charCodeAt(0))));
   if(!input||Array.isArray(input)||input.mode!=='publish'||!Number.isSafeInteger(input.bytes)||input.bytes<0||input.bytes>SUBSCRIPTION_ARTIFACT_BYTES)return null;
   return {...input,rawCalendar:true};
  }
  const bytes=await readBoundedSubscriptionBytes(request,SUBSCRIPTION_ARTIFACT_BYTES+META_LIMIT);
  const input=JSON.parse(new TextDecoder().decode(bytes));return input&&typeof input==='object'&&!Array.isArray(input)?{...input,rawCalendar:false}:null;
 }catch{return null;}
}
