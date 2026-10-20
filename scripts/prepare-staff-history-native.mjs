// Synthetic-only Free-tier probe. No D1/R2 bindings or real roster data.
import {mkdir,writeFile} from 'node:fs/promises';
import {randomBytes} from 'node:crypto';
import {gzipSync} from 'node:zlib';
const out=process.argv[2];if(!out?.startsWith('/'))throw Error('Absolute disposable directory required.');
await mkdir(out,{recursive:true,mode:0o700});const gate=randomBytes(32).toString('hex');await writeFile(out+'/gate',gate,{mode:0o600});
const root=process.cwd(),revision='a'.repeat(64),objects={};
const encode=value=>gzipSync(JSON.stringify(value)).toString('base64');
const members=Array.from({length:512},(_,i)=>({doctorKey:'SYNTHETIC '+i,displayName:'Synthetic Doctor '+i,seniority:'HMO',contributionCount:1,membershipSource:'roster'}));
const term={termStart:'2026-08-03',termEnd:'2026-11-01',staffKey:'fixture/staff',staffRevision:revision};
objects['facility-overview/v1/mmc/manifest.json']=encode({terms:[term],days:Object.fromEntries(Array.from({length:120},(_,i)=>['day'+i,{key:'fixture/'+i,revision}]))});objects['fixture/staff']=encode({members,seniorityOverrides:[]});
objects[`staff-history/v1/${revision}/names/0.json.gz`]=encode({revision,next:null,total:128,missing:[],people:Array.from({length:128},(_,slot)=>({slot,displayName:'Synthetic Doctor '+slot,names:Array.from({length:16},(_,i)=>'Synthetic alias '+slot+' '+i)}))});
objects[`staff-history/v1/${revision}/histories/0.json.gz`]=encode({people:Array.from({length:16},(_,slot)=>({slot,displayName:'Synthetic Doctor '+slot,history:Array.from({length:32},(_,i)=>({sourceType:'mmc',termStart:'2026-02-02',termEnd:'2026-05-03',grade:'HMO',rosterNames:['Synthetic Doctor '+slot],upcoming:false}))}))});
await writeFile(out+'/worker.mjs',`import {readStaffHistoryView,describeStaffHistory,staffHistorySourceTerms,readStaffHistoryFragment,staffHistoryWatermark} from ${JSON.stringify(root+'/functions/_lib/staff-history-cache.js')};
import {storeCachedSnapshot} from ${JSON.stringify(root+'/functions/_lib/d1-calendar.js')};
const blobs=${JSON.stringify(objects)};
export default {async fetch(request){
 if(request.method!=='POST'||request.headers.get('Authorization')!=='Bearer ${gate}')return new Response(null,{status:404});
 const mode=new URL(request.url).pathname.slice(1),r2={async head(key){return blobs[key]?{etag:key}:null;},async get(key){if(!blobs[key])return null;const bytes=Uint8Array.from(atob(blobs[key]),c=>c.charCodeAt(0));return {arrayBuffer:async()=>bytes.buffer};},async put(){return {etag:'synthetic'};}};
 const today='2026-10-20',watermark=await staffHistoryWatermark(r2,today);
 const get=r2.get;r2.get=async key=>key==='staff-history/v1/current.json'?{json:async()=>({inputRevision:watermark.inputRevision,revision:'${revision}',pages:1,count:16,missing:[]})}:get(key);
 try{
 let result;
 if(mode==='names')result=await readStaffHistoryView(r2,{mode:'names'},today);
 else if(mode==='histories')result=await readStaffHistoryView(r2,{mode:'histories',slots:[0,1,2,3,4,5,6,7,8,9,10,11]},today);
 else if(mode==='terms')result=await staffHistorySourceTerms(r2,'mmc',today);
 else if(mode==='fragment')result=await readStaffHistoryFragment(r2,{...${JSON.stringify(term)},sourceType:'mmc'});
 else if(mode==='store'){const object=await r2.get('staff-history/v1/${revision}/histories/0.json.gz');const text=await new Response(new Blob([await object.arrayBuffer()]).stream().pipeThrough(new DecompressionStream('gzip'))).json();await storeCachedSnapshot(r2,'synthetic/output',text);result={ok:true};}
 else return new Response(null,{status:404});
 // Include the same JSON serialization cost as the real response.
 const encoded=JSON.stringify(result);return Response.json({ok:true,bytes:encoded.length});
 }catch{return Response.json({ok:false},{status:500});}
}};`);
await writeFile(out+'/wrangler.toml','name = "roster-staff-history-native-20261020"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[observability]\nenabled = true\nhead_sampling_rate = 1\n');
console.log('Prepared synthetic staff-history CPU probe without production bindings.');
