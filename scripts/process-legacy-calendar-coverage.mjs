import {gunzipSync} from 'node:zlib';
import {summarizeLegacyCoverage,legacyCoverageKey,LEGACY_COVERAGE_VERSION} from '../functions/_lib/legacy-calendar-coverage.js';
export async function prepareLegacyCalendarCoverage({request}){
 let prepared=0;
 for(;;){
  const page=await request({mode:'legacy-page'});if(page.status==='complete')return {status:'complete',prepared};
  if(page.status!=='running')throw Error('Legacy coverage inventory failed.');
  for(const item of page.pending){
   const bytes=Buffer.from(await request({...item,mode:'legacy-object'},{binary:true}));
   let summary;
   try{
    if(bytes.length>64*1024*1024)throw Error('Coverage input exceeds the preparation bound.');
    const snapshot=JSON.parse(bytes.subarray(0,2).equals(Buffer.from([31,139]))?gunzipSync(bytes,{maxOutputLength:64*1024*1024}).toString('utf8'):bytes.toString('utf8'));
    summary=summarizeLegacyCoverage(snapshot,item.key,item.revision);
   }catch{summary={version:LEGACY_COVERAGE_VERSION,key:item.key,revision:item.revision,safe:false,bounds:[]};}
   summary.scope=await legacyCoverageKey(item.key,item.revision);delete summary.revision;
   const result=await request({...item,mode:'legacy-publish',summary});
   if(result.status!=='complete')throw Error('Legacy coverage did not complete.');prepared++;
  }
  await request({mode:'legacy-advance',cursor:page.cursor,complete:page.complete});
  if(page.complete)return {status:'complete',prepared};
 }
}
