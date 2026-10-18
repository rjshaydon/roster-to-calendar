import fs from 'node:fs';
const directory=process.argv[2],mode=process.argv[3]||'normal';if(!directory?.startsWith('/tmp/')||!['normal','100','500','expanded','creation','review'].includes(mode))throw Error('Provide private /tmp test directory and normal/100/500/expanded/review.');
const base='https://roster-requests-native-20261018.rhaydon.workers.dev',gate=fs.readFileSync(directory+'/gate','utf8');
const evidence=[];async function raw(stage,input){const start=Date.now(),r=await fetch(base+'/'+stage,{method:'POST',headers:{Authorization:'Bearer '+gate,'Content-Type':'application/json'},body:JSON.stringify(input)});const data=await r.json();evidence.push({at:new Date(start).toISOString(),stage,status:r.status,meter:data.meter,wallMs:Date.now()-start});if(!r.ok)throw Error(stage+': '+JSON.stringify({status:r.status,code:data.code,error:data.error}));return data;}
const fixtures=(await raw('seed',{})).fixtures;if(['100','500'].includes(mode))await raw('scale',{count:Number(mode)});if(mode==='expanded')await raw('expand',{});
if(mode==='review'){await raw('scale',{count:500});await raw('expand',{});}
async function call(actor,stage,body){return raw(stage,{cookie:fixtures[actor].cookie,body:{...body,expectedAccount:fixtures[actor].email}})}
if(mode==='creation'){
 for(let i=0;i<5;i++)for(const kind of ['support','identity'])await call('user','create-'+kind,{action:'create',kind,subject:'Synthetic new '+kind,body:'Synthetic explanation. '+('x'.repeat(i===0?3900:40)),mutationId:crypto.randomUUID()});
 fs.writeFileSync(directory+'/evidence-'+mode+'.json',JSON.stringify(evidence,null,2));console.log(JSON.stringify({mode,requests:evidence.length,allPassed:true}));process.exit(0);
}
const created=await call('user','create',{action:'create',kind:'identity',subject:'Synthetic review',body:'Synthetic ownership review only',alias:{sourceType:'mmc',key:'SYNTHETIC DOCTOR'},mutationId:crypto.randomUUID()});const caseId=created.caseId;
for(let i=0;i<5;i++){
 if(mode==='review'&&i>=3)await raw('bounds',{count:i===3?8:16});
 if(mode==='review')await call('user','create-identity',{action:'create',kind:'identity',subject:'Synthetic duplicate review',body:'Synthetic duplicate only',alias:{sourceType:'mmc',key:'SYNTHETIC DOCTOR'},mutationId:crypto.randomUUID()});
 await call('admin','summary',{action:'summary'});await call('user','list',{action:'list'});
 let detail=await call('admin','detail',{action:'detail',caseId});
 let preview=(await call('admin','transfer-preview',{action:'previewResolution',caseId,version:detail.case.version,kind:'transfer'})).preview;
 await call('admin','transfer-commit',{action:'commitResolution',caseId,version:detail.case.version,kind:'transfer',previewToken:preview.previewToken,reason:'Synthetic reviewed transfer',mutationId:crypto.randomUUID()});
 detail=await call('admin','detail',{action:'detail',caseId});
 preview=(await call('admin','reverse-preview',{action:'previewResolution',caseId,version:detail.case.version,kind:'reverse'})).preview;
 await call('admin','reverse-commit',{action:'commitResolution',caseId,version:detail.case.version,kind:'reverse',previewToken:preview.previewToken,reason:'Synthetic reviewed reversal',mutationId:crypto.randomUUID()});
 await call('user','reply',{action:'reply',caseId,body:'Synthetic reply '+i,mutationId:crypto.randomUUID()});
 detail=await call('admin','detail',{action:'detail',caseId});await call('admin','read',{action:'read',caseId,channelId:detail.channel.channelId,sequence:detail.channel.lastSeq});
}
fs.writeFileSync(directory+'/evidence-'+mode+'.json',JSON.stringify(evidence,null,2));console.log(JSON.stringify({mode,requests:evidence.length,allPassed:true}));
