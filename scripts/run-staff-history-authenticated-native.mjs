import fs from 'node:fs';
const dir=process.argv[2],base='https://roster-staff-history-auth-native-20261020.rhaydon.workers.dev';
if(!dir?.startsWith('/'))throw Error('Private disposable directory required.');
const gate=fs.readFileSync(dir+'/gate','utf8'),cookies={};
for(const kind of ['creator','sms','trainee','disabled']){
 const r=await fetch(base+'/'+(kind==='creator'?'issue':'issue-'+kind),{method:'POST',headers:{Authorization:'Bearer '+gate}});
 cookies[kind]=r.headers.get('Set-Cookie')?.split(';')[0];if(!r.ok||!cookies[kind])throw Error('Synthetic session unavailable.');
}
const from=new Date().toISOString(),results=[];
for(const kind of ['creator','sms'])for(const mode of ['full-names','full-histories']){
 const requests=[];
 for(let i=0;i<12;i++){
  const r=await fetch(base+'/'+mode,{method:'POST',headers:{Authorization:'Bearer '+gate,Cookie:cookies[kind]}}),data=await r.json();
  const meter=JSON.parse(r.headers.get('X-Test-Meter')||'{}');
  if(!r.ok||data.preparing||data.people?.length!==(mode==='full-names'?128:12))throw Error('Synthetic authenticated history failed.');
  requests.push({status:r.status,people:data.people.length,meter});
 }
 results.push({kind,mode,requests});
}
const denied=[];
for(const [kind,target]of [['trainee',''],['disabled',''],['creator','staff-history-trainee@example.test'],['sms','staff-history-probe@example.test']]){
 const r=await fetch(base+'/full-names',{method:'POST',headers:{Authorization:'Bearer '+gate,Cookie:cookies[kind],...(target?{'X-Test-Target':target}:{})}}),data=await r.json();
 if(r.status!==403||data.people?.length)throw Error('Synthetic permission denial failed.');
 denied.push({kind,enteredAccount:!!target,status:r.status});
}
const evidence={from,until:new Date().toISOString(),syntheticOnly:true,results,denied};fs.writeFileSync(dir+'/stream-evidence.json',JSON.stringify(evidence,null,2));
console.log(JSON.stringify({from,until:evidence.until,requests:results.map(r=>({kind:r.kind,mode:r.mode,passed:r.requests.length,firstMeter:r.requests[0].meter,lastMeter:r.requests.at(-1).meter})),denied}));
