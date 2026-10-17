// Private cookies/gate remain in memory; evidence contains aggregate costs only.
import fs from 'node:fs';import assert from 'node:assert/strict';
const [origin,gatePath,output]=process.argv.slice(2);
if(!/^https:\/\/roster-session-native-validation\.[a-z0-9-]+\.workers\.dev$/.test(origin||'')||!gatePath?.startsWith('/')||!output?.startsWith('/'))throw Error('Provide isolated HTTPS origin, private gate file and evidence file.');
const gate=fs.readFileSync(gatePath,'utf8').trim(),cookies=new Map();
async function pool(n,operation){let next=0;const results=Array(n);await Promise.all(Array.from({length:Math.min(10,n)},async()=>{while(next<n){const i=next++;results[i]=await operation(i);}}));return results;}
async function call(mode,i){const start=Date.now();const response=await fetch(origin+'/'+mode+'?i='+i,{method:'POST',redirect:'error',headers:{Authorization:'Bearer '+gate,...(cookies.has(i)?{Cookie:cookies.get(i)}:{})},signal:AbortSignal.timeout(30000)});const data=await response.json();if(response.headers.has('Set-Cookie'))cookies.set(i,response.headers.get('Set-Cookie').split(';')[0]);return{...data,wallMs:Date.now()-start};}
function summary(samples){const times=samples.map(x=>x.wallMs).sort((a,b)=>a-b);return{requests:samples.length,failures:samples.filter(x=>!x.ok||x.status!==200).length,wallMs:{p50:times[Math.floor(times.length*.5)],p95:times[Math.floor(times.length*.95)],max:times.at(-1)},maxStatements:Math.max(...samples.map(x=>x.meter?.statements||0)),maxRowsRead:Math.max(...samples.map(x=>x.meter?.rowsRead||0)),rowsWritten:samples.reduce((n,x)=>n+(x.meter?.rowsWritten||0),0)};}
const evidence={at:new Date().toISOString(),compatibilityDate:'2026-04-22',syntheticOnly:true,providerCalls:false,productionChanged:false,cpuMeasured:false,workloads:[]};
try{
 for(const users of [100,500]){
  const first=cookies.size;const issued=await pool(users-first,i=>call('issue',first+i));assert.ok(issued.every(x=>x.ok));
  const reads=await pool(users*2,i=>call('state',i%users));const result={users,...summary(reads)};evidence.workloads.push(result);assert.equal(result.failures,0);assert.equal(result.rowsWritten,0);
 }
 assert.equal((await call('forged',0)).status,403);
 assert.equal((await call('logout',1)).status,200);
 assert.equal((await call('state',1)).status,401);
 evidence.actorAndRevocationChecks=true;evidence.passed=true;
}catch{evidence.passed=false;}
fs.writeFileSync(output,JSON.stringify(evidence,null,2)+'\n',{mode:0o600});console.log(JSON.stringify(evidence));if(!evidence.passed)process.exitCode=1;
