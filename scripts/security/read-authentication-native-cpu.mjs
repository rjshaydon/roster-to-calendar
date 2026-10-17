// Read-only aggregate telemetry for disposable tests. Never reads Wrangler's
// broad OAuth token, invocation events, request headers or provider credentials.
import fs from 'node:fs';import {execFileSync} from 'node:child_process';import crypto from 'node:crypto';
const [worker,from,output,until]=process.argv.slice(2),start=Date.parse(from),end=until?Date.parse(until):Date.now();
if(!['roster-auth-live-acceptance-20261017','roster-session-native-validation'].includes(worker)||!Number.isFinite(start)||!Number.isFinite(end)||end<start||end>Date.now()||start<Date.now()-86400000||!output?.startsWith('/'))throw Error('Provide a disposable Worker, recent ISO interval and absolute evidence path.');
const token=execFileSync('/usr/bin/security',['find-generic-password','-s','roster-d1-account-analytics','-w'],{encoding:'utf8',stdio:['ignore','pipe','ignore']}).trim();
const body={queryId:crypto.randomUUID(),view:'calculations',calculationMode:'aggregate',ignoreSeries:true,dry:true,timeframe:{from:start,to:end},parameters:{filters:[{key:'$workers.scriptName',operation:'eq',type:'string',value:worker}],groupBys:[{type:'string',value:'$metadata.trigger'}],calculations:[{operator:'count',alias:'requests'},...['median','p95','max'].map(operator=>({key:'$workers.cpuTimeMs',keyType:'number',operator,alias:operator+'CpuMs'}))]}};
const response=await fetch('https://api.cloudflare.com/client/v4/accounts/653047f50279c40888ef9715aa7c4c5f/workers/observability/telemetry/query',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify(body)});
const data=await response.json();if(!response.ok||data.success!==true)throw Error('Read-only aggregate telemetry unavailable.');
const allowed=new Set(['seed','signup','login','upgradeLogin','context','change','oldPassword','returning','logout','cleanup','configuration','inspect','issue','state','forged','status','activity']);
const stages=new Map();
for(const calculation of data.result?.calculations||[])for(const aggregate of calculation.aggregates||[]){const raw=String(aggregate.groupKey||'').replace(/^POST \//,'');const stage=allowed.has(raw)?raw:'other';const current=stages.get(stage)||{stage};const value=Number(aggregate.value);if(['requests','medianCpuMs','p95CpuMs','maxCpuMs'].includes(calculation.alias)&&Number.isFinite(value))current[calculation.alias]=value;stages.set(stage,current);}
const evidence={at:new Date().toISOString(),from,until:new Date(end).toISOString(),source:'Cloudflare Workers Observability aggregate CPU in milliseconds',credential:'existing-read-only-analytics',syntheticOnly:true,stages:[...stages.values()]};
fs.writeFileSync(output,JSON.stringify(evidence,null,2)+'\n',{mode:0o600});console.log(JSON.stringify(evidence));
