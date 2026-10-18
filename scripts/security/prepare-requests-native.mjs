// Disposable synthetic DB only. No production/provider/mail bindings.
import fs from 'node:fs';import crypto from 'node:crypto';
const root=process.cwd(),out=process.argv[2],id=process.argv[3],databaseName=process.argv[4]||'roster-requests-native-20261018';if(!out?.startsWith('/tmp/')||!/^[a-f0-9-]{36}$/.test(id||'')||!/^roster-requests-(?:native|review)-20261018$/.test(databaseName))throw Error('Provide a /tmp directory and a NEW empty D1 UUID.');
fs.mkdirSync(out,{recursive:true,mode:0o700});const gate=fs.existsSync(out+'/gate')?fs.readFileSync(out+'/gate','utf8'):crypto.randomBytes(32).toString('hex');fs.writeFileSync(out+'/gate',gate,{mode:0o600});
fs.writeFileSync(out+'/schema.sql',fs.readdirSync('migrations').filter(n=>n.endsWith('.sql')).sort().map(n=>fs.readFileSync('migrations/'+n,'utf8')).join('\n'));
fs.writeFileSync(out+'/wrangler.toml',`name = "roster-requests-native-20261018"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[observability]\nenabled = true\nhead_sampling_rate = 1\n[[d1_databases]]\nbinding = "TEST_DB"\ndatabase_name = "${databaseName}"\ndatabase_id = "${id}"\n`);
fs.writeFileSync(out+'/worker.mjs',`import {seedSupportFixture,expandSupportFixture,scaleSupportFixture,boundSupportFixture} from ${JSON.stringify(root+'/scripts/support-request-fixture.mjs')};
import {onRequestPost} from ${JSON.stringify(root+'/functions/api/requests.js')};
import {onRequest as middleware} from ${JSON.stringify(root+'/functions/_middleware.js')};
export default{async fetch(request,bindings){
if(request.method!=='POST'||request.headers.get('Authorization')!==${JSON.stringify('Bearer '+gate)})return new Response(null,{status:404});
if(new URL(request.url).pathname==='/seed'){const now=new Date();await bindings.TEST_DB.prepare("INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,4000000,80000,?) ON CONFLICT(utc_day) DO UPDATE SET valid_until=excluded.valid_until").bind(now.toISOString().slice(0,10),new Date(Date.now()+3600000).toISOString()).run();return Response.json({ok:true,fixtures:await seedSupportFixture(bindings.TEST_DB,{reuse:true})});}
if(new URL(request.url).pathname==='/expand'){await expandSupportFixture(bindings.TEST_DB);return Response.json({ok:true});}
if(new URL(request.url).pathname==='/scale')return Response.json({ok:true,...await scaleSupportFixture(bindings.TEST_DB,(await request.json()).count)});
if(new URL(request.url).pathname==='/bounds'){await boundSupportFixture(bindings.TEST_DB,(await request.json()).count);return Response.json({ok:true});}
const meter={statements:0,rowsRead:0,rowsWritten:0};function record(r){meter.statements++;meter.rowsRead+=Number(r.meta?.rows_read||0);meter.rowsWritten+=Number(r.meta?.rows_written||0);return r;}
const db={prepare(sql){return{stmt:bindings.TEST_DB.prepare(sql),bind(...a){this.stmt=this.stmt.bind(...a);return this;},async all(){return record(await this.stmt.all());},async first(){return(await this.all()).results?.[0]||null;},async run(){return record(await this.stmt.run());}};},async batch(items){return(await bindings.TEST_DB.batch(items.map(s=>s.stmt))).map(record);}};
const input=await request.json(),origin=new URL(request.url).origin;
const internal=new Request(origin+'/api/requests',{method:'POST',headers:{Origin:origin,'Content-Type':'application/json',Cookie:input.cookie||''},body:JSON.stringify(input.body)});
const context={env:{ROSTER_DB:db,ROSTER_ACCOUNT_BUDGET_ENABLED:'true',ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,requestsEnabled:true,requestTransfersEnabled:true}),AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}}},request:internal,data:{}};context.next=()=>onRequestPost(context);const response=await middleware(context);
return Response.json({...await response.json(),meter},{status:response.status,headers:{'Cache-Control':'no-store'}});
}};`);
