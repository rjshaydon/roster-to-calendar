// Prepare only: no network, credentials or staff records. Supply a NEW empty DB.
import fs from 'node:fs';import crypto from 'node:crypto';
const root=process.cwd(),out=process.argv[2],id=process.argv[3];
if(!out?.startsWith('/')||!/^[a-f0-9-]{36}$/.test(id||''))throw Error('Provide an absolute disposable directory and NEW disposable D1 UUID.');
fs.mkdirSync(out,{recursive:true,mode:0o700});
const gate=crypto.randomBytes(32).toString('hex'),fixture=crypto.randomUUID();
fs.writeFileSync(out+'/gate',gate,{mode:0o600});
fs.writeFileSync(out+'/worker.mjs',`import {authenticationLiveProbe} from ${JSON.stringify(root+'/scripts/security/authentication-live-probe-worker.mjs')};export default authenticationLiveProbe(${JSON.stringify(gate)},${JSON.stringify(fixture)});`);
fs.writeFileSync(out+'/schema.sql',fs.readdirSync(root+'/migrations').filter(n=>n.endsWith('.sql')).sort().map(n=>fs.readFileSync(root+'/migrations/'+n,'utf8')).join('\n'));
fs.writeFileSync(out+'/wrangler.toml',`name = "roster-auth-live-acceptance-20261017"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[[d1_databases]]\nbinding = "TEST_DB"\ndatabase_name = "roster-auth-live-acceptance-20261017"\ndatabase_id = "${id}"\n`);
