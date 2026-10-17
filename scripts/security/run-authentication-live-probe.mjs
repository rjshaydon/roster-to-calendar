// Synthetic fixtures only. Secrets/cookies stay in memory and never enter the
// evidence file. Always attempts exact fixture cleanup, including after failure.
import fs from 'node:fs';import assert from 'node:assert/strict';
const [origin,gateFile,evidenceFile]=process.argv.slice(2);
if(!/^https:\/\/roster-auth-live-acceptance-20261017\.[a-z0-9-]+\.workers\.dev$/.test(origin||'')||!gateFile?.startsWith('/')||!evidenceFile?.startsWith('/'))throw Error('Provide the exact isolated Worker HTTPS origin, private gate path and absolute evidence path.');
const gate=fs.readFileSync(gateFile,'utf8').trim();if(!/^[a-f0-9]{64}$/.test(gate))throw Error('Invalid private gate.');
const evidence={test:'isolated-synthetic-real-provider',at:new Date().toISOString(),productionChanged:false,emailsSent:false,checks:[],cleanup:false};
async function step(path,kind='new',cookie=''){
 const start=Date.now();const response=await fetch(origin+'/'+path+'?kind='+kind,{method:'POST',redirect:'error',headers:{Authorization:'Bearer '+gate,...(cookie?{Cookie:cookie}:{})},signal:AbortSignal.timeout(30000)});
 const result=await response.json();
 // Copy a fixed aggregate whitelist, never arbitrary provider/application data.
 const row={step:path,kind,status:response.status,wallMs:Date.now()-start,ok:result.ok===true};
 for(const key of ['legacyHashRetired','identityReady','legacyFeedPreserved'])if(typeof result[key]==='boolean')row[key]=result[key];
 if(result.meter)row.meter=Object.fromEntries(['statements','rowsRead','rowsWritten'].map(k=>[k,Number(result.meter[k])||0]));
 evidence.checks.push(row);
 return{response,result,cookie:response.headers.get('Set-Cookie')?.split(';')[0]||''};
}
function accepted(r){assert.equal(r.response.status,200);assert.equal(r.result.ok,true);}
try{
 accepted(await step('seed','legacy'));
 const preparedSignup=await step('signup');assert.equal(preparedSignup.response.status,202);assert.equal(preparedSignup.result.code,'auth-account-prepared');assert.equal(preparedSignup.cookie,'');
 assert.equal((await step('context')).response.status,401,'pending registration grants no calendar access');
 const signup=await step('login');accepted(signup);const signupProof=await step('inspect');accepted(signupProof);assert.equal(signupProof.result.legacyHashRetired,true);assert.equal(signupProof.result.identityReady,true);
 assert.match(signup.response.headers.get('Set-Cookie'),/^__Host-roster_session=[a-f0-9]{64}; Path=\/; Secure; HttpOnly; SameSite=Lax; Max-Age=2592000$/);
 const context=await step('context','new',signup.cookie);accepted(context);assert.equal(context.result.meter.rowsWritten,0);
 accepted(await step('logout','new',signup.cookie));
 assert.equal((await step('context','new',signup.cookie)).response.status,401);
 accepted(await step('login'));
 const legacy=await step('login','legacy');accepted(legacy);const legacyProof=await step('inspect','legacy');accepted(legacyProof);assert.equal(legacyProof.result.legacyHashRetired,true);assert.equal(legacyProof.result.legacyFeedPreserved,true);
 const legacyContext=await step('context','legacy',legacy.cookie);accepted(legacyContext);assert.equal(legacyContext.result.meter.rowsWritten,0);
 const changed=await step('change','legacy',legacy.cookie);accepted(changed);const changeProof=await step('inspect','legacy');accepted(changeProof);assert.equal(changeProof.result.identityReady,true);
 assert.equal((await step('context','legacy',legacy.cookie)).response.status,401);
 assert.equal((await step('oldPassword','legacy')).response.status,401);
 const returning=await step('returning','legacy');accepted(returning);
 accepted(await step('context','legacy',returning.cookie));
 accepted(await step('logout','legacy',returning.cookie));
 accepted(await step('seed','upgrade'));
 const requiredUpgrade=await step('login','upgrade');assert.equal(requiredUpgrade.response.status,409);assert.equal(requiredUpgrade.result.code,'auth-password-upgrade-required');
 const preparedUpgrade=await step('upgradeLogin','upgrade');assert.equal(preparedUpgrade.response.status,202);assert.equal(preparedUpgrade.cookie,'');
 const upgraded=await step('returning','upgrade');accepted(upgraded);const upgradedProof=await step('inspect','upgrade');accepted(upgradedProof);assert.equal(upgradedProof.result.legacyHashRetired,true);assert.equal(upgradedProof.result.legacyFeedPreserved,true);assert.equal(upgradedProof.result.identityReady,true);
 assert.equal((await step('oldPassword','upgrade')).response.status,401);
 accepted(await step('logout','upgrade',upgraded.cookie));
 evidence.acceptancePassed=true;
}catch{evidence.acceptancePassed=false;}
finally{
 try{const cleaned=await step('cleanup');evidence.cleanup=cleaned.response.status===200&&cleaned.result.syntheticFixturesAbsent===true;}catch{evidence.cleanup=false;}
 fs.writeFileSync(evidenceFile,JSON.stringify(evidence,null,2)+'\n',{mode:0o600});
}
console.log(JSON.stringify({acceptancePassed:evidence.acceptancePassed,syntheticCleanupComplete:evidence.cleanup,steps:evidence.checks.length}));
if(!evidence.acceptancePassed||!evidence.cleanup)process.exitCode=1;
