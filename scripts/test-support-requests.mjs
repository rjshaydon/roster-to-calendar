import assert from 'node:assert/strict';import {DatabaseSync} from 'node:sqlite';import {readdir,readFile} from 'node:fs/promises';
import {seedSupportFixture,fixtureEmails} from './support-request-fixture.mjs';
import {onRequestPost as route} from '../functions/api/requests.js';
import {accountIdentityAliases} from '../functions/_lib/doctor-identity.js';
import {markSupportRead} from '../functions/_lib/support-requests.js';
const sqlite=new DatabaseSync(':memory:');sqlite.exec('PRAGMA foreign_keys=ON');
for(const n of(await readdir('migrations')).filter(x=>x.endsWith('.sql')).sort())sqlite.exec(await readFile('migrations/'+n,'utf8'));
const statements=[];let beforeBatch,inject='';
const db={prepare(sql){return{sql,args:[],bind(...args){assert.ok(args.length<=100);this.args=args;return this;},async all(){statements.push(sql);return{results:sqlite.prepare(sql).all(...this.args)};},async first(){statements.push(sql);return sqlite.prepare(sql).get(...this.args)||null;},async run(){statements.push(sql);if(inject&&sql.includes(inject))throw Error('Injected failure');if(/^\s*SELECT/.test(sql))return this.all();const result=sqlite.prepare(sql).run(...this.args);return{meta:{changes:Number(result.changes)}};}};},async batch(items){if(beforeBatch&&items.some(item=>!/^\s*SELECT/.test(item.sql))){const f=beforeBatch;beforeBatch=null;f();}sqlite.exec('SAVEPOINT request_test');try{const result=[];for(const item of items)result.push(await item.run());sqlite.exec('RELEASE request_test');return result;}catch(e){sqlite.exec('ROLLBACK TO request_test;RELEASE request_test');throw e;}}};
const fixtures=await seedSupportFixture(db),config={enabled:true,sessionsEnabled:true,requestsEnabled:true,requestTransfersEnabled:true};
const env={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify(config),AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}}};
const mutation=body=>({...body,mutationId:crypto.randomUUID()});
async function call(actor,body,environment=env){const r=await route({env:environment,request:new Request('https://fixture.test/api/requests',{method:'POST',headers:{Origin:'https://fixture.test','Content-Type':'application/json',Cookie:fixtures[actor].cookie},body:JSON.stringify({...body,expectedAccount:fixtures[actor].email})})});return{r,data:await r.json()};}
const ok=result=>{assert.equal(result.r.status,200,JSON.stringify(result.data));return result.data;};
let response=await call('user',{action:'summary'},{...env,ROSTER_AUTH_CONFIG:'{}'});assert.equal(response.r.status,503);
const createdInput=mutation({action:'create',kind:'identity',subject:'Private requester subject',body:'Private requester explanation',alias:{sourceType:'mmc',key:'SYNTHETIC DOCTOR'}});
const created=ok(await call('user',createdInput)),caseId=created.caseId;
assert.equal(sqlite.prepare('SELECT email FROM account_people WHERE person_id=?').get('person:fixture-doctor').email,fixtureEmails.holder,'pending requests never change ownership');
assert.equal(ok(await call('admin',{action:'summary'})).summary.unread,1);
assert.equal(ok(await call('user',{action:'summary'})).summary.unread,0);
assert.equal(ok(await call('user',createdInput)).caseId,caseId);assert.equal(sqlite.prepare('SELECT count(*) n FROM support_cases').get().n,1);
assert.equal((await call('user',{...createdInput,body:'changed'})).r.status,409);
assert.equal((await call('stranger',{action:'detail',caseId})).r.status,404);
let detail=ok(await call('admin',{action:'detail',caseId}));assert.equal(detail.messages[0].body,'Private requester explanation');
assert.equal((await markSupportRead(db,fixtures.admin,{caseId,channelId:detail.channel.channelId,sequence:1},{admitWrite:async()=>false})).readDeferred,true);
assert.equal(ok(await call('admin',{action:'summary'})).summary.unread,1,'a read-budget deferral does not falsely acknowledge unread messages');
await call('admin',{action:'read',caseId,channelId:detail.channel.channelId,sequence:1});assert.equal(ok(await call('admin',{action:'summary'})).summary.unread,0);
const userReply=mutation({action:'reply',caseId,channelId:created.channelId,body:'A second explanation'});ok(await call('user',userReply));ok(await call('user',userReply));
assert.equal(sqlite.prepare('SELECT count(*) n FROM support_messages').get().n,2);
detail=ok(await call('admin',{action:'detail',caseId}));
ok(await call('admin',mutation({action:'inviteHolder',caseId,version:detail.case.version,participantEmail:fixtureEmails.holder,body:'Private holder question'})));
const holderDetail=ok(await call('holder',{action:'detail',caseId}));
assert.equal(holderDetail.case.subject,'Roster identity review');assert.equal(holderDetail.channels,undefined);assert.equal(holderDetail.case.requesterEmail,undefined);
assert.equal(JSON.stringify(holderDetail).includes('Private requester'),false);assert.equal(JSON.stringify(holderDetail).includes(fixtureEmails.user),false);
assert.equal((await call('user',{action:'detail',caseId,channelId:holderDetail.channel.channelId})).data.channel.channelId,created.channelId,'a non-admin supplied channel cannot select another party');
ok(await call('admin',mutation({action:'reply',caseId,channelId:created.channelId,body:'Private admin response to requester'})));
assert.equal(JSON.stringify(ok(await call('holder',{action:'detail',caseId}))).includes('Private admin response'),false);
detail=ok(await call('admin',{action:'detail',caseId}));
sqlite.prepare("INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,'mmc','SYNTHETIC DOCTOR','Synthetic Doctor')").run(fixtureEmails.stranger);
const transfer={action:'previewResolution',caseId,version:detail.case.version,kind:'transfer'},preview=ok(await call('admin',transfer)).preview;
assert.match(preview.previewToken,/^prepared:/);
const previewId=preview.previewToken.slice(9),expiry=sqlite.prepare('SELECT expires_at FROM support_ownership_previews WHERE preview_id=?').get(previewId).expires_at;
sqlite.prepare('UPDATE support_ownership_previews SET expires_at=0 WHERE preview_id=?').run(previewId);
assert.equal((await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'transfer',previewToken:preview.previewToken,reason:'Expired review'}))).r.status,409);
sqlite.prepare('UPDATE support_ownership_previews SET expires_at=? WHERE preview_id=?').run(expiry,previewId);
const resolution=mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'transfer',previewToken:preview.previewToken,reason:'Reviewed ownership correction'});
const privateBefore=JSON.stringify(sqlite.prepare('SELECT email,password_hash,password_salt,subscription_token FROM account_profiles ORDER BY email').all());
const originalClaims=JSON.stringify(sqlite.prepare('SELECT * FROM account_claims').all());
inject='INSERT INTO support_case_events';assert.equal((await call('admin',resolution)).r.status,503);inject='';
assert.ok(sqlite.prepare('SELECT preview_id FROM support_ownership_previews WHERE preview_id=?').get(previewId),'a failed transaction does not consume the reviewed evidence');
assert.equal(JSON.stringify(sqlite.prepare('SELECT * FROM account_claims').all()),originalClaims,'case/message failure rolls back ownership and snapshot changes');
assert.equal(sqlite.prepare('SELECT status FROM support_cases WHERE case_id=?').get(caseId).status,'open');
const commitStart=statements.length,resolved=ok(await call('admin',resolution));assert.equal(sqlite.prepare('SELECT preview_id FROM support_ownership_previews WHERE preview_id=?').get(previewId),undefined,'successful transaction consumes the review');assert.equal(statements.slice(commitStart).some(s=>s.includes('SELECT * FROM roster_person_aliases WHERE person_id=')),false,'commit does not rebuild the reviewed snapshot');assert.equal(JSON.stringify(sqlite.prepare('SELECT email,password_hash,password_salt,subscription_token FROM account_profiles ORDER BY email').all()),privateBefore,'ownership leaves credentials and feed tokens intact');ok(await call('admin',resolution));
assert.equal(sqlite.prepare('SELECT count(*) n FROM roster_identity_operations').get().n,1,'replay does not transfer twice');
const outcomeStart=statements.length;
assert.equal(ok(await call('admin',{action:'outcome',mutationId:resolution.mutationId})).result.operationId,resolved.operationId);
assert.equal(statements.slice(outcomeStart).some(s=>/^\s*(INSERT|UPDATE|DELETE)/.test(s)),false,'acknowledgement recovery is read-only');
assert.equal(ok(await call('admin',{action:'outcome',mutationId:resolution.mutationId},{...env,ROSTER_ACCOUNT_BUDGET_ENABLED:'true'})).result.operationId,resolved.operationId,'an accepted result remains recoverable without a fresh write grant');
assert.equal(ok(await call('holder',{action:'outcome',mutationId:resolution.mutationId})).found,false,'other accounts cannot recover an admin mutation');
const automaticNotice=ok(await call('stranger',{action:'detail',caseId}));assert.equal(automaticNotice.messages[0].body,'Reviewed ownership correction');assert.equal(JSON.stringify(automaticNotice).includes('Private requester'),false);assert.equal(JSON.stringify(automaticNotice).includes(fixtureEmails.user),false);
assert.equal((await accountIdentityAliases(db,fixtureEmails.user,[])).length,2);assert.equal((await accountIdentityAliases(db,fixtureEmails.holder,[])).length,0);
assert.equal(sqlite.prepare("SELECT count(*) n FROM snapshot_registry WHERE status='stale'").get().n,2,'old and new cache invalidation is part of commit');
detail=ok(await call('admin',{action:'detail',caseId}));assert.equal(detail.case.status,'resolved');
const reverse=ok(await call('admin',{action:'previewResolution',caseId,version:detail.case.version,kind:'reverse'})).preview;
sqlite.prepare("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:another-fixture','Another Synthetic Doctor')").run();
beforeBatch=()=>sqlite.prepare("INSERT INTO account_people(email,person_id) VALUES(?,'person:another-fixture')").run(fixtureEmails.stranger);
assert.equal((await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'reverse',previewToken:reverse.previewToken,reason:'Reject obsolete legacy ownership'}))).r.status,409,'a formerly legacy-only owner linked elsewhere cannot regain claims through reversal');
assert.equal(sqlite.prepare('SELECT email FROM account_people WHERE person_id=?').get('person:fixture-doctor').email,fixtureEmails.user,'failed reversal keeps current access intact');
sqlite.prepare('DELETE FROM account_people WHERE email=?').run(fixtureEmails.stranger);
ok(await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'reverse',previewToken:reverse.previewToken,reason:'Reviewed reversal'})));
assert.equal((await accountIdentityAliases(db,fixtureEmails.holder,[])).length,2);assert.equal((await accountIdentityAliases(db,fixtureEmails.user,[])).length,0);
assert.equal(ok(await call('user',{action:'detail',caseId})).case.status,'open');
sqlite.prepare('DELETE FROM account_claims WHERE email=?').run(fixtureEmails.stranger);
// A stale ownership preview, revoked session or private actor substitution
// cannot produce an accepted mutation.
detail=ok(await call('admin',{action:'detail',caseId}));const stale=ok(await call('admin',{...transfer,version:detail.case.version})).preview;
sqlite.prepare("UPDATE roster_people SET version=version+1 WHERE person_id='person:fixture-doctor'").run();
assert.equal((await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'transfer',previewToken:stale.previewToken,reason:'Stale change'}))).r.status,409);
const ownerOnly={...env,ROSTER_AUTH_CONFIG:JSON.stringify({...config,requestsAdminOnly:true})};assert.equal((await call('user',{action:'list'},ownerOnly)).r.status,403);assert.equal((await call('admin',{action:'list'},ownerOnly)).r.status,200);
const denied={...env,ROSTER_AUTH_CONFIG:JSON.stringify({...config,requestTransfersEnabled:false})};assert.equal((await call('admin',resolution,denied)).r.status,403);
const readBefore=statements.length;await call('user',{action:'summary'});assert.equal(statements.slice(readBefore).some(s=>/^\s*(UPDATE|INSERT|DELETE)/.test(s)),false);
// Authority and legacy claims are checked again inside the actual transaction.
const caseCount=sqlite.prepare('SELECT count(*) n FROM support_cases').get().n;
beforeBatch=()=>sqlite.prepare('DELETE FROM account_sessions WHERE token_hash=?').run(fixtures.stranger.token_hash);
assert.equal((await call('stranger',mutation({action:'create',kind:'support',subject:'Revoked session',body:'Must not be saved'}))).r.status,409);
assert.equal(sqlite.prepare('SELECT count(*) n FROM support_cases').get().n,caseCount);
detail=ok(await call('admin',{action:'detail',caseId}));
const concurrent=ok(await call('admin',{...transfer,version:detail.case.version})).preview;
beforeBatch=()=>sqlite.prepare("INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,'mmc','SYNTHETIC DOCTOR','Synthetic Doctor')").run(fixtureEmails.stranger);
assert.equal((await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'transfer',previewToken:concurrent.previewToken,reason:'Concurrent claim test'}))).r.status,409);
assert.equal(sqlite.prepare('SELECT email FROM account_people').get().email,fixtureEmails.holder);
sqlite.prepare('DELETE FROM account_claims WHERE email=?').run(fixtureEmails.stranger);
// Sending can prepare a grant on a quiet roster day without provider/mail or
// mutating the case. The actual write still reserves against the daily cap.
const day=new Date().toISOString().slice(0,10);
sqlite.prepare("INSERT INTO roster_account_budget(utc_day,valid_until,maximum_reads,maximum_writes) VALUES(?,?,4000000,80000)").run(day,new Date(Date.now()+60000).toISOString());
const budgetEnv={...env,ROSTER_ACCOUNT_BUDGET_ENABLED:'true'};
assert.ok(ok(await call('user',{action:'prepare'},budgetEnv)).preparedUntil>Date.now());
sqlite.prepare('UPDATE roster_account_budget SET valid_until=? WHERE utc_day=?').run('',day);
assert.equal((await call('user',{action:'prepare'},budgetEnv)).r.status,503,'missing internal admission cannot silently skip the cap');
const originalFetch=globalThis.fetch;let admissions=0;
globalThis.fetch=async(url,init)=>{assert.equal(String(url),'https://fixture.test/api/automation/account-budget');assert.equal(init.headers.Authorization,'Bearer synthetic-only');assert.deepEqual(JSON.parse(init.body),{reuseCurrentGrant:true});admissions++;return Response.json({ok:true,deferred:true,reason:'private analytics detail'});};
try{const paused=await call('user',{action:'prepare'},{...budgetEnv,ROSTER_AUTOMATION_TOKEN:'synthetic-only'});assert.equal(paused.r.status,503);assert.equal(JSON.stringify(paused.data).includes('private analytics detail'),false);assert.equal(admissions,1);}finally{globalThis.fetch=originalFetch;}
// An unrelated identity operation cannot close this request.
assert.equal((await call('admin',{action:'previewResolution',caseId,version:detail.case.version,kind:'identity',operation:{kind:'account-link',personIds:['person:fixture-doctor'],accountEmail:fixtureEmails.user}})).r.status,409);
const legacy=[{id:'identity:fixture',message:'Unknown historic roster context',firstSeenAt:'2026-01-01T00:00:00.000Z'},{id:'parser:fixture',message:'Leave parser issues alone'}];
sqlite.prepare('UPDATE account_profiles SET admin_issues_json=? WHERE email=?').run(JSON.stringify(legacy),fixtureEmails.user);
const historic=ok(await call('admin',{action:'legacy'}));assert.equal(historic.reports.length,1);assert.equal(historic.reports[0].id,'identity:fixture');
const converted=ok(await call('admin',mutation({action:'convertLegacy',email:fixtureEmails.user,reportId:'identity:fixture'})));
assert.equal(ok(await call('admin',mutation({action:'convertLegacy',email:fixtureEmails.user,reportId:'identity:fixture'}))).caseId,converted.caseId);
assert.equal(ok(await call('admin',{action:'detail',caseId:converted.caseId})).case.personId,'','legacy MMC label is not inferred');
// Large local history/directory fixture: indexed page bounds do not depend on roster rows.
const now=new Date().toISOString();sqlite.exec('BEGIN');
for(let i=0;i<500;i++){const caseKey=crypto.randomUUID(),channelKey=crypto.randomUUID();sqlite.prepare("INSERT INTO account_profiles(email,real_name) VALUES(?,?)").run('scale'+i+'@example.test','Synthetic scale '+i);sqlite.prepare("INSERT INTO support_cases(case_id,requester_email,kind,subject,status,created_at,updated_at) VALUES(?,?,'support','Synthetic historic case','resolved',?,?)").run(caseKey,fixtureEmails.user,now,now);sqlite.prepare("INSERT INTO support_channels(channel_id,case_id,participant_email,case_status,updated_at) VALUES(?,?,?,'resolved',?)").run(channelKey,caseKey,fixtureEmails.user,now);}
sqlite.exec('COMMIT');
const duplicateStart=statements.length,duplicate=ok(await call('user',mutation({...createdInput})));assert.equal(duplicate.alreadyOpen,true);assert.equal(duplicate.caseId,caseId);
assert.equal(statements.slice(duplicateStart).some(s=>/^\s*(INSERT|UPDATE|DELETE)/.test(s)),false,'an already-open case lookup saves no new message or mutation receipt');
const page1=ok(await call('user',{action:'list',status:'resolved'})),page2=ok(await call('user',{action:'list',status:'resolved',cursor:page1.nextCursor}));assert.equal(page1.items.length,20);assert.equal(page2.items.length,20);assert.equal(page1.items.some(a=>page2.items.some(b=>a.case_id===b.case_id)),false);
const plans=[
 ["SELECT c.case_id,h.channel_id FROM support_cases c INDEXED BY idx_support_identity_open JOIN support_channels h ON h.case_id=c.case_id AND h.participant_email=c.requester_email WHERE c.requester_email=? AND c.person_id=? AND c.person_id<>'' AND c.kind='identity' AND c.status='open'",[fixtureEmails.user,'person:fixture-doctor']],
 ["SELECT request_id FROM roster_maintenance_receipts WHERE utc_day=? AND reconciled_at='' AND (finished_at='' OR metadata_complete<>1) LIMIT 3",['2026-10-18']],
 ["SELECT * FROM support_cases WHERE status=? AND has_admin_unread=1 ORDER BY updated_at DESC,case_id DESC LIMIT 21",['open']],
 ["SELECT * FROM support_channels WHERE participant_email=? AND case_status=? ORDER BY updated_at DESC,channel_id DESC LIMIT 21",[fixtureEmails.user,'resolved']],
 ["SELECT * FROM support_channels WHERE participant_email=? AND case_status=? AND waiting=? ORDER BY updated_at DESC,channel_id DESC LIMIT 21",[fixtureEmails.user,'open','admin']],
 ["SELECT * FROM support_messages WHERE channel_id=? AND sequence<? ORDER BY sequence DESC LIMIT 21",[created.channelId,99999]]
];
for(const [sql,args]of plans){const plan=sqlite.prepare('EXPLAIN QUERY PLAN '+sql).all(...args).map(r=>r.detail).join(' ');assert.match(plan,/SEARCH .* (?:INDEX|PRIMARY KEY)/);assert.doesNotMatch(plan,/USE TEMP B-TREE|SCAN support_/);}
// Generic identity correction and reversal are bound to the reviewed case too.
detail=ok(await call('admin',{action:'detail',caseId}));
const rename={kind:'name',personIds:['person:fixture-doctor'],targetId:'person:fixture-doctor',preferredName:'Reviewed Synthetic Doctor',reason:'Name correction after review'};
const namePreview=ok(await call('admin',{action:'previewResolution',caseId,version:detail.case.version,kind:'identity',operation:rename})).preview;
const other=ok(await call('holder',mutation({action:'create',kind:'identity',subject:'Another synthetic case',body:'Separate case',alias:{sourceType:'mmc',key:'SYNTHETIC DOCTOR'}})));
sqlite.prepare('UPDATE support_cases SET version=? WHERE case_id=?').run(detail.case.version,other.caseId);
assert.equal((await call('admin',mutation({action:'commitResolution',caseId:other.caseId,version:detail.case.version,kind:'identity',operation:rename,previewToken:namePreview.previewToken,reason:rename.reason}))).r.status,409,'an identical version/person cannot reuse another case’s preview');
ok(await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'identity',operation:rename,previewToken:namePreview.previewToken,reason:rename.reason})));
assert.equal(sqlite.prepare("SELECT preferred_display_name FROM roster_people WHERE person_id='person:fixture-doctor'").get().preferred_display_name,'Reviewed Synthetic Doctor');
detail=ok(await call('admin',{action:'detail',caseId}));const undoName=ok(await call('admin',{action:'previewResolution',caseId,version:detail.case.version,kind:'reverse'})).preview;
ok(await call('admin',mutation({action:'commitResolution',caseId,version:detail.case.version,kind:'reverse',previewToken:undoName.previewToken,reason:'Reverse synthetic name review'})));
assert.equal(sqlite.prepare("SELECT preferred_display_name FROM roster_people WHERE person_id='person:fixture-doctor'").get().preferred_display_name,'Synthetic Doctor');
// Account deletion clears its private messages, not the other party's thread.
sqlite.prepare('DELETE FROM account_profiles WHERE email=?').run(fixtureEmails.user);
assert.equal(sqlite.prepare('SELECT count(*) n FROM support_channels WHERE participant_email=?').get(fixtureEmails.user).n,0);
assert.equal(sqlite.prepare('SELECT count(*) n FROM support_messages WHERE channel_id=?').get(created.channelId).n,0);
assert.ok(sqlite.prepare('SELECT count(*) n FROM support_messages WHERE channel_id=?').get(holderDetail.channel.channelId).n>0);
assert.equal(ok(await call('holder',{action:'detail',caseId})).case.status,'resolved');
assert.equal((await call('user',{action:'list'})).r.status,401);
assert.equal(statements.some(s=>/FROM roster_events|FROM roster_files/.test(s)),false,'request journeys never scan clinical event tables');
console.log('Support requests passed private-channel isolation, unread/read cursors, idempotency, atomic ownership/cache/case/messages, rollback, safe reversal, stale preview rejection, feature gates and account deletion. Also passed concurrent session/legacy-claim rejection, legacy conversion, indexed 500-account/case pagination. No services or staff data.');
const {default:admission}=await import('../workers/auth-admission/index.js');
const buckets=[];const limiter=name=>({async limit({key}){buckets.push({name,key});return{success:true};}});
for(const mode of ['support','login','passkey'])assert.equal((await admission.fetch(new Request('https://admission.internal/check',{method:'POST',body:mode+':'+('a'.repeat(64))}),{LOGIN_RATE:limiter('login'),PASSKEY_RATE:limiter('passkey'),SUPPORT_RATE:limiter('support')})).status,204);
assert.deepEqual(buckets.map(b=>b.name),['support','login','passkey']);
assert.equal((await admission.fetch(new Request('https://admission.internal/check',{method:'POST',body:'support:'+('a'.repeat(64))}),{SUPPORT_RATE:{async limit(){return{success:false};}}})).status,429);
console.log('Support admission uses its own private bucket and preserves password/passkey admission.');

// The app wrapper must deliver foreground summaries to Requests, not Passkeys.
const vm=await import('node:vm'),app=await readFile('public/static/app.js','utf8');const received=[];
const bridge=vm.createContext({accountTransport:{email:()=>fixtureEmails.admin},received});
await vm.runInContext(app.slice(app.indexOf('let requestController='),app.indexOf('function accountFetch('))+"\nrequestUI.controller=async()=>({receive:data=>received.push(data)});requestUI.receive({summary:{unread:1}})",bridge);
assert.equal(received[0].summary.unread,1);
