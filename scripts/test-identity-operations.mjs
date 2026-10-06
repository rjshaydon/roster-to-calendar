import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {previewIdentityOperation,commitIdentityOperation,previewIdentityReversal,reverseIdentityOperation,expandApprovedIdentityAliases,resolvePersonId,queryIdentityPeople,queryIdentityPerson,publishIdentityOperation,accountIdentityAliases,calendarIdentityAliases} from '../functions/_lib/doctor-identity.js';
import {auditIdentityBatch,queryIdentityCandidates,rejectIdentityCandidate,restoreIdentityCandidate} from '../functions/_lib/identity-discovery.js';
import {loadPublishedDoctorCalendar,filterSnapshotByIdentityAliases} from '../functions/_lib/published-doctor-calendar.js';
import {onRequestPost as identityMaintenance} from '../functions/api/automation/identity-maintenance.js';
import {onRequest as middleware} from '../functions/_middleware.js';
import {melbourneDateKey} from '../public/static/roster-term-policy.js';
import {identityAuditWindow} from '../public/static/identity-audit-policy.js';
import {onRequestPost as stateHandler} from '../functions/api/state.js';
import {onRequestGet as feedHandler} from '../functions/api/feed.js';
import {createHash} from 'node:crypto';
const sqlite=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort()) sqlite.exec(await readFile(new URL(`../migrations/${name}`,import.meta.url),'utf8'));
// Existing production index, required by the bounded source calendar reader.
sqlite.exec('CREATE INDEX IF NOT EXISTS idx_roster_events_file_doctor ON roster_events(file_id,doctor_key)');
const sql=[];
const db={prepare(query){sql.push(query);return {args:[],bind(...args){assert.ok(args.length<=100,'D1 bind parameter ceiling');this.args=args;return this;},async all(){return {results:sqlite.prepare(query).all(...this.args)};},async first(){return sqlite.prepare(query).get(...this.args)||null;},async run(){if(db.fail && query.includes(db.fail)) throw Error('Injected failure'); return sqlite.prepare(query).run(...this.args);}};},async batch(statements){if(db.beforeBatch){const fn=db.beforeBatch;db.beforeBatch=null;fn();} sqlite.exec('BEGIN');try{for(const s of statements) await s.run();sqlite.exec('COMMIT');}catch(e){sqlite.exec('ROLLBACK');throw e;}}};
sqlite.exec(`INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:a','Aeshan KULARATNE'),('person:b','Aeshan KULURATNE');
 INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('vhh','AESHAN KULARATNE','Aeshan KULARATNE','person:a'),('ddh','AESHAN KULURATNE','Aeshan KULURATNE','person:b');
 INSERT INTO account_profiles(email,real_name) VALUES('one@test','One'),('two@test','Two');
 INSERT INTO account_people(email,person_id) VALUES('one@test','person:a'),('two@test','person:b');
 INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES('one@test','vhh','AESHAN KULARATNE','Aeshan KULARATNE'),('two@test','ddh','AESHAN KULURATNE','Aeshan KULURATNE');
 INSERT INTO subscription_tokens(token,email,created_at) VALUES('keep-token','one@test','fixture');`);
const input={kind:'merge',personIds:['person:a','person:b'],targetId:'person:kularatne-aeshan',preferredName:'Aeshan KULARATNE',reason:'Reviewed spelling variation',actor:'creator@test'};
const identityState=()=>JSON.stringify(sqlite.prepare('SELECT * FROM roster_people ORDER BY person_id').all());
const original=identityState(); const preview=await previewIdentityOperation(db,input);
assert.equal(identityState(),original,'preview is read-only');
assert.equal(preview.requiresAccountConfirmation,true);
await assert.rejects(()=>commitIdentityOperation(db,{...input,previewToken:preview.previewToken}),/confirm the named accounts/);
db.fail='INSERT INTO roster_identity_jobs';
await assert.rejects(()=>commitIdentityOperation(db,{...input,previewToken:preview.previewToken,confirmAccountEmails:preview.accountEmails}),/Injected failure/);
assert.equal(identityState(),original,'failed operation rolls back all identity state'); db.fail='';
const merged=await commitIdentityOperation(db,{...input,previewToken:preview.previewToken,confirmAccountEmails:preview.accountEmails});
assert.equal((await resolvePersonId(db,'person:a')).person_id,input.targetId);
assert.equal((await expandApprovedIdentityAliases(db,[{sourceType:'vhh',key:'AESHAN KULARATNE'}])).length,2);
assert.equal(sqlite.prepare('SELECT email FROM subscription_tokens WHERE token=?').get('keep-token').email,'one@test','subscription URL ownership is unchanged');
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM account_profiles').get().n,2,'accounts are never joined/deleted');
await assert.rejects(()=>commitIdentityOperation(db,{...input,previewToken:preview.previewToken,confirmAccountEmails:preview.accountEmails}),/reserved|stale|retired/);
const edit={kind:'name',personIds:[input.targetId],targetId:input.targetId,preferredName:'Corrected Name',reason:'Name edit',actor:'creator@test'};
const editPreview=await previewIdentityOperation(db,edit);
const edited=await commitIdentityOperation(db,{...edit,previewToken:editPreview.previewToken,confirmAccountEmails:editPreview.accountEmails});
await assert.rejects(()=>previewIdentityReversal(db,merged.operationId),/Later identity changes/);
let reversal=await previewIdentityReversal(db,edited.operationId);
await reverseIdentityOperation(db,{operationId:edited.operationId,previewToken:reversal.previewToken,reason:'Undo name',actor:'creator@test',confirmAccountEmails:reversal.accountEmails});
// Versions remain monotonic, but a fully undone dependency permits a fresh undo.
assert.ok((await previewIdentityReversal(db,merged.operationId)).previewToken);

// Independent fixture verifies immediate exact restoration and permanent IDs.
sqlite.exec("DELETE FROM roster_identity_operation_people; DELETE FROM roster_identity_jobs; DELETE FROM roster_identity_operations; DELETE FROM roster_person_redirects; DELETE FROM account_people; DELETE FROM roster_person_aliases; DELETE FROM roster_people;");
sqlite.exec("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:a','Original'),('person:b','Variant'); INSERT INTO roster_person_aliases(source_type,doctor_key,person_id) VALUES('mmc','ORIGINAL','person:a'),('mch','VARIANT','person:b');");
const fresh=await previewIdentityOperation(db,input);
const operation=await commitIdentityOperation(db,{...input,previewToken:fresh.previewToken});
reversal=await previewIdentityReversal(db,operation.operationId);
db.fail='INSERT INTO roster_identity_jobs'; const mergedState=identityState();
await assert.rejects(()=>reverseIdentityOperation(db,{operationId:operation.operationId,previewToken:reversal.previewToken,reason:'Wrong match',actor:'creator@test'}),/Injected failure/);
assert.equal(identityState(),mergedState); db.fail='';
await reverseIdentityOperation(db,{operationId:operation.operationId,previewToken:reversal.previewToken,reason:'Wrong match',actor:'creator@test'});
assert.equal((await resolvePersonId(db,'person:a')).preferred_display_name,'Original');
assert.equal(sqlite.prepare('SELECT person_id FROM roster_person_aliases WHERE source_type=?').get('mch').person_id,'person:b');
assert.equal(sqlite.prepare('SELECT status FROM roster_people WHERE person_id=?').get(input.targetId).status,'retired');
await assert.rejects(()=>previewIdentityOperation(db,input),/reserved/);
assert.ok((await queryIdentityPerson(db,'person:a')).history.length>=2);
assert.equal(sql.some(q=>/\b(roster_events|roster_daily_presence|roster_file_doctors)\b|CREATE TABLE|ALTER TABLE/.test(q)),false,'identity operations never scan or mutate roster history or do runtime DDL');
const doctors=[
 {sourceType:'vhh',key:'TOBY O BRIEN',displayName:'Toby O BRIEN'},
 {sourceType:'ddh',key:'TOBY OBRIEN',displayName:'Toby O’Brien'},
 {sourceType:'vhh',key:'AESHAN KULARATNE',displayName:'Aeshan KULARATNE'},
 {sourceType:'ddh',key:'AESHAN KULURATNE',displayName:'Aeshan KULURATNE'},
 ...Array.from({length:28},(_,i)=>({sourceType:'mmc',key:`DOCTOR${i} TEST${i}`,displayName:`Doctor${i} Test${i}`})),
];
let audit=await auditIdentityBatch(db,doctors,{register:true,actor:'creator@test'});
assert.equal(audit.examined,5,'registry preparation checkpoints within the production request limit');
while(audit.status!=='complete') audit=await auditIdentityBatch(db,doctors,{register:true,actor:'creator@test',runId:audit.runId});
const toby=sqlite.prepare("SELECT person_id FROM roster_person_aliases WHERE doctor_key LIKE 'TOBY%' ORDER BY doctor_key").all();
assert.equal(toby[0].person_id,toby[1].person_id,'only harmless formatting is automatically linked');
const aeshan=sqlite.prepare("SELECT person_id FROM roster_person_aliases WHERE doctor_key LIKE 'AESHAN%' ORDER BY doctor_key").all();
assert.notEqual(aeshan[0].person_id,aeshan[1].person_id,'spelling changes are never automatically linked');
const beforeScan=identityState(), jobsBefore=sqlite.prepare('SELECT COUNT(*) n FROM roster_identity_jobs').get().n;
audit=await auditIdentityBatch(db,doctors);
while(audit.status!=='complete') audit=await auditIdentityBatch(db,doctors,{runId:audit.runId});
assert.equal(identityState(),beforeScan,'candidate discovery does not change people');
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_identity_jobs').get().n,jobsBefore,'discovery creates no snapshot jobs');
const candidate=(await queryIdentityCandidates(db)).candidates.find(c=>c.left.key.startsWith('AESHAN') && c.right.key.startsWith('AESHAN'));
assert.ok(candidate,'bounded discovery suggests the misspelling for human review');
await rejectIdentityCandidate(db,{pairKey:candidate.pair_key,fingerprint:candidate.fingerprint,actor:'creator@test'});
audit=await auditIdentityBatch(db,doctors);
while(audit.status!=='complete') audit=await auditIdentityBatch(db,doctors,{runId:audit.runId});
assert.equal(sqlite.prepare('SELECT status FROM roster_identity_candidates WHERE pair_key=?').get(candidate.pair_key).status,'rejected','unchanged rejected evidence remains suppressed');
assert.ok((await queryIdentityCandidates(db,{status:'rejected'})).candidates.some(c=>c.pair_key===candidate.pair_key),'dismissed suggestions remain reviewable');
await assert.rejects(()=>restoreIdentityCandidate(db,{pairKey:candidate.pair_key,fingerprint:'stale'}),/changed/);
await restoreIdentityCandidate(db,{pairKey:candidate.pair_key,fingerprint:candidate.fingerprint});
assert.ok((await queryIdentityCandidates(db)).candidates.some(c=>c.pair_key===candidate.pair_key),'undo dismissal restores the pending suggestion without changing identity links');
await rejectIdentityCandidate(db,{pairKey:candidate.pair_key,fingerprint:candidate.fingerprint,actor:'creator@test'});
const pendingPlan=sqlite.prepare("EXPLAIN QUERY PLAN SELECT * FROM roster_identity_features INDEXED BY idx_identity_feature_pending WHERE audited_fingerprint<>fingerprint ORDER BY source_type,doctor_key LIMIT 26").all();
assert.ok(pendingPlan.some(row=>row.detail.includes('idx_identity_feature_pending')));
assert.equal(audit.examined,0,'an unchanged weekly audit examines zero identities');
const repeatedAuditQueries=sql.slice(-200);
assert.ok(repeatedAuditQueries.filter(q=>q.includes('WHERE f.given_block')).length<doctors.length,'unchanged features skip candidate comparisons');
const featurePlan=sqlite.prepare("EXPLAIN QUERY PLAN SELECT * FROM roster_identity_features WHERE given_block=? LIMIT 25").all('AESHAN:KUL');
assert.ok(featurePlan.some(row=>row.detail.includes('idx_identity_feature_given')));
assert.equal(identityAuditWindow(new Date('2026-10-03T16:30:00Z')).eligible,true,'Melbourne Sunday DST change stays inside the audit window');
assert.equal(identityAuditWindow(new Date('2026-04-04T17:30:00Z')).eligible,true,'Melbourne winter Sunday remains 03:30 after DST ends');
assert.equal(identityAuditWindow(new Date('2026-10-03T18:00:00Z')).eligible,false,'no work outside the Melbourne window');
assert.equal(identityAuditWindow(new Date('2026-10-05T16:00:00Z')).eligible,false,'no weekday discovery');

// A merge changes calendar identity selection without changing grades, roster
// rows, personal overrides, source names or subscription URLs.
sqlite.exec(`INSERT INTO roster_files(id,name,source_type,active) VALUES('grade-vhh','VHH history','vhh',1),('grade-ddh','DDH current','ddh',1);
 INSERT INTO roster_events(id,file_id,source_type,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,seniority,event_json) VALUES
 ('grade-old','grade-vhh','vhh','AESHAN KULARATNE','Aeshan KULARATNE','2022-01-01','2022-01-01','2022-01-01','2022-01-01','ED','Junior Registrar','{"id":"grade-old","source":"VHH","doctor":"Aeshan KULARATNE","title":"ED","start":"2022-01-01T09:00:00","end":"2022-01-01T17:00:00","seniority":"Junior Registrar"}'),
 ('grade-new','grade-ddh','ddh','AESHAN KULURATNE','Aeshan KULURATNE','2026-10-05','2026-10-05','2026-10-05','2026-10-05','ED','SMS','{"id":"grade-new","source":"DDH","doctor":"Aeshan KULURATNE","title":"ED","start":"2026-10-05T09:00:00","end":"2026-10-05T17:00:00","seniority":"SMS"}');`);
assert.equal((await queryIdentityPeople(db,{search:'aeshan'})).people.length,3);
sqlite.exec("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:haydon-richard','Richard Haydon'),('person:haydon-other','Other Haydon'),('person:literal-percent','Literal % Name');");
assert.ok((await queryIdentityPeople(db,{search:'hAyDoN'})).people.some(p=>p.person_id==='person:haydon-richard'),'surname substring finds Richard Haydon');
assert.ok((await queryIdentityPeople(db,{search:'chard Hay'})).people.some(p=>p.person_id==='person:haydon-richard'),'substring can span given name and surname');
assert.equal((await queryIdentityPeople(db,{search:'%'})).people.length,1,'search characters are literal, not SQL wildcards');
const page1=await queryIdentityPeople(db,{search:'Haydon',limit:1});
const page2=await queryIdentityPeople(db,{search:'Haydon',limit:1,after:page1.next});
assert.notEqual(page1.people[0].person_id,page2.people[0].person_id,'substring pagination does not repeat or skip the second matching person');
const substringPlan=sqlite.prepare('EXPLAIN QUERY PLAN SELECT * FROM roster_people INDEXED BY idx_identity_person_name WHERE (preferred_display_name,person_id)>(? COLLATE NOCASE,?) ORDER BY preferred_display_name COLLATE NOCASE,person_id LIMIT 251').all('Richard','person:a');
assert.ok(substringPlan.some(p=>p.detail.includes('SEARCH roster_people USING INDEX idx_identity_person_name')),'substring windows seek into the name index');

assert.equal((await queryIdentityPeople(db,{search:'AESHAN KULARATNE',searchType:'alias',sourceType:'vhh'})).people.length,1);
assert.equal((await queryIdentityPeople(db,{search:'AESHAN KULARATNE',searchType:'alias'})).people.length,1,'exact roster-name search can query all five indexed sites without a selector');
const allSitePlan=sqlite.prepare("EXPLAIN QUERY PLAN SELECT p.* FROM roster_person_aliases a JOIN roster_people p ON p.person_id=a.person_id WHERE a.source_type IN ('mmc','mch','ddh','vhh','casey') AND a.doctor_key=?").all('AESHAN KULARATNE');
assert.ok(allSitePlan.some(r=>r.detail.includes('source_type=? AND doctor_key=?')),'all-site search uses the alias primary key');
const namePlan=sqlite.prepare('EXPLAIN QUERY PLAN SELECT * FROM roster_people WHERE preferred_display_name COLLATE NOCASE>=? AND preferred_display_name COLLATE NOCASE<? ORDER BY preferred_display_name COLLATE NOCASE,person_id LIMIT 26').all('Aeshan','Aeshan\uffff');
assert.ok(namePlan.some(row=>row.detail.includes('idx_identity_person_name')));
const eventBefore=JSON.stringify(sqlite.prepare('SELECT * FROM roster_events ORDER BY id').all());
const aliasOwners=sqlite.prepare("SELECT person_id FROM roster_person_aliases WHERE doctor_key LIKE 'AESHAN%' ORDER BY doctor_key").all().map(r=>r.person_id);
const feedMerge={kind:'merge',personIds:aliasOwners,targetId:aliasOwners[0],preferredName:'Aeshan KULARATNE',reason:'Reviewed identity',actor:'creator@test'};
let feedPreview=await previewIdentityOperation(db,feedMerge);
const feedOperation=await commitIdentityOperation(db,{...feedMerge,previewToken:feedPreview.previewToken,confirmAccountEmails:feedPreview.accountEmails});
sqlite.prepare('INSERT INTO account_people(email,person_id) VALUES(?,?)').run('one@test',aliasOwners[0]);
assert.equal((await accountIdentityAliases(db,'one@test',[])).length,2);
const creatorAliases=await calendarIdentityAliases(db,{role:'creator',email:'one@test',doctorKey:'AESHAN KULURATNE',aliases:[]});
assert.ok(creatorAliases.some(a=>a.key==='AESHAN KULARATNE'&&a.sourceType==='vhh'),'Creator without claims or cached aliases retains selected calendar and expands approved aliases');
const unrelatedSelection=await calendarIdentityAliases(db,{role:'creator',email:'one@test',doctorKey:'UNRELATED SELECTED DOCTOR',aliases:[]});
assert.ok(unrelatedSelection.every(a=>a.key==='UNRELATED SELECTED DOCTOR'),'Creator account linkage cannot substitute the selected doctor');

const puts=[]; const r2={async put(key,value){puts.push({key,value});}};
assert.equal((await publishIdentityOperation(db,r2,feedOperation.operationId)).status,'complete');
const completedPuts=puts.length;
assert.equal((await publishIdentityOperation(db,r2,feedOperation.operationId)).status,'complete');
assert.equal(puts.length,completedPuts,'completed publication is idempotent');
assert.equal(puts[0].key,'identity/revision.json');
assert.equal(JSON.stringify(sqlite.prepare('SELECT * FROM roster_events ORDER BY id').all()),eventBefore);
const feed=await feedHandler({env:{ROSTER_DB:db,IDENTITY_REVIEW_ENABLED:'true'},request:new Request('https://example.test/api/feed?token=keep-token')});
const ics=await feed.text(); assert.equal(feed.status,200,ics); assert.ok(ics.includes('20220101') && ics.includes('20261005'),'existing subscription delivers both approved aliases: '+ics);
assert.equal(sqlite.prepare('SELECT seniority FROM roster_events WHERE id=?').get('grade-old').seniority,'Junior Registrar');
sqlite.prepare("UPDATE account_profiles SET role='creator' WHERE email='one@test'").run();
sqlite.prepare("INSERT INTO account_states(email,session_json) VALUES(?,?)").run('one@test',JSON.stringify({doctorKey:'AESHAN KULURATNE'}));
const creatorFeed=await feedHandler({env:{ROSTER_DB:db,IDENTITY_REVIEW_ENABLED:'true'},request:new Request('https://example.test/api/feed?token=keep-token')});
const creatorIcs=await creatorFeed.text();
assert.equal(creatorFeed.status,200,creatorIcs);
assert.ok(creatorIcs.includes('20220101') && creatorIcs.includes('20261005'),'Creator selected subscription expands approved aliases across sites');
sqlite.prepare("UPDATE account_profiles SET role='' WHERE email='one@test'").run();
sqlite.prepare("DELETE FROM account_states WHERE email='one@test'").run();


// Other supported edits use the same transaction and reversal boundary.
sqlite.exec("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:edit-fixture','Fixture'); INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('mmc','FIXTURE','Fixture','person:edit-fixture');");
for(const change of [
 {kind:'name',personIds:['person:edit-fixture'],targetId:'person:edit-fixture',preferredName:'Corrected Fixture'},
 {kind:'id',personIds:['person:edit-fixture'],targetId:'person:fixture-corrected',preferredName:'Fixture'},
 {kind:'alias-move',personIds:['person:edit-fixture'],targetId:'person:fixture-split',preferredName:'Separate Fixture',alias:{sourceType:'mmc',key:'FIXTURE'}},
 {kind:'account-link',personIds:['person:edit-fixture'],targetId:'person:edit-fixture',accountEmail:'two@test'},
]) {
 const editInput={...change,actor:'creator@test',reason:'Synthetic edit verification'};
 const p=await previewIdentityOperation(db,editInput);
 const changed=await commitIdentityOperation(db,{...editInput,previewToken:p.previewToken,confirmAccountEmails:p.accountEmails});
 const detail=await queryIdentityPerson(db,editInput.targetId);
 assert.ok(detail.history.find(h=>h.operation_id===changed.operationId)?.summary.length,'person history explains the actual roster/name/account change');
 const undo=await previewIdentityReversal(db,changed.operationId);
 await reverseIdentityOperation(db,{operationId:changed.operationId,previewToken:undo.previewToken,reason:'Undo fixture edit',actor:'creator@test',confirmAccountEmails:undo.accountEmails});
 assert.equal((await resolvePersonId(db,'person:edit-fixture')).preferred_display_name,'Fixture');
 assert.equal(sqlite.prepare("SELECT person_id FROM roster_person_aliases WHERE source_type='mmc' AND doctor_key='FIXTURE'").get().person_id,'person:edit-fixture');
}
sqlite.exec("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:existing-destination','Existing destination'); INSERT INTO account_people(email,person_id) VALUES('one@test','person:edit-fixture') ON CONFLICT(email) DO UPDATE SET person_id=excluded.person_id;");
const separate={kind:'alias-move',personIds:['person:edit-fixture','person:existing-destination'],targetId:'person:existing-destination',alias:{sourceType:'mmc',key:'FIXTURE'},reason:'Separate wrong roster name',actor:'creator@test'};
const separatePreview=await previewIdentityOperation(db,separate);
const separated=await commitIdentityOperation(db,{...separate,previewToken:separatePreview.previewToken,confirmAccountEmails:separatePreview.accountEmails});
assert.equal(sqlite.prepare("SELECT person_id FROM account_people WHERE email='one@test'").get().person_id,'person:edit-fixture','separation does not silently move the source account');
assert.equal(sqlite.prepare("SELECT person_id FROM roster_person_aliases WHERE source_type='mmc' AND doctor_key='FIXTURE'").get().person_id,'person:existing-destination');
const separateUndo=await previewIdentityReversal(db,separated.operationId);
await reverseIdentityOperation(db,{operationId:separated.operationId,previewToken:separateUndo.previewToken,reason:'Undo separation',actor:'creator@test',confirmAccountEmails:separateUndo.accountEmails});
// A new legacy claim arriving after preview must abort the whole operation.
const raceInput={kind:'name',personIds:['person:edit-fixture'],targetId:'person:edit-fixture',preferredName:'Raced Fixture',actor:'creator@test',reason:'Race verification'};
const racePreview=await previewIdentityOperation(db,raceInput);
db.beforeBatch=()=>sqlite.exec("INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES('two@test','mmc','FIXTURE','Fixture');");
await assert.rejects(()=>commitIdentityOperation(db,{...raceInput,previewToken:racePreview.previewToken}),/changed/);
assert.equal((await resolvePersonId(db,'person:edit-fixture')).preferred_display_name,'Fixture');
sqlite.exec("DELETE FROM account_claims WHERE email='two@test' AND source_type='mmc' AND doctor_key='FIXTURE';");
const salt='fixture',password='fixture-password',hash=createHash('sha256').update(`${salt}:${password}`).digest('hex');
sqlite.prepare('UPDATE account_profiles SET password_salt=?,password_hash=? WHERE email=?').run(salt,hash,'one@test');
sqlite.prepare("INSERT INTO account_profiles(email,real_name,role,password_salt,password_hash) VALUES('creator@test','Creator','creator',?,?)").run(salt,hash);
const request=body=>new Request('https://example.test/api/state',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify(body)});
const disabledQueries=sql.length;
assert.equal((await stateHandler({env:{ROSTER_DB:db},request:request({action:'identity',operation:'list',email:'creator@test',password})})).status,503);
assert.equal(sql.length,disabledQueries,'disabled identity surface must fail before authentication/D1');
assert.equal((await stateHandler({env:{ROSTER_DB:db,IDENTITY_REVIEW_ENABLED:'true'},request:request({action:'identity',operation:'list',email:'one@test',password})})).status,403,'ordinary accounts cannot inspect cross-account identity evidence');
const listResponse=await stateHandler({env:{ROSTER_DB:db,IDENTITY_REVIEW_ENABLED:'true'},request:request({action:'identity',operation:'list',email:'creator@test',password})});
assert.equal(listResponse.status,200); assert.equal((await listResponse.json()).people.length,25);
const oldSnapshot={session:{doctorKey:'OLD'},preview:{events:[{id:'old-owner',source:'VHH',doctor:'Old Doctor',start:'2026-10-05T09:00:00',end:'2026-10-05T17:00:00'},{id:'personal',source:'Custom',title:'Personal',start:'2026-10-05T09:00:00',end:'2026-10-05T17:00:00'}]}};
assert.deepEqual(filterSnapshotByIdentityAliases(oldSnapshot,[]).preview.events.map(e=>e.id),['personal'],'reassignment removes former doctor cache and preserves personal events');
const emptyCalendar=await loadPublishedDoctorCalendar({}, {doctorKey:'OLD',sourceTypes:[],aliases:[],identityAliasesEnforced:true,state:{session:{}}},{range:{startDate:'2026-10-01',endDate:'2026-10-31'},today:'2026-10-05',previousSnapshot:oldSnapshot});
assert.equal(emptyCalendar.snapshotAvailable,true);assert.equal(emptyCalendar.snapshot.preview.events.length,0,'an unlinked identity receives an empty roster instead of its earlier cache');
const eightIds=Array.from({length:8},(_,i)=>'person:eight-'+i);
for(const id of eightIds){sqlite.prepare('INSERT INTO roster_people(person_id,preferred_display_name) VALUES(?,?)').run(id,'Eight fixture');sqlite.prepare('INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES(?,?,?,?)').run('mmc',id,'Eight fixture',id);}
const eightInput={kind:'merge',personIds:eightIds,targetId:'person:eight-target',preferredName:'Eight fixture',reason:'Maximum people scope',actor:'creator@test'};
const eightPreview=await previewIdentityOperation(db,eightInput);
const eightOperation=await commitIdentityOperation(db,{...eightInput,previewToken:eightPreview.previewToken});
const eightUndo=await previewIdentityReversal(db,eightOperation.operationId);
await reverseIdentityOperation(db,{operationId:eightOperation.operationId,previewToken:eightUndo.previewToken,reason:'Undo maximum scope',actor:'creator@test'});
assert.equal((await resolvePersonId(db,eightIds[7])).status,'active','a new canonical ninth row can be reversed for the maximum eight-person selection');
const maintenanceRequest=(mode)=>new Request('https://example.test/api/automation/identity-maintenance',{method:'POST',headers:{authorization:'Bearer fixture-token','content-type':'application/json'},body:JSON.stringify({mode,operationId:feedOperation.operationId})});
const maintenanceContext={env:{ROSTER_DB:db,ROSTER_FILES:r2,IDENTITY_REVIEW_ENABLED:'true',ROSTER_AUTOMATION_TOKEN:'fixture-token',ROSTER_AUTOMATION_WRITES_ENABLED:'false'},request:maintenanceRequest('publish'),waitUntil(){}};
maintenanceContext.next=()=>identityMaintenance(maintenanceContext);
assert.equal((await middleware(maintenanceContext)).status,200,'identity publication does not enable roster imports');
const registryQueries=sql.length;
const unchangedRegistry={async head(){return {etag:'fixture'};},async get(){return {async json(){return {revision:JSON.stringify([melbourneDateKey().slice(0,7),'fixture','fixture','fixture','fixture']),runId:''};}};}};
const unchanged=await identityMaintenance({env:{...maintenanceContext.env,IDENTITY_REGISTRY_ENABLED:'true',ROSTER_FILES:unchangedRegistry},request:maintenanceRequest('register')});
assert.equal((await unchanged.json()).status,'unchanged');assert.equal(sql.length,registryQueries,'unchanged published metadata causes zero D1 work');
const insertSearch=sqlite.prepare('INSERT INTO roster_people(person_id,preferred_display_name) VALUES(?,?)');
sqlite.exec('BEGIN');for(let i=0;i<2500;i++)insertSearch.run('person:search-fixture-'+i,'Directory Fixture '+String(i).padStart(4,'0'));sqlite.exec('COMMIT');
const searchStatements=sql.length;
const incompleteSearch=await queryIdentityPeople(db,{search:'zzzzz-unmatched'});
assert.equal(incompleteSearch.people.length,0);
assert.equal(incompleteSearch.searchIncomplete,true,'large substring searches expose a resumable cursor rather than an unlimited scan');
assert.ok(incompleteSearch.next);
assert.equal(sql.slice(searchStatements).filter(q=>q.includes('LIMIT 251')).length,8,'a request reads no more than eight fixed directory windows');
const resumedSearch=await queryIdentityPeople(db,{search:'zzzzz-unmatched',after:incompleteSearch.next});
assert.equal(resumedSearch.next,'','the remaining bounded directory window completes the search');
console.log('Identity preview, atomic merge/reversal, account conflicts, reserved IDs and bounded history tests passed.');
console.log('Bounded registry, formatting-only linking, suggestion-only discovery and durable rejection tests passed.');
console.log('Creator authorization, disabled gates, unchanged grades, calendar aliases and subscription continuity passed.');
