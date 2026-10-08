import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { onRequestPost as stateHandler } from '../functions/api/state.js';
import { onRequest as middleware } from '../functions/_middleware.js';
import { storeCachedSnapshot, loadAccountMirror } from '../functions/_lib/d1-calendar.js';
import { publishedIdentityDirectory, publishedIdentityAuditDirectory, publishedClaimSeniorities, saveBoundedAccountClaims } from '../functions/_lib/bounded-identity.js';
class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; this.sql = []; this.failRunIncludes = ""; }
  prepare(sql) {
    const owner = this;
    owner.sql.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() {
        if (owner.failRunIncludes && sql.includes(owner.failRunIncludes)) {
          owner.failRunIncludes = "";
          throw new Error("Injected D1 statement failure");
        }
        const result = owner.sqlite.prepare(sql).run(...this.args);
        owner.rowsWritten += Number(result.changes || 0);
        return { success: true, meta: { changes: Number(result.changes || 0), rows_read: Number(result.changes || 0) + 1, rows_written: Number(result.changes || 0) } };
      },
      // Metadata is simulated here to test reservation/refund logic, not to
      // claim SQLite result counts measure Cloudflare index billing.
      async all() { const results = owner.sqlite.prepare(sql).all(...this.args); return { success: true, results, meta: { rows_read: results.length + 1, rows_written: 0 } }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) {
    if (this.beforeBatch) { const fn = this.beforeBatch; this.beforeBatch = null; fn(statements); }
    const results = [];
    const ownsTransaction = !this.sqlite.isTransaction;
    if (ownsTransaction) this.sqlite.exec("BEGIN");
    try {
      for (const statement of statements) results.push(await statement.run());
      if (ownsTransaction) this.sqlite.exec("COMMIT");
      return results;
    } catch (error) {
      if (ownsTransaction) this.sqlite.exec("ROLLBACK");
      throw error;
    }
  }
}

class LocalR2 {
  constructor() { this.objects = new Map(); this.puts = 0; this.gets = 0; this.version = 0; this.failPointerOnce = false; }
  async put(key, value, options = {}) {
    if (this.failPointerOnce && key.endsWith("/manifest.json")) { this.failPointerOnce = false; throw new Error("Injected manifest failure"); }
    const current = this.objects.get(key);
    if (options.onlyIf?.etagMatches && current?.etag !== options.onlyIf.etagMatches) throw new Error("Precondition failed");
    if (options.onlyIf?.etagDoesNotMatch === "*" && current) throw new Error("Precondition failed");
    this.puts += 1;
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.version += 1;
    this.objects.set(key, { bytes, options, etag: `local-${this.version}` });
  }
  async get(key) {
    this.gets += 1;
    const item = this.objects.get(key);
    if (!item) return null;
    return { etag: item.etag, arrayBuffer: async () => item.bytes.buffer.slice(item.bytes.byteOffset, item.bytes.byteOffset + item.bytes.byteLength) };
  }
}


const sqlite = new DatabaseSync(':memory:');
for (const name of (await readdir(new URL('../migrations', import.meta.url))).filter(name => name.endsWith('.sql')).sort()) sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), 'utf8'));
const db = new LocalD1(sqlite), r2 = new LocalR2();
const today = new Intl.DateTimeFormat('en-CA', {timeZone:'Australia/Melbourne',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const password = 'fixture-password', salt = 'fixture-salt';
function account(email, realName, role='user', nonClinical=0) {
  sqlite.prepare('INSERT INTO account_profiles(email,real_name,role,password_salt,password_hash,non_clinical) VALUES(?,?,?,?,?,?)').run(email,realName,role,salt,createHash('sha256').update(`${salt}:${password}`).digest('hex'),nonClinical);
  sqlite.prepare('INSERT INTO account_states(email,session_json) VALUES(?,?)').run(email,JSON.stringify({settings:{includeLeave:false},customEvents:[]}));
}
account('alice@example.test','Alice Test'); account('other@example.test','Other Person');
account('creator@example.test','Creator','creator'); account('nonclinical@example.test','Alice Test','user',1);
const alice={key:'ALICE TEST',displayName:'Alice Test',sourceType:'mmc',matchedAt:'fixture'};
const old={key:'OLD NAME',displayName:'Old Name',sourceType:'mch',matchedAt:'fixture'};
// The database has substantial history and unrelated accounts/profiles. None
// should be traversed by login, suggestions or a directory page.
sqlite.exec("BEGIN");
const event=sqlite.prepare('INSERT INTO roster_events(id,file_id,source_type,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,event_json) VALUES(?,?,?,?,?,?,?,?,?,?,?)');
for(let i=0;i<100001;i++) event.run('history-'+i,'old-file','mmc','HISTORICAL','Historical','2020-01-01','2020-01-01','2020-01-01','2020-01-01','Old','{}');
for(let i=0;i<1000;i++) {
  const email=`z${String(i).padStart(4,'0')}@example.test`;
  sqlite.prepare('INSERT INTO account_profiles(email,real_name) VALUES(?,?)').run(email,'Unrelated');
  sqlite.prepare('INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,?,?,?)').run(email,'mmc','UNRELATED '+i,'Unrelated');
  sqlite.prepare('INSERT INTO doctor_profiles(profile_id,doctor_key,display_name) VALUES(?,?,?)').run('unrelated-'+i,'UNRELATED '+i,'Unrelated');
}
sqlite.exec("COMMIT");
sqlite.prepare('INSERT INTO roster_files(id,name,source_type,active) VALUES(?,?,?,?)').run('old-file','Historical.xlsx','mmc',0);
sqlite.prepare('INSERT INTO roster_files(id,name,source_type,active) VALUES(?,?,?,?)').run('active-file','Current.xlsx','mmc',1);
sqlite.exec("INSERT INTO roster_file_doctors(file_id,source_type,doctor_key,display_name) SELECT 'old-file','mmc',id,display_name FROM roster_events");
sqlite.prepare('INSERT INTO roster_file_doctors(file_id,source_type,doctor_key,display_name) VALUES(?,?,?,?)').run('active-file','mmc',alice.key,alice.displayName);
for(const source of ['mmc','mch','ddh','vhh']) {
  const staffKey=`fixture/${source}/staff`, futureKey=`fixture/${source}/future`;
  await storeCachedSnapshot(r2,`facility-overview/v1/${source}/manifest.json`,{terms:[
    {termStart:'2000-01-01',termEnd:'2099-01-01',visibleFrom:'2000-01-01',staffKey},
    {termStart:'2099-01-02',termEnd:'2099-04-01',visibleFrom:'2000-01-01',staffKey:futureKey}],coverage:[]});
  await storeCachedSnapshot(r2,staffKey,{members:source==='mmc' ? [{doctorKey:'ALICE TEST',displayName:'Alice Test',seniority:'HMO'},{doctorKey:'ALICE T TEST',displayName:'Alice T Test',seniority:'Intern'}] : [{doctorKey:'SITE DOCTOR',displayName:'Site Doctor',seniority:'SMS'}],seniorityOverrides:source==='mmc'?[{doctorKey:'ALICE TEST',sourceType:'mmc',seniority:'Junior Registrar',termStart:'2000-01-01'}]:[]});
  await storeCachedSnapshot(r2,futureKey,{members:source==='mmc'?[{doctorKey:'ALICE TEST',displayName:'Alice Test',seniority:'Senior Registrar'}]:[]});
}
const directory=await publishedIdentityDirectory(r2,today);
assert.equal(directory.preparing,false); assert.deepEqual(directory.missingSources,[]);
assert.deepEqual(publishedClaimSeniorities([alice],directory.doctors),['Junior Registrar']);
sqlite.prepare('INSERT INTO roster_doctors(source_type,doctor_key,display_name,updated_at) VALUES(?,?,?,?)').run('ddh','AESHAN KULURATNE','Aeshan KULURATNE',today);
const historicalAuditDirectory=await publishedIdentityAuditDirectory(db,r2,today);
assert(historicalAuditDirectory.doctors.some(d=>d.key==='AESHAN KULURATNE'),'historical-only names reach the audit');
assert(!directory.doctors.some(d=>d.key==='AESHAN KULURATNE'),'historical registration does not expand current department/signup membership');
assert.deepEqual(publishedClaimSeniorities([alice],historicalAuditDirectory.doctors),['Junior Registrar'],'current grades are not replaced by historical records');
const scheduled=[];
async function api(body,email='alice@example.test',flags={}) {
  const request=new Request('https://fixture.test/api/state',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({email,password,...body})});
  const context={request,env:{ROSTER_DB:db,ROSTER_FILES:r2,IDENTITY_DISCOVERY_ENABLED:'true',CREATOR_DIRECTORY_ENABLED:'true',CREATOR_STARTUP_HYDRATION_ENABLED:'false',ACCOUNT_SNAPSHOT_BUILD_ENABLED:'false',...flags},waitUntil(work){scheduled.push(work)}};
  context.next=()=>stateHandler(context);
  const response=await middleware(context);
  return {status:response.status,data:await response.json()};
}
const writesBefore=db.rowsWritten, sqlStart=db.sql.length;
const login=await api({action:'login',responseMode:'fast'});
assert.equal(login.status,200,JSON.stringify(login));
assert.ok(login.data.suggestedClaims.some(claim=>claim.key===alice.key),'unlinked fast login supplies matches before the shell appears');
assert.ok(login.data.availableDoctors.length,'unlinked fast login supplies the name dropdown');
const discovered=await api({action:'loadAccountContext'});
assert.equal(discovered.status,200,JSON.stringify(discovered));
assert.deepEqual(discovered.data.claims,[],'discovery must not silently acquire identities');
assert.ok(discovered.data.suggestedClaims.some(claim=>claim.key===alice.key));
assert.ok(discovered.data.suggestedClaims.some(claim=>claim.key==='ALICE T TEST'),'ambiguous names require human confirmation');
const resolvedDiscovery = await api({action:'resolveAccountClaims'});
assert.ok(resolvedDiscovery.data.suggestedClaims.some(claim=>claim.key===alice.key),'claim resolution must retain discovered matches');
assert.equal(db.rowsWritten,writesBefore,'ordinary discovery writes nothing');
assert.equal(scheduled.length,0,'ordinary discovery schedules no work');
assert.ok(db.sql.length-sqlStart<48,'login and context have a fixed small statement count');
assert.equal(db.sql.slice(sqlStart).some(sql=>/roster_events|canonical_doctors|roster_doctors|sqlite_master/.test(sql)),false);
assert.equal(db.sql.slice(sqlStart).some(sql=>/FROM doctor_profiles ORDER BY/.test(sql)),false,'no full profile scan');
const nonclinical=await api({action:'loadAccountContext'},'nonclinical@example.test');
assert.deepEqual(nonclinical.data.suggestedClaims,[]);
const disabled=await api({action:'claimRosterName',claim:alice},'alice@example.test',{IDENTITY_DISCOVERY_ENABLED:'false'});
assert.equal(disabled.status,503);
const claimResult=await api({action:'claimRosterName',claim:alice});
assert.equal(claimResult.status,200,JSON.stringify(claimResult));
assert.deepEqual(claimResult.data.state.imports.map(file=>file.id),['active-file']);
const claimsBeforeReplay=db.rowsWritten;
assert.equal((await api({action:'claimRosterName',claim:alice})).status,200);
assert.equal(db.rowsWritten,claimsBeforeReplay,'claim replay writes nothing');
assert.equal((await api({action:'claimRosterName',claim:alice},'other@example.test')).status,409);
assert.equal((await api({action:'setAccountRosterClaims',targetEmail:'other@example.test',claims:[alice]},'creator@example.test')).status,409,'Creator cannot create a duplicate owner');
assert.equal((await api({action:'setAccountRosterClaims',targetEmail:'other@example.test',claims:[{...alice,key:'MISSING'}]},'creator@example.test')).status,400,'unknown identities are rejected rather than silently dropping claims');
assert.equal((await api({action:'setAccountRosterClaims',targetEmail:'other@example.test',claims:Array(17).fill(alice)},'creator@example.test')).status,400);
assert.equal((await api({action:'loadAccountContext'},'other@example.test')).data.suggestedClaims.some(claim=>claim.key===alice.key),false);
const directorySql=db.sql.length;
const users=await api({action:'listUsers'},'creator@example.test');
assert.equal(users.status,200,JSON.stringify(users));
assert.deepEqual(users.data.users.find(user=>user.email==='alice@example.test').seniorities,['Junior Registrar']);
assert.ok(users.data.nextCursor);
assert.ok(db.sql.length-directorySql<=4,'directory page has no per-user D1 queries');
assert.equal(db.sql.slice(directorySql).some(sql=>/roster_events|canonical_doctors|doctor_profiles/.test(sql)),false);
// CAS guards must fail inside the same transaction as the changes.
const other=await loadAccountMirror(db,'other@example.test');
const bob={...alice,key:'BOB TEST',displayName:'Bob Test'};
db.beforeBatch=()=>sqlite.prepare('INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,?,?,?)').run('alice@example.test','mmc',bob.key,bob.displayName);
await assert.rejects(saveBoundedAccountClaims(db,other,[bob]),error=>error.code==='IDENTITY_CLAIM_CONFLICT');
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM account_claims WHERE email=?').get(other.email).n,0);
db.beforeBatch=()=>sqlite.prepare('INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,?,?,?)').run(other.email,old.sourceType,old.key,old.displayName);
await assert.rejects(saveBoundedAccountClaims(db,other,[{...bob,key:'NEW'}]),error=>error.code==='IDENTITY_CLAIM_CONFLICT');
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM account_claims WHERE email=?').get(other.email).n,1);
const otherNow=await loadAccountMirror(db,other.email);
db.failRunIncludes='INSERT INTO account_claims';
await assert.rejects(saveBoundedAccountClaims(db,otherNow,[{...bob,key:'NEW'}]),/Injected/);
assert.equal(sqlite.prepare('SELECT doctor_key FROM account_claims WHERE email=?').get(other.email).doctor_key,old.key,'failed replacement retains old links');
// A clear full name goes directly from login to a published multi-site calendar.
account('blocked@example.test','Blocked Exact');
account('reviewed@example.test','Reviewed Exact');
const year=Number(today.slice(0,4)), month=today.slice(0,7);
for (const source of ['mmc','mch']) {
  const staffKey=`automatic/${source}/staff`, months={};
  for(let i=0;i<13;i++) {
    const date=new Date(Date.UTC(year,i,1)), key=date.toISOString().slice(0,7), objectKey=`automatic/${source}/${key}`;
    months[key]={key:objectKey,revision:'fixture'};
    await storeCachedSnapshot(r2,objectKey,{rows:key===month?[{sourceType:source,doctorKey:'ZORA EXACT',event:{id:'zora-'+source,doctorKey:'ZORA EXACT',source:source.toUpperCase(),title:'Day shift',start:today+'T08:00:00',end:today+'T16:00:00'}}]:[]});
  }
  await storeCachedSnapshot(r2,staffKey,{members:[{doctorKey:'ZORA EXACT',displayName:source==='mmc'?'Dr Zora EXACT':'EXACT, Zora',seniority:'HMO'},{doctorKey:'BLOCKED EXACT',displayName:'Blocked Exact',seniority:'HMO'},{doctorKey:'REVIEWED EXACT',displayName:'Reviewed Exact',seniority:'HMO'}]});
  await storeCachedSnapshot(r2,`facility-overview/v1/${source}/manifest.json`,{terms:[{termStart:`${year}-01-01`,termEnd:`${year+1}-01-31`,visibleFrom:'2000-01-01',staffKey}],coverage:[{startDate:`${year}-01-01`,endDate:`${year+1}-01-31`}],months});
}
const automaticStart=db.sql.length;
const automatic=await api({action:'login',mode:'create',realName:'Exact, Zora',responseMode:'fast'},'zora@example.test',{IDENTITY_REVIEW_ENABLED:'true',BOUNDED_MANUAL_ROSTER_ENABLED:'true'});
assert.equal(automatic.status,200,JSON.stringify(automatic));
assert.equal(automatic.data.created,true,'actual signup performs the automatic link');
assert.equal(automatic.data.claims.length,2,'clear full-name matches at both sites are linked');
assert.equal(automatic.data.defaultDoctorKey,'ZORA EXACT');
assert.equal(automatic.data.snapshotAvailable,true,JSON.stringify(automatic));
assert.deepEqual(new Set(automatic.data.snapshot.preview.events.map(e=>e.id)),new Set(['zora-mmc','zora-mch']),'login delivers the full published calendar without confirmation');
assert.equal(db.sql.slice(automaticStart).some(sql=>/roster_events|canonical_doctors|roster_doctors/.test(sql)),false,'automatic login never scans roster history');
const replayWrites=db.rowsWritten;
assert.equal((await api({action:'login',responseMode:'fast'},'zora@example.test',{IDENTITY_REVIEW_ENABLED:'true',BOUNDED_MANUAL_ROSTER_ENABLED:'true'})).status,200);
assert.equal(db.rowsWritten,replayWrites,'linked logins do not write again');
sqlite.prepare("INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES('other@example.test','mmc','BLOCKED EXACT','Blocked Exact')").run();
const occupied=await api({action:'login',responseMode:'fast'},'blocked@example.test',{IDENTITY_REVIEW_ENABLED:'true'});
assert.equal(occupied.data.claims.length,0,'a partial available match cannot acquire the other site');
sqlite.prepare("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('reviewed-person','Reviewed Exact')").run();
sqlite.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('mmc','REVIEWED EXACT','Reviewed Exact','reviewed-person')").run();
sqlite.prepare("INSERT INTO account_people(email,person_id) VALUES('other@example.test','reviewed-person')").run();
const reviewedOwner=await api({action:'login',responseMode:'fast'},'reviewed@example.test',{IDENTITY_REVIEW_ENABLED:'true'});
assert.equal(reviewedOwner.status,200,JSON.stringify(reviewedOwner));
assert.equal(reviewedOwner.data.claims.length,0,'approved ownership cannot be seized');
// The durable ownership assertion also runs inside the claim transaction.
const candidate={sourceType:'mch',key:'REVIEW RACE',displayName:'Reviewed Exact'};
db.beforeBatch=()=>{
 sqlite.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('mch','REVIEW RACE','Reviewed Exact','reviewed-person')").run();
};
await assert.rejects(saveBoundedAccountClaims(db,await loadAccountMirror(db,'reviewed@example.test'),[candidate],{guardAutomaticIdentity:true}),e=>e.code==='IDENTITY_CLAIM_CONFLICT');
assert.equal((await loadAccountMirror(db,'reviewed@example.test')).claims.length,0);
console.log('Automatic clear-name login, multi-site calendar, replay, uncertainty and durable ownership checks passed.');
// Existing links survive missing publications, and removal is still available.
r2.objects.clear();
const missing=await api({action:'loadAccountContext'});
assert.equal(missing.status,200,JSON.stringify(missing));
assert.equal(missing.data.identityDiscoveryUnavailable,true);
assert.equal(missing.data.claims[0].key,alice.key);
assert.deepEqual(missing.data.suggestedClaims,[]);
const removal=await api({action:'removeRosterClaim',claim:alice}); assert.equal(removal.status,200,JSON.stringify(removal));
assert.deepEqual((await loadAccountMirror(db,'alice@example.test')).claims.filter(claim=>claim.key===alice.key),[]);
for(const sql of db.sql.filter(sql=>/INDEXED BY idx_account_claims_source_doctor_email|INDEXED BY idx_doctor_profiles_doctor|INDEXED BY idx_roster_files_source_active|FROM roster_files WHERE id IN/.test(sql))) {
  const bindings=(sql.match(/\?/g)||[]).length;
  const plan=sqlite.prepare('EXPLAIN QUERY PLAN '+sql).all(...Array(bindings).fill('fixture'));
  assert.ok(plan.some(row=>/SEARCH/.test(row.detail)),JSON.stringify(plan));
  assert.equal(plan.some(row=>/SCAN (account_claims|doctor_profiles|roster_file_doctors|roster_files)/.test(row.detail)),false,JSON.stringify(plan));
}
console.log('Bounded identity: 100,001 historical events, read-only suggestions, ambiguity, grades, paginated enrichment, gates, ownership races, replay, rollback, stale links and missing publications passed.');

const permissionWrites = db.rowsWritten;
const permissionSqlStart = db.sql.length;
const permissions = await api({action:'setUserFacilityOverviewEnabled',targetEmail:'alice@example.test',facilityOverviewEnabled:true},'creator@example.test');
assert.equal(permissions.status,200,JSON.stringify(permissions));
assert.equal(permissions.data.user.facilityOverviewEnabled,true);
assert.equal(db.rowsWritten-permissionWrites,1,'access toggle updates one profile row only');
assert.equal(db.sql.slice(permissionSqlStart).some(sql=>/DELETE FROM account_claims|INSERT INTO account_states|DELETE FROM subscription_tokens/.test(sql)),false,'access toggle leaves identity links, calendars and tokens untouched');

// Approved identity links must be visible without inventing legacy claims.
const directoryClaimCount = sqlite.prepare("SELECT COUNT(*) AS n FROM account_claims WHERE email='alice@example.test'").get().n;
sqlite.prepare("INSERT INTO roster_people(person_id,preferred_display_name,created_at,updated_at,status) VALUES('directory-approved','Approved Directory Name',?,?,'active')").run(today,today);
sqlite.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id,created_at,updated_at) VALUES('mch','DIRECTORY APPROVED','Approved Directory Name','directory-approved',?,?)").run(today,today);
sqlite.prepare("INSERT INTO account_people(email,person_id,created_at,updated_at) VALUES('alice@example.test','directory-approved',?,?) ON CONFLICT(email) DO UPDATE SET person_id=excluded.person_id").run(today,today);
let cursor='', found;
for(let page=0;page<10;page++) {
  const result = await api({action:'listUsers',cursor},'creator@example.test',{IDENTITY_REVIEW_ENABLED:'true'});
  assert.equal(result.status,200,JSON.stringify(result));
  found ||= result.data.users.find(user=>user.email==='alice@example.test');
  cursor=result.data.nextCursor;if(!cursor)break;
}
assert.ok(found.rosterLinks.some(link=>link.key==='DIRECTORY APPROVED'),'directory displays approved identity aliases');
assert.deepEqual(found.sites,['MCH']);
assert.equal(found.claims.length,directoryClaimCount,'display does not rewrite or relabel editable legacy claims');
const toggleLinked = await api({action:'setUserInsightsEnabled',targetEmail:'alice@example.test',insightsEnabled:true},'creator@example.test',{IDENTITY_REVIEW_ENABLED:'true'});
assert.equal(toggleLinked.status,200,JSON.stringify(toggleLinked));
assert.ok(toggleLinked.data.user.rosterLinks.some(link=>link.key==='DIRECTORY APPROVED'),'permission response retains approved labels');

for(const sql of db.sql.filter(sql=>sql.includes('JOIN account_people a ON a.person_id=r.person_id'))) {
 const plan=sqlite.prepare('EXPLAIN QUERY PLAN '+sql).all(...Array((sql.match(/\?/g)||[]).length).fill('fixture'));
 assert.equal(plan.some(row=>/SCAN (r|a)\b/.test(row.detail)),false,'automatic ownership lookup must use indexes: '+JSON.stringify(plan));
}
