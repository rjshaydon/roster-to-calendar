import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import { onRequestPost as handler } from '../functions/api/state.js';
import { facilityMetadataManifestKey } from '../functions/_lib/facility-overview-cache.js';
import { verifyFacilityContactAccessToken } from '../functions/_lib/facility-contact-access.js';

const RealDate = Date;
let instant = RealDate.parse('2026-10-05T15:30:00+11:00');
globalThis.Date = class extends RealDate {
  constructor(...args) { super(...(args.length ? args : [instant])); }
  static now() { return instant; }
};
const sqlite = new DatabaseSync(':memory:');
for (const name of (await readdir(new URL('../migrations', import.meta.url))).filter(n => n.endsWith('.sql')).sort()) {
  sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), 'utf8'));
}
let sql = [], returnedRows = 0;
const db = { prepare(query) {
  sql.push(query.replace(/\s+/g, ' ').trim());
  const statement = { args: [], bind(...args) { this.args = args; return this; },
    async all() { const results = sqlite.prepare(query).all(...this.args); returnedRows += results.length; return { results }; },
    async first() { const result = sqlite.prepare(query).get(...this.args) || null; returnedRows += Number(Boolean(result)); return result; },
    async run() { throw Error(`Unexpected D1 mutation: ${query}`); },
  }; return statement;
} };
const passwordHash = createHash('sha256').update('salt:test').digest('hex');
for (const email of ['staff@example.test', 'off@example.test', 'alias@example.test', 'unlinked@example.test', 'director@example.test', 'creator@example.test']) {
  sqlite.prepare('INSERT INTO account_profiles(email,real_name,role,password_salt,password_hash,non_clinical) VALUES(?,?,?,?,?,?)')
    .run(email, 'Test Staff', email.startsWith('creator') ? 'creator' : 'user', 'salt', passwordHash, Number(email.startsWith('director')));
}
for (const [email,key] of [['staff@example.test','TEST STAFF'],['off@example.test','DAY OFF'],['alias@example.test','OLD NAME']]) {
  sqlite.prepare('INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,?,?,?)').run(email,'mmc',key,key);
}
sqlite.prepare("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:test-alias','New Name')").run();
for (const key of ['OLD NAME','NEW NAME']) sqlite.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('mmc',?,?, 'person:test-alias')").run(key,key);
sqlite.prepare("INSERT INTO account_people(email,person_id) VALUES('alias@example.test','person:test-alias')").run();
const objects = new Map();
let r2Gets = [];
const r2 = { async get(key) { r2Gets.push(key); const value = objects.get(key); if (!value) return null;
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  return { arrayBuffer: async () => bytes.buffer, text: async () => JSON.stringify(value) };
} };
const date = '2026-10-05';
const shift = { source: 'mmc', kind: 'shift', title: 'MMC: PM', start: `${date}T15:00:00`, end: '2026-10-06T00:00:00' };
const staffRow = key => ({ doctorKey:key, displayName:key, sourceType:'mmc', date, seniority:'HMO', event:{...shift,id:key} });
function publish(rows) {
  objects.set(facilityMetadataManifestKey('mmc'), { sourceType:'mmc', terms:[{termStart:'2026-08-03',termEnd:'2026-11-01',staffKey:'staff',staffRevision:'s1',coverage:[{startDate:'2026-08-03',endDate:'2026-11-01'}]}],
    months:{'2026-10':{key:'month',revision:'m1'}}, days:{[date]:{key:'day',revision:`d${rows.length}`}} });
  objects.set('staff',{seniorityOverrides:[]}); objects.set('month',{rows}); objects.set('day',{date,rows});
}
publish([staffRow('TEST STAFF'),staffRow('NEW NAME')]);
const secret = 'test-on-shift-contact-secret-long-enough';
const env = { ROSTER_DB:db,ROSTER_FILES:r2, ON_SHIFT_FOR_ALL_ENABLED:'true', IDENTITY_REVIEW_ENABLED:'true',
  FACILITY_OVERVIEW_MAINTENANCE_MODE:'false',FACILITY_SHARED_ROLLOUT_ACTIVE:'true',FACILITY_SHARED_EMERGENCY_PAUSED:'false',
  FACILITY_LEGACY_READS_PAUSED:'true',FACILITY_SHARED_READER_SOURCE_ALLOWLIST:'mmc',FACILITY_SHARED_READER_COHORT:'all',
  FACILITY_SHARED_METADATA_ENABLED:'true',FACILITY_SHARED_DAYS_ENABLED:'true',FACILITY_SHARED_CONTACTS_ENABLED:'true',
  FACILITY_SHARED_CONTACTS_SOURCE_ALLOWLIST:'mmc',FACILITY_CONTACT_ACCESS_SECRET:secret };
async function request(action, body={}, override={}) {
  sql=[]; returnedRows=0; r2Gets=[];
  const response=await handler({env:{...env,...override}, request:new Request('https://example.test/api/state',{method:'POST',headers:{'content-type':'application/json'},
    body:JSON.stringify({email:'staff@example.test',password:'test',action,...body})}),waitUntil(){throw Error('No maintenance should be scheduled');}});
  const data=await response.json();
  assert.ok(!sql.some(q => /\b(?:roster_events|facility_term_staff_contributions|roster_daily_presence|facility_access_sessions)\b/.test(q)), `Restricted request performed roster/membership SQL: ${sql.join('\n')}`);
  return {status:response.status,data,statements:sql.length,returnedRows,r2Reads:r2Gets.length};
}
try {
  const login = await request('login', { responseMode: 'fast' }, { BOUNDED_MANUAL_ROSTER_ENABLED: 'true' });
  assert.equal(login.status, 200);
  assert.equal(login.r2Reads, 0, 'authentication must not parse roster or snapshot artifacts');
  assert.equal(login.data.snapshotSource, 'login-deferred');
  const launch=await request('queryFacilityOverviewLaunchWindow');
  assert.equal(launch.status,200); assert.equal(launch.data.shiftWindow.facilityKey,'MMC');
  assert.equal(launch.statements,3,'account authentication and two bounded identity reads only');
  assert.ok(!r2Gets.some(key => /month/.test(key)), 'eligibility never parses a whole roster month');
  assert.equal(r2Gets.filter(key => key === facilityMetadataManifestKey('mmc')).length, 1, 'three day reads share one manifest parse');
  const query={date,facilityKey:'MMC'};
  const fullCreator=await request('queryFacilityOverviewOnShift',{...query,email:'creator@example.test'});
  assert.equal(fullCreator.status,200,'existing full entitlement continues with the new flag enabled');
  assert.equal(fullCreator.data.shiftWindow,undefined);
  assert.equal((await verifyFacilityContactAccessToken(secret,fullCreator.data.contactAccessToken)).date,undefined,'full overview contacts retain ordinary navigation');
  const opened=await request('queryFacilityOverviewOnShift',query);
  assert.equal(opened.status,200); assert.equal(opened.data.events.length,2); assert.equal(opened.data.facilityOverviewAccess.mode,'denied');
  assert.equal(opened.statements,3,'temporary On shift must bypass the full access resolver');
  const token=opened.data.contactAccessToken; assert.ok(token);
  const tokenClaims=await verifyFacilityContactAccessToken(secret,token);
  assert.equal(tokenClaims.date,date); assert.equal(tokenClaims.facilityKey,'MMC');
  assert.ok(Date.parse(tokenClaims.expiresAt)-instant <= 15*60*1000);
  const repeat=await request('queryFacilityOverviewOnShift',{...query,cachedRevision:opened.data.revision});
  assert.equal(repeat.status,200); assert.equal(repeat.data.rosterUnchanged,true); assert.equal(repeat.data.events,undefined);
  const contact=await request('queryFacilityOverviewContactList',{...query,contactAccessToken:token});
  assert.equal(contact.status,200); assert.equal(contact.statements,0,'contact refresh bypasses D1');
  assert.equal((await request('queryFacilityOverviewContactList',{...query,date:'2026-10-06',contactAccessToken:token})).status,403);
  assert.equal(sql.length,0,'wrong-date contact token fails without D1');
  assert.equal((await request('queryFacilityOverviewContactList',{...query,contactAccessToken:token},{ON_SHIFT_FOR_ALL_ENABLED:'false'})).status,403);
  for (const forbidden of [{...query,date:'2026-10-06'},{...query,facilityKey:'DDH'},{...query,facilityKey:'ALL'}, {...query,targetEmail:'alias@example.test'}, {...query,password:'incorrect'}]) {
    assert.ok([401,403].includes((await request('queryFacilityOverviewOnShift',forbidden)).status));
  }
  for (const action of ['queryFacilityOverviewTerms','queryFacilityOverviewMetadata','queryFacilityOverviewByStream','queryFacilityOverviewStaff','queryFacilityOverviewWorkingTogether','queryFacilityOverviewTogetherContext','setContactAllocationResolution']) {
    assert.equal((await request(action,query)).status,403,`${action} still requires full overview permission`);
    assert.equal(r2Gets.length,0,'broader views reject before R2');
  }
  for (const email of ['off@example.test','unlinked@example.test','director@example.test']) {
    const result=await request('queryFacilityOverviewLaunchWindow',{email});
    assert.ok(result.status===403 || result.data.shiftWindow===null);
    assert.equal((await request('queryFacilityOverviewOnShift',{...query,email})).status,403);
  }
  const alias=await request('queryFacilityOverviewLaunchWindow',{email:'alias@example.test'});
  assert.equal(alias.status,200); assert.equal(alias.data.shiftWindow.facilityKey,'MMC'); assert.equal(alias.statements,5);
  sqlite.prepare("UPDATE roster_person_aliases SET review_state='pending' WHERE doctor_key='NEW NAME'").run();
  assert.equal((await request('queryFacilityOverviewOnShift',{...query,email:'alias@example.test'})).status,403,'unapproved aliases cannot establish eligibility');
  sqlite.prepare("UPDATE roster_person_aliases SET review_state='approved' WHERE doctor_key='NEW NAME'").run();
  assert.equal((await request('queryFacilityOverviewOnShift',{...query,email:'creator@example.test',targetEmail:'staff@example.test'})).status,200,'impersonation uses the viewed staff shift');
  assert.equal((await request('queryFacilityOverviewStaff',{...query,email:'creator@example.test',targetEmail:'staff@example.test'})).status,403,'Creator access must not leak to the viewed account');
  assert.equal((await request('queryFacilityOverviewLaunchWindow',{}, {ON_SHIFT_FOR_ALL_ENABLED:'false'})).status,403);
  assert.equal((await request('queryFacilityOverviewOnShift',query,{FACILITY_SHARED_DAYS_ENABLED:'false'})).status,403);
  assert.equal((await request('queryFacilityOverviewOnShift',query,{FACILITY_SHARED_READER_SOURCE_ALLOWLIST:''})).status,403);
  assert.equal((await request('queryFacilityOverviewOnShift',query,{FACILITY_SHARED_ROLLOUT_ACTIVE:'false',FACILITY_LEGACY_READS_PAUSED:'false'})).status,403,'restricted users cannot trigger legacy D1 reads');
  publish(Array.from({length:500},(_,i)=>staffRow(i===0?'TEST STAFF':`COLLEAGUE ${i}`)));
  const large=await request('queryFacilityOverviewOnShift',query);
  assert.equal(large.status,200); assert.equal(large.data.events.length,500);
  assert.equal(large.statements,opened.statements); assert.equal(large.returnedRows,opened.returnedRows,'D1 result volume is independent of daily staff count');
  assert.equal(large.r2Reads,opened.r2Reads,'shared-read operation count is independent of daily staff count');
  instant=RealDate.parse('2026-10-05T13:59:59+11:00'); assert.equal((await request('queryFacilityOverviewOnShift',query)).status,403);
  instant=RealDate.parse('2026-10-05T14:00:00+11:00'); assert.equal((await request('queryFacilityOverviewOnShift',query)).status,200);
  instant=RealDate.parse('2026-10-06T00:59:00+11:00');
  const night=await request('queryFacilityOverviewOnShift',query); assert.equal(night.status,200); assert.equal(night.data.shiftWindow.rosterDate,date);
  assert.equal((await verifyFacilityContactAccessToken(secret,night.data.contactAccessToken)).expiresAt,'2026-10-05T14:00:00.000Z','contact token expires at the grace-window end');
  instant=RealDate.parse('2026-10-06T01:00:01+11:00'); assert.equal((await request('queryFacilityOverviewOnShift',query)).status,403);
  assert.equal((await request('queryFacilityOverviewContactList',{...query,contactAccessToken:night.data.contactAccessToken})).status,403); assert.equal(sql.length,0);
  instant=RealDate.parse('2026-10-05T15:30:00+11:00'); publish([{...staffRow('TEST STAFF'),event:{...shift,kind:'leave'}}]);
  assert.equal((await request('queryFacilityOverviewLaunchWindow')).data.shiftWindow,null,'leave grants no access');
  objects.clear(); assert.equal((await request('queryFacilityOverviewOnShift',query)).status,403,'missing publications fail closed');
  console.log(JSON.stringify({restrictedStartup:{eligibilityStatements:launch.statements,rosterStatements:opened.statements,aliasEligibilityStatements:alias.statements},rosterRows:500,contactRefreshStatements:contact.statements, note:'Statement counts are traced; SQLite result counts do not measure Cloudflare billed rows.'},null,2));
} finally { globalThis.Date=RealDate; sqlite.close(); }

// Exercise client permission, startup and expiry with the actual app functions.
const app=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const section=(start,end)=>app.slice(app.indexOf(start),app.indexOf(end,app.indexOf(start)));
const noop=()=>{};
let timers=[], cleared=0, closed=0, full=false;
const client={Date:class extends RealDate { static now(){return RealDate.parse('2026-10-05T15:30:00+11:00');} },currentOnShiftForAllEnabled:true,currentOnShiftAccessWindow:null,onShiftAccessExpiryTimer:0,
 currentNonClinical:false,currentFacilityOverviewMaintenance:false,canUseFullFacilityOverview:()=>full,window:{setTimeout(fn){timers.push(fn);return timers.length;},clearTimeout(){cleared++;}},
 facilityOverviewState:{onShiftData:[1],content:'cached',contactList:{},previousNightRoster:{},contactAccessToken:'token'},
 cancelFacilityOverviewDataRequest:noop,stopFacilityOverviewContactRefresh:noop,syncFacilityOverviewAccess:()=>{closed++;}};
runInNewContext(`${section('function currentOnShiftWindow()', 'function sanitizeFacilityOverviewAccess(')};this.apply=applyOnShiftAccessWindow;this.allowed=canUseFacilityOverview;this.clear=clearOnShiftAccess;`,client);
assert.equal(client.allowed(),false);
const window={facilityKey:'MMC',rosterDate:'2026-10-05',start:RealDate.parse('2026-10-05T15:00:00+11:00'),end:RealDate.parse('2026-10-06T00:00:00+11:00')};
client.apply(window); assert.equal(client.allowed(),true);
client.facilityAccessKeys=()=>[]; client.currentFacilityOverviewAccess={mode:'denied'};
runInNewContext(`${section('function applyFacilityOverviewSiteScope()', 'function resetFacilityOverviewAccessForEnteredUser(')};this.scope=applyFacilityOverviewSiteScope;`,client);
client.scope(); assert.equal(client.facilityOverviewState.tab,'on-shift'); assert.equal(client.facilityOverviewState.facilityKey,'MMC'); assert.equal(client.facilityOverviewState.date,'2026-10-05');
const contexts=section('function facilityOverviewSnapshotContext()', 'function loadFacilityOverviewSnapshot(');
Object.assign(client,{facilityOverviewTargetEmail:()=>'',currentUserEmail:'staff@example.test',normalizeEmail:v=>v,FACILITY_ACCESS_VERSION:'test'});
runInNewContext(`${contexts};this.cacheContext=facilityOverviewSnapshotContext`,client);
assert.match(client.cacheContext().scopeKey,/on-shift:MMC:2026-10-05/); assert.ok(client.cacheContext().accessExpiresAt);
full=true;client.clear();assert.equal(client.allowed(),true,'full overview access survives temporary expiry');full=false;
client.apply(window);timers.at(-1)(); assert.equal(client.allowed(),false); assert.equal(closed,1);assert.equal(client.facilityOverviewState.onShiftData,null);assert.equal(client.facilityOverviewState.contactAccessToken,'');
Object.assign(client,{cancelFacilityOverviewAccessWait:noop,rememberFacilityOverviewTabForCurrentAccount:noop});
client.apply(window);
runInNewContext(`${section('function beginFacilityOverviewAccountSession()', 'function resetFacilityOverviewSessionState(')};beginFacilityOverviewAccountSession()`,client);
assert.equal(client.currentOnShiftAccessWindow,null);assert.equal(client.currentOnShiftForAllEnabled,false);assert.equal(client.facilityOverviewState.contactAccessToken,'');assert.ok(cleared);
console.log('Restricted client access passed server-confirmed eligibility, ED/date scope, cache expiry, full-access preservation and account-transition cleanup.');

let eligibilityRequests=0, automaticallyOpened=0, resolveEligibility;
Object.assign(client,{currentOnShiftForAllEnabled:true,currentFacilityOverviewAutomaticLaunchEnabled:true,currentNonClinical:false,
 currentSnapshot:{preview:{events:[shift]}},currentSnapshotStale:false,calendarSnapshotMatchesActiveContext:()=>true,
 calendarTransitionRunId:1,activeCalendarTransitionKey:()=> 'staff',calendarTransitionStillCurrent:()=>true,
 activeCalendarMode:()=> 'claimed-account',authUserEmail:'staff@example.test',authUserPassword:'test',
 facilityOverviewSessionNeedsInitialization:false,markLoginPhase:noop,setStatus:noop,normalizeAuthMessage:v=>v,
 openFacilityOverview:async()=>{automaticallyOpened++;},
 readJsonResponse:async data=>data,fetch:()=>{eligibilityRequests++;return new Promise(resolve=>{resolveEligibility=resolve;});}});
runInNewContext(`${section('function requestClinicalStartupShiftWindow(', 'async function loginWithEmail(')};this.launch=launchClinicalOnShiftWorkspace`,client);
assert.equal(client.launch(),false,'restricted startup must confirm a personal snapshot with the server');
assert.equal(client.launch(),false);assert.equal(eligibilityRequests,1,'startup calls coalesce');
resolveEligibility({shiftWindow:window});await client.clinicalOnShiftWindowPromise;
assert.equal(automaticallyOpened,1);assert.equal(client.facilityOverviewState.date,window.rosterDate);assert.equal(client.clinicalOnShiftStartupPending,false);
client.clear();client.clinicalOnShiftStartupPending=true;client.clinicalOnShiftWindowPromise=null;
client.launch();client.clinicalOnShiftStartupPending=false;resolveEligibility({shiftWindow:window});await client.clinicalOnShiftWindowPromise;
assert.equal(automaticallyOpened,1,'manual navigation cancels delayed automatic opening');assert.equal(client.allowed(),true,'manual navigation retains permission to reopen');
client.clear();client.clinicalOnShiftStartupPending=true;client.clinicalOnShiftWindowPromise=null;
client.launch();client.calendarTransitionRunId=2;resolveEligibility({shiftWindow:window});await client.clinicalOnShiftWindowPromise;
assert.equal(client.allowed(),false,'a previous account response cannot grant the new account access');
client.activeCalendarMode=()=> 'doctor-profile';client.clinicalOnShiftWindowPromise=null;
assert.equal(client.launch(),false);assert.equal(eligibilityRequests,3,'unlinked profile must not query the Creator shift');
client.activeCalendarMode=()=> 'claimed-account';client.currentOnShiftForAllEnabled=false;
assert.equal(client.launch(),false);assert.equal(eligibilityRequests,3,'feature disablement prevents restricted startup');

let termsLoaded=0,dailyLoads=0;
client.currentOnShiftForAllEnabled=true;client.apply(window);
Object.assign(client,{facilityOverviewOpeningPromise:null,facilityOverviewOpeningRunId:0,facilityOverviewNavigationLocked:false,facilityOverviewIgnoreToggleUntil:0,
 isFacilityOverviewOpen:()=>true,currentFacilityOverviewAccessReady:true,refreshFacilityOverviewPreferredFacility:noop,
 currentFacilityOverviewShiftWindow:()=>window,contactOperationalDate:()=> '2026-10-06',formatDateKey:()=> '2026-10-06',
 facilityOverviewState:{tab:'together',date:'2026-10-06',facilityKey:'DDH'},
 form:{classList:{add:noop}},previewSection:{classList:{add:noop}},facilityOverviewSection:{classList:{remove:noop}},
 resetFacilityOverviewScroll:noop,syncFacilityOverviewNavigationState:noop,renderFacilityOverview:noop,
 loadFacilityOverviewAvailableTerms:async()=>{termsLoaded++;},loadFacilityOverviewOnShift:async()=>{dailyLoads++;}});
runInNewContext(`${section('function openFacilityOverview(options = {})', 'async function openFacilityOverviewByStream(')};this.open=openFacilityOverview`,client);
await client.open();assert.equal(dailyLoads,1);assert.equal(termsLoaded,0,'restricted reopening must skip term queries');
assert.equal(client.facilityOverviewState.facilityKey,'MMC');assert.equal(client.facilityOverviewState.date,window.rosterDate,'reopening after midnight keeps the shift date');
console.log('Restricted startup/opening passed server confirmation, one fallback, auto-open, manual navigation, stale accounts, unlinked profiles, rollout disablement and zero term loads.');
