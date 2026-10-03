import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, readdir } from 'node:fs/promises';
import { DatabaseSync } from 'node:sqlite';
import { runInNewContext } from 'node:vm';
import { resolveFacilityOverviewAccess, resolveFacilityOverviewRangeAccess, onRequestPost } from '../functions/api/state.js';
import { facilityAccessAllows, filterFacilityRowsBySegments, validFacilityDateRange } from '../public/static/facility-access-policy.js';

const RealDate = Date;
globalThis.Date = class extends RealDate { constructor(...args) { super(...(args.length ? args : ['2026-10-03T01:00:00Z'])); } static now() { return RealDate.parse('2026-10-03T01:00:00Z'); } };
const sqlite = new DatabaseSync(':memory:');
for (const name of (await readdir(new URL('../migrations', import.meta.url))).filter(name => name.endsWith('.sql')).sort()) sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url),'utf8'));
const queries = [];
const db = { prepare(sql) { queries.push(sql); return { args: [], bind(...args) { this.args = args; return this; }, async first() { return sqlite.prepare(sql).get(...this.args) || null; }, async all() { return { results: sqlite.prepare(sql).all(...this.args) }; }, async run() { const r = sqlite.prepare(sql).run(...this.args); return { success: true, meta: { changes: Number(r.changes) } }; } }; } };
const objects = new Map();
const r2 = { async get(key) { const value = objects.get(key); return value ? { async arrayBuffer() { return new TextEncoder().encode(JSON.stringify(value)).buffer; } } : null; } };
const claims = sources => sources.map(sourceType => ({ sourceType, key: 'TRAINEE', displayName: 'Trainee' }));
const record = { email:'trainee@example.com', role:'user', facilityOverviewEnabled:true, realName:'Trainee', claims:claims(['mmc','ddh','mch']) };
const opts = { today:'2026-10-03', materializedOnly:true };
function membership(source, term, key='TRAINEE', grade='HMO', file=`${source}-${term}`) {
  sqlite.prepare('INSERT OR IGNORE INTO roster_files (id,name,source_type,active) VALUES (?, ?, ?, 1)').run(file,file,source);
  sqlite.prepare(`INSERT OR REPLACE INTO facility_term_staff_contributions (source_type,term_start,doctor_key,file_id,display_name,seniority,first_applicable_date,last_applicable_date,fact_digest,updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(source,term,key,file,key,grade,term,term,`${file}-${key}-${grade}`,'2026-10-02');
}
function row(source, key, date, grade='HMO') { return { sourceType:source, doctorKey:key, displayName:key, event:{id:`${source}-${key}-${date}`,sourceType:source,source:source.toUpperCase(),doctorKey:key,start:`${date}T08:00:00+10:00`,end:`${date}T17:00:00+10:00`,title:'Day',rawValue:'AM',seniority:grade} }; }
function publish(source, term, end, dates) {
  const manifestKey = `facility-overview/v1/${source}/manifest.json`;
  const manifest = objects.get(manifestKey) || {sourceType:source,terms:[],months:{},days:{}};
  const staffKey = `${source}-${term}-staff`;
  manifest.terms.push({termStart:term,termEnd:end,visibleFrom:term,staffKey,staffRevision:staffKey});
  const rows = dates.flatMap(date => [row(source,'TRAINEE',date),row(source,`${source.toUpperCase()} PEER`,date)]);
  objects.set(staffKey,{members:[{doctorKey:'TRAINEE',displayName:'Trainee',sourceType:source},{doctorKey:`${source.toUpperCase()} PEER`,displayName:`${source.toUpperCase()} Peer`,sourceType:source}],seniorityOverrides:[]});
  for (const month of new Set(dates.map(date=>date.slice(0,7)))) {
    const key = `${source}-${month}-month`;
    const existing = objects.get(key)?.rows || [];
    objects.set(key,{rows:[...existing,...rows.filter(r=>r.event.start.startsWith(month))]});
    manifest.months[month] = {key,revision:`${key}-${objects.get(key).rows.length}`};
  }
  for (const date of dates) {
    const key = `${source}-${date}-day`;
    objects.set(key,{rows:rows.filter(row=>row.event.start.startsWith(date))});
    manifest.days[date]={key,revision:key};
  }
  objects.set(manifestKey,manifest);
}
function account(email, accountClaims, role='user', enabled=1) {
  sqlite.prepare('INSERT INTO account_profiles (email,real_name,role,facility_overview_enabled,insights_enabled,password_salt,password_hash) VALUES (?,?,?,?,1,?,?)')
    .run(email,'Trainee',role,enabled,'salt',createHash('sha256').update('salt:password').digest('hex'));
  for (const claim of accountClaims) sqlite.prepare('INSERT INTO account_claims (email,source_type,doctor_key,display_name) VALUES (?,?,?,?)').run(email,claim.sourceType,claim.key,claim.displayName);
}
const env = { ROSTER_DB:db, ROSTER_FILES:r2, FACILITY_OVERVIEW_MAINTENANCE_MODE:'false', FACILITY_SHARED_ROLLOUT_ACTIVE:'true', FACILITY_SHARED_EMERGENCY_PAUSED:'false', FACILITY_LEGACY_READS_PAUSED:'true', FACILITY_SHARED_READER_COHORT:'all', FACILITY_SHARED_READER_SOURCE_ALLOWLIST:'mmc,ddh,mch', FACILITY_SHARED_METADATA_ENABLED:'true', FACILITY_SHARED_DAYS_ENABLED:'true' };
async function call(body, status=200, maximumQueries=32) {
  queries.length=0;
  const response = await onRequestPost({env,waitUntil(){},request:new Request('http://localhost/api/state',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({email:record.email,password:'password',...body})})});
  const payload = await response.json();
  assert.equal(response.status,status,`${body.action}: ${JSON.stringify(payload)}`);
  assert.equal(queries.some(q=>/\broster_events\b/.test(q)),false,'shared access must not scan full roster events');
  assert.ok(queries.length<=maximumQueries,`bounded query count: ${queries.length}`);
  return payload;
}

membership('ddh','2026-05-04');
membership('mmc','2026-08-03');
membership('ddh','2026-08-03'); // a current-term locum
membership('mch','2026-11-02'); // linked upcoming rotation, not present access
publish('ddh','2026-05-04','2026-08-02',['2026-06-03']);
publish('ddh','2026-08-03','2026-11-01',['2026-10-02']);
publish('mmc','2026-08-03','2026-11-01',['2026-10-02']);
publish('mch','2026-08-03','2026-11-01',['2026-10-02']);
account(record.email,record.claims);
account('creator@example.com',[],'creator');
account('disabled@example.com',record.claims,'user',0);
let access = await resolveFacilityOverviewAccess(db,record,opts);
assert.equal(access.mode,'sites');
assert.deepEqual(access.facilityKeys,['DDH','MMC']);
assert.equal(facilityAccessAllows(access,'mch'),false);
assert.equal(facilityAccessAllows(access,'mmc'),true);
const nextTerm = await resolveFacilityOverviewAccess(db,record,{today:'2026-11-02',materializedOnly:true});
assert.deepEqual(nextTerm.facilityKeys,['MCH'],'term change removes former hospitals and activates the next rotation');
const legacy = await resolveFacilityOverviewAccess(null,record,{today:opts.today,events:[row('mmc','TRAINEE','2026-10-02').event,row('ddh','TRAINEE','2026-10-02').event,row('mch','TRAINEE','2026-11-03').event]});
assert.deepEqual(legacy.facilityKeys,access.facilityKeys,'legacy and compact current-term scope agree');

for (const grade of ['SMS','CMO']) {
  membership('mmc','2026-08-03',grade,grade);
  const senior = { ...record,email:`${grade}@example.com`,claims:[{sourceType:'mmc',key:grade,displayName:grade}] };
  assert.equal((await resolveFacilityOverviewAccess(db,senior,opts)).mode,'all');
  assert.equal((await resolveFacilityOverviewAccess(null,senior,{today:opts.today,events:[{...row('mmc',grade,'2026-10-02').event,seniority:grade}]})).mode,'all');
  assert.equal((await resolveFacilityOverviewAccess(db,{...senior,facilityOverviewEnabled:false},opts)).mode,'denied');
}
// A permanent CMO can be absent from this term, as a permanent SMS can.
membership('mmc','2026-05-04','CMO ON LEAVE','CMO');
const cmoLeave = {...record,email:'leave@example.com',claims:[{sourceType:'mmc',key:'CMO ON LEAVE',displayName:'CMO On Leave'}]};
assert.equal((await resolveFacilityOverviewAccess(db,cmoLeave,opts)).mode,'all');
assert.equal((await resolveFacilityOverviewAccess(db,cmoLeave,{today:opts.today,materializedOnly:false})).mode,'all','legacy payload routing retains the same compact permission policy');
// Effective grade overrides, including resetting to roster grade, revoke scope.
sqlite.prepare(`INSERT INTO facility_staff_seniority_overrides (id,source_type,doctor_key,seniority,term_start) VALUES ('grade','mmc','TRAINEE','CMO','2026-08-03')`).run();
assert.equal((await resolveFacilityOverviewAccess(db,record,opts)).mode,'all');
sqlite.prepare("UPDATE facility_staff_seniority_overrides SET seniority='HMO' WHERE id='grade'").run();
assert.equal((await resolveFacilityOverviewAccess(db,record,opts)).mode,'sites');
sqlite.prepare("UPDATE facility_staff_seniority_overrides SET use_roster_seniority=1,seniority='' WHERE id='grade'").run();
assert.equal((await resolveFacilityOverviewAccess(db,record,opts)).mode,'sites');

const range = {startDate:'2026-06-01',endDate:'2026-10-03',today:opts.today};
const historical = await resolveFacilityOverviewRangeAccess(db,record,access,range);
assert.deepEqual(historical.segments.map(s=>`${s.sourceType}|${s.termStart}`).sort(),['ddh|2026-05-04','ddh|2026-08-03','mmc|2026-08-03']);
assert.equal((await resolveFacilityOverviewRangeAccess(db,record,access,{...range,sourceTypes:['mch']})).mode,'denied');
assert.equal((await resolveFacilityOverviewRangeAccess(db,record,access,{...range,startDate:'2026-02-30'})).invalid,true);
assert.equal(validFacilityDateRange('2026-02-30','2026-03-01'),false);

const context = await call({action:'queryFacilityOverviewTogetherContext',startDate:'2026-06-03',endDate:'2026-06-03'});
assert.deepEqual(context.sourceTypes,['ddh']);
assert.ok(context.members.some(member=>member.doctorKey==='DDH PEER'));
assert.equal(context.members.some(member=>member.doctorKey==='MMC PEER'),false);
const dateFirst = await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-06-03',endDate:'2026-06-03',doctorKeys:[]});
assert.ok(dateFirst.events.some(event=>event.doctorKey==='DDH PEER'));
await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-06-03',endDate:'2026-06-03',sourceTypes:['mmc'],doctorKeys:['MMC PEER']},403);
await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-06-03',endDate:'2026-06-03',sourceTypes:['not-a-hospital']},400);
const crossTerm = await call({action:'queryFacilityOverviewWorkingTogether',...range,doctorKeys:[]});
assert.ok(crossTerm.events.some(row=>row.sourceType==='ddh'&&row.event.start.startsWith('2026-06')));
assert.ok(crossTerm.events.some(row=>row.sourceType==='mmc'&&row.event.start.startsWith('2026-10')));
assert.equal(crossTerm.events.some(row=>row.sourceType==='mch'),false);
await call({action:'queryFacilityOverviewOnShift',date:'2026-10-02',facilityKey:'MCH'},403);
const onShift = await call({action:'queryFacilityOverviewOnShift',date:'2026-10-02',facilityKey:'ALL'});
assert.ok(onShift.events.some(row=>row.sourceType==='mmc'));
assert.ok(onShift.events.some(row=>row.sourceType==='ddh'));
assert.equal(onShift.events.some(row=>row.sourceType==='mch'),false);
await call({action:'queryFacilityOverviewStaff',facilityKey:'mmc',termStart:'2026-05-04',termEnd:'2026-08-02'},403);
const metadata = await call({action:'queryFacilityOverviewMetadata',sourceTypes:['mch']});
assert.equal(metadata.facilities.some(f=>f.sourceType==='mch'||f.facilityKey==='MCH'),false);
// Who/When discovery is filtered to historical scope before computing peers.
const peers = await call({action:'queryRosterOverlapDoctors',startDate:'2026-06-03',endDate:'2026-06-03',overlapDoctorKeys:['TRAINEE'],excludeDoctorKeys:['TRAINEE']});
assert.deepEqual(peers.doctors.map(d=>d.doctorKey),['DDH PEER']);
await call({action:'queryRosterOverlapDoctors',startDate:'2026-06-03',endDate:'2026-06-03',sourceTypes:['mmc'],overlapDoctorKeys:['MMC PEER']},403);
// Creator entering a trainee must use that trainee's scope.
await call({action:'queryFacilityOverviewWorkingTogether',email:'creator@example.com',targetEmail:record.email,startDate:'2026-06-03',endDate:'2026-06-03',sourceTypes:['mmc']},403);
await call({action:'queryFacilityOverviewWorkingTogether',email:'creator@example.com',targetEmail:'disabled@example.com',startDate:'2026-06-03',endDate:'2026-06-03'},403);
// Scope changes must invalidate cached decisions and historical revision keys.
const beforeCorrection = historical.revision;
sqlite.prepare("DELETE FROM facility_term_staff_contributions WHERE source_type='ddh' AND term_start='2026-08-03' AND doctor_key='TRAINEE'").run();
access = await resolveFacilityOverviewAccess(db,record,opts);
assert.equal(access.facilityKey,'MMC');
assert.equal(access.cache,'refreshed');
assert.notEqual((await resolveFacilityOverviewRangeAccess(db,record,access,range)).revision,beforeCorrection);
await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-10-02',endDate:'2026-10-02',sourceTypes:['ddh']},403);
await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-06-03',endDate:'2026-06-03'});
// A trainee with no present assignment can still use verified history.
sqlite.prepare("UPDATE roster_files SET active=0 WHERE id='mmc-2026-08-03'").run();
access = await resolveFacilityOverviewAccess(db,record,opts);
assert.equal(access.mode,'denied');
assert.equal(access.canSearchHistory,true);
await call({action:'queryFacilityOverviewTogetherContext',startDate:'2026-06-03',endDate:'2026-06-03'});
const missingManifest=objects.get('facility-overview/v1/ddh/manifest.json');
objects.delete('facility-overview/v1/ddh/manifest.json');
assert.ok((await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-06-03',endDate:'2026-06-03'})).missing.length);
objects.set('facility-overview/v1/ddh/manifest.json',missingManifest);
// Deactivated historical files cease to authorise the old rotation.
sqlite.prepare("UPDATE roster_files SET active=0 WHERE id='ddh-2026-05-04'").run();
await call({action:'queryFacilityOverviewWorkingTogether',startDate:'2026-06-03',endDate:'2026-06-03'},403);

const overnight={sourceType:'ddh',doctorKey:'NIGHT',event:{id:'night',start:'2026-08-02T22:00:00+10:00',end:'2026-08-03T08:00:00+10:00'}};
const clipped=filterFacilityRowsBySegments([overnight],[{sourceType:'ddh',startDate:'2026-08-02',endDate:'2026-08-02'}]);
assert.equal(clipped[0].event.end,'2026-08-03T00:00:00+10:00');
const dst=filterFacilityRowsBySegments([{...overnight,event:{...overnight.event,start:'2026-10-03T23:00:00+10:00',end:'2026-10-04T08:00:00+11:00'}}],[{sourceType:'ddh',startDate:'2026-10-04',endDate:'2026-10-04'}]);
assert.equal(dst[0].event.start,'2026-10-04T00:00:00+10:00'); // midnight is before the DST change
// Query planner must use the new bounded identity/term index.
const plan=sqlite.prepare(`EXPLAIN QUERY PLAN SELECT term_start FROM facility_term_staff_contributions WHERE source_type=? AND doctor_key=? AND term_start>=? AND term_start<=?`).all('ddh','TRAINEE','2026-05-04','2026-08-03');
assert.ok(plan.some(row=>row.detail.includes('idx_facility_staff_identity_terms')));
// Exercise browser scope normalization and empty restricted scope without DOM.
const app=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const sanitise=app.slice(app.indexOf('function sanitizeFacilityOverviewAccess(value)'),app.indexOf('function facilityOverviewSnapshotContext()'));
const { facilityAccessKeys, restrictedFacilityScope, FACILITY_ACCESS_VERSION } = await import('../public/static/facility-access-policy.js');
const browserScope=runInNewContext(`${sanitise}; sanitizeFacilityOverviewAccess(value)`,{value:{mode:'sites',facilityKeys:['MMC','DDH'],canSearchHistory:true},facilityAccessKeys,restrictedFacilityScope,FACILITY_ACCESS_VERSION});
assert.equal(browserScope.mode,'sites');
assert.equal(browserScope.facilityKeys.length,2);
// Browser colleague options come only from the verified historical period,
// even if a global picker or pinned result contains unrelated current staff.
const directoryFunction = app.slice(app.indexOf('function facilityOverviewTogetherStaffOptions()'),app.indexOf('function initializeFacilityOverviewTogetherState()'));
const historicalContext = {key:'historic-ddh',members:[{doctorKey:'FORMER PEER',displayName:'Former Peer',sourceType:'ddh'},{doctorKey:'LOCUM PEER',displayName:'Locum Peer',sourceType:'mmc'}]};
const browserState = {togetherContext:historicalContext,togetherFacilityKey:'DDH',togetherPinnedDoctors:[{key:'PRIVATE PEER',displayName:'Private Peer',sourceType:'mch'}]};
const browserGlobals = {facilityOverviewState:browserState,facilityOverviewTogetherContextKey:()=> 'historic-ddh',dedupeDoctorOptions:items=>items,doctorIdentityKey:doctor=>doctor.key,
  availableRosterDoctors:[{key:'PRIVATE PEER',displayName:'Private Peer',sourceType:'mch'}]};
assert.deepEqual(JSON.parse(JSON.stringify(runInNewContext(`${directoryFunction}; facilityOverviewTogetherStaffOptions().map(doctor=>doctor.key)`,browserGlobals))),['FORMER PEER']);
assert.equal(runInNewContext(`${directoryFunction}; facilityOverviewTogetherStaffOptions().length`,{...browserGlobals,facilityOverviewTogetherContextKey:()=> 'different-term'}),0,'stale period names must disappear while a new scope loads');
// All 16 supported linked roster identities stay inside the default 64-statement
// guard even on a cold current-access cache followed by historical context.
const aliases = Array.from({length:16},(_,i)=>({sourceType:'mmc',key:`ALIAS ${i}`,displayName:`Alias ${i}`}));
for (const claim of aliases) membership('mmc','2026-08-03',claim.key,'HMO',`alias-file-${claim.key}`);
account('aliases@example.com',aliases);
await call({action:'queryFacilityOverviewTogetherContext',email:'aliases@example.com',startDate:'2026-10-02',endDate:'2026-10-02'},200,64);
assert.ok(queries.length < 64);
globalThis.Date=RealDate;
console.log('Clinician access passed CMO parity, locums, historical rotations, direct API denial, impersonation, corrections, cache invalidation, related Insights, missing history and bounded indexed reads.');
