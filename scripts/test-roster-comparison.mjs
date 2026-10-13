import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, readdir } from 'node:fs/promises';
import { DatabaseSync } from 'node:sqlite';
import { runInNewContext } from 'node:vm';
import { gzipSync } from 'node:zlib';
import { onRequestPost as statePost } from '../functions/api/state.js';
import { onRequestPost as publishPost } from '../functions/api/automation/doctor-calendars.js';
import { publishDoctorSource } from './process-doctor-source.mjs';
import { doctorSourceKey, doctorSourceCurrentKey, doctorCatalogueShard, DOCTOR_SOURCE_VERSION, DOCTOR_SOURCE_LAYOUT } from '../functions/_lib/doctor-source-artifacts.js';
import { facilityMetadataManifestKey, facilityMonthObjectKey } from '../functions/_lib/facility-overview-cache.js';
import { loadRosterComparison, loadRosterComparisonDoctors, COMPARISON_INDEX_VERSION } from '../functions/_lib/roster-comparison.js';
import { comparisonWorkingEvent, comparisonIntervalMatches, normalizeComparisonEvent, comparisonPresence } from '../public/static/comparison-policy.js';
import { createRosterComparisonCache } from '../public/static/roster-comparison-cache.js';

const hash = value => createHash('sha256').update(value).digest('hex');
class Bucket {
  objects = new Map(); reads = []; heads = [];
  async put(key, value, options = {}) {
    if (options.onlyIf?.etagMatches && hash(this.objects.get(key) || '') !== options.onlyIf.etagMatches || options.onlyIf?.etagDoesNotMatch === '*' && this.objects.has(key)) return null;
    this.objects.set(key, typeof value === 'string' ? Buffer.from(value) : value instanceof ReadableStream ? Buffer.from(await new Response(value).arrayBuffer()) : Buffer.from(value)); return { key };
  }
  async head(key) { this.heads.push(key); return this.objects.has(key) ? { etag: hash(this.objects.get(key)) } : null; }
  async get(key) { this.reads.push(key); const bytes = this.objects.get(key); return bytes ? { body: bytes, etag: hash(bytes), arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length) } : null; }
}
const sourceType = 'mmc', month = '2026-10', startDate = '2026-10-01', endDate = '2026-10-31', today = '2026-10-14';
const event = (key, start, end, rawValue = 'AM') => ({ sourceType, doctorKey: key, displayName: key, event: { id: key + start + rawValue, doctorKey: key, source: 'MMC', kind: 'shift', title: rawValue, rawValue, start, end } });
const fixtureRows = [event('A', '2026-10-07T08:00:00+11:00', '2026-10-07T17:30:00+11:00'), event('B', '2026-10-07T18:00:00+11:00', '2026-10-07T23:00:00+11:00'),
  event('C', '2026-10-07T10:00:00+11:00', '2026-10-07T15:00:00+11:00'), event('NIGHT', '2026-09-30T23:00:00+10:00', '2026-10-01T08:00:00+10:00'),
  event('DAY ONE', '2026-10-01T08:00:00+10:00', '2026-10-01T17:00:00+10:00'), event('LEAVE', '2026-10-07T08:00:00+11:00', '2026-10-07T16:00:00+11:00', 'Annual leave')];
const bucket = new Bucket();
async function request(input, options) { return artifactRequest(bucket,input,options); }
async function artifactRequest(target,input, { binary = false, metadata = false, body } = {}) {
  const headers = { authorization: 'Bearer fixture', 'Content-Type': 'application/json' };
  let payload = JSON.stringify(input);
  if (body !== undefined) { payload = Buffer.from(body); headers['Content-Type'] = 'application/octet-stream'; headers['Content-Length'] = String(payload.length); headers['x-doctor-artifact'] = Buffer.from(JSON.stringify({ ...input, bytes: payload.length })).toString('base64'); }
  const response = await publishPost({ env: { ROSTER_FILES: target, ROSTER_AUTOMATION_TOKEN: 'fixture', ROSTER_AUTOMATION_WRITES_ENABLED: 'true' }, request: new Request('https://fixture', { method: 'POST', headers, body: payload }) });
  assert.equal(response.status, 200);
  if (binary) { const bytes = new Uint8Array(await response.arrayBuffer()); return metadata ? { bytes, etag: response.headers.get('x-object-etag'), key: input.key } : bytes; } return response.json();
}
async function publish(rows) {
  const revision = hash(JSON.stringify(rows));
  await bucket.put(facilityMonthObjectKey(sourceType, month, revision), gzipSync(JSON.stringify({ sourceType, month, rows })));
  const staffKey = `facility-overview/v1/mmc/staff/2026-08-03/${hash('staff')}.json.gz`;
  await bucket.put(staffKey, JSON.stringify({ members: rows.map(row => ({ ...row, event: undefined })), coverage: [{ startDate, endDate }], seniorityOverrides: [] }));
  await bucket.put(facilityMetadataManifestKey(sourceType), JSON.stringify({ sourceType, revision, terms: [{ termStart: '2026-08-03', termEnd: '2026-11-01', staffKey, coverage: [{ startDate, endDate }] }], months: { [month]: { key: facilityMonthObjectKey(sourceType, month, revision), revision } } }));
  await publishDoctorSource({ sourceType, request });
}
await publish(fixtureRows);
const options = { startDate, endDate, sourceTypes: ['mmc'], authorisedSegments: [{ sourceType, startDate, endDate }], doctorKeys: ['A', 'B'], scopeRevision: 'authorized-period' };
let result = await loadRosterComparison(bucket, options, today, comparisonWorkingEvent);
assert.deepEqual(result.events.map(row => row.doctorKey), ['A', 'B']);
assert.equal(bucket.reads.some(key => key.includes('/people/')), true);
const selectedRevision = result.revision, selectedPeople = result.personRevisions;
bucket.reads = [];
assert.equal((await loadRosterComparison(bucket, { ...options, cachedRevision: result.revision }, today, comparisonWorkingEvent)).unchanged, true);
assert.equal(bucket.reads.some(key => /\/(people|rows|months)\//.test(key)), false, 'unchanged check must never decode event bodies');
const discovery = { ...options, doctorKeys: [], overlapDoctorKeys: ['A'], excludeDoctorKeys: ['A'] };
const doctors = await loadRosterComparisonDoctors(bucket, discovery, today);
assert.deepEqual(doctors.doctors.map(row => row.doctorKey), ['B', 'C']);
bucket.reads = [];
assert.equal((await loadRosterComparisonDoctors(bucket, { ...discovery, cachedRevision: doctors.revision }, today)).unchanged, true);
assert.equal(bucket.reads.some(key => key.includes('/people/') || key.includes('/rows/')), false);
const nightRange = { ...options, endDate: startDate, authorisedSegments: [{ sourceType, startDate, endDate: startDate }], doctorKeys: ['NIGHT'] };
assert.equal((await loadRosterComparison(bucket, nightRange, today, comparisonWorkingEvent)).events[0].event.start, '2026-10-01T00:00:00+10:00');
assert.deepEqual((await loadRosterComparisonDoctors(bucket, { ...nightRange, overlapDoctorKeys: ['NIGHT'], excludeDoctorKeys: ['NIGHT'] }, today)).doctors.map(row => row.doctorKey), ['DAY ONE']);
const changedC = fixtureRows.map(row => row.doctorKey === 'C' ? { ...row, event: { ...row.event, title: 'Changed colleague' } } : row);
await publish(changedC);
assert.equal((await loadRosterComparison(bucket, { ...options, cachedRevision: selectedRevision }, today, comparisonWorkingEvent)).unchanged, true, 'unrelated colleague edit leaves selected bodies reusable');
await publish(changedC.map(row => row.doctorKey === 'B' ? { ...row, event: { ...row.event, end: '2026-10-07T23:30:00+11:00' } } : row));
assert.notEqual((await loadRosterComparison(bucket, { ...options, cachedRevision: selectedRevision }, today, comparisonWorkingEvent)).unchanged, true, 'peer-only edit invalidates pair');
bucket.reads=[];
const peerDelta=await loadRosterComparison(bucket,{...options,cachedPersonRevisions:selectedPeople},today,comparisonWorkingEvent);
assert.equal(peerDelta.delta,true);assert.deepEqual(peerDelta.updatedDoctorKeys,['B']);assert.equal(bucket.reads.filter(key=>key.includes('/people/')).length,1,'peer-only edit reads only the changed peer');

const originalCurrent = bucket.objects.get(doctorSourceCurrentKey(sourceType));
const current = JSON.parse(originalCurrent);
delete current.comparisonIndexes[month];
await bucket.put(doctorSourceCurrentKey(sourceType), JSON.stringify(current));
assert.equal((await loadRosterComparisonDoctors(bucket, discovery, today)).unavailable, true, 'missing monthly index never means no colleagues');
bucket.objects.set(doctorSourceCurrentKey(sourceType), originalCurrent);
const originalManifest = bucket.objects.get(facilityMetadataManifestKey(sourceType));
await bucket.put(facilityMetadataManifestKey(sourceType), JSON.stringify({ ...JSON.parse(originalManifest), revision: 'not-prepared' }));
assert.equal((await loadRosterComparison(bucket, options, today, comparisonWorkingEvent)).unavailable, true, 'stale parent cannot be served');
bucket.objects.set(facilityMetadataManifestKey(sourceType), originalManifest);

// Published previous-term nights stay visible without releasing the next term.
const boundary=new Bucket(),november='2026-11';
const boundaryRows=[event('PREVIOUS NIGHT','2026-11-01T23:00:00+11:00','2026-11-02T08:00:00+11:00'),event('OTHER NIGHT','2026-11-01T23:00:00+11:00','2026-11-02T08:00:00+11:00'),event('UNRELEASED','2026-11-02T08:00:00+11:00','2026-11-02T17:00:00+11:00')];
const boundaryRevision=hash(JSON.stringify(boundaryRows));
await boundary.put(facilityMonthObjectKey(sourceType,november,boundaryRevision),JSON.stringify({sourceType,month:november,rows:boundaryRows}));
const boundaryTerms=[];
for(const [termStart,termEnd] of [['2026-08-03','2026-11-01'],['2026-11-02','2027-01-31']]){
  const staffKey=`facility-overview/v1/mmc/staff/${termStart}/${hash(termStart)}.json.gz`;
  await boundary.put(staffKey,JSON.stringify({coverage:[{startDate:termStart,endDate:termEnd}],seniorityOverrides:[]}));boundaryTerms.push({termStart,termEnd,staffKey});
}
await boundary.put(facilityMetadataManifestKey(sourceType),JSON.stringify({sourceType,revision:boundaryRevision,terms:boundaryTerms,months:{[november]:{key:facilityMonthObjectKey(sourceType,november,boundaryRevision),revision:boundaryRevision}}}));
await publishDoctorSource({sourceType,request:(input,options)=>artifactRequest(boundary,input,options)});
const boundaryQuery={...options,startDate:'2026-11-02',endDate:'2026-11-02',authorisedSegments:[{sourceType,startDate:'2026-11-02',endDate:'2026-11-02'}]};
const boundaryDoctors=await loadRosterComparisonDoctors(boundary,{...boundaryQuery,overlapDoctorKeys:['PREVIOUS NIGHT'],excludeDoctorKeys:['PREVIOUS NIGHT']},'2026-09-30');
assert.deepEqual(boundaryDoctors.doctors.map(row=>row.doctorKey),['OTHER NIGHT'],'carry-in discovery never discloses an unreleased term');
const boundaryEvents=await loadRosterComparison(boundary,{...boundaryQuery,doctorKeys:['PREVIOUS NIGHT','UNRELEASED']},'2026-09-30',comparisonWorkingEvent);
assert.deepEqual(boundaryEvents.events.map(row=>row.doctorKey),['PREVIOUS NIGHT']);
assert.equal(normalizeComparisonEvent({start:'2026-10-04T15:00:00+11:00',end:'2026-10-04T00:00:00+10:00'}).end,'2026-10-05T00:00:00+11:00');
assert.equal(comparisonPresence([{title:'PM',start:'2026-10-07T15:00:00+11:00',end:'2026-10-07T00:00:00+11:00'}],'mmc')['2026-10'],192);
assert.equal((await loadRosterComparison(bucket,{...options,doctorKeys:['A'],doctorIdentities:[{key:'A',sourceType:'mch'}]},today,comparisonWorkingEvent)).events.length,0,'matching names at another source cannot select this identity');

// Execute real authentication, historical scope and both endpoint adapters.
const sqlite = new DatabaseSync(':memory:');
for (const name of (await readdir(new URL('../migrations', import.meta.url))).filter(name => name.endsWith('.sql')).sort()) sqlite.exec(await readFile(new URL('../migrations/' + name, import.meta.url), 'utf8'));
for (const [email, role] of [['creator@example.test', 'creator'], ['staff@example.test', 'user']]) sqlite.prepare('INSERT INTO account_profiles(email,real_name,role,password_salt,password_hash) VALUES(?,?,?,?,?)').run(email, email, role, 'salt', hash('salt:test'));
let sql = [];
const db = { prepare(query) { sql.push(query); let args = []; return { bind(...values) { args = values; return this; }, async all() { return { results: sqlite.prepare(query).all(...args) }; }, async first() { return sqlite.prepare(query).get(...args) || null; }, async run() { throw Error('Comparison must never write D1'); } }; } };
const env = { ROSTER_DB: db, ROSTER_FILES: bucket, ROSTER_COMPARISON_ENABLED: 'true', FACILITY_OVERVIEW_MAINTENANCE_MODE: 'false', FACILITY_SHARED_ROLLOUT_ACTIVE: 'true', FACILITY_SHARED_EMERGENCY_PAUSED: 'false', FACILITY_LEGACY_READS_PAUSED: 'true', FACILITY_SHARED_READER_SOURCE_ALLOWLIST: 'mmc', FACILITY_SHARED_READER_COHORT: 'all', FACILITY_SHARED_METADATA_ENABLED: 'true', FACILITY_SHARED_DAYS_ENABLED: 'true' };
async function api(action, extra = {}) {
  return statePost({ env, request: new Request('https://fixture/api/state', { method: 'POST', body: JSON.stringify({ action, email: 'creator@example.test', password: 'test', startDate, endDate, sourceTypes: ['mmc'], doctorKeys: ['A', 'B'], overlapDoctorKeys: ['A'], excludeDoctorKeys: ['A'], comparison: true, ...extra }) }) });
}
const togetherResponse = await api('queryFacilityOverviewWorkingTogether');
const together = await togetherResponse.json(); assert.equal(togetherResponse.status, 200, JSON.stringify(together));
const whenResponse = await api('queryRosterInsights'); assert.equal(whenResponse.status, 200);
const when = await whenResponse.json();
assert.deepEqual(when.coworkers, together.events, 'both selected interfaces read exactly the same facts');
assert.equal(when.scopeRevision, together.scopeRevision, 'same authorized period permits common facts cache');
assert.equal((await api('queryFacilityOverviewTogetherContext')).status, 200);
bucket.reads = []; const conditional = await api('queryFacilityOverviewWorkingTogether', { cachedRevision: together.revision });
assert.equal((await conditional.json()).unchanged, true); assert.equal(bucket.reads.some(key => key.includes('/people/')), false);
assert.equal((await api('queryRosterOverlapDoctors')).status, 200);
assert.equal((await api('queryRosterInsights', { targetEmail: 'staff@example.test' })).status, 403, 'God mode cannot inherit Creator access');
assert.equal((await api('queryFacilityOverviewWorkingTogether', { targetEmail: 'staff@example.test' })).status, 403);
assert.equal(sql.some(query => /\b(roster_events|roster_event_sources)\b/.test(query)), false, 'no D1 event searching');
assert.equal((await api('queryRosterInsights', { endDate: '2027-06-01' })).status, 400, 'retain six-month insight limit');

// The existing day matcher and the shared time matcher deliberately differ.
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
const dayMatch = app.slice(app.indexOf('function buildOverlapDays('), app.indexOf('function indexEventsByDay('));
const dayGlobals = { indexEventsByDay: events => new Map(events.map(event => [event.start.slice(0, 10), [event]])), eventSourceCode: event => event.source };
runInNewContext(`${dayMatch};this.match=buildOverlapDays`, dayGlobals);
assert.equal(dayGlobals.match([fixtureRows[0].event], [fixtureRows[1].event]).length, 1);
const intervals = rows => rows.map(row => ({ facility: row.sourceType, start: new Date(row.event.start), end: new Date(row.event.end), event: row.event }));
const indexed = rows => rows.map(row => ({ doctor: row.doctorKey, intervals: intervals([row]) }));
assert.equal(comparisonIntervalMatches(indexed(fixtureRows.slice(0, 2)), options, () => 0).length, 0);
assert.equal(comparisonIntervalMatches(indexed([fixtureRows[0], fixtureRows[2]]), options, () => 0).length, 1);
assert.equal(comparisonWorkingEvent(fixtureRows.at(-1).event), false);
assert.equal(comparisonWorkingEvent({ title: 'DDH HITH' }, { facilityKey: 'DDH' }), false);
const dstRows = [event('X', '2026-10-03T23:00:00+10:00', '2026-10-04T08:00:00+11:00'), event('Y', '2026-10-04T07:00:00+11:00', '2026-10-04T10:00:00+11:00')];
assert.equal(comparisonIntervalMatches(indexed(dstRows), { startDate: '2026-10-03', endDate: '2026-10-05' }, () => 0).length, 1);

// Real cache behavior: fresh reuse, cross-interface reuse after authorization,
// persistent results, conditional requests, expiry, denial and stale responses.
let time = 1_000_000, calls = 0, writes = 0;
const disk = new Map(), storage = { get: async key => disk.get(key), put: async (key, record) => { writes++; disk.set(key, record); } };
const cache = createRosterComparisonCache({ storage, now: () => time });
const query = { startDate, endDate, sourceTypes: ['mmc'], doctorKeys: ['A', 'B'] }, leaseQuery = { startDate, endDate, sourceTypes: ['mmc'] };
const send = async (_, revision) => { calls++; return { ok: true, revision: 'facts', scopeRevision: 'scope', accessExpiresAt: new Date(time + 600_000).toISOString(), ...(revision ? { unchanged: true } : { events: together.events }) }; };
const authorize = (client, feature) => client.request({ owner: 'creator:staff', feature, kind: 'context', query: leaseQuery, send });
await authorize(cache, 'when');
const args = { owner: 'creator:staff', feature: 'when', kind: 'selected', query, leaseQuery, send };
const first = await cache.request(args); assert.equal(calls, 2);
assert.deepEqual((await cache.request(args)).events, first.events); assert.equal(calls, 2, 'fresh reopen spends no authenticated request');
await authorize(cache, 'together'); await cache.request({ ...args, feature: 'together' }); assert.equal(calls, 3, 'other feature reuses facts after its own authorization');
time += 61_000; await authorize(cache, 'when'); let shown = 0; const priorWrites = writes;
await cache.request({ ...args, onCached: () => shown++ }); assert.equal(shown, 1); assert.equal(writes, priorWrites, 'unchanged context and selected payloads are not rewritten');
const restarted = createRosterComparisonCache({ storage, now: () => time }); await authorize(restarted, 'when'); const beforeRestart = calls;
await restarted.request(args); assert.equal(calls, beforeRestart + 1, 'verified new session conditionally validates older persistent facts');
cache.clear(); let release;
const pending = cache.request({ ...args, send: () => new Promise(resolve => { release = resolve; }) });
await Promise.resolve(); cache.clear(); release({ ok: true, revision: 'late', scopeRevision: 'scope', accessExpiresAt: new Date(time + 600000).toISOString(), events: [] });
await assert.rejects(pending, error => error.name === 'AbortError');
await authorize(cache, 'when');
await assert.rejects(cache.request({ ...args, force: true, send: async () => { throw Object.assign(Error('Denied'), { status: 403 }); } }), /Denied/);
const beforeDenied = calls; await cache.request(args); assert.equal(calls, beforeDenied + 1, 'denial clears authorization and memory');
const factCache=createRosterComparisonCache({storage:{get:async()=>null,put:async()=>{}},now:()=>time});
let factRequests=0;
const realSend=async(query,cachedRevision,cachedPersonRevisions)=>{factRequests++;return {...await loadRosterComparison(bucket,{...options,...query,cachedRevision,cachedPersonRevisions},today,comparisonWorkingEvent),accessExpiresAt:new Date(time+600000).toISOString()};};
await factCache.request({owner:'facts',feature:'together',kind:'context',query:leaseQuery,send:async()=>({ok:true,scopeRevision:options.scopeRevision,sourceTypes:['mmc'],accessExpiresAt:new Date(time+600000).toISOString()})});
const factArgs={owner:'facts',feature:'together',kind:'selected',query,leaseQuery,send:realSend};
await factCache.request(factArgs);bucket.reads=[];
assert.deepEqual((await factCache.request({...factArgs,query:{...query,doctorKeys:['A','C']}})).events.map(row=>row.doctorKey).sort(),['A','C']);
assert.equal(bucket.reads.filter(key=>key.includes('/people/')).length,1,'new pair fetches only the person not already cached');
const beforePair=factRequests;
assert.deepEqual((await factCache.request({...factArgs,query:{...query,doctorKeys:['B','C']}})).events.map(row=>row.doctorKey).sort(),['B','C']);
assert.equal(factRequests,beforePair,'new pairing of cached people needs no network');
time+=100;factCache.invalidate();
await factCache.request({owner:'facts',feature:'together',kind:'context',query:leaseQuery,send:async()=>({ok:true,scopeRevision:options.scopeRevision,sourceTypes:['mmc'],accessExpiresAt:new Date(time+600000).toISOString()})});
bucket.reads=[];
await factCache.request({...factArgs,query:{...query,doctorKeys:['B','C']}});
assert.equal(bucket.reads.some(key=>key.includes('/people/')),false,'assembled pair conditionally validates its per-person revisions without event bodies');

// A fresh authorized broad result can be narrowed locally.
const filterCache=createRosterComparisonCache({storage:{get:async()=>null,put:async()=>{}},now:()=>time});let filteredRequests=0;
const broad={startDate,endDate,sourceTypes:[]};
const discoverySend=async()=>{filteredRequests++;return {ok:true,scopeRevision:'scope',sourceTypes:['mmc','ddh'],accessExpiresAt:new Date(time+600000).toISOString(),doctors:[{doctorKey:'B',displayName:'B',sourceType:'mmc'},{doctorKey:'D',displayName:'D',sourceType:'ddh'}]};};
await filterCache.request({owner:'filters',feature:'when',kind:'discover',query:{...broad,overlapDoctorKeys:['A']},send:discoverySend});
const filtered=await filterCache.request({owner:'filters',feature:'when',kind:'discover',query:{...broad,sourceTypes:['mmc'],overlapDoctorKeys:['A']},send:discoverySend});
assert.deepEqual(filtered.doctors.map(row=>row.doctorKey),['B']);assert.equal(filteredRequests,1,'narrowing a fresh authorized directory is local');
// Exercise the actual UI adapter and loader with the real cache implementation.
let uiCalls=0,uiRenders=0;
const uiState={requestId:0,tab:'together',togetherContent:'',togetherStaffKeys:['A','B'],togetherFacilityKey:'ALL'};
const ui={rosterComparisonCache:createRosterComparisonCache({storage:{get:async()=>null,put:async()=>{}}}),
  normalizeEmail:value=>value,currentUserEmail:'creator@test',currentUserPassword:'fixture',authUserEmail:'creator@test',authUserPassword:'fixture',
  activeWorkspaceOwnerKey:()=> 'profile:A',facilityOverviewTargetEmail:()=> 'viewed@test',
  canUseFullFacilityOverview:()=>true,canUseFacilityOverview:()=>true,currentFacilityOverviewMaintenance:false,facilityOverviewState:uiState,
  facilityOverviewTogetherDateRange:()=>({startDate,endDate}),facilityOverviewTogetherContextKey:()=> 'context:viewed',
  renderFacilityOverview:()=>uiRenders++,refreshFacilityOverviewSnapshotAccess:()=>{},initializeFacilityOverviewTogetherState:()=>{},
  facilityOverviewTogetherStaffOptions:()=>[{key:'A',identity:'A',displayName:'A',sourceTypes:['mmc']},{key:'B',identity:'B',displayName:'B',sourceTypes:['mmc']}],
  facilityOverviewTogetherDoctorKeys:doctor=>[doctor.key],normalizedDoctorSourceTypes:doctor=>doctor.sourceTypes,normalizeRosterName:value=>value,
  validFacilityDateRange:()=>true,renderFacilityOverviewCoverageNotice:()=>'',renderFacilityOverviewTogetherResults:rows=>`Matched ${rows.length} shifts`,
  facilityOverviewRequestWasCancelled:error=>error.name==='AbortError',escapeHtml:value=>value,
  readJsonResponse:response=>response.json(),fetch:async(_,request)=>{uiCalls++;const body=JSON.parse(request.body);assert.equal(body.targetEmail,'viewed@test');assert.equal(body.comparison,true);
    return Response.json({ok:true,sourceTypes:['mmc'],scopeRevision:'ui',revision:'ui-facts',accessExpiresAt:new Date(Date.now()+600000).toISOString(),...(body.action==='queryFacilityOverviewTogetherContext'?{members:[]}:{events:together.events})});},
};
const adapters=app.slice(app.indexOf('function comparisonDoctorIdentities('),app.indexOf('async function fetchRosterInsightRows('));
const togetherLoader=app.slice(app.indexOf('async function loadFacilityOverviewTogether()'),app.indexOf('function renderFacilityOverviewTogetherResults('));
runInNewContext(`${adapters};${togetherLoader};this.load=loadFacilityOverviewTogether`,ui);
assert.equal(await ui.load(),true);assert.equal(uiCalls,2);assert.equal(uiState.togetherContent,'Matched 2 shifts');
const firstRenders=uiRenders;assert.equal(await ui.load(),true);assert.equal(uiCalls,2);assert.equal(uiRenders,firstRenders,'fresh reopen performs neither network nor result repaint');
uiState.togetherContent='';assert.equal(await ui.load(),true);assert.equal(uiState.togetherContent,'Matched 2 shifts','returning to an earlier picker selection cannot strand the loading message');
console.log('Shared comparisons pass real publication/API parity, metadata-only checks, peer invalidation, carry-in, access isolation, explicit matching modes and cache lifecycle.');

// Synthetic scale check: selected event reads stay fixed as the department grows.
for (const count of [100, 500, 1000]) {
  const scaled = new Bucket(), groups = Array.from({ length: 16 }, () => Object.create(null));
  const annualStart = '2026-01-01', annualEnd = '2026-12-31';
  const months = Array.from({length:12}, (_, n) => `2026-${String(n + 1).padStart(2, '0')}`);
  const terms = [['01-01','03-31'],['04-01','06-30'],['07-01','09-30'],['10-01','12-31']].map(([start, end]) => ({termStart:'2026-'+start, termEnd:'2026-'+end, coverage:[{startDate:'2026-'+start,endDate:'2026-'+end}]}));
  const manifestKey = facilityMetadataManifestKey(sourceType); await scaled.put(manifestKey, JSON.stringify({ terms, months: Object.fromEntries(months.map(month => [month, {}])) }));
  const members = [];
  for (let n = 0; n < count; n++) {
    const doctorKey = 'PERSON ' + n;
    const events = months.flatMap(month => Array.from({length:16}, (_, day) => { const date = `${month}-${String(day + 1).padStart(2, '0')}`, offset = date < '2026-04-05' || date >= '2026-10-04' ? '+11:00' : '+10:00'; return event(doctorKey, date+'T08:00:00'+offset, date+'T17:00:00'+offset).event; }));
    const body = JSON.stringify({ version: DOCTOR_SOURCE_VERSION, sourceType, doctorKey, events });
    const digest = hash(body), key = doctorSourceKey(sourceType, 'people', digest); await scaled.put(key, body);
    groups[parseInt(await doctorCatalogueShard(doctorKey), 16)][doctorKey] = { key, digest }; members.push([doctorKey, doctorKey, 65535]);
  }
  const shards = {};
  for (let n = 0; n < 16; n++) { const shard = n.toString(16), body = JSON.stringify({ version: DOCTOR_SOURCE_VERSION, sourceType, shard, doctors: groups[n] }), digest = hash(body), key = doctorSourceKey(sourceType, 'catalogues', digest); await scaled.put(key, body); shards[shard] = { key, digest }; }
  const comparisonIndexes = {}; let indexBytes = 0;
  for (const month of months) {
    const indexBody = JSON.stringify({version:COMPARISON_INDEX_VERSION,sourceType,month,doctors:members.map(([key,name,mask])=>[key,name,mask,{[terms.find(term=>term.termStart.slice(0,7)<=month&&term.termEnd.slice(0,7)>=month).termStart]:mask}])}), digest=hash(indexBody), key=doctorSourceKey(sourceType,'catalogues',digest);
    await scaled.put(key,indexBody); comparisonIndexes[month]={digest,key}; indexBytes=Buffer.byteLength(indexBody);
  }
  await scaled.put(doctorSourceCurrentKey(sourceType), JSON.stringify({ version: DOCTOR_SOURCE_VERSION, layout: DOCTOR_SOURCE_LAYOUT, parentEtag: (await scaled.head(manifestKey)).etag, digest: hash('catalogue'), catalogueKey: doctorSourceKey(sourceType, 'catalogues', hash('catalogue')), metadata: { terms, months }, shards, comparisonIndexes }));
  const annualOptions = {...options,startDate:annualStart,endDate:annualEnd,authorisedSegments:[{sourceType,startDate:annualStart,endDate:annualEnd}],doctorKeys:['PERSON 0','PERSON 1']};
  const began = performance.now();
  const selected = await loadRosterComparison(scaled, annualOptions, today, comparisonWorkingEvent);
  assert.equal(selected.events.length,384); assert.equal(selected.missing.length,0); assert.equal(scaled.reads.filter(key => key.includes('/people/')).length,2);
  const selectedNodeMs = Math.round(performance.now()-began);
  scaled.reads=[];
  assert.equal((await loadRosterComparison(scaled,{...annualOptions,cachedRevision:selected.revision},today,comparisonWorkingEvent)).unchanged,true);
  assert.equal(scaled.reads.some(key=>key.includes('/people/')),false);
  const discovered = await loadRosterComparisonDoctors(scaled, { ...discovery, overlapDoctorKeys: ['PERSON 0'], excludeDoctorKeys: ['PERSON 0'] }, today);
  assert.equal(discovered.doctors.length, count - 1);
  console.log(JSON.stringify({ staff: count, selectedEventBodies: 2, annualEventsReturned:384, selectedNodeMs, discoveryIndexBytes:indexBytes, discoveryBodies:1 }));
}
