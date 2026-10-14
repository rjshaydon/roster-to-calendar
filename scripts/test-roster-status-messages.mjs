import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {coverageNoticeRanges,routineCalendarStatus} from '../public/static/roster-coverage-notice.js';
import {rosterWritePausedResponse} from '../functions/_lib/roster-automation-guard.js';
const missing=[{sourceType:'ddh',startDate:'2026-01-01',endDate:'2026-02-01'},{sourceType:'mmc',startDate:'2026-01-01',endDate:'2026-02-01'},{sourceType:'mmc',startDate:'2026-11-02',endDate:'2027-01-31'}];
const term3={startDate:'2026-08-03',endDate:'2026-11-01',upcoming:{startDate:'2026-11-02',endDate:'2027-01-31'}};
assert.deepEqual(coverageNoticeRanges(missing,term3),{current:[],nextSources:['mmc']});
assert.deepEqual(coverageNoticeRanges(missing,{startDate:'2026-11-02',endDate:'2027-01-31'}).current,[missing[2]]);
assert.deepEqual(coverageNoticeRanges(missing,{startDate:'2026-01-20',endDate:'2026-01-31'}).current,[{sourceType:'ddh',startDate:'2026-01-20',endDate:'2026-01-31'},{sourceType:'mmc',startDate:'2026-01-20',endDate:'2026-01-31'}]);
assert.deepEqual(coverageNoticeRanges(missing,{...term3,sourceType:'DDH'}),{current:[],nextSources:[]});
assert.deepEqual(coverageNoticeRanges([{sourceType:'mch',month:'2026-10',startDate:'2026-01-01',endDate:'2027-01-31'}],{startDate:'2026-10-15',endDate:'2026-10-20'}).current,[{sourceType:'mch',startDate:'2026-10-15',endDate:'2026-10-20'}]);
assert.equal(coverageNoticeRanges([{sourceType:'mch',reason:'staff-unavailable'}],term3).current.length,1,'undated failures are not silently discarded');
const source=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const extract=(start,end)=>source.slice(source.indexOf(start),source.indexOf(end,source.indexOf(start)));
// Link the browser's actual imports. Inventing a clock stub here previously
// hid a missing browser binding and let a startup-breaking change pass.
const browserImports={};
for(const match of source.matchAll(/import\s*\{([^}]+)\}\s*from\s*['"](\.\/[^'"]+)['"]/g)) {
  const module=await import(new URL('../public/static/'+match[2].slice(2),import.meta.url));
  for(const specifier of match[1].split(',')) {
    const [exported,local]=specifier.trim().split(/\s+as\s+/);
    if(exported)browserImports[local||exported]=module[exported];
  }
}

const g={...browserImports,settings:{dateFrom:'2026-08-03',dateTo:'2026-11-01',hospitalFilter:'all'},parseDateOnly:d=>new Date(d+'T12:00:00'),australianTermForDate:()=>({start:new Date('2026-08-03T12:00:00')}),nextAustralianTerm:()=>({start:new Date('2026-11-02T12:00:00'),end:new Date('2027-02-01T12:00:00')}),formatDateKey:d=>d.toISOString().slice(0,10),addDays:(d,n)=>new Date(d.getTime()+n*86400000),eventSourceCodes:e=>[e.source.toUpperCase()],displaySourceCode:s=>s.toUpperCase(),escapeHtml:s=>String(s).replaceAll('<','&lt;'),formatDate:s=>s};
if(typeof g.australianDateKey==='function'){const clock=g.australianDateKey;g.australianDateKey=()=>clock(new Date('2026-10-14T12:00:00Z'));}
runInNewContext(extract('function renderCalendarCoverageNotice(', 'function renderFacilityOverviewStaffResults(')+';this.render=renderCalendarCoverageNotice;',g);
let html=g.render({previewStart:'2026-08-03',previewEnd:'2027-01-31',publicationMissing:missing});assert.match(html,/Next term: MMC/);assert.doesNotMatch(html,/2026-01|2026-02|incomplete/);
assert.match(g.render({publicationMissing:missing,coverageEvents:[{source:'mmc',start:'2026-11-03T08:00:00'}]}),/Next term: MMC roster coverage is incomplete/,'retained future shifts must not be described as wholly unavailable');
g.settings.dateFrom='2026-11-02';g.settings.dateTo='2027-01-31';html=g.render({publicationMissing:missing});assert.match(html,/coverage is incomplete.*MMC/);assert.doesNotMatch(html,/Next term/);
g.settings.dateFrom='2026-08-03';g.settings.hospitalFilter='DDH';assert.equal(g.render({publicationMissing:missing}),'');
const repair={currentAdvancedRosterRecoveryEnabled:false,hasPendingRosterAutomation:()=>false};runInNewContext(extract('function renderAdvancedRosterRecovery(', 'function renderLoginPerformanceCard(')+';this.render=renderAdvancedRosterRecovery;',repair);assert.equal(repair.render(),'');repair.currentAdvancedRosterRecoveryEnabled=true;assert.match(repair.render(),/data-replace-active-rosters/);repair.hasPendingRosterAutomation=()=>true;assert.doesNotMatch(repair.render(),/data-replace-active-rosters/);
assert.equal(routineCalendarStatus('Calendar loaded from cache.'),true);assert.equal(routineCalendarStatus('Changes couldn’t be saved.'),false);assert.equal(routineCalendarStatus('Calendar refreshed.'),false,'explicit refresh feedback stays visible');
let saveErrors=0;const saveFailure={savePayloadMatchesActiveCalendar:()=>false};
runInNewContext(extract('function reportCloudSaveFailure(', 'function scheduleCloudStateSave(')+';this.report=reportCloudSaveFailure;',saveFailure);
saveFailure.report(new Error('old save'),{},{reportErrors:false,onError:()=>saveErrors++});assert.equal(saveErrors,0,'an earlier account save must not show an error on the current account');
saveFailure.savePayloadMatchesActiveCalendar=()=>true;saveFailure.report(new Error('current save'),{},{reportErrors:false,onError:()=>saveErrors++});assert.equal(saveErrors,1);
const statusDOM={status:{},routineCalendarStatus,removeSupersededStatusMessages:()=>{},document:{createElement(){throw Error('Routine success must not add DOM messages');}}};
runInNewContext(extract('function setStatus(', 'function removeSupersededStatusMessages(')+';this.set=setStatus;',statusDOM);
statusDOM.set('Calendar loaded.');statusDOM.set('Calendar loaded from cache.');
assert.throws(()=>statusDOM.set('Changes couldn’t be saved.',true),/Routine success must not add DOM/,'real errors still reach the normal message rendering path');
assert.match((await rosterWritePausedResponse('Advanced roster repair is not enabled.').json()).error,/Advanced roster repair/);
assert.doesNotMatch((await rosterWritePausedResponse().json()).error,/quota|all roster updates/i);
console.log('Roster notices pass term/date clipping, next-term notices, hospital filters, incomplete coverage, disabled recovery and actionable errors.');
