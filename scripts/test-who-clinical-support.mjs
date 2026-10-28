import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {findmyshiftRowsWorkbook} from '../functions/_lib/findmyshift.js';
import {buildAutomatedDerivedRosterPayload} from '../functions/_lib/automation-import.js';
import {clinicalSupportRosterMode} from '../public/static/roster.js';

const app=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const grouping=app.slice(app.indexOf('function whoTeamLabel('),app.indexOf('function whoTeamRank('));
const team=runInNewContext(grouping+';whoTeamLabel',{clinicalSupportRosterMode});
// The real provider-workbook adapter and automated parser preserve the raw
// long label. Who-is-working must group that correctly, even without an
// explicit per-grade parser extension in the browser.
const rows=[
 {name:'Alex Brown',label:'Clinical Support',start:'08:00',end:'17:30'},
 {name:'Jamie Smith',label:'Clinical Support',start:'',end:''},
 {name:'Sam Jones',label:'CS',start:'',end:''},
 {name:'Pat Green',label:'CS onsite',start:'08:00',end:'17:30'},
].map(row=>({...row,date:'2026-10-29',seniority:'SMS',facility:'',comment:''}));
const bytes=await findmyshiftRowsWorkbook(rows);
const payload=await buildAutomatedDerivedRosterPayload({sourceId:'dandenong-findmyshift',contentHash:'synthetic-cs',file:new File([bytes],'Dandenong-FindMyShift-cs.xlsx')});
const events=Object.values(payload.eventsByDoctor).flat();
assert.equal(events.length,4);
for(const event of events){
 const original=structuredClone(event);
 assert.equal(team(event),event.title.replace(/^DDH:\s*/,''),'Who grouping preserves the parsed CS label');
 assert.notEqual(team(event),'Clinic');assert.equal(event.seniority,'SMS');assert.deepEqual(event,original,'grouping never mutates roster facts');
}
assert.ok(events.some(e=>e.rawValue==='Clinical Support'&&!e.allDay));
assert.ok(events.some(e=>e.rawValue==='Clinical Support'&&e.allDay));
assert.equal(team({source:'MMC',title:'MMC: Clinic PM',rawValue:'PCC'}),'Clinic','real MMC clinic remains a clinic');
assert.equal(team({source:'MMC',title:'MMC: CS Office',rawValue:'CSO'}),'CS Office');
assert.equal(team({source:'MCH',title:'MCH: SMS',rawValue:'0800-1730 OCS'}),'SMS','legacy CS classifications retain their existing title rather than becoming Clinic');
assert.equal(team({source:'DDH',title:'DDH: Clinical Assistant',rawValue:'Clinical Assistant'}),'Clinical Assistant','the word clinical is not a Clinic stream');
assert.equal(team({source:'DDH',title:'DDH: CS Exam',rawValue:'Clinical Support ACEM OSCE'}),'CS Exam');
assert.equal(team({source:'DDH',title:'DDH: TOX Clinical Support',rawValue:'TOX Clinical Support'}),'TOX Clinical Support');
assert.equal(team({source:'DDH',title:'DDH: Orange AM',rawValue:'Orange AM'}),'Orange');
console.log('DDH provider → parser → Who grouping: timed/all-day Clinical Support, CS/onsite, grade/time preservation; MMC Clinic and specialised support labels passed.');
