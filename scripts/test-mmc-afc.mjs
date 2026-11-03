import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import {parseUploadForm,buildRosterView,doctorOptions,filterCalendarRosterEvents,setParserExtensions,normalizeMmcFloatRosterEvent} from '../public/static/roster.js';
const rows=Array.from({length:9},()=>[]);
rows[2]=['Role','Pager No','','Cost Centre','Name (Not Used)','Emp No','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
rows[3]=['Date','','','','','',...Array.from({length:7},(_,i)=>new Date(Date.UTC(2026,10,2+i)))];
rows[5]=['','','','SMS'];rows[6]=['','','','','Alex Brown','test','AFC','0900-1800 AFC','','','','',''];
const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([[]]),'Whole thing');XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(rows,{cellDates:true}),'Week 1');
const form=new FormData();form.append('rosterFiles',new File([XLSX.write(book,{type:'buffer',bookType:'xlsx'})],'AdultTerm4.2026.xlsx'));
const parsed=await parseUploadForm(new Request('https://test.invalid',{method:'POST',body:form}));setParserExtensions({});
const doctor=doctorOptions(parsed.sources.mmc,[]).find(d=>d.key==='ALEX BROWN');
const events=buildRosterView(parsed.sources.mmc,[],doctor.key).events;assert.equal(events.length,2);
assert.ok(events.every(e=>e.title==='MMC: Float AM'&&e.seniority==='SMS'));
assert.ok(events[0].start.includes('08:00'));assert.ok(events[0].end.includes('17:30'));
assert.ok(events[1].start.includes('09:00'));assert.ok(events[1].end.includes('18:00'));
const cached={...events[1],title:'MMC: AFC',status:'unknown',warning:'MMC shift code not recognised; using explicit roster time.'};
const fixed=filterCalendarRosterEvents([cached])[0];assert.equal(fixed.title,'MMC: Float AM');assert.equal(fixed.status,'ok');assert.equal(fixed.warning,'');assert.equal(fixed.start,cached.start);assert.equal(fixed.end,cached.end);assert.equal(fixed.rawValue,cached.rawValue);
for(const changed of [{source:'DDH'},{seniority:'HMO'},{rawValue:'AGC'}]){const other={...cached,...changed};assert.equal(normalizeMmcFloatRosterEvent(other),other);}
console.log('MMC AFC: SMS Float AM parser, explicit/default hours, cached correction and source/grade/code isolation passed.');
