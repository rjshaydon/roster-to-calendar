import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
import {onShiftLaunchWindow} from '../public/static/shift-launch-policy.js';

const source=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const start=source.indexOf('function facilityOverviewInitialShiftPeriod(');
const end=source.indexOf('function facilityOverviewContactRefreshIsActive(',start);
const state={facilityKey:'MMC',date:'2026-10-11',tab:'on-shift',requestId:1,onShiftData:[]};
let callbacks=[],preferred='ME',open=true,now=new Date('2026-10-11T15:00:00+11:00');
const body={scrollTop:0,getBoundingClientRect:()=>({top:100}),querySelectorAll:()=>[
 {dataset:{onShiftPeriod:'PM'},getBoundingClientRect:()=>({top:600})},
 {dataset:{onShiftPeriod:'Night'},getBoundingClientRect:()=>({top:1100})}
]};
const ui=vm.createContext({facilityOverviewState:state,facilityOverviewBody:body,currentNonClinical:false,
 currentRosterClaims:[{key:'MY ALIAS',sourceType:'mmc'}],preferredDoctorKeyForCurrentAccount:()=>preferred,
 normalizeRosterName:v=>String(v||'').toUpperCase(),eventRosterDateKey:e=>e.start.slice(0,10),
 isRosterShiftEvent:e=>!e.allDay&&e.kind!=='leave',
 onShiftLaunchWindow:events=>onShiftLaunchWindow(events,now),
 buildWhoAssignment:(_doctor,_metadata,event)=>({period:event.period}),
 isFacilityOverviewOpen:()=>open,requestAnimationFrame:callback=>callbacks.push(callback)});
vm.runInContext(source.slice(start,end),ui);
function row(period,{key='ME',site='mmc',date=state.date}={}) {
 const hours={AM:['07:00','16:00'],PM:['14:00','23:00'],Night:['23:00','08:00']};
 const [from,to]=hours[period];
 return {doctorKey:key,sourceType:site,event:{period,start:`${date}T${from}:00+11:00`,
  end:`${period==='Night'?'2026-10-12':date}T${to}:00+11:00`,title:period}};
}
for(const period of ['AM','PM','Night']) {
 state.onShiftData=[row(period)];body.scrollTop=0;callbacks=[];
 assert.equal(ui.facilityOverviewInitialShiftPeriod(),period);
 ui.scrollFacilityOverviewToInitialShift(1,0);callbacks.forEach(fn=>fn());
 assert.equal(body.scrollTop,{AM:0,PM:500,Night:1000}[period]);
}
state.onShiftData=[row('PM',{key:'MY ALIAS'})];
assert.equal(ui.facilityOverviewInitialShiftPeriod(),'PM','site roster aliases resolve the viewer');
state.onShiftData=[row('PM',{key:'OTHER'}),row('Night',{site:'ddh'}),row('AM',{date:'2026-10-10'})];
assert.equal(ui.facilityOverviewInitialShiftPeriod(),'','other doctors, sites and dates cannot choose the jump');
state.onShiftData=[row('AM'),row('PM')];now=new Date('2026-10-11T17:00:00+11:00');
assert.equal(ui.facilityOverviewInitialShiftPeriod(),'PM','active shift wins over a finished AM shift');
now=new Date('2026-10-12T03:00:00+11:00');state.onShiftData=[row('Night')];
assert.equal(ui.facilityOverviewInitialShiftPeriod(),'Night','overnight view keeps the starting roster date');
ui.currentNonClinical=true;assert.equal(ui.facilityOverviewInitialShiftPeriod(),'');ui.currentNonClinical=false;
preferred='';ui.currentRosterClaims=[];assert.equal(ui.facilityOverviewInitialShiftPeriod(),'','unlinked users do not jump to an arbitrary doctor');
preferred='ME';
for(const cancel of ['scroll','navigation','close']) {
 body.scrollTop=0;state.requestId=1;callbacks=[];open=true;
 ui.scrollFacilityOverviewToInitialShift(1,0);
 if(cancel==='scroll')body.scrollTop=123;
 if(cancel==='navigation')state.requestId=2;
 if(cancel==='close')open=false;
 callbacks.forEach(fn=>fn());
 assert.equal(body.scrollTop,cancel==='scroll'?123:0,`${cancel} supersedes the pending jump`);
}
assert.match(source,/if \(refreshed && !background\) scrollFacilityOverviewToInitialShift/,'background contact renewal never triggers another jump');
assert.match(source,/data-on-shift-period="\$\{period\}"/);
console.log('On shift initial scroll passed: AM/PM/Night, aliases, active/overnight shifts, unlinked users and manual/navigation precedence.');
