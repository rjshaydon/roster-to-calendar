import assert from 'node:assert/strict';
import {serverMelbourneDateKey} from '../functions/_lib/melbourne-clock.js';
import {melbourneClockData as clock} from '../functions/_lib/melbourne-clock-data.js';
const fmt=new Intl.DateTimeFormat('en-CA',{timeZone:'Australia/Melbourne',year:'numeric',month:'2-digit',day:'2-digit'});
function expected(ms){const parts=fmt.formatToParts(new Date(ms)),get=t=>parts.find(p=>p.type===t)?.value;return `${get('year')}-${get('month')}-${get('day')}`;}
const Original=Intl.DateTimeFormat;let constructors=0;Intl.DateTimeFormat=class extends Original{constructor(...args){super(...args);constructors++;}};
let checks=0;
try{
 for(let ms=clock.start;ms<clock.end;ms+=6*3600000){assert.equal(serverMelbourneDateKey(new Date(ms)),expected(ms));checks++;}
 for(const [boundary]of clock.transitions)for(const delta of [-86400001,-1,0,1,86399999]){const ms=boundary+delta;if(ms>=clock.start&&ms<clock.end){assert.equal(serverMelbourneDateKey(new Date(ms)),expected(ms));checks++;}}
 // Every local midnight around all DST changes, including the 2006 Games exception.
 for(const[boundary]of clock.transitions)for(let delta=-48*3600000;delta<=48*3600000;delta+=15*60000){const ms=boundary+delta;if(ms>=clock.start&&ms<clock.end){assert.equal(serverMelbourneDateKey(new Date(ms)),expected(ms));checks++;}}
 assert.equal(constructors,0,'current and covered historical dates never initialize Worker ICU');
 for(const ms of [Date.UTC(1990,0,1),clock.end,Date.UTC(2050,6,1)])assert.equal(serverMelbourneDateKey(new Date(ms)),expected(ms));
 assert.equal(constructors,1,'out-of-range dates retain a cached ICU fallback');
 assert.throws(()=>serverMelbourneDateKey(new Date(NaN)),RangeError);
}finally{Intl.DateTimeFormat=Original;}
console.log('Melbourne clock matches ICU in '+checks+' cases across 2000–2040 and DST/midnight boundaries, with zero covered-range ICU constructions and correct out-of-range fallback.');
