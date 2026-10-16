import { melbourneClockData as clock } from './melbourne-clock-data.js';
let fallback;
export function serverMelbourneDateKey(now = new Date()) {
  const ms=now.getTime();
  if(ms>=clock.start&&ms<clock.end){
    let lo=0,hi=clock.transitions.length-1;
    while(lo<hi){const mid=Math.ceil((lo+hi)/2);if(clock.transitions[mid][0]<=ms)lo=mid;else hi=mid-1;}
    return new Date(ms+clock.transitions[lo][1]*60000).toISOString().slice(0,10);
  }
  // Preserve dates outside the generated range; ordinary current-day requests
  // never initialize ICU on the Worker. The browser keeps its own Intl path.
  fallback ||= new Intl.DateTimeFormat('en-CA',{timeZone:'Australia/Melbourne',year:'numeric',month:'2-digit',day:'2-digit'});
  const parts=fallback.formatToParts(now),get=type=>parts.find(p=>p.type===type)?.value;
  return `${get('year')}-${get('month')}-${get('day')}`;
}
