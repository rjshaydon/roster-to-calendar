import { normaliseVhhRosterExtract } from '../functions/_lib/vhh-roster.js';
import { australianTermStartForDate } from '../functions/_lib/d1-calendar.js';
const next=d=>new Date(Date.parse(d+'T12:00:00Z')+86400000).toISOString().slice(0,10);
// Only hidden, past headers are eligible. Hidden future drafts remain excluded.
export function planVhhHistoryWindows(extract,active,today=new Date().toISOString().slice(0,10)) {
  const roster=normaliseVhhRosterExtract(extract);
  if(!roster || active.length>32) throw Error('Invalid VHH historical inspection.');
  const visible=roster.blocks.filter(b=>b.visible).flatMap(b=>b.dates.map(d=>d.date)).sort();
  if(!visible.length) throw Error('VHH has no visible roster window.');
  const first=visible[0],last=visible.at(-1);
  const retiring=active.filter(f=>f.firstDate<=last&&f.lastDate>=first);
  const dates=[...new Set(roster.blocks.filter(b=>!b.visible).flatMap(b=>b.dates.map(d=>d.date)))]
    .filter(d=>d<first&&d<today&&(!active.some(f=>f.firstDate<=d&&f.lastDate>=d)||retiring.some(f=>f.firstDate<=d&&d<first))).sort();
  const windows=[];
  for(const date of dates) {
    const previous=windows.at(-1), overlap=retiring.find(f=>f.firstDate<=date&&f.lastDate>=date)?.id||'';
    if(previous&&next(previous.to)===date&&previous.term===australianTermStartForDate(date)&&previous.overlap===overlap) previous.to=date;
    else windows.push({from:date,to:date,term:australianTermStartForDate(date),overlap});
  }
  if(windows.length>3) throw Error('VHH historical recovery exceeds three bounded windows; retain the old roster and review.');
  return windows.map(({from,to,overlap})=>({from,to,overlap}));
}
