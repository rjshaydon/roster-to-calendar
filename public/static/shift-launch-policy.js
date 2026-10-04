const HOUR_MS = 60 * 60 * 1000;
const formatter = new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Australia/Melbourne', year: 'numeric', month: '2-digit', day: '2-digit',
  hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
});
function wallParts(epoch) {
  return Object.fromEntries(formatter.formatToParts(new Date(epoch)).map(part => [part.type, part.value]));
}
function shiftEpoch(value) {
  const text = String(value || '');
  if (/T.*(?:Z|[+-]\d{2}:?\d{2})$/i.test(text)) return Date.parse(text);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return NaN;
  const [,y,m,d,h,min,sec='00'] = match;
  const wall = Date.UTC(+y,+m-1,+d,+h,+min,+sec);
  let epoch = wall;
  for (let attempt=0;attempt<4;attempt++) {
    const p=wallParts(epoch);
    const delta=wall-Date.UTC(+p.year,+p.month-1,+p.day,+p.hour,+p.minute,+p.second);
    if (!delta) return epoch;
    epoch+=delta;
  }
  return NaN; // A nonexistent local time during the daylight-saving jump.
}
export function onShiftLaunchWindow(events = [], now = new Date()) {
  const instant = now.getTime();
  const matches = events.flatMap(event => {
    if (event?.allDay || ['unknown','custom','leave'].includes(String(event?.kind || '').toLowerCase())
      || String(event?.status || '').toLowerCase()==='unknown'
      || /\b(?:leave|conference|cme|annual|sick|personal|study|exam|sabbatical|parental|phnw|public holiday)\b/i.test(`${event?.title || ''} ${event?.rawValue || ''}`)) return [];
    const facilityKey=String(event?.sourceType || event?.source || event?.facilityKey || '').toUpperCase();
    if (!['MMC','DDH','MCH','VHH','CASEY'].includes(facilityKey)) return [];
    const start=shiftEpoch(event.start),end=shiftEpoch(event.end);
    if (!Number.isFinite(start) || !Number.isFinite(end) || end<=start || instant<start-HOUR_MS || instant>end+HOUR_MS) return [];
    const p=wallParts(start);
    return [{facilityKey,rosterDate:`${p.year}-${p.month}-${p.day}`,start,end,active:start<=instant && instant<end}];
  });
  matches.sort((a,b)=>Number(b.active)-Number(a.active) || Math.abs(instant-a.start)-Math.abs(instant-b.start));
  return matches[0] || null;
}
