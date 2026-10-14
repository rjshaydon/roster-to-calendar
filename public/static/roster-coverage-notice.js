// Present existing coverage metadata; never fetch data or infer complete
// coverage merely because an individual has some retained shifts.
const dateKey = value => /^\d{4}-\d{2}-\d{2}$/.test(value || '') ? value : '';
export function coverageNoticeRanges(missing = [], {startDate = '', endDate = '', sourceType = '', upcoming = null} = {}) {
  const current = [], nextSources = new Set();
  for (const item of missing) {
    const source = String(item.sourceType || '').toLowerCase();
    if (sourceType && source !== sourceType.toLowerCase()) continue;
    let from = dateKey(item.startDate), to = dateKey(item.endDate);
    if (item.month && /^\d{4}-\d{2}$/.test(item.month)) {
      from = item.month + '-01';
      to = new Date(Date.UTC(Number(item.month.slice(0,4)), Number(item.month.slice(5)), 0)).toISOString().slice(0,10);
    }
    to ||= from;
    if (upcoming && from && from >= upcoming.startDate && from <= upcoming.endDate && to >= upcoming.startDate) {
      nextSources.add(source); continue;
    }
    if (from && endDate && from > endDate || to && startDate && to < startDate) continue;
    if (from && startDate && from < startDate) from = startDate;
    if (to && endDate && to > endDate) to = endDate;
    current.push({sourceType:source,startDate:from,endDate:to});
  }
  const grouped = [];
  for (const item of current.sort((a,b)=>a.sourceType.localeCompare(b.sourceType)||a.startDate.localeCompare(b.startDate))) {
    const last = grouped.at(-1);
    const next = last?.endDate ? new Date(last.endDate + 'T12:00:00Z') : null;
    if (next) next.setUTCDate(next.getUTCDate()+1);
    if (last && last.sourceType===item.sourceType && last.startDate && item.startDate && next && item.startDate<=next.toISOString().slice(0,10)) {
      if (item.endDate>last.endDate) last.endDate=item.endDate;
    } else if (!last || JSON.stringify(last)!==JSON.stringify(item)) grouped.push({...item});
  }
  return {current:grouped,nextSources:[...nextSources].filter(Boolean).sort()};
}
export function routineCalendarStatus(message) {
  return /^Calendar loaded(?: from cache)?\.$/.test(message) || message==='Calendar loaded. Checking for updates...';
}
