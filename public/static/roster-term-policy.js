// Eligibility is a Melbourne calendar date, independent of the term's weekday.
export function rosterTermAvailableFrom(termStart) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(termStart || ''))) return '';
  const [year, month] = termStart.split('-').map(Number);
  if (month < 1 || month > 12) return '';
  return `${month === 1 ? year - 1 : year}-${String(month === 1 ? 12 : month - 1).padStart(2, '0')}-01`;
}

export function rosterTermVisible(term, today) {
  const available = rosterTermAvailableFrom(term?.termStart);
  return Boolean(available && available <= today);
}

export function melbourneDateKey(now = new Date()) {
  const parts = new Intl.DateTimeFormat('en-AU', { timeZone: 'Australia/Melbourne', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(now);
  const value = type => parts.find(part => part.type === type)?.value;
  return `${value('year')}-${value('month')}-${value('day')}`;
}
