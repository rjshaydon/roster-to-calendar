import { normaliseVhhRosterExtract } from '../functions/_lib/vhh-roster.js';
import { australianTermStartForDate } from '../functions/_lib/d1-calendar.js';

// Explicit recovery only. Live extraction continues to respect visibility.
// Callers must independently verify that this interval is absent in the active
// source set before activating its bounded contribution.
export function extractVhhHistoryWindow(extract, { from, to, fileName } = {}) {
  const roster = normaliseVhhRosterExtract(extract);
  const valid = value => /^\d{4}-\d{2}-\d{2}$/.test(value) && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
  if (!roster || !valid(from) || !valid(to) || from > to || (Date.parse(to) - Date.parse(from)) / 86400000 > 92) throw new Error('A valid VHH historical window of at most one term is required.');
  if (australianTermStartForDate(from) !== australianTermStartForDate(to)) throw new Error('Historical recovery must be split at the medical term boundary.');
  if (!/\.json$/i.test(fileName || '')) throw new Error('Historical extracts require a distinct JSON source filename.');
  const blocks = roster.blocks.map(block => {
    const dates = block.dates.filter(item => from <= item.date && item.date <= to);
    const rows = block.rows.map(row => ({ ...row, assignments: row.assignments.filter(item => from <= item.date && item.date <= to) })).filter(row => row.assignments.length);
    return { ...block, visible: true, dates, rows };
  }).filter(block => block.dates.length && block.rows.length);
  const dates = new Set(blocks.flatMap(block => block.dates.map(item => item.date)));
  for (let stamp = Date.parse(from); stamp <= Date.parse(to); stamp += 86400000) {
    if (!dates.has(new Date(stamp).toISOString().slice(0, 10))) throw new Error('Historical workbook headers do not cover the complete requested window.');
  }
  // This local recovery is not a new SharePoint revision.
  return { ...roster, fileName, providerVersion: '', blocks };
}
