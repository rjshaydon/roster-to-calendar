import {authorityGuard} from './roster-authority.js';
import {australianTermStartForDate} from './d1-calendar.js';
import { rosterTermOwnershipForFiles } from './roster-term-ownership.js';
import { facilitySmsMembershipStatement } from './d1-calendar.js';

export const VHH_HISTORY_SOURCE = 'vhh-active-medical-roster';
export const historyJobId = (parent, from, to) => `${parent}:history:${from}:${to}`;
export const historyFileName = (from, to) => `VHH-history-${from}-to-${to}.json`;
export function historyWindow(from, to) {
  if (![from,to].every(v=>/^\d{4}-\d{2}-\d{2}$/.test(v) && new Date(v+'T00:00:00Z').toISOString().slice(0,10)===v)
    || from>to || (Date.parse(to)-Date.parse(from))/86400000>92) throw Error('Invalid bounded historical window.');
  return {from,to};
}
export async function activeVhhOwnership(db) {
  const files=(await db.prepare("SELECT id,parsed_at FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type='vhh' AND active=1 ORDER BY id LIMIT 33").all()).results;
  const ownership=await rosterTermOwnershipForFiles(db,files.map(f=>f.id));
  return files.map(f=>({...f,...ownership.get(f.id)}));
}

// Archives are prepared invisibly by the external processor. Their promotion
// and the live file's removal of those dates must share one D1 transaction.
export async function prepareVhhHistoryPromotion(db,parentId,windows,excludedFiles=[]) {
  if (!Array.isArray(windows) || windows.length>3) throw Error('Historical promotion exceeds its window bound.');
  const active=await activeVhhOwnership(db), jobs=[];
  for (const window of windows) {
    const {from,to}=historyWindow(window.from,window.to);
    const id=historyJobId(parentId,from,to);
    const job=await db.prepare('SELECT * FROM roster_import_jobs WHERE run_id=?').bind(id).first();
    if(!job || !job.compact_ready || job.activated || job.source_id!==VHH_HISTORY_SOURCE) throw Error('Historical preservation is not fully prepared.');
    const manifest=JSON.parse(job.manifest_json);
    if(manifest.startDate!==from || manifest.rosterEndDate!==to) throw Error('Historical preservation dates changed.');
    if(active.some(f=>!excludedFiles.includes(f.id)&&f.firstDate<=to&&f.lastDate>=from)) throw Error('Historical preservation overlaps another active input.');
    if(jobs.some(j=>j.from<=to&&j.to>=from)) throw Error('Historical preservation windows overlap.');
    jobs.push({...job,from,to,manifest});
  }
  if(active.length+jobs.length>32) throw Error('Historical preservation exceeds the active file bound.');
  const pin=JSON.stringify(active.map(({id,parsed_at})=>({id,parsed_at})));
  const guards=[db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type='vhh' AND active=1)=json_array_length(?)
    AND NOT EXISTS(SELECT 1 FROM roster_files f INDEXED BY idx_roster_files_source_active WHERE f.source_type='vhh' AND f.active=1 AND NOT EXISTS(SELECT 1 FROM json_each(?) p WHERE json_extract(p.value,'$.id')=f.id AND json_extract(p.value,'$.parsed_at')=f.parsed_at))
    THEN 1 ELSE json('vhh-history-baseline-conflict') END`).bind(pin,pin)];
  const statements=[];
  for(const job of jobs) {
    guards.push(authorityGuard(db,job.run_id,'vhh',[australianTermStartForDate(job.from)]));
    guards.push(db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_import_jobs j JOIN roster_files f ON f.id=j.file_id WHERE j.run_id=? AND j.plan_revision=? AND j.compact_ready=1 AND j.activated=0 AND f.active=0)
      THEN 1 ELSE json('vhh-history-not-ready') END`).bind(job.run_id,job.plan_revision));
    statements.push(db.prepare('UPDATE roster_files SET active=1 WHERE id=? AND active=0').bind(job.file_id),
      db.prepare("UPDATE roster_file_status_summaries SET active=1,derived_state='ready',status_revision=?,updated_at=? WHERE file_id=?").bind(crypto.randomUUID(),new Date().toISOString(),job.file_id),
      db.prepare('UPDATE roster_import_jobs SET activated=1 WHERE run_id=? AND plan_revision=?').bind(job.run_id,job.plan_revision),
      facilitySmsMembershipStatement(db,job.file_id,true));
  }
  return {guards,statements,jobs,writes:jobs.reduce((n,j)=>n+j.manifest.doctors.length*4+64,0)};
}
export function requirePreservedPrefix(oldStart,newStart,windows) {
  if(!oldStart || !newStart || newStart<=oldStart) return;
  let cursor=oldStart;
  for(const {from,to} of [...windows].sort((a,b)=>a.from.localeCompare(b.from))) {
    if(from>cursor) break;
    if(to>=cursor) cursor=new Date(Date.parse(to+'T12:00:00Z')+86400000).toISOString().slice(0,10);
  }
  if(cursor<newStart) throw Error('VHH rollover would remove historical shifts without a prepared archive.');
}
