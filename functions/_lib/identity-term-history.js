import {loadCachedSnapshot} from './d1-calendar.js';
import {facilityMetadataManifestKey,FACILITY_PUBLICATION_LIMITS} from './facility-overview-cache.js';
import {rosterTermVisible} from '../../public/static/roster-term-policy.js';

// Inspect published staff lists only: no event SQL or source workbook retrieval.
export async function identityTermHistory(r2,aliases,today,after=0) {
 const sources=[...new Set(aliases.map(a=>a.source_type))].filter(s=>['mmc','mch','ddh','vhh','casey'].includes(s));
 const plans=[],missingSites=[];
 for(const sourceType of sources) {
  const manifest=await loadCachedSnapshot(r2,facilityMetadataManifestKey(sourceType));
  if(!manifest){missingSites.push(sourceType);continue;}
  if((manifest.terms||[]).length>64) throw Error('Roster term history exceeds the manifest inspection limit.');
  for(const term of manifest.terms||[]) if(term.staffKey&&rosterTermVisible(term,today)) plans.push({...term,sourceType});
 }
 plans.sort((a,b)=>b.termStart.localeCompare(a.termStart)||a.sourceType.localeCompare(b.sourceType));
 const offset=Number(after);
 if(!Number.isInteger(offset)||offset<0||offset>plans.length) throw Error('Invalid roster term history page.');
 const terms=[],missingTerms=[];
 for(const term of plans.slice(offset,offset+12)) {
  const staff=await loadCachedSnapshot(r2,term.staffKey);
  if(!staff){missingTerms.push({sourceType:term.sourceType,termStart:term.termStart});continue;}
  if((staff.members||[]).length>FACILITY_PUBLICATION_LIMITS.staffRows) throw Error('Roster term history exceeds the staff inspection limit.');
  const keys=new Set(aliases.filter(a=>a.source_type===term.sourceType).map(a=>a.doctor_key));
  const matches=(staff.members||[]).filter(m=>keys.has(m.doctorKey));
  if((staff.seniorityOverrides||[]).length>FACILITY_PUBLICATION_LIMITS.overrideRows) throw Error('Roster term history exceeds the grade override limit.');
  const overrides=new Map((staff.seniorityOverrides||[]).filter(o=>o.termStart===term.termStart&&!o.useRosterSeniority).map(o=>[o.doctorKey,o.seniority]));
  const grades=[...new Set(matches.map(m=>String(overrides.get(m.doctorKey)||m.seniority||'').trim()||'Not recorded'))];
  if(matches.length) terms.push({sourceType:term.sourceType,termStart:term.termStart,termEnd:term.termEnd,rosterNames:[...new Set(matches.map(m=>m.displayName||m.doctorKey))],grades,upcoming:term.termStart>today});
 }
 return {terms,missingSites,missingTerms,next:offset+12<plans.length?offset+12:null};
}
