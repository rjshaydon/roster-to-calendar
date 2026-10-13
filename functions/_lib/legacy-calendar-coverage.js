export const LEGACY_COVERAGE_VERSION='legacy-coverage-v1';
const validDate=date=>/^\d{4}-\d{2}-\d{2}$/.test(date)&&Number.isFinite(Date.parse(date+'T12:00:00Z'))&&new Date(date+'T12:00:00Z').toISOString().slice(0,10)===date;
export function validLegacySnapshotKey(key){return /^snapshots\/(user-account|creator-account|doctor-profile)\/.+\.json\.gz$/.test(key||'')&&key.length<2048;}
export async function legacyCoverageKey(key,revision){
 const digest=new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify([key,revision]))));
 return 'legacy-calendar/v1/coverage/'+[...digest].map(v=>v.toString(16).padStart(2,'0')).join('')+'.json';
}
export function summarizeLegacyCoverage(snapshot,key,revision){
 const bounds=new Map();let safe=true;
 for(const event of snapshot?.preview?.events||[]){
  const sourceType=String(event.source||'').toLowerCase();
  if(!['mmc','mch','ddh','vhh','casey'].includes(sourceType))continue;
  const date=String(event.start||'').slice(0,10);
  if(!validDate(date)){safe=false;continue;}
  const bound=bounds.get(sourceType)||{sourceType,startDate:date,endDate:date};
  if(date<bound.startDate)bound.startDate=date;if(date>bound.endDate)bound.endDate=date;bounds.set(sourceType,bound);
 }
 return {version:LEGACY_COVERAGE_VERSION,key,revision,safe,bounds:[...bounds.values()]};
}
// Conservative bounds can require an unnecessary history read, but cannot
// omit a historical shift: every date in the intersecting interval is covered.
export function legacyHistoryIsCovered(summary,key,revision,range,visibleTerms,scope=''){
 if(summary?.version!==LEGACY_COVERAGE_VERSION||summary.key!==key||!(summary.revision===revision||scope&&summary.scope===scope)||summary.safe!==true||!Array.isArray(summary.bounds)||summary.bounds.length>5)return false;
 const next=date=>{const d=new Date(date+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+1);return d.toISOString().slice(0,10);};
 for(const bound of summary.bounds){
  if(!['mmc','mch','ddh','vhh','casey'].includes(bound.sourceType)||!validDate(bound.startDate)||!validDate(bound.endDate)||bound.endDate<bound.startDate)return false;
  let cursor=bound.startDate>range.startDate?bound.startDate:range.startDate;
  const end=bound.endDate<range.endDate?bound.endDate:range.endDate;if(cursor>end)continue;
  const terms=visibleTerms.filter(t=>t.sourceType===bound.sourceType).sort((a,b)=>a.termStart<b.termStart?-1:a.termStart>b.termStart?1:0);
  for(const term of terms){if(term.termEnd<cursor)continue;if(term.termStart>cursor)break;if(term.termEnd>=end){cursor='covered';break;}cursor=next(term.termEnd);}
  if(cursor!=='covered')return false;
 }
 return true;
}
