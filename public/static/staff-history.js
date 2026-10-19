// Shared by the offline catalogue builder and the browser. No roster events,
// account details or current-grade inference belong in this projection.
export const STAFF_HISTORY_VERSION = 1;
export const STAFF_HISTORY_PAGE = 128;
export const STAFF_HISTORY_SHARD = 16;
export function staffHistoryName(value) {
  return String(value || '').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase('en').replace(/[^\p{L}\p{N}]+/gu, ' ').trim().replace(/\s+/g, ' ');
}
export function matchesStaffHistory(person, query, words=staffHistoryName(query).split(' ').filter(Boolean)) {
  return words.length > 0 && (person.searchNames || (person.names || [person.displayName]).map(staffHistoryName)).some(normal => {
    return words.every(word => normal.includes(word));
  });
}
export function coalesceStaffHistory(rows) {
  const groups = new Map();
  for (const row of rows) {
    const key = JSON.stringify([row.sourceType, row.grade, row.upcoming]);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const result = [];
  for (const group of groups.values()) {
    for (const row of group.sort((a,b) => a.termStart.localeCompare(b.termStart))) {
      const previous = result.at(-1);
      const adjacent = previous && new Date(Date.parse(previous.termEnd+'T00:00:00Z')+86400000).toISOString().slice(0,10) >= row.termStart;
      if (previous && previous.sourceType===row.sourceType && previous.grade===row.grade && previous.upcoming===row.upcoming && adjacent) {
        previous.termEnd = previous.termEnd > row.termEnd ? previous.termEnd : row.termEnd;
        previous.terms = [...new Set([...previous.terms, row.termStart])];
        previous.rosterNames = [...new Set([...previous.rosterNames, ...row.rosterNames])];
      } else result.push({...row, terms:[row.termStart], rosterNames:[...row.rosterNames]});
    }
  }
  return result.sort((a,b) => a.sourceType.localeCompare(b.sourceType)||a.termStart.localeCompare(b.termStart)||a.grade.localeCompare(b.grade));
}
export function buildStaffHistoryCatalogue(fragments, aliases, today) {
  const approved = new Map();
  for (const alias of aliases) if(alias.personId) approved.set(JSON.stringify([alias.sourceType,alias.doctorKey]), alias);
  const people = new Map();
  for (const fragment of fragments) for (const member of fragment.members || []) {
    if (!member.doctorKey || member.contributionCount<=0 || member.membershipSource==='sms-continuity' || member.membershipSource==='provider-only') continue;
    const aliasKey=JSON.stringify([fragment.sourceType,member.doctorKey]), alias=approved.get(aliasKey);
    const id=alias?.personId ? 'approved:'+alias.personId : 'roster:'+aliasKey;
    if(!people.has(id))people.set(id,{id,displayName:alias?.preferredName||member.displayName||member.doctorKey,names:new Set(),history:[]});
    const person=people.get(id);
    person.names.add(member.displayName||member.doctorKey);
    person.names.add(alias?.displayName||member.displayName||member.doctorKey);
    person.names.add(person.displayName);
    person.history.push({sourceType:fragment.sourceType,termStart:fragment.termStart,termEnd:fragment.termEnd,grade:member.grade||'Not recorded',upcoming:fragment.termStart>today,rosterNames:[member.displayName||member.doctorKey]});
  }
  // Approved aliases without a shift in this particular spelling still search
  // the person's real history; names alone never establish identity membership.
  for(const alias of aliases){const person=people.get('approved:'+alias.personId);if(person){person.names.add(alias.displayName);person.names.add(alias.preferredName);}}
  return [...people.values()].sort((a,b)=>a.displayName.localeCompare(b.displayName)||a.id.localeCompare(b.id)).map((p,index)=>({slot:index,displayName:p.displayName,names:[...p.names].filter(Boolean).sort(),history:p.history.sort((a,b)=>a.sourceType.localeCompare(b.sourceType)||a.termStart.localeCompare(b.termStart)||a.grade.localeCompare(b.grade))}));
}

// Per-account in-memory cache only. Logout/access changes destroy the instance.
// Catalogue fetches coalesce; typing filters locally and only missing matching
// histories are fetched. No source polls or persistent browser roster cache.
export function createStaffHistoryClient(request) {
  let generation=0, catalogue=[], revision='', histories=new Map(), pending=null, currentMissing=[];
  return {
    clear(){generation++;catalogue=[];revision='';histories.clear();pending=null;currentMissing=[];},
    get revision(){return revision;},
    matches(query){const words=staffHistoryName(query).split(' ').filter(Boolean);return catalogue.filter(p=>matchesStaffHistory(p,query,words));},
    history(slot){return histories.get(slot);},
    async load(){
      if(pending)return pending;
      const own=generation;
      pending=(async()=>{
        let page=0,people=[],missing=[],expectedRevision='';
        do{
          const data=await request({mode:'names',page,cachedRevision:revision});
          if(own!==generation)return false;
          if(data.preparing||data.stale)throw Error('Staff history is being updated. Try again shortly.');
          if(data.unchanged)return {missing:currentMissing};
          if(page&&data.revision!==expectedRevision)throw Error('Staff history changed while loading. Try again.');
          if(!page)expectedRevision=data.revision;
          people.push(...data.people.map(p=>({...p,searchNames:p.names.map(staffHistoryName)})));missing=data.missing||[];page=data.next;
          if(people.length>4096)throw Error('Staff history catalogue exceeds its limit.');
        }while(page!==null);
        if(revision!==expectedRevision)histories.clear();revision=expectedRevision;catalogue=people;currentMissing=missing;return {missing};
      })().finally(()=>{if(own===generation)pending=null;});
      return pending;
    },
    async details(people){
      const own=generation,expected=revision;
      const slots=people.map(p=>p.slot).filter(slot=>!histories.has(slot)).slice(0,12);
      if(!slots.length)return;
      const data=await request({mode:'histories',revision,slots});
      if(own!==generation)return;
      if(data.stale||data.preparing||data.revision!==expected){this.clear();throw Object.assign(Error('Staff history is being updated. Try again shortly.'),{code:'STAFF_HISTORY_STALE'});}
      for(const p of data.people)histories.set(p.slot,p.history);
    }
  };
}
