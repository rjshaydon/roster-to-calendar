const VERSION = 1;
const MAX_ENTRIES = 60;
const MAX_BYTES = 5 * 1024 * 1024;
const stable = value => JSON.stringify(value, (_, item) => item && typeof item === 'object' && !Array.isArray(item) ? Object.fromEntries(Object.entries(item).sort(([a], [b]) => a.localeCompare(b))) : item);
const canonical = query => ({ ...query, sourceTypes: [...new Set(query.sourceTypes || [])].map(s => s.toLowerCase()).sort(), doctorKeys: [...new Set(query.doctorKeys || [])].sort(), ...(query.doctorIdentities ? { doctorIdentities: [...new Map(query.doctorIdentities.map(identity => [identity.sourceType + '|' + identity.key, identity])).values()].sort((a, b) => (a.sourceType + '|' + a.key).localeCompare(b.sourceType + '|' + b.key)) } : {}) });
const narrowData = (data, sources) => {
  if (!sources.length) return data;
  const includes = row => sources.includes(String(row.sourceType || '').toLowerCase());
  const missing = data.missing?.filter(includes);
  return { ...data, ...(data.events ? { events:data.events.filter(includes) } : {}), ...(data.doctors ? { doctors:data.doctors.filter(includes) } : {}), ...(missing ? { missing, incomplete:!!missing.length } : {}) };
};

// Permission leases are memory-only. A persisted result never authenticates a
// new session. All network requests still authenticate on the server.
export function createRosterComparisonCache({ storage = comparisonStorage(), now = Date.now, freshness = 60_000 } = {}) {
  let generation = 0, invalidatedAt = 0;
  const memory = new Map(), people = new Map(), pending = new Map(), leases = new Map();
  const keyFor = (owner, kind, query) => stable([VERSION, owner, kind, canonical(query)]);
  const valid = record => record && record.expiresAt > now();
  const leaseKey = (owner, feature, query) => keyFor(owner, feature, { startDate: query.startDate, endDate: query.endDate, sourceTypes: query.sourceTypes });
  const remember = (key, record) => { memory.delete(key); memory.set(key, record); while (memory.size > MAX_ENTRIES) memory.delete(memory.keys().next().value); };
  return {
    clear() { generation++; memory.clear(); people.clear(); leases.clear(); pending.clear(); },
    invalidate() { generation++; invalidatedAt = now(); pending.clear(); for (const entry of [...memory.values(), ...people.values()]) entry.at = 0; leases.clear(); },
    async request({ owner, feature, kind, query, send, onCached = () => {}, force = false, leaseQuery = query }) {
      const epoch = generation;
      const normalized = canonical(query), leaseId = leaseKey(owner, feature, leaseQuery);
      let lease = leases.get(leaseId);
      if (!lease && leaseQuery.sourceTypes?.length) {
        lease = [...leases.values()].find(value => value.owner===owner && value.feature===feature && value.query.startDate===leaseQuery.startDate && value.query.endDate===leaseQuery.endDate && !value.query.sourceTypes.length
          && valid(value) && now()-value.at < freshness && leaseQuery.sourceTypes.every(source=>value.sources.includes(source.toLowerCase())));
        if(lease)leases.set(leaseId,{...lease,query:canonical(leaseQuery)});
      }
      const scope = kind === 'selected' ? lease?.scopeRevision : '';
      const key = keyFor(owner, kind, { ...normalized, ...(kind === 'selected' ? { scopeRevision: scope || '' } : {}) });
      let record = memory.get(key);
      if (!record && scope) record = await storage.get(key).catch(() => null);
      if (record) record = { ...record, data:narrowData(record.data, normalized.sourceTypes) };
      if (epoch !== generation) throw Object.assign(Error('Comparison context changed.'), { name: 'AbortError' });
      const authorized = valid(lease) && now() - lease.at < freshness;
      if (!record && kind==='discover' && authorized && normalized.sourceTypes.length) {
        const broader=[...memory.values()].find(value=>value.owner===owner && value.kind===kind && !value.query.sourceTypes.length && valid(value) && value.at>=invalidatedAt && now()-value.at<freshness
          && stable({...value.query,sourceTypes:[]})===stable({...normalized,sourceTypes:[]}));
        if(broader){record={...broader,query:normalized,data:narrowData(broader.data,normalized.sourceTypes)};remember(key,record);}
      }
      const personKey = doctorKey => keyFor(owner, 'person', { ...normalized, doctorKeys: [doctorKey], scopeRevision: scope || '',
        ...(normalized.doctorIdentities ? {doctorIdentities:normalized.doctorIdentities.filter(identity=>identity.key===doctorKey)} : {}) });
      if (kind === 'selected' && authorized && !force && !record) {
        const facts = await Promise.all(normalized.doctorKeys.map(async doctorKey => {
          const key = personKey(doctorKey); let value = people.get(key) || await storage.get(key).catch(()=>null);
          if(!value && normalized.sourceTypes.length){
            value=[...people.values()].find(fact=>fact.owner===owner && fact.personKey===doctorKey && fact.scopeRevision===scope && fact.query.startDate===normalized.startDate && fact.query.endDate===normalized.endDate && !fact.query.sourceTypes.length
              && (!normalized.doctorIdentities || normalized.doctorIdentities.filter(identity=>identity.key===doctorKey&&normalized.sourceTypes.includes(identity.sourceType)).every(identity=>fact.query.doctorIdentities?.some(item=>item.key===identity.key&&item.sourceType===identity.sourceType))));
            if(value){const data=narrowData(value.data,normalized.sourceTypes);value={...value,complete:!data.incomplete,data};}
          }
          return valid(value) && now()-value.at < freshness && value.at >= invalidatedAt ? value : null;
        }));
        if (epoch !== generation) throw Object.assign(Error('Comparison context changed.'), {name:'AbortError'});
        if(facts.some(Boolean)&&!facts.every(Boolean)){
          const known=facts.filter(Boolean);
          record={partial:true,at:Math.min(...known.map(fact=>fact.at)),expiresAt:Math.min(...known.map(fact=>fact.expiresAt)),data:{ok:true,events:known.flatMap(fact=>fact.data.events),revision:'',personRevisions:Object.assign({},...known.map(fact=>fact.data.personRevisions))}};
        }
        if (facts.length && facts.every(Boolean)) {
          const data = { ok:true, events:facts.flatMap(fact=>fact.data.events), revision:facts.map(fact=>fact.data.revision).join(':'), personRevisions:Object.assign({},...facts.map(fact=>fact.data.personRevisions)), scopeRevision:scope,
            accessExpiresAt:new Date(Math.min(...facts.map(fact=>fact.expiresAt))).toISOString(), missing:[...new Map(facts.flatMap(fact=>fact.data.missing||[]).map(missing=>[stable(missing),missing])).values()], incomplete:facts.some(fact=>fact.complete===false) };
          record={at:Math.min(...facts.map(fact=>fact.at)),expiresAt:Math.min(...facts.map(fact=>fact.expiresAt)),data};
          remember(key,record);
          return data;
        }
      }
      if (valid(record) && authorized && !record.partial) {
        remember(key, record);
        if (!force && now() - record.at < freshness && record.at >= invalidatedAt) return record.data;
        onCached(record.data);
      }
      // Different capabilities cannot share an authorization request. Selected
      // facts become reusable after each feature verifies its own period lease.
      const requestKey = kind === 'selected' && authorized ? key : key + '|' + feature;
      if (pending.has(requestKey)) return pending.get(requestKey);
      const work = (async () => {
        try {
          const response = await send(normalized, valid(record) ? record.data.revision || '' : '', valid(record) ? record.data.personRevisions || {} : {});
          if (!response?.ok || response.unavailable) throw Object.assign(Error(response?.error || 'Roster comparison is temporarily unavailable.'), { status: response?.status });
          if (epoch !== generation) throw Object.assign(Error('Comparison context changed.'), { name: 'AbortError' });
          const expiresAt = Date.parse(response.accessExpiresAt || '');
          if (kind !== 'selected' && response.scopeRevision && expiresAt > now()) leases.set(leaseId, { owner, feature, query:canonical(leaseQuery), sources:response.sourceTypes||normalized.sourceTypes, at: now(), expiresAt, scopeRevision: response.scopeRevision });
          let data = response.unchanged && valid(record) ? { ...record.data, ...response, unchanged: false } : response;
          if (response.unchanged && !valid(record)) throw Error('The saved comparison expired. Please reopen this search.');
          if(response.delta){
            if(!valid(record))throw Error('The saved comparison expired. Please reopen this search.');
            const updated=new Set(response.updatedDoctorKeys||[]);
            data={...response,delta:false,events:[...(record.data.events||[]).filter(row=>!updated.has(row.doctorKey)),...(response.events||[])]};
          }
          // A selected request requires the same already-verified scope. A
          // changed scope is accepted from the server but stored under its new key.
          const savedKey = keyFor(owner, kind, { ...normalized, ...(kind === 'selected' ? { scopeRevision: scope || response.scopeRevision || '' } : {}) });
          if (expiresAt > now()) {
            const entry = { complete:!data.missing?.length, owner, kind, query:normalized, scopeRevision:scope||response.scopeRevision, at: now(), expiresAt, data };
            if (kind === 'selected') for (const doctorKey of normalized.doctorKeys) {
              const key = personKey(doctorKey), fact = { ...entry, kind:'person',personKey:doctorKey,query:{...normalized,doctorKeys:[doctorKey],...(normalized.doctorIdentities?{doctorIdentities:normalized.doctorIdentities.filter(identity=>identity.key===doctorKey)}:{})}, data: { ok:true, events:(data.events || []).filter(row=>row.doctorKey===doctorKey), revision:data.revision, personRevisions:{[doctorKey]:data.personRevisions?.[doctorKey]}, scopeRevision:scope, accessExpiresAt:data.accessExpiresAt, missing:data.missing||[], incomplete:!!data.missing?.length } };
              people.delete(key);people.set(key,fact);while(people.size>MAX_ENTRIES)people.delete(people.keys().next().value);
              if (!response.unchanged) void storage.put(key,fact).catch(()=>{});
            }
            remember(savedKey, entry);
            if (!response.unchanged) void storage.put(savedKey, entry).catch(() => {});
          }
          return data;
        } catch (error) {
          if (error.status === 401 || error.status === 403) { generation++; memory.clear(); people.clear(); leases.clear(); }
          else if (epoch === generation && authorized && valid(record) && !record.partial && error.name !== 'AbortError') return { ...record.data, refreshError: error.message };
          throw error;
        }
      })();
      pending.set(requestKey, work);
      try { return await work; } finally { if (pending.get(requestKey) === work) pending.delete(requestKey); }
    },
  };
}

function comparisonStorage() {
  let opening;
  const database = () => {
    if (typeof indexedDB === 'undefined') return Promise.resolve(null);
    return opening ||= new Promise((resolve, reject) => {
      const request = indexedDB.open('roster-comparisons', VERSION);
      request.onupgradeneeded = () => request.result.createObjectStore('facts', { keyPath: 'key' });
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => { opening = null; reject(request.error); };
      request.onblocked = () => { opening = null; resolve(null); };
    });
  };
  const result = request => new Promise((resolve, reject) => { request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); });
  return {
    async get(key) { const db = await database(); return db ? (await result(db.transaction('facts').objectStore('facts').get(key)))?.record : null; },
    async put(key, record) {
      const bytes = JSON.stringify(record).length * 2;
      if (bytes > MAX_BYTES) return;
      const db = await database(); if (!db) return;
      // Prune in the same transaction as the write. No unbounded disk history.
      await new Promise((resolve, reject) => {
        const transaction = db.transaction('facts', 'readwrite'), store = transaction.objectStore('facts');
        const request = store.getAll();
        request.onsuccess = () => {
          const entries = request.result.filter(entry => entry.key !== key).sort((a, b) => b.record.at - a.record.at);
          let total = bytes, count = 1;
          for (const entry of entries) { total += entry.bytes || 0; count++; if (count > MAX_ENTRIES || total > MAX_BYTES || entry.record.expiresAt <= Date.now()) store.delete(entry.key); }
          store.put({ key, record, bytes });
        };
        transaction.oncomplete = resolve; transaction.onerror = () => reject(transaction.error); transaction.onabort = () => reject(transaction.error);
      });
    },
  };
}
