// Explicit, bounded identity operations. Raw roster rows and grades are never
// mutated here. The global revision is a cheap transactional CAS, not a scan.
const MAX_PEOPLE = 9, MAX_ALIASES = 32, MAX_ACCOUNTS = 16, MAX_REDIRECTS = 64;
const SOURCES = new Set(['mmc','mch','ddh','vhh','casey']);
const specs = {
 people: { table:'roster_people', keys:['person_id'] },
 aliases: { table:'roster_person_aliases', keys:['source_type','doctor_key'] },
 accounts: { table:'account_people', keys:['email'] },
 redirects: { table:'roster_person_redirects', keys:['old_person_id'] },
};
const canonical = value => JSON.stringify(value, Object.keys(value || {}).sort());
const marker = (type,row) => specs[type].keys.map(key => row[key]).join('|');
const copy = value => JSON.parse(JSON.stringify(value));
function error(message,code='IDENTITY_CONFLICT') { const e=new Error(message); e.code=code; return e; }
function bounded(rows,limit) { if(rows.length>limit) throw error('Identity scope is too large; narrow the selection.','IDENTITY_SCOPE'); return rows; }
async function rows(db,sql,args=[]) { return (await db.prepare(sql).bind(...args).all()).results || []; }
export function harmlessIdentityKey(value) {
 return String(value||'').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase()
  .replace(/^(?:DR|DOCTOR|MR|MRS|MS|MISS|PROF|PROFESSOR|A PROF|ASSOC PROF)[\s.]+/,'').replace(/[^A-Z0-9]/g,'');
}
export function validatePersonId(value) {
 if(!/^person:[a-z0-9]+(?:-[a-z0-9]+)*$/.test(value) || value.length>160) throw error('Use a person ID such as person:surname-given-name.','IDENTITY_INPUT');
 return value;
}
export async function resolvePersonId(db,id) {
 const redirect=await db.prepare('SELECT person_id FROM roster_person_redirects WHERE old_person_id=? AND active=1').bind(id).first();
 const resolved=redirect?.person_id || id;
 const person=await db.prepare('SELECT * FROM roster_people WHERE person_id=?').bind(resolved).first();
 if(!person || person.status!=='active') throw error('Active person not found.','IDENTITY_NOT_FOUND');
 return person;
}
async function scope(db,ids) {
 ids=[...new Set(ids)].sort(); if(!ids.length || ids.length>MAX_PEOPLE) throw error('Select a smaller identity scope.','IDENTITY_SCOPE');
 const placeholders=ids.map(()=>'?').join(',');
 const result={
  people:await rows(db,`SELECT * FROM roster_people WHERE person_id IN (${placeholders}) ORDER BY person_id`,ids),
  aliases:bounded(await rows(db,`SELECT * FROM roster_person_aliases WHERE person_id IN (${placeholders}) ORDER BY source_type,doctor_key LIMIT 33`,ids),MAX_ALIASES),
  accounts:bounded(await rows(db,`SELECT * FROM account_people WHERE person_id IN (${placeholders}) ORDER BY email LIMIT 17`,ids),MAX_ACCOUNTS),
  redirects:bounded(await rows(db,`SELECT * FROM roster_person_redirects WHERE person_id IN (${placeholders}) OR old_person_id IN (${placeholders}) ORDER BY old_person_id LIMIT 65`,[...ids,...ids]),MAX_REDIRECTS),
 };
 if(result.people.length!==ids.length) throw error('Person not found.','IDENTITY_NOT_FOUND');
 // Existing claimed source aliases must participate in the preview even when
 // older account_people links were never seeded. Never silently seize a claim.
 if(result.aliases.length) {
  const predicates=result.aliases.map(()=>'(source_type=? AND doctor_key=?)').join(' OR ');
  result.claimOwners=bounded(await rows(db,`SELECT email,source_type,doctor_key FROM account_claims INDEXED BY idx_account_claims_source_doctor_email WHERE ${predicates} ORDER BY email,source_type,doctor_key LIMIT 65`,result.aliases.flatMap(a=>[a.source_type,a.doctor_key])),64);
 } else result.claimOwners=[];
 return result;
}
async function token(value) { const bytes=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value))); return [...new Uint8Array(bytes)].map(b=>b.toString(16).padStart(2,'0')).join(''); }
export async function previewIdentityOperation(db,input) {
 const kind=input.kind;
 if(!['merge','name','id','alias-move','account-link'].includes(kind)) throw error('Unsupported identity operation.','IDENTITY_INPUT');
 const ids=[...new Set(input.personIds||[])];
 if(!ids.length || ids.length>8) throw error('Select one to eight people.','IDENTITY_SCOPE');
 const before=await scope(db,ids);
 if(before.people.some(p=>p.status!=='active')) throw error('Resolve retired identities before editing.');
 const after=copy(before), now='PREVIEW';
 let target=String(input.targetId||ids[0]);
 const name=String(input.preferredName||'').trim();
 if(['merge','name','id'].includes(kind) && (!name || name.length>200)) throw error('Provide the preferred full name.','IDENTITY_INPUT');
 let targetRow=after.people.find(p=>p.person_id===target);
 if(kind==='id' || (['merge','alias-move'].includes(kind) && !targetRow)) {
  if(!name || name.length>200) throw error('Provide the preferred full name for the new person.','IDENTITY_INPUT');
  validatePersonId(target);
  const occupied=await db.prepare('SELECT person_id FROM roster_people WHERE person_id=? UNION ALL SELECT old_person_id FROM roster_person_redirects WHERE old_person_id=? LIMIT 1').bind(target,target).first();
  if(occupied) throw error('That person ID is reserved and cannot be reused.');
  targetRow={...copy(before.people[0]),person_id:target,preferred_display_name:name,status:'active',merged_into_person_id:'',version:1,created_at:now,updated_at:now,provenance:'admin-approved',approved_by:input.actor,review_state:'approved'};
  after.people.push(targetRow);
 }
 if(!targetRow) throw error('Target must be a selected person.','IDENTITY_INPUT');
 if(kind==='name' && ids.length!==1 || kind==='id' && ids.length!==1 || kind==='merge' && ids.length<2) throw error('Invalid identity selection.','IDENTITY_INPUT');
 if(['merge','name','id'].includes(kind)) targetRow.preferred_display_name=name;
 if(['merge','id'].includes(kind)) {
  for(const person of after.people) if(person.person_id!==target) {
   person.status='merged'; person.merged_into_person_id=target;
   const previous=after.redirects.find(r=>r.old_person_id===person.person_id);
   if(previous) Object.assign(previous,{person_id:target,active:1,operation_id:'PREVIEW'});
   else after.redirects.push({old_person_id:person.person_id,person_id:target,active:1,operation_id:'PREVIEW',created_at:now});
  }
  for(const alias of after.aliases) if(alias.person_id!==target) Object.assign(alias,{person_id:target,provenance:'admin-approved',confidence:'confirmed',review_state:'approved',approved_by:input.actor,updated_at:now});
  for(const account of after.accounts) if(account.person_id!==target) Object.assign(account,{person_id:target,updated_at:now});
  for(const redirect of after.redirects) if(redirect.active) redirect.person_id=target;
 }
 if(kind==='alias-move') {
  const selected=after.aliases.find(a=>a.source_type===input.alias?.sourceType && a.doctor_key===input.alias?.key);
  if(!selected || selected.person_id===target) throw error('Select an alias belonging to another selected person.','IDENTITY_INPUT');
  Object.assign(selected,{person_id:target,provenance:'admin-approved',confidence:'confirmed',review_state:'approved',approved_by:input.actor,updated_at:now});
 }
 if(kind==='account-link') {
  const email=String(input.accountEmail||'').trim().toLowerCase();
  if(!email || email.length>254 || !await db.prepare('SELECT email FROM account_profiles WHERE email=?').bind(email).first()) throw error('Account not found.','IDENTITY_INPUT');
  const old=await db.prepare('SELECT * FROM account_people WHERE email=?').bind(email).first();
  if(old && !ids.includes(old.person_id)) throw error('Include the account’s existing person in the preview.');
  const account=after.accounts.find(a=>a.email===email);
  if(account) Object.assign(account,{person_id:target,updated_at:now});
  else after.accounts.push({email,person_id:target,created_at:now,updated_at:now});
 }
 bounded(after.redirects,MAX_REDIRECTS);
 if(after.aliases.filter(a=>a.person_id===target && a.review_state==='approved').length>16) throw error('This person would exceed the 16-alias calendar limit; review the identity scope first.','IDENTITY_SCOPE');
 for(const person of after.people) {
  const old=before.people.find(p=>p.person_id===person.person_id);
  if(old) { person.version=old.version+1; person.updated_at=now; person.approved_by=input.actor; }
 }
 const accountEmails=bounded([...new Set([...after.accounts.map(a=>a.email),...before.claimOwners.map(a=>a.email)])].sort(),MAX_ACCOUNTS);
 const revision=(await db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1').first()).revision;
 const value={kind,target,before,after,revision,accountEmails,requiresAccountConfirmation:accountEmails.length>1,reason:String(input.reason||'').trim()};
 value.previewToken=await token(value);
 return value;
}
function assertion(db,condition,args) { return db.prepare(`SELECT CASE WHEN ${condition} THEN 1 ELSE json('identity-conflict') END AS identity_guard`).bind(...args); }
function writeRow(db,type,row) {
 const spec=specs[type], columns=Object.keys(row), updates=columns.filter(k=>!spec.keys.includes(k));
 return db.prepare(`INSERT INTO ${spec.table} (${columns.join(',')}) VALUES (${columns.map(()=>'?').join(',')}) ON CONFLICT(${spec.keys.join(',')}) DO UPDATE SET ${updates.map(k=>`${k}=excluded.${k}`).join(',')}`).bind(...columns.map(k=>row[k]));
}
function changes(db,before,after) {
 const statements=[];
 for(const type of Object.keys(specs)) {
  const old=new Map(before[type].map(row=>[marker(type,row),row]));
  for(const row of after[type]) if(canonical(old.get(marker(type,row)))!==canonical(row)) statements.push(writeRow(db,type,row));
 }
 return statements;
}
function affected(before,after) {
 return {personIds:[...new Set([...before.people,...after.people].map(p=>p.person_id))],accountEmails:[...new Set([...before.accounts,...after.accounts,...before.claimOwners,...after.claimOwners].map(a=>a.email))],aliases:[...new Map([...before.aliases,...after.aliases].map(a=>[marker('aliases',a),{sourceType:a.source_type,key:a.doctor_key}])).values()]};
}
export async function commitIdentityOperation(db,input) {
 const preview=await previewIdentityOperation(db,input);
 if(input.previewToken!==preview.previewToken) throw error('The preview is stale. Review the identity changes again.');
 if(!preview.reason || preview.reason.length>1000) throw error('Enter a reason for this change.','IDENTITY_INPUT');
 if(preview.requiresAccountConfirmation && JSON.stringify([...(input.confirmAccountEmails||[])].sort())!==JSON.stringify(preview.accountEmails)) throw error('Explicitly confirm the named accounts; accounts themselves will remain separate.','IDENTITY_ACCOUNT_CONFIRMATION');
 const id=crypto.randomUUID(), now=new Date().toISOString(), after=copy(preview.after);
 for(const type of Object.keys(specs)) for(const row of after[type]) for(const key of Object.keys(row)) if(row[key]==='PREVIEW') row[key]=key==='operation_id'?id:now;
 const affectedPeople=affected(preview.before,after);
 const statements=[assertion(db,'(SELECT revision FROM roster_identity_revision WHERE id=1)=?',[preview.revision])];
 // Guard exact rows and claim ownership inside the same transaction. This also
 // catches legacy linking paths which do not yet increment the revision.
 for(const type of Object.keys(specs)) {
  const spec=specs[type];
  for(const row of preview.before[type]) statements.push(assertion(db,`EXISTS(SELECT 1 FROM ${spec.table} WHERE ${Object.keys(row).map(k=>`${k} IS ?`).join(' AND ')})`,Object.values(row)));
 }
 for(const person of preview.before.people) {
  for(const type of ['aliases','accounts','redirects']) statements.push(assertion(db,`(SELECT COUNT(*) FROM ${specs[type].table} WHERE person_id=?)=?`,[person.person_id,preview.before[type].filter(r=>r.person_id===person.person_id).length]));
 }
 for(const alias of preview.before.aliases) {
  const emails=preview.before.claimOwners.filter(c=>c.source_type===alias.source_type && c.doctor_key===alias.doctor_key).map(c=>c.email);
  statements.push(assertion(db,`(SELECT COUNT(*) FROM account_claims WHERE source_type=? AND doctor_key=?)=?`,[alias.source_type,alias.doctor_key,emails.length]));
  for(const email of emails) statements.push(assertion(db,'EXISTS(SELECT 1 FROM account_claims WHERE email=? AND source_type=? AND doctor_key=?)',[email,alias.source_type,alias.doctor_key]));
 }
 for(const row of after.people) if(!preview.before.people.some(p=>p.person_id===row.person_id)) statements.push(assertion(db,'NOT EXISTS(SELECT 1 FROM roster_people WHERE person_id=?) AND NOT EXISTS(SELECT 1 FROM roster_person_redirects WHERE old_person_id=?)',[row.person_id,row.person_id]));
 for(const row of after.accounts) if(!preview.before.accounts.some(a=>a.email===row.email)) statements.push(assertion(db,'NOT EXISTS(SELECT 1 FROM account_people WHERE email=?)',[row.email]));
 statements.push(...changes(db,preview.before,after), db.prepare('UPDATE roster_identity_revision SET revision=revision+1 WHERE id=1'));
 statements.push(db.prepare(`INSERT INTO roster_identity_operations(operation_id,kind,status,actor,reason,created_at,before_json,after_json,affected_json) VALUES(?,?,'committed',?,?,?,?,?,?)`).bind(id,preview.kind,input.actor,preview.reason,now,JSON.stringify(preview.before),JSON.stringify(after),JSON.stringify(affectedPeople)));
 statements.push(db.prepare('INSERT INTO roster_identity_jobs(operation_id,affected_json,created_at) VALUES(?,?,?)').bind(id,JSON.stringify(affectedPeople),now));
 for(const personId of affectedPeople.personIds) statements.push(db.prepare('INSERT INTO roster_identity_operation_people(person_id,operation_id,created_at) VALUES(?,?,?)').bind(personId,id,now));
 if(statements.length>400) throw error('Operation exceeds the transaction budget.','IDENTITY_SCOPE');
 try { await db.batch(statements); } catch(e) { if(/malformed JSON/i.test(e.message)) throw error('Identity links changed; reload the preview.'); throw e; }
 return {operationId:id,affected:affectedPeople,status:'committed'};
}

function sameRows(left,right,type) {
 const order=list=>list.map(row=>{ const comparable={...row}; if(type==='people') { delete comparable.version; delete comparable.updated_at; } return [marker(type,row),canonical(comparable)]; }).sort((a,b)=>a[0].localeCompare(b[0]));
 return JSON.stringify(order(left))===JSON.stringify(order(right));
}
export async function previewIdentityReversal(db,operationId) {
 const op=await db.prepare('SELECT * FROM roster_identity_operations WHERE operation_id=?').bind(operationId).first();
 if(!op || op.status!=='committed' || op.kind==='reverse') throw error('This operation cannot be reversed.');
 const expected=JSON.parse(op.after_json), original=JSON.parse(op.before_json);
 const current=await scope(db,expected.people.map(p=>p.person_id));
 for(const type of Object.keys(specs)) if(!sameRows(current[type],expected[type],type)) throw error('Later identity changes depend on this operation. Reverse or reassign those changes first.','IDENTITY_DEPENDENCY');
 if(JSON.stringify(current.claimOwners)!==JSON.stringify(expected.claimOwners)) throw error('Account claims changed after this operation. Review those links before reversal.','IDENTITY_DEPENDENCY');
 const restored=copy(original);
 for(const person of restored.people) { person.version=current.people.find(p=>p.person_id===person.person_id).version+1; person.updated_at='PREVIEW'; }
 // IDs stay reserved permanently, including a canonical ID created by a merge.
 for(const person of current.people) if(!restored.people.some(p=>p.person_id===person.person_id)) restored.people.push({...person,status:'retired',merged_into_person_id:'',version:person.version+1,updated_at:'PREVIEW'});
 for(const redirect of current.redirects) if(!restored.redirects.some(r=>r.old_person_id===redirect.old_person_id)) restored.redirects.push({...redirect,active:0});
 const revision=(await db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1').first()).revision;
 const accountEmails=bounded([...new Set([...current.accounts,...current.claimOwners].map(a=>a.email))].sort(),MAX_ACCOUNTS);
 const preview={operationId,revision,before:current,after:restored,affected:JSON.parse(op.affected_json),accountEmails,requiresAccountConfirmation:accountEmails.length>1};
 preview.previewToken=await token(preview);
 return preview;
}
export async function reverseIdentityOperation(db,input) {
 const preview=await previewIdentityReversal(db,input.operationId);
 if(input.previewToken!==preview.previewToken) throw error('The reversal preview is stale.');
 if(preview.requiresAccountConfirmation && JSON.stringify([...(input.confirmAccountEmails||[])].sort())!==JSON.stringify(preview.accountEmails)) throw error('Explicitly confirm the named accounts before reversal.','IDENTITY_ACCOUNT_CONFIRMATION');
 const reason=String(input.reason||'').trim(); if(!reason || reason.length>1000) throw error('Enter a reversal reason.','IDENTITY_INPUT');
 const id=crypto.randomUUID(),now=new Date().toISOString(),after=copy(preview.after);
 for(const person of after.people) if(person.updated_at==='PREVIEW') person.updated_at=now;
 const statements=[assertion(db,'(SELECT revision FROM roster_identity_revision WHERE id=1)=?',[preview.revision])];
 for(const type of Object.keys(specs)) for(const row of preview.before[type]) statements.push(assertion(db,`EXISTS(SELECT 1 FROM ${specs[type].table} WHERE ${Object.keys(row).map(k=>`${k} IS ?`).join(' AND ')})`,Object.values(row)));
 for(const person of preview.before.people) for(const type of ['aliases','accounts','redirects']) {
  const column=type==='redirects'?'person_id':'person_id';
  statements.push(assertion(db,`(SELECT COUNT(*) FROM ${specs[type].table} WHERE ${column}=?)=?`,[person.person_id,preview.before[type].filter(r=>r.person_id===person.person_id).length]));
 }
 for(const alias of preview.before.aliases) {
  const owners=preview.before.claimOwners.filter(c=>c.source_type===alias.source_type && c.doctor_key===alias.doctor_key);
  statements.push(assertion(db,'(SELECT COUNT(*) FROM account_claims WHERE source_type=? AND doctor_key=?)=?',[alias.source_type,alias.doctor_key,owners.length]));
  for(const owner of owners) statements.push(assertion(db,'EXISTS(SELECT 1 FROM account_claims WHERE email=? AND source_type=? AND doctor_key=?)',[owner.email,alias.source_type,alias.doctor_key]));
 }
 // Remove only links created by the original operation, after the full-state
 // comparison. Reversal is never best-effort and cannot delete user accounts.
 for(const type of ['aliases','accounts']) for(const row of preview.before[type]) if(!after[type].some(r=>marker(type,r)===marker(type,row))) {
  const spec=specs[type]; statements.push(db.prepare(`DELETE FROM ${spec.table} WHERE ${spec.keys.map(k=>`${k}=?`).join(' AND ')}`).bind(...spec.keys.map(k=>row[k])));
 }
 statements.push(...changes(db,preview.before,after),db.prepare('UPDATE roster_identity_revision SET revision=revision+1 WHERE id=1'));
 statements.push(db.prepare("UPDATE roster_identity_operations SET status='reversed',reversed_by=? WHERE operation_id=? AND status='committed'").bind(id,input.operationId));
 statements.push(db.prepare(`INSERT INTO roster_identity_operations(operation_id,kind,status,actor,reason,created_at,reversed_operation_id,before_json,after_json,affected_json) VALUES(?,'reverse','committed',?,?,?,?,?,?,?)`).bind(id,input.actor,reason,now,input.operationId,JSON.stringify(preview.before),JSON.stringify(after),JSON.stringify(preview.affected)));
 statements.push(db.prepare('INSERT INTO roster_identity_jobs(operation_id,affected_json,created_at) VALUES(?,?,?)').bind(id,JSON.stringify(preview.affected),now));
 for(const personId of preview.affected.personIds) statements.push(db.prepare('INSERT INTO roster_identity_operation_people(person_id,operation_id,created_at) VALUES(?,?,?)').bind(personId,id,now));
 if(statements.length>400) throw error('Reversal exceeds the transaction budget.','IDENTITY_SCOPE');
 try { await db.batch(statements); } catch(e) { if(/malformed JSON/i.test(e.message)) throw error('Identity changed; reload the reversal preview.'); throw e; }
 return {operationId:id,status:'committed',affected:preview.affected};
}

export async function queryIdentityPeople(db,{after='',limit=25,search='',searchType='name',sourceType=''}={}) {
 limit=Math.min(25,Math.max(1,Number(limit)||25));
 search=String(search).trim(); after=String(after);
 if(search.length>200 || after.length>160) throw error('Search is too long.','IDENTITY_INPUT');
 let people;
 if(!search) people=await rows(db,'SELECT * FROM roster_people WHERE person_id>? ORDER BY person_id LIMIT ?',[after,limit+1]);
 else if(searchType==='person') people=await rows(db,'SELECT p.* FROM roster_people p WHERE p.person_id=? UNION SELECT p.* FROM roster_person_redirects r JOIN roster_people p ON p.person_id=r.person_id WHERE r.old_person_id=? AND r.active=1',[search,search]);
 else if(searchType==='account') people=await rows(db,'SELECT p.* FROM account_people a JOIN roster_people p ON p.person_id=a.person_id WHERE a.email=?',[search.toLowerCase()]);
 else if(searchType==='alias') {
  if(!SOURCES.has(sourceType)) throw error('Select the alias site.','IDENTITY_INPUT');
  people=await rows(db,'SELECT p.* FROM roster_person_aliases a JOIN roster_people p ON p.person_id=a.person_id WHERE a.source_type=? AND a.doctor_key=?',[sourceType,search]);
 } else if(searchType==='name') {
  const cursor=after?await db.prepare('SELECT preferred_display_name FROM roster_people WHERE person_id=?').bind(after).first():null;
  if(after && !cursor) throw error('Refresh the search before continuing.','IDENTITY_INPUT');
  people=await rows(db,`SELECT * FROM roster_people WHERE preferred_display_name COLLATE NOCASE>=? AND preferred_display_name COLLATE NOCASE<? ${cursor?'AND (preferred_display_name COLLATE NOCASE>? OR (preferred_display_name COLLATE NOCASE=? AND person_id>?))':''} ORDER BY preferred_display_name COLLATE NOCASE,person_id LIMIT ?`,[search,search+'\uffff',...(cursor?[cursor.preferred_display_name,cursor.preferred_display_name,after]:[]),limit+1]);
 } else throw error('Unsupported search.','IDENTITY_INPUT');
 return {people:people.slice(0,limit),next:people.length>limit?people[limit-1].person_id:''};
}
export async function queryIdentityPerson(db,id) {
 const data=await scope(db,[String(id)]);
 const history=await rows(db,`SELECT o.operation_id,o.kind,o.status,o.actor,o.reason,o.created_at,o.reversed_operation_id,o.reversed_by,j.status AS publication_status,j.last_error AS publication_error FROM roster_identity_operation_people h JOIN roster_identity_operations o ON o.operation_id=h.operation_id LEFT JOIN roster_identity_jobs j ON j.operation_id=o.operation_id WHERE h.person_id=? ORDER BY h.created_at DESC,h.operation_id DESC LIMIT 26`,[String(id)]);
 // No roster history scan. Coverage comes from published artifacts in the API.
 return {...data,history:history.slice(0,25),historyTruncated:history.length>25};
}
export async function expandApprovedIdentityAliases(db,aliases) {
 if(!aliases.length || aliases.length>16) return aliases;
 const predicates=aliases.map(()=>'(a.source_type=? AND a.doctor_key=?)').join(' OR ');
 const result=bounded(await rows(db,`SELECT DISTINCT b.source_type,b.doctor_key,b.display_name,p.preferred_display_name,p.person_id
 FROM roster_person_aliases a JOIN roster_people p ON p.person_id=a.person_id
 JOIN roster_person_aliases b ON b.person_id=p.person_id
 WHERE (${predicates}) AND a.review_state='approved' AND b.review_state='approved' AND p.status='active'
 ORDER BY b.source_type,b.doctor_key LIMIT 17`,aliases.flatMap(a=>[a.sourceType,a.key])),16);
 return bounded([...new Map([...aliases,...result.map(r=>({sourceType:r.source_type,key:r.doctor_key,displayName:r.display_name,personId:r.person_id,preferredName:r.preferred_display_name}))].map(a=>[`${a.sourceType}|${a.key}`,a])).values()],16);
}

export async function accountIdentityAliases(db,email,aliases) {
 const link=await db.prepare("SELECT a.person_id,p.status FROM account_people a JOIN roster_people p ON p.person_id=a.person_id WHERE a.email=?").bind(String(email||'').toLowerCase()).first();
 if(!link) return expandApprovedIdentityAliases(db,aliases);
 const person=await resolvePersonId(db,link.person_id);
 const list=bounded(await rows(db,"SELECT source_type,doctor_key,display_name FROM roster_person_aliases WHERE person_id=? AND review_state='approved' ORDER BY source_type,doctor_key LIMIT 17",[person.person_id]),16);
 return list.map(a=>({sourceType:a.source_type,key:a.doctor_key,displayName:a.display_name,personId:person.person_id,preferredName:person.preferred_display_name}));
}

// Rebuilds are served from the existing published roster artifacts. Identity
// jobs invalidate only exact owners; they never dispatch roster parsing or a
// database-wide snapshot warm-up. The R2 revision wakes open calendar tabs.
export async function publishIdentityOperation(db,r2,operationId) {
 const job=await db.prepare('SELECT * FROM roster_identity_jobs WHERE operation_id=?').bind(operationId).first();
 if(!job) return {status:'missing'};
 if(job.status==='complete') return {status:'complete'};
 const impacted=JSON.parse(job.affected_json);
 try {
  const statements=[];
  const owners=impacted.accountEmails.map(email=>({type:'user-account',id:email}));
  owners.push(...impacted.accountEmails.map(email=>({type:'creator-account',id:email})));
  // Profile lookup uses the doctor-key index. Filter exact source affiliation
  // before invalidating, and refuse an oversized profile scope.
  const keys=[...new Set(impacted.aliases.map(a=>a.key))];
  let profiles=[];
  if(keys.length) profiles=bounded(await rows(db,`SELECT profile_id,doctor_key,source_types_json FROM doctor_profiles WHERE doctor_key IN (${keys.map(()=>'?').join(',')}) LIMIT 17`,keys),16);
  for(const profile of profiles) if(JSON.parse(profile.source_types_json||'[]').some(source=>impacted.aliases.some(a=>a.sourceType===source && a.key===profile.doctor_key))) owners.push({type:'doctor-profile',id:profile.profile_id});
  for(const owner of owners) {
    const entries=bounded(await rows(db,'SELECT doctor_key,range_key FROM snapshot_registry WHERE owner_type=? AND owner_id=? LIMIT 9',[owner.type,owner.id]),8);
    for(const entry of entries) statements.push(db.prepare("UPDATE snapshot_registry SET status='stale',requested_revision=?,updated_at=? WHERE owner_type=? AND owner_id=? AND doctor_key=? AND range_key=?").bind('identity:'+operationId,new Date().toISOString(),owner.type,owner.id,entry.doctor_key,entry.range_key));
  }
  if(statements.length) await db.batch(statements);
  // Exact changed aliases enter the indexed weekly suggestion queue. A name
  // or alias change never causes the audit to rescan unchanged identities.
  if(impacted.aliases.length) {
   const changed=await rows(db,`SELECT source_type,doctor_key,display_name,person_id FROM roster_person_aliases WHERE ${impacted.aliases.map(()=>'(source_type=? AND doctor_key=?)').join(' OR ')} LIMIT 33`,impacted.aliases.flatMap(a=>[a.sourceType,a.key]));
   const features=[];
   for(const alias of changed) { const fingerprint=await token([alias.source_type,alias.doctor_key,alias.display_name,alias.person_id]);features.push(db.prepare('UPDATE roster_identity_features SET fingerprint=? WHERE source_type=? AND doctor_key=? AND fingerprint<>?').bind(fingerprint,alias.source_type,alias.doctor_key,fingerprint)); }
   if(features.length) await db.batch(features);
  }
  // Random wake-up marker is safe out of order: calendar assembly always reads
  // current identity state. No account identifiers leave the public endpoint.
  await r2.put('identity/revision.json',JSON.stringify({revision:crypto.randomUUID()}),{httpMetadata:{contentType:'application/json'}});
  await db.prepare("UPDATE roster_identity_jobs SET status='complete',attempts=attempts+1,last_error='' WHERE operation_id=?").bind(operationId).run();
  return {status:'complete'};
 } catch(e) {
  await db.prepare("UPDATE roster_identity_jobs SET status='failed',attempts=attempts+1,last_error=? WHERE operation_id=?").bind(String(e.message).slice(0,500),operationId).run();
  return {status:'failed',error:e.message};
 }
}
