// Disposable database only. This module never connects to a service.
import {issueAccountSession} from '../functions/_lib/account-sessions.js';
export const fixtureEmails={admin:'rhaydon@gmail.com',user:'requester@example.test',holder:'holder@example.test',stranger:'stranger@example.test'};
export async function seedSupportFixture(db,{reuse=false}={}){
 const result={};let n=1;
 for(const [name,email]of Object.entries(fixtureEmails)){
  const subject=`11111111-2222-4333-8444-${String(n++).padStart(12,'0')}`;
  if(!reuse||!await db.prepare('SELECT email FROM account_profiles WHERE email=?').bind(email).first())await db.batch([db.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,?,?,'','',?)").bind(email,'Synthetic '+name,name==='admin'?'creator':'user','synthetic-feed-'+name),
   db.prepare("INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at) VALUES(?,'supabase',?,'ready','fixture','now','now')").bind(email,subject)]);
  const issued=await issueAccountSession(db,{email,subject});result[name]={cookie:issued.setCookie.split(';')[0],email,token_hash:issued.tokenHash,role:name==='admin'?'creator':'user'};
 }
 if(!reuse||!await db.prepare("SELECT person_id FROM roster_people WHERE person_id='person:fixture-doctor'").first())await db.batch([db.prepare("INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:fixture-doctor','Synthetic Doctor')"),
  db.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('mmc','SYNTHETIC DOCTOR','Synthetic Doctor','person:fixture-doctor')"),
  db.prepare("INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES('mch','SYNTHETIC DR','Synthetic Dr','person:fixture-doctor')"),
  db.prepare("INSERT INTO account_people(email,person_id) VALUES(?,'person:fixture-doctor')").bind(fixtureEmails.holder),
  db.prepare("INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,'mmc','SYNTHETIC DOCTOR','Synthetic Doctor')").bind(fixtureEmails.holder),
  db.prepare("INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,'mch','SYNTHETIC DR','Synthetic Dr')").bind(fixtureEmails.holder),
  ...[fixtureEmails.user,fixtureEmails.holder].map(email=>db.prepare("INSERT INTO snapshot_registry(owner_type,owner_id,range_key,status) VALUES('user-account',?,'fixture','ready')").bind(email))]);
 return result;
}
export async function expandSupportFixture(db){
 const aliases=[['ddh','SYNTHETIC DDH'],['vhh','SYNTHETIC VHH']];
 await db.batch(aliases.map(([site,key])=>db.prepare("INSERT OR IGNORE INTO roster_person_aliases(source_type,doctor_key,display_name,person_id) VALUES(?,?,?,'person:fixture-doctor')").bind(site,key,'Synthetic Doctor')));
 const allAliases=(await db.prepare("SELECT source_type,doctor_key FROM roster_person_aliases WHERE person_id='person:fixture-doctor'").all()).results;
 await db.batch([db.prepare("INSERT OR IGNORE INTO account_people(email,person_id) VALUES(?,'person:fixture-doctor')").bind(fixtureEmails.stranger),...allAliases.flatMap(a=>[fixtureEmails.holder,fixtureEmails.stranger].map(email=>db.prepare('INSERT OR IGNORE INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,?,?,?)').bind(email,a.source_type,a.doctor_key,'Synthetic Doctor')))]);
}
export async function scaleSupportFixture(db,count=500){
 if(![100,500].includes(count))throw Error('Synthetic scale must be 100 or 500.');
 const have=Number((await db.prepare("SELECT count(*) n FROM account_profiles WHERE email LIKE 'scale%@example.test'").first())?.n||0),now=new Date().toISOString();
 for(let offset=have;offset<count;offset+=20){const statements=[];for(let i=offset;i<Math.min(offset+20,count);i++){const caseId=crypto.randomUUID(),channelId=crypto.randomUUID();statements.push(db.prepare('INSERT INTO account_profiles(email,real_name) VALUES(?,?)').bind('scale'+i+'@example.test','Synthetic scale '+i),db.prepare("INSERT INTO support_cases(case_id,requester_email,kind,subject,status,created_at,updated_at) VALUES(?,?,'support','Synthetic historic case','resolved',?,?)").bind(caseId,fixtureEmails.user,now,now),db.prepare("INSERT INTO support_channels(channel_id,case_id,participant_email,case_status,updated_at) VALUES(?,?,?,'resolved',?)").bind(channelId,caseId,fixtureEmails.user,now));}await db.batch(statements);}
 return {syntheticAccounts:count};
}
