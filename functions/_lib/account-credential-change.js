import {normalizeEmail} from './email.js';
import {AuthenticationError,supabaseAuthenticationConfig,signInWithSupabasePassword,validateNewPassword} from './supabase-auth.js';
import {finishAuthenticationMigration,finishNewAccountAuthentication,loadAuthenticationLink} from './authentication-migration.js';

const pending=()=>new AuthenticationError('auth-password-pending','The password update has not been confirmed. Try signing in with your new password; contact the administrator if it still fails.',409);
function version(link){const v=link?.credential_version;return Number.isSafeInteger(v)&&v>=0?v:0;}
export function credentialProof(proof,link){return{...proof,credentialVersion:version(link),credentialOperationId:link.credential_operation_id||''};}
async function targetDigest(env,link,id,password){
  if(typeof password!=='string'||!password||new TextEncoder().encode(password).length>1024)throw pending();
  const config=supabaseAuthenticationConfig(env);
  const key=await crypto.subtle.importKey('raw',new TextEncoder().encode(config.secretKey),{name:'HMAC',hash:'SHA-256'},false,['sign']);
  const value=JSON.stringify(['roster-password-operation-v1',normalizeEmail(link.email),link.provider_subject,id,password]);
  return Array.from(new Uint8Array(await crypto.subtle.sign('HMAC',key,new TextEncoder().encode(value))),b=>b.toString(16).padStart(2,'0')).join('');
}
function equalDigest(a,b){if(typeof a!=='string'||typeof b!=='string'||a.length!==64||b.length!==64)return false;let diff=0;for(let i=0;i<64;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);return diff===0;}

// Atomic admission revokes sessions before any provider mutation. No new
// session can be issued until the intended new password has provider proof.
export async function beginCredentialChange(env,link,newPassword){
  validateNewPassword(newPassword);
  if(!link||link.credential_operation_id)throw pending();
  const id=crypto.randomUUID(),nextVersion=version(link)+1,now=new Date().toISOString();
  const digest=await targetDigest(env,link,id,newPassword);
  const result=await env.ROSTER_DB.prepare(`UPDATE account_authentication_links
    SET credential_version=credential_version+1,credential_operation_id=?,credential_target_digest=?,credential_operation_started_at=?,updated_at=?
    WHERE email=? AND provider_subject=? AND credential_version=? AND credential_operation_id=''`)
    .bind(id,digest,now,now,link.email,link.provider_subject,version(link)).run();
  // D1 includes trigger side effects in changes; this uniquely keyed CAS can
  // update at most one direct row. Any positive count means it matched.
  if(!(Number(result.meta?.changes)>=1))throw pending();
  return{...link,credential_version:nextVersion,credential_operation_id:id,credential_target_digest:digest,credential_operation_started_at:now};
}

// Only a definitive rejection of the original mutation can cancel admission.
// A timeout/5xx keeps the intended target, so a late provider write cannot
// restore access under an old password after the operation was abandoned.
export async function cancelRejectedCredentialChange(env,operation,error){
  if(error?.credentialChangeNotApplied!==true)return;
  await env.ROSTER_DB.prepare(`UPDATE account_authentication_links SET credential_version=credential_version+1,
    credential_operation_id='',credential_target_digest='',credential_operation_started_at='',updated_at=?
    WHERE email=? AND provider_subject=? AND credential_version=? AND credential_operation_id=?`)
    .bind(new Date().toISOString(),operation.email,operation.provider_subject,version(operation),operation.credential_operation_id).run();
}

// Call only for a freshly admitted operation. If the target already works,
// never submit a redundant password write: a late identical write could undo a
// later password change. Each operation has at most ONE actual provider write.
export async function applyCredentialChange(env,operation,password,mutate){
  if(!operation?.credential_operation_id||!equalDigest(operation.credential_target_digest,await targetDigest(env,operation,operation.credential_operation_id,password)))throw pending();
  let needsMutation=false;
  try{await signInWithSupabasePassword(env,{email:operation.email,password,expectedSubject:operation.provider_subject});}
  catch(error){
    if(['invalid-credentials','email-verification-required'].includes(error?.code))needsMutation=true;
    else{await cancelRejectedCredentialChange(env,operation,{credentialChangeNotApplied:true});error.credentialChangeNotApplied=true;throw error;}
  }
  if(needsMutation){
    try{await mutate();}
    catch(error){await cancelRejectedCredentialChange(env,operation,error);throw error;}
  }
  return verifyAndCompleteCredentialChange(env,operation,password);
}

// A lost response can be completed only when the intended new password already
// works at the provider. Never reissue an uncertain password mutation: its late
// write could otherwise overwrite a subsequent, completed password change.
export async function verifyAndCompleteCredentialChange(env,operation,password,{resume=false}={}){
  if(!operation?.credential_operation_id||!equalDigest(operation.credential_target_digest,await targetDigest(env,operation,operation.credential_operation_id,password)))throw pending();
  const db=env.ROSTER_DB;
  let verified;
  try{verified=await signInWithSupabasePassword(env,{email:operation.email,password,expectedSubject:operation.provider_subject});}
  catch(error){if(resume&&['invalid-credentials','email-verification-required'].includes(error?.code))throw pending();throw error;}
  const proof=credentialProof(verified,operation);
  if(operation.phase==='pending'){
    if(operation.credential_revision==='new-account')await finishNewAccountAuthentication(db,{link:operation,proof});
    else{
      const record=await db.prepare('SELECT email,password_hash,password_salt FROM account_profiles WHERE email=?').bind(operation.email).first();
      await finishAuthenticationMigration(db,{reservation:operation,record:{email:record?.email,passwordHash:record?.password_hash,passwordSalt:record?.password_salt},proof});
    }
  }
  await db.prepare(`UPDATE account_authentication_links SET credential_version=credential_version+1,
    credential_operation_id='',credential_target_digest='',credential_operation_started_at='',updated_at=?
    WHERE email=? AND provider_subject=? AND credential_version=? AND credential_operation_id=? AND phase='ready'`)
    .bind(new Date().toISOString(),operation.email,operation.provider_subject,version(operation),operation.credential_operation_id).run();
  const current=await loadAuthenticationLink(db,operation.email);
  if(current?.provider_subject!==operation.provider_subject||current.phase!=='ready'||current.credential_operation_id||version(current)!==version(operation)+1)throw pending();
  return credentialProof(proof,current);
}
