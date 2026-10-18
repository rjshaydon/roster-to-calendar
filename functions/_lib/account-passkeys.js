import { authenticationSettings } from './authentication-settings.js';
import { AuthenticationError, providerRequest, signInWithSupabasePassword, supabaseAuthenticationConfig } from './supabase-auth.js';
import { loadAuthenticationLink } from './authentication-migration.js';

export const PASSKEY_ORIGIN = 'https://rtc.curiousmind.app';
export const PASSKEY_RP = 'rtc.curiousmind.app';
const COOKIE = '__Host-roster_passkey';
const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const PURPOSE = new TextEncoder().encode('roster-passkey-ceremony-v1:' + PASSKEY_ORIGIN);
const fail = () => new AuthenticationError('auth-passkey-expired', 'Please start passkey setup or sign-in again.', 400);
export function passkeysEnabled(env, request) {
  try { const c = authenticationSettings(env); return c.enabled === true && c.sessionsEnabled === true && c.passkeysEnabled === true && new URL(request.url).origin === PASSKEY_ORIGIN; }
  catch { return false; }
}
function encode(bytes) { return btoa(String.fromCharCode(...bytes)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,''); }
function decode(value) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]+$/.test(value) || value.length > 3700) throw fail();
  return Uint8Array.from(atob(value.replace(/-/g,'+').replace(/_/g,'/')), c => c.charCodeAt(0));
}
async function key(env) {
  const secret = supabaseAuthenticationConfig(env).secretKey;
  const material = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('roster-passkey-cookie-v1:' + secret));
  return crypto.subtle.importKey('raw', material, 'AES-GCM', false, ['encrypt','decrypt']);
}
// Short-lived provider proof never enters JS or D1. It is encrypted, host-only,
// purpose-bound and tied to one exact app session/version. Logout invalidates it.
export async function sealPasskeyCookie(env, value, now = Date.now()) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const plain = new TextEncoder().encode(JSON.stringify({...value, expires:now+300000}));
  const cipher = new Uint8Array(await crypto.subtle.encrypt({name:'AES-GCM',iv,additionalData:PURPOSE},await key(env),plain));
  const bytes = new Uint8Array(iv.length+cipher.length);bytes.set(iv);bytes.set(cipher,iv.length);
  const encoded = encode(bytes);if(encoded.length>3700)throw fail();
  return `${COOKIE}=${encoded}; Path=/; Secure; HttpOnly; SameSite=Lax; Max-Age=300`;
}
export function clearPasskeyCookie() { return `${COOKIE}=; Path=/; Secure; HttpOnly; SameSite=Lax; Max-Age=0`; }
export async function readPasskeyCookie(env, request, now = Date.now()) {
  try {
    const values=(request.headers.get('Cookie')||'').split(';').map(x=>x.trim()).filter(x=>x.startsWith(COOKIE+'='));
    if(values.length!==1)throw fail();
    const bytes=decode(values[0].slice(COOKIE.length+1));
    const plain=await crypto.subtle.decrypt({name:'AES-GCM',iv:bytes.slice(0,12),additionalData:PURPOSE},await key(env),bytes.slice(12));
    const data=JSON.parse(new TextDecoder().decode(plain));
    if(!Number.isSafeInteger(data.expires)||data.expires<=now||data.expires>now+300000)throw fail();
    return data;
  } catch { throw fail(); }
}
export async function passkeyLoginGrant(env, request, proof, issued) {
  if(!passkeysEnabled(env,request)||typeof proof.accessToken!=='string')return '';
  return sealPasskeyCookie(env,{kind:'grant',email:proof.email,subject:proof.subject,version:proof.credentialVersion,
    session:issued.tokenHash,accessToken:proof.accessToken});
}
export async function passkeyAccount(db, session) {
  const link=await loadAuthenticationLink(db,session.email);
  if(!link||link.phase!=='ready'||link.provider!=='supabase'||!UUID.test(link.provider_subject)||link.credential_operation_id)throw fail();
  return link;
}
export function grantMatches(grant,session,link) {
  return ['grant','registration'].includes(grant?.kind)&&grant.email===session.email&&grant.subject===link.provider_subject
    &&grant.version===link.credential_version&&grant.session===session.token_hash&&typeof grant.accessToken==='string';
}
export async function registrationProof(env,request,session,link,password) {
  try { const grant=await readPasskeyCookie(env,request);if(grantMatches(grant,session,link))return grant.accessToken; } catch {}
  if(!password)throw new AuthenticationError('auth-passkey-password-required','Confirm your current password to add a passkey.',401);
  const proof=await signInWithSupabasePassword(env,{email:session.email,password,expectedSubject:link.provider_subject});
  return proof.accessToken;
}
export function validatePasskeyOptions(data,registration=false) {
  if(!UUID.test(data?.challenge_id||'')||!data.options||typeof data.options.challenge!=='string'
    ||(registration?data.options.rp?.id:data.options.rpId)!==PASSKEY_RP)throw fail();
  // Supabase verifies signatures, RP/origin, challenge, credential ownership and
  // user verification; require a biometric/PIN-verified credential in the browser.
  if(registration){data.options.authenticatorSelection={...data.options.authenticatorSelection,userVerification:'required'};}
  else data.options.userVerification='required';
  return {challengeId:data.challenge_id,options:data.options};
}
export function validateCredential(credential) {
  if(!credential||credential.type!=='public-key'||typeof credential.id!=='string'||credential.id.length>2048
    ||!credential.response||typeof credential.response.clientDataJSON!=='string')throw fail();
}
export async function listAccountPasskeys(env, subject) {
  if(!UUID.test(subject))throw fail();
  const data=await providerRequest(env,`admin/users/${subject}/passkeys`,{admin:true,method:'GET'});
  if(!Array.isArray(data)||data.length>100)throw fail();
  return data.map(item=>{
    if(!UUID.test(item.id||''))throw fail();
    return {id:item.id,name:String(item.friendly_name||'Passkey').slice(0,120),createdAt:item.created_at,lastUsedAt:item.last_used_at||null};
  });
}
export async function removeAccountPasskey(env,subject,id) {
  if(!UUID.test(subject)||!UUID.test(id||''))throw fail();
  await providerRequest(env,`admin/users/${subject}/passkeys/${id}`,{admin:true,method:'DELETE'});
}
export async function verifyPasskeySignIn(env,challengeId,credential) {
  validateCredential(credential);
  // The provider's current default is verification "preferred". Enforce our
  // stronger policy on the signed authenticator-data flag too, rather than
  // relying only on browser options which an HTTP client could omit.
  let authenticatorData;try{authenticatorData=decode(credential.response.authenticatorData);}catch{throw fail();}
  if(authenticatorData.length<37||(authenticatorData[32]&0x05)!==0x05)throw fail();
  const data=await providerRequest(env,'passkeys/authentication/verify',{body:{challenge_id:challengeId,credential}});
  const user=data?.user;
  if(!UUID.test(user?.id||'')||typeof user.email!=='string'||!user.email_confirmed_at||user.is_anonymous
    ||user.banned_until&&Date.parse(user.banned_until)>Date.now()||typeof data.access_token!=='string')throw fail();
  // Exact provider subject AND email must already have a ready local link.
  // Passkeys never create accounts or adopt an existing account by email alone.
  const link=await env.ROSTER_DB.prepare(`SELECT email,credential_version FROM account_authentication_links
    WHERE provider_subject=? AND email=? AND provider='supabase' AND phase='ready' AND credential_operation_id=''`)
    .bind(user.id,user.email.trim().toLowerCase()).first();
  if(!link)throw new AuthenticationError('auth-passkey-unlinked','Use your password to sign in to this account.',401);
  return {email:link.email,subject:user.id,credentialVersion:link.credential_version,accessToken:data.access_token};
}
