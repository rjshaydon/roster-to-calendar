import { normalizeEmail } from './email.js';
import { AuthenticationError } from './supabase-auth.js';

function cohort(config) {
  if(config.migrationEmails === undefined)return null;
  if(!Array.isArray(config.migrationEmails)||config.migrationEmails.length>100
    || config.migrationEmails.some(email=>typeof email!=='string'||!email.includes('@')||email.length>254||/[\s\x00-\x1f\x7f]/.test(email))) {
    throw new AuthenticationError('auth-configuration','Authentication rollout is not configured.');
  }
  return new Set(config.migrationEmails.map(normalizeEmail));
}

// An omitted cohort means all accounts; an empty cohort means no new migrations.
// This controls rollout eligibility, never authentication or roster ownership.
export function requirePasswordMigrationAdmission(config,email,now=Date.now()) {
  const deadline=config.legacyMigrationUntil===undefined?null:Date.parse(config.legacyMigrationUntil);
  if(config.migrationEnabled!==true||config.migrationPaused===true
    ||deadline!==null&&(!Number.isFinite(deadline)||deadline<=now)) {
    throw new AuthenticationError('auth-migration-paused','The account upgrade is temporarily unavailable. Please try again later.');
  }
  const emails=cohort(config);
  if(emails&&!emails.has(normalizeEmail(email)))throw new AuthenticationError('auth-migration-paused','The account upgrade is not yet enabled for this account.');
}

export function allowLegacyStateDuringRollout(env,request,body,now=Date.now()) {
  // Any session cookie, including malformed/expired cookies, stays on the cookie
  // verifier. Never turn a rejected cookie into legacy password authentication.
  if((request.headers.get('Cookie')||'').split(';').some(p=>p.trim().split('=')[0]==='__Host-roster_session'))return false;
  if(body?.action==='login'&&body?.mode==='create'||body?.action==='acceptInvite')return false;
  let config;try{config=JSON.parse(env.ROSTER_AUTH_CONFIG||'{}');}catch{return false;}
  const emails=cohort(config),email=normalizeEmail(body?.email);
  if(emails&&email&&!emails.has(email))return true;
  // An explicitly dated, short grace period preserves already-open clients.
  // Actual sign-in still upgrades; this cannot create accounts or bypass the
  // pending-operation and retired-hash checks in the legacy verifier.
  const until=typeof config.legacyStateUntil==='string'?Date.parse(config.legacyStateUntil):NaN;
  return Boolean(email&&body?.action&&body.action!=='login'
    &&Number.isFinite(until)&&until>now&&until-now<=7*86400000);
}

export async function requireLegacyCredentialNotPending(db,email) {
  const link=await db.prepare('SELECT credential_operation_id FROM account_authentication_links WHERE email=?').bind(normalizeEmail(email)).first();
  if(link?.credential_operation_id)throw new AuthenticationError('auth-upgrade-required','Sign in with your new password to finish the account upgrade.',409);
}
