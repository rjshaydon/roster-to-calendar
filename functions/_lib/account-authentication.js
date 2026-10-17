import {applyCredentialChange,beginCredentialChange,credentialProof,verifyAndCompleteCredentialChange} from './account-credential-change.js';
import {requirePasswordMigrationAdmission} from './authentication-rollout.js';
import { normalizeEmail } from './email.js';
import { AuthenticationError, supabaseAuthenticationConfig, signInWithSupabasePassword, createSupabaseMigrationUser, getSupabaseMigrationUser, preconfirmProvisionedAccount,validateNewPassword,setSupabasePasswordAfterApprovedProof } from './supabase-auth.js';
import { loadAuthenticationStatus, reserveAuthenticationMigration, finishAuthenticationMigration, finishNewAccountAuthentication } from './authentication-migration.js';

function denied() { return new AuthenticationError('invalid-credentials', 'Sign-in could not be completed.', 401); }

// Existing accounts only. This never creates a local account, claims a roster or
// changes roster ownership. Provider preconfirmation preserves existing account
// trust; it does not establish mailbox ownership.
export async function authenticateExistingAccount(env, input,{deferProvisioned=false}={}) {
  const config = supabaseAuthenticationConfig(env), db = env.ROSTER_DB;
  const email = normalizeEmail(input.email), password = input.password;
  if (!email || email.length > 254 || typeof password !== 'string' || !password || new TextEncoder().encode(password).length > 1024) throw denied();
  const status=await loadAuthenticationStatus(db,email);if(status.deletionQueued)throw denied();
  const link = status.link;
  if(link?.credential_operation_id)return verifyAndCompleteCredentialChange(env,link,input.newPassword||password,{resume:true});
  if (link?.phase === 'ready') return credentialProof(await signInWithSupabasePassword(env, { email, password, expectedSubject: link.provider_subject }),link);
  if (link?.phase === 'pending' && link.credential_revision === 'new-account') {
    let verified;
    try{verified=await signInWithSupabasePassword(env,{email,password,expectedSubject:link.provider_subject});}
    catch(error){if(error.code!=='email-verification-required')throw error;await preconfirmProvisionedAccount(env,{email,subject:link.provider_subject});verified=await signInWithSupabasePassword(env,{email,password,expectedSubject:link.provider_subject});}
    const proof=credentialProof(verified,link);
    await finishNewAccountAuthentication(db,{link,proof});return proof;
  }
  // Admission does not apply to ready provider accounts or new signup. Pausing
  // migration must never disable the already-migrated login path.
  requirePasswordMigrationAdmission(config,email);
  const stored = await db.prepare('SELECT email,password_hash,password_salt FROM account_profiles WHERE email=?').bind(email).first();
  if (!stored || !/^[a-f0-9]{64}$/i.test(stored.password_hash || '') || !/^[a-f0-9]{32}$/i.test(stored.password_salt || '')) throw denied();
  const hash = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(stored.password_salt + ':' + password))), x => x.toString(16).padStart(2, '0')).join('');
  if (hash !== stored.password_hash) throw denied();
  const record = { email, passwordHash: stored.password_hash, passwordSalt: stored.password_salt };
  const reservation = await reserveAuthenticationMigration(db, record);
  const upgradeRequired=()=>new AuthenticationError('auth-password-upgrade-required','Your current password needs an upgrade. Enter a new password of at least 12 characters to continue.',409);
  if(input.newPassword){
    validateNewPassword(input.newPassword);
    const existing=reservation.newlyReserved?null:await getSupabaseMigrationUser(env,{subject:reservation.provider_subject,email,allowMissing:true});
    if(!existing)await createSupabaseMigrationUser(env,{subject:reservation.provider_subject,email,password:input.newPassword,preconfirm:true});
    if(existing&&!existing.emailVerified)await preconfirmProvisionedAccount(env,{email,subject:reservation.provider_subject});
    const operation=await beginCredentialChange(env,reservation,input.newPassword);
    if(!existing){
      if(deferProvisioned)return{prepared:true};
      return verifyAndCompleteCredentialChange(env,operation,input.newPassword);
    }
    return applyCredentialChange(env,operation,input.newPassword,()=>setSupabasePasswordAfterApprovedProof(env,{email,subject:reservation.provider_subject,password:input.newPassword}));
  }
  if(new TextEncoder().encode(password).length>72)throw upgradeRequired();
  // A retry first looks up its reserved UUID. It never searches by email then
  // resets somebody else's provider password to resolve a conflict.
  let provider = reservation.newlyReserved?null:await getSupabaseMigrationUser(env, { subject: reservation.provider_subject, email, allowMissing: true });
  if (!provider){try{provider = await createSupabaseMigrationUser(env, { subject: reservation.provider_subject, email, password, preconfirm: true });}
    catch(error){if(error.code==='auth-password-policy')throw upgradeRequired();throw error;}}
  if (provider.disabled) throw new AuthenticationError('auth-account-unavailable', 'This account is unavailable. Contact the administrator.', 403);
  if (!provider.emailVerified) await preconfirmProvisionedAccount(env,{email,subject:reservation.provider_subject});
  const proof = credentialProof(await signInWithSupabasePassword(env, { email, password, expectedSubject: reservation.provider_subject }),reservation);
  await finishAuthenticationMigration(db, { reservation, record, proof });
  return proof;
}
