import {requirePasswordMigrationAdmission} from './authentication-rollout.js';
import { normalizeEmail } from './email.js';
import { AuthenticationError, supabaseAuthenticationConfig, signInWithSupabasePassword, createSupabaseMigrationUser, getSupabaseMigrationUser, preconfirmProvisionedAccount } from './supabase-auth.js';
import { loadAuthenticationLink, reserveAuthenticationMigration, finishAuthenticationMigration, finishNewAccountAuthentication } from './authentication-migration.js';

function denied() { return new AuthenticationError('invalid-credentials', 'Sign-in could not be completed.', 401); }

// Existing accounts only. This never creates a local account, claims a roster or
// changes roster ownership. Provider preconfirmation preserves existing account
// trust; it does not establish mailbox ownership.
export async function authenticateExistingAccount(env, input) {
  const config = supabaseAuthenticationConfig(env), db = env.ROSTER_DB;
  const email = normalizeEmail(input.email), password = input.password;
  if (!email || email.length > 254 || typeof password !== 'string' || !password || new TextEncoder().encode(password).length > 1024) throw denied();
  if(await db.prepare("SELECT 1 FROM account_auth_deletions WHERE email=? AND phase='queued' LIMIT 1").bind(email).first())throw denied();
  const link = await loadAuthenticationLink(db, email);
  if (link?.phase === 'ready') return signInWithSupabasePassword(env, { email, password, expectedSubject: link.provider_subject });
  if (link?.phase === 'pending' && link.credential_revision === 'new-account') {
    await preconfirmProvisionedAccount(env,{email,subject:link.provider_subject});
    const proof=await signInWithSupabasePassword(env,{email,password,expectedSubject:link.provider_subject});
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
  // A retry first looks up its reserved UUID. It never searches by email then
  // resets somebody else's provider password to resolve a conflict.
  let provider = await getSupabaseMigrationUser(env, { subject: reservation.provider_subject, email, allowMissing: true });
  if (!provider) provider = await createSupabaseMigrationUser(env, { subject: reservation.provider_subject, email, password, preconfirm: true });
  if (provider.disabled) throw new AuthenticationError('auth-account-unavailable', 'This account is unavailable. Contact the administrator.', 403);
  if (!provider.emailVerified) await preconfirmProvisionedAccount(env,{email,subject:reservation.provider_subject});
  const proof = await signInWithSupabasePassword(env, { email, password, expectedSubject: reservation.provider_subject });
  await finishAuthenticationMigration(db, { reservation, record, proof });
  return proof;
}
