import {credentialProof} from './account-credential-change.js';
import { normalizeEmail } from './email.js';
import { AuthenticationError, supabaseAuthenticationConfig, getSupabaseMigrationUser, createSupabaseMigrationUser, preconfirmProvisionedAccount, signInWithSupabasePassword,validateNewPassword } from './supabase-auth.js';
import { loadAuthenticationLink,loadAuthenticationStatus, finishNewAccountAuthentication } from './authentication-migration.js';

export async function registerAccount(env,{email,password,realName,origin,prepareOnly=false}) {
  const config=supabaseAuthenticationConfig(env),db=env.ROSTER_DB;email=normalizeEmail(email);realName=String(realName||'').trim();
  validateNewPassword(password);
  if(config.signupEnabled!==true)throw new AuthenticationError('auth-registration-unavailable','Account registration is currently unavailable. Please contact the administrator.');
  if(!email||email.length>254||!email.includes('@')||/[\s\x00-\x1f\x7f]/.test(email)||!realName||realName.length>120||typeof password!=='string'||password.length<12||new TextEncoder().encode(password).length>1024)throw new AuthenticationError('auth-request','Enter your name, email and a password of at least 12 characters.',400);
  const status=await loadAuthenticationStatus(db,email);
  if(status.deletionQueued)throw new AuthenticationError('auth-registration-unavailable','Account registration is currently unavailable. Please contact the administrator.');
  let link=status.link,freshSubject='';
  if(link&&!(link.phase==='pending'&&link.credential_revision==='new-account'))throw new AuthenticationError('auth-registration-conflict','If you already have an account, sign in or reset your password.',409);
  if(!link){
    // Independent creation admission bounds account writes without using any
    // email allowance. This is not a mail reservation.
    const signupLimit=config.signupMonthlyLimit ?? 1000;
    if(!Number.isInteger(signupLimit)||signupLimit<1||signupLimit>10000)throw new AuthenticationError('auth-registration-unavailable','Account registration is currently unavailable.');
    const reserve=await db.prepare(`INSERT INTO account_auth_email_budget(month,reserved) VALUES(?,1)
      ON CONFLICT(month) DO UPDATE SET reserved=reserved+1 WHERE reserved<?`).bind('signup:'+new Date().toISOString().slice(0,7),signupLimit).run();
    if(Number(reserve.meta?.changes)!==1)throw new AuthenticationError('auth-registration-unavailable','Account registration is currently unavailable. Please contact the administrator.');
    const subject=crypto.randomUUID(),subscription=crypto.randomUUID()+crypto.randomUUID(),now=new Date().toISOString();
    const inserted=await db.batch([
      db.prepare(`INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token,created_at,updated_at)
        VALUES(?,?,'user','','',?,?,?) ON CONFLICT(email) DO NOTHING`).bind(email,realName,subscription,now,now),
      db.prepare(`INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at)
        SELECT email,'supabase',?,'pending','new-account',?,? FROM account_profiles WHERE email=? AND subscription_token=?
        AND password_hash='' AND password_salt='' ON CONFLICT(email) DO NOTHING`).bind(subject,now,now,email,subscription),
    ]);
    // Only our nonce-bound insert can produce this positive result. A conflict
    // still reads/validates the existing reservation; never adopt it by email.
    link=Number(inserted[1]?.meta?.changes)>0
      ?{email,provider:'supabase',provider_subject:subject,phase:'pending',credential_revision:'new-account',credential_version:0,credential_operation_id:'',credential_target_digest:''}
      :await loadAuthenticationLink(db,email);
    if(!link||link.phase!=='pending'||link.credential_revision!=='new-account')throw new AuthenticationError('auth-registration-conflict','If you already have an account, sign in or reset your password.',409);
    if(link.provider_subject===subject)freshSubject=subject;
  }
  if(link.credential_operation_id)throw new AuthenticationError('auth-password-pending','Sign in with your new password to finish the password update.',409);
  let user=freshSubject?null:await getSupabaseMigrationUser(env,{subject:link.provider_subject,email,allowMissing:true});
  // A retry cannot reset an existing provider password, even if email is still
  // unconfirmed. Email recovery is the explicit way to choose another password.
  if(!user)user=await createSupabaseMigrationUser(env,{email,subject:link.provider_subject,password,preconfirm:true});
  if(!user.emailVerified)await preconfirmProvisionedAccount(env,{email,subject:link.provider_subject});
  // A browser registration provisions first, then uses the existing login
  // coordinator for actual password proof and cookie issuance. Pending is not
  // authority; do not retire/finalize anything or grant a session here.
  if(prepareOnly)return{prepared:true};
  const proof=credentialProof(await signInWithSupabasePassword(env,{email,password,expectedSubject:link.provider_subject}),link);
  await finishNewAccountAuthentication(db,{link,proof});
  return proof;
}
