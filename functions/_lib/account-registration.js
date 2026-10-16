import { normalizeEmail } from './email.js';
import { AuthenticationError, supabaseAuthenticationConfig, getSupabaseMigrationUser, createSupabaseMigrationUser, preconfirmProvisionedAccount, signInWithSupabasePassword } from './supabase-auth.js';
import { loadAuthenticationLink, finishNewAccountAuthentication } from './authentication-migration.js';

export async function registerAccount(env,{email,password,realName,origin}) {
  const config=supabaseAuthenticationConfig(env),db=env.ROSTER_DB;email=normalizeEmail(email);realName=String(realName||'').trim();
  if(config.signupEnabled!==true)throw new AuthenticationError('auth-registration-unavailable','Account registration is currently unavailable. Please contact the administrator.');
  if(!email||email.length>254||!email.includes('@')||/[\s\x00-\x1f\x7f]/.test(email)||!realName||realName.length>120||typeof password!=='string'||password.length<12||new TextEncoder().encode(password).length>1024)throw new AuthenticationError('auth-request','Enter your name, email and a password of at least 12 characters.',400);
  if(await db.prepare("SELECT 1 FROM account_auth_deletions WHERE email=? AND phase='queued' LIMIT 1").bind(email).first())throw new AuthenticationError('auth-registration-unavailable','Account registration is currently unavailable. Please contact the administrator.');
  let link=await loadAuthenticationLink(db,email);
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
    await db.batch([
      db.prepare(`INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token,created_at,updated_at)
        VALUES(?,?,'user','','',?,?,?) ON CONFLICT(email) DO NOTHING`).bind(email,realName,subscription,now,now),
      db.prepare(`INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at)
        SELECT email,'supabase',?,'pending','new-account',?,? FROM account_profiles WHERE email=? AND subscription_token=?
        AND password_hash='' AND password_salt='' ON CONFLICT(email) DO NOTHING`).bind(subject,now,now,email,subscription),
    ]);
    link=await loadAuthenticationLink(db,email);
    if(!link||link.phase!=='pending'||link.credential_revision!=='new-account')throw new AuthenticationError('auth-registration-conflict','If you already have an account, sign in or reset your password.',409);
  }
  const user=await getSupabaseMigrationUser(env,{subject:link.provider_subject,email,allowMissing:true});
  // A retry cannot reset an existing provider password, even if email is still
  // unconfirmed. Email recovery is the explicit way to choose another password.
  if(!user)await createSupabaseMigrationUser(env,{email,subject:link.provider_subject,password,preconfirm:true});
  await preconfirmProvisionedAccount(env,{email,subject:link.provider_subject});
  const proof=await signInWithSupabasePassword(env,{email,password,expectedSubject:link.provider_subject});
  await finishNewAccountAuthentication(db,{link,proof});
  return proof;
}
