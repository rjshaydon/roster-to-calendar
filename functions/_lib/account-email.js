import { normalizeEmail } from './email.js';
import { guardedFetch } from './outbound-network.js';
import { AuthenticationError, supabaseAuthenticationConfig, generateSupabaseEmailProof, verifySupabaseEmailProof, changeSupabasePassword, signInWithSupabasePassword, getSupabaseMigrationUser, createSupabaseMigrationUser } from './supabase-auth.js';
import { loadAuthenticationLink, reserveAuthenticationMigration, finishAuthenticationMigration, finishNewAccountAuthentication } from './authentication-migration.js';

async function digest(value) { return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value))), x=>x.toString(16).padStart(2,'0')).join(''); }
function deliveryConfig(env) {
  const config=supabaseAuthenticationConfig(env);
  if(config.emailDeliveryReady!==true || !Number.isInteger(config.emailMonthlyLimit) || config.emailMonthlyLimit<1 || config.emailMonthlyLimit>100 || !env.POSTMARK_API_TOKEN) throw new AuthenticationError('auth-email-unavailable','Email verification is currently unavailable. Please contact the administrator.');
  return config;
}

export async function reserveAccountEmailCapacity(env) {
  const config=deliveryConfig(env);
  const reserved=await env.ROSTER_DB.prepare(`INSERT INTO account_auth_email_budget(month,reserved) VALUES(?,1)
    ON CONFLICT(month) DO UPDATE SET reserved=reserved+1 WHERE reserved<?`).bind(new Date().toISOString().slice(0,7),config.emailMonthlyLimit).run();
  if(Number(reserved.meta?.changes)!==1)throw new AuthenticationError('auth-email-unavailable','Email verification is currently unavailable. Please contact the administrator.');
}

export async function sendAccountEmailProof(env,{email,purpose,origin}) {
  const config=deliveryConfig(env),db=env.ROSTER_DB,now=Date.now();email=normalizeEmail(email);
  if(!['confirm','recovery'].includes(purpose)||!email||email.length>254||/[\s\x00-\x1f\x7f]/.test(email))throw new AuthenticationError('auth-request','Check your email address.',400);
  // Server supplies this origin after an exact-origin request check. Never use a
  // redirect, URL or recipient supplied in an arbitrary request body.
  if(origin!=='https://rtc.curiousmind.app' && !(config.testOrigin===origin && /^https:\/\/[a-z0-9-]+\.roster-to-calendar\.pages\.dev$/.test(origin)))throw new AuthenticationError('auth-origin','Open the app directly and try again.',403);
  const record=await db.prepare('SELECT email,password_hash,password_salt FROM account_profiles WHERE email=?').bind(email).first();
  if(!record)return; // Same public response whether an account exists or not.
  const recent=await db.prepare('SELECT id FROM account_email_challenges WHERE email=? AND purpose=? AND created_at>? LIMIT 1').bind(email,purpose,now-10*60000).first();if(recent)return;
  const month=new Date(now).toISOString().slice(0,7);
  const budget=await db.prepare('SELECT reserved FROM account_auth_email_budget WHERE month=?').bind(month).first();
  if(Number(budget?.reserved||0)>=config.emailMonthlyLimit)throw new AuthenticationError('auth-email-unavailable','Email verification is currently unavailable. Please contact the administrator.');
  const lock=await db.prepare(`INSERT INTO account_email_delivery_locks(email,purpose,next_at) VALUES(?,?,?)
    ON CONFLICT(email,purpose) DO UPDATE SET next_at=excluded.next_at WHERE next_at<=?`).bind(email,purpose,now+10*60000,now).run();
  if(Number(lock.meta?.changes)!==1)return;
  await reserveAccountEmailCapacity(env);
  let link=await loadAuthenticationLink(db,email);
  if(!link||link.phase!=='ready'){
    if(link?.credential_revision!=='new-account'&&(!Number.isFinite(Date.parse(config.legacyMigrationUntil||''))||Date.parse(config.legacyMigrationUntil)<=now))return;
    // Email recovery may reserve a UUID without the old password. It may create
    // only a random, undisclosed provider password. Email proof plus a new
    // provider password proof are still required before any credential retirement.
    if(link?.credential_revision!=='new-account')link=await reserveAuthenticationMigration(db,{email,passwordHash:record.password_hash,passwordSalt:record.password_salt});
    const user=await getSupabaseMigrationUser(env,{email,subject:link.provider_subject,allowMissing:true});
    if(!user)await createSupabaseMigrationUser(env,{email,subject:link.provider_subject,password:crypto.randomUUID()+crypto.randomUUID()});
  }
  const type=purpose==='recovery'&&link.phase==='ready'?'recovery':'magiclink';
  const proof=await generateSupabaseEmailProof(env,{email,subject:link.provider_subject,type}),id=crypto.randomUUID();
  // Do not reuse a link after an uncertain email send. A reservation is never
  // refunded automatically; this bounds both provider traffic and quota spend.
  await db.prepare(`INSERT INTO account_email_challenges(id,email,provider_subject,purpose,verification_type,token_digest,expires_at,phase,created_at)
    VALUES(?,?,?,?,?,?,?,'issued',?)`).bind(id,email,link.provider_subject,purpose,type,await digest(proof.tokenHash),now+30*60000,now).run();
  const url=origin+'/#'+new URLSearchParams({account:purpose,challenge:id,token:proof.tokenHash});
  const recovery=purpose==='recovery';
  const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),10000);
  try{
    const response=await guardedFetch(env,'https://api.postmarkapp.com/email',{method:'POST',redirect:'manual',signal:controller.signal,
      headers:{'Content-Type':'application/json','X-Postmark-Server-Token':env.POSTMARK_API_TOKEN},
      body:JSON.stringify({From:'Roster Converter <invites@curiousmind.app>',To:email,
        Subject:recovery?'Reset your Roster to Calendar password':'Confirm your Roster to Calendar email',
        TextBody:(recovery?'Choose a new password':'Confirm your email address')+' using this one-time link:\n'+url+'\n\nThis link expires in 30 minutes. If you did not request this, ignore this email.',
        MessageStream:'outbound',TrackOpens:false,TrackLinks:'None'})},{label:'Account verification email'});
    if(!response.ok)throw new AuthenticationError('auth-email-unavailable','The email could not be sent. Please try again later.');
    await response.body?.cancel();
  }catch(error){if(error instanceof AuthenticationError||error?.code==='LOCAL_OUTBOUND_BLOCKED')throw error;throw new AuthenticationError('auth-email-unavailable','The email could not be sent. Please try again later.');}
  finally{clearTimeout(timeout);}
}

export async function completeAccountEmailProof(env,{id,token,newPassword}) {
  deliveryConfig(env);const db=env.ROSTER_DB,now=Date.now();
  if(!/^[a-f0-9-]{36}$/i.test(id||'')||!/^[A-Za-z0-9_-]{32,256}$/.test(token||''))throw new AuthenticationError('auth-email-invalid','This link is invalid or has expired. Request a new one.',400);
  const challenge=await db.prepare(`SELECT * FROM account_email_challenges WHERE id=? AND token_digest=? AND expires_at>? AND phase='issued'`)
    .bind(id,await digest(token),now).first();
  if(!challenge)throw new AuthenticationError('auth-email-invalid','This link is invalid or has expired. Request a new one.',400);
  if(challenge.purpose==='recovery'&&(typeof newPassword!=='string'||newPassword.length<12||new TextEncoder().encode(newPassword).length>1024))throw new AuthenticationError('auth-password','Use a password of at least 12 characters.',400);
  const link=await loadAuthenticationLink(db,challenge.email);
  if(!link||link.provider_subject!==challenge.provider_subject)throw new AuthenticationError('auth-email-invalid','This link is no longer valid. Request a new one.',409);
  const claim=await db.prepare(`UPDATE account_email_challenges SET phase='using' WHERE id=? AND phase='issued' AND expires_at>?`).bind(id,now).run();
  if(Number(claim.meta?.changes)!==1)throw new AuthenticationError('auth-email-invalid','This link has already been used. Request a new one.',409);
  const proof=await verifySupabaseEmailProof(env,{email:challenge.email,subject:link.provider_subject,tokenHash:token,type:challenge.verification_type});
  if(challenge.purpose==='recovery'){
    // Revoke first, including when provider mutation has an uncertain outcome.
    await db.prepare('DELETE FROM account_sessions WHERE email=?').bind(challenge.email).run();
    await changeSupabasePassword(env,{email:challenge.email,subject:link.provider_subject,password:newPassword,accessToken:proof.accessToken});
    const verified=await signInWithSupabasePassword(env,{email:challenge.email,password:newPassword,expectedSubject:link.provider_subject});
    if(link.phase==='pending'&&link.credential_revision==='new-account')await finishNewAccountAuthentication(db,{link,proof:verified});
    else if(link.phase==='pending'){
      const record=await db.prepare('SELECT email,password_hash,password_salt FROM account_profiles WHERE email=?').bind(challenge.email).first();
      await finishAuthenticationMigration(db,{reservation:link,record:{email:record?.email,passwordHash:record?.password_hash,passwordSalt:record?.password_salt},proof:verified});
    }
  }
  await db.prepare(`UPDATE account_email_challenges SET phase='complete' WHERE id=? AND phase='using'`).bind(id).run();
  // Email proofs never become an application session. Explicit password sign-in
  // follows, so GET/scanner activity cannot silently sign in as another person.
  return{ok:true,email:challenge.email,purpose:challenge.purpose};
}
