import {applyCredentialChange,beginCredentialChange} from './account-credential-change.js';
import { AuthenticationError, createSupabaseMigrationUser, getSupabaseMigrationUser, setSupabasePasswordAfterApprovedProof, preconfirmProvisionedAccount,validateNewPassword } from './supabase-auth.js';
import { loadAuthenticationLink, reserveAuthenticationMigration } from './authentication-migration.js';

export async function acceptAccountInvitation(env,{inviteToken,newPassword}) {
  if(typeof inviteToken!=='string'||inviteToken.length<24||inviteToken.length>256||typeof newPassword!=='string'||newPassword.length<12||new TextEncoder().encode(newPassword).length>1024)throw new AuthenticationError('auth-invite','Use a valid invitation and a password of at least 12 characters.',400);
  validateNewPassword(newPassword);
  const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(inviteToken))),x=>x.toString(16).padStart(2,'0')).join('');
  const db=env.ROSTER_DB,now=new Date().toISOString();
  const invite=await db.prepare(`SELECT email FROM account_invites WHERE token_hash=? AND accepted_at='' AND revoked_at='' AND expires_at>?`).bind(hash,now).first();
  if(!invite)throw new AuthenticationError('auth-invite','This invitation has expired or has been used. Request a new one.',400);
  const record=await db.prepare('SELECT email,password_hash,password_salt FROM account_profiles WHERE email=?').bind(invite.email).first();
  if(!record)throw new AuthenticationError('auth-invite','The invited account no longer exists.',409);
  // Consuming proof precedes any provider credential mutation. A failure/lost
  // response cannot leave a reusable invitation that can reset the account.
  const claimed=await db.prepare(`UPDATE account_invites SET accepted_at=?,updated_at=? WHERE token_hash=? AND accepted_at='' AND revoked_at='' AND expires_at>?`).bind(now,now,hash,now).run();
  if(Number(claimed.meta?.changes)!==1)throw new AuthenticationError('auth-invite','This invitation has already been used. Request a new one.',409);
  let link=await loadAuthenticationLink(db,record.email);
  const legacy={email:record.email,passwordHash:record.password_hash,passwordSalt:record.password_salt};
  if(!link||link.phase!=='ready'&&link.credential_revision!=='new-account')link=await reserveAuthenticationMigration(db,legacy);
  const operation=await beginCredentialChange(env,link,newPassword);
  const current=await getSupabaseMigrationUser(env,{email:record.email,subject:link.provider_subject,allowMissing:true});
  if(!current)await createSupabaseMigrationUser(env,{email:record.email,subject:link.provider_subject,password:newPassword,preconfirm:true});
  await preconfirmProvisionedAccount(env,{email:record.email,subject:link.provider_subject});
  await applyCredentialChange(env,operation,newPassword,()=>setSupabasePasswordAfterApprovedProof(env,{email:record.email,subject:link.provider_subject,password:newPassword}));
  return{ok:true,email:record.email};
}
