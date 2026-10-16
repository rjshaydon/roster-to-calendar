import { deleteSupabaseAccount } from './supabase-auth.js';
import { normalizeEmail } from './email.js';

export async function finishAccountAuthenticationDeletion(env,email) {
  const db=env.ROSTER_DB;
  const job=await db.prepare(`SELECT provider_subject,email FROM account_auth_deletions WHERE email=? AND phase='queued' LIMIT 1`).bind(normalizeEmail(email)).first();
  if(!job)return{pending:false};
  const exists=await db.prepare('SELECT email FROM account_profiles WHERE email=?').bind(job.email).first();
  if(exists)return{pending:true}; // Never delete a provider while its app account exists.
  const attempt=await db.prepare(`UPDATE account_auth_deletions SET attempts=attempts+1,updated_at=? WHERE provider_subject=? AND phase='queued' AND attempts<5`).bind(new Date().toISOString(),job.provider_subject).run();
  if(Number(attempt.meta?.changes)!==1)return{pending:true};
  try{
    await deleteSupabaseAccount(env,{email:job.email,subject:job.provider_subject});
    await db.prepare(`UPDATE account_auth_deletions SET phase='complete',updated_at=? WHERE provider_subject=? AND phase='queued'`).bind(new Date().toISOString(),job.provider_subject).run();
    return{pending:false};
  }catch{return{pending:true};} // The durable exact-subject job survives; no blind retries.
}
