import {boundedBody,admit} from '../../_lib/authentication-request.js';
import {authenticationSettings} from '../../_lib/authentication-settings.js';
import {applyCredentialChange,beginCredentialChange} from '../../_lib/account-credential-change.js';
import {loadAuthenticationLink} from '../../_lib/authentication-migration.js';
import { acceptAccountInvitation } from '../../_lib/account-invitations.js';
import { finishAccountAuthenticationDeletion } from '../../_lib/account-auth-deletions.js';
import { registerAccount } from '../../_lib/account-registration.js';
import { sendAccountEmailProof, completeAccountEmailProof } from '../../_lib/account-email.js';
import { AuthenticationError, changeSupabasePassword, validateNewPassword } from '../../_lib/supabase-auth.js';
import { authenticateExistingAccount } from '../../_lib/account-authentication.js';
import { normalizeEmail } from '../../_lib/email.js';
import {
  accountSessionsEnabled, requireSameOrigin, requireExpectedAccount, issueAccountSession, readAccountSession,
  renewAccountSession, revokeAccountSession, revokeOtherAccountSessions, clearSessionCookie, hasSessionCookie, requireRecentAuthentication,
} from '../../_lib/account-sessions.js';

const headers = { 'Cache-Control': 'no-store', 'Vary': 'Cookie' };
function reply(body, status = 200, setCookie) {
  return Response.json(body, { status, headers: { ...headers, ...(setCookie ? { 'Set-Cookie': setCookie } : {}) } });
}

export async function onRequestPost(context) {
  if (!accountSessionsEnabled(context.env)) return reply({ error: 'The sign-in upgrade is not enabled.', code: 'auth-disabled' }, 503);
  try {
    requireSameOrigin(context.request);
    if (!context.env.ROSTER_DB) throw new AuthenticationError('auth-unavailable', 'Sign-in is currently unavailable. Please try again later.');
    const body = await boundedBody(context.request), action = String(body.action || '');
    const db = context.env.ROSTER_DB;
    if (action === 'acceptInvite') {
      if(typeof body.inviteToken!=='string'||body.inviteToken.length>256)throw new AuthenticationError('auth-invite','Invalid invitation.',400);
      await admit(context.env,body.inviteToken);
      return reply(await acceptAccountInvitation(context.env,{inviteToken:body.inviteToken,newPassword:body.newPassword}));
    }
    if (action === 'signup') {
      const email=normalizeEmail(body.email);if(!email||email.length>254)throw new AuthenticationError('auth-request','Check your email address.',400);
      await admit(context.env,email);
      await registerAccount(context.env,{email,password:body.password,realName:body.realName,origin:new URL(context.request.url).origin,prepareOnly:true});
      return reply({ok:false,code:'auth-account-prepared'},202);
    }
    if (action === 'requestRecovery') {
      const email=normalizeEmail(body.email);
      if(!email||email.length>254||!email.includes('@')||/[\s\x00-\x1f\x7f]/.test(email))throw new AuthenticationError('auth-request','Check your email address.',400);
      await admit(context.env,email);
      try { await sendAccountEmailProof(context.env,{email,purpose:'recovery',origin:new URL(context.request.url).origin}); }
      catch(error) { if(error?.code==='LOCAL_OUTBOUND_BLOCKED')throw error; console.warn('Account recovery email unavailable.'); }
      return reply({ok:true,message:'If this account can be recovered, a one-time link will be emailed to it.'});
    }
    if (action === 'completeEmail') {
      if(typeof body.id!=='string'||body.id.length!==36)throw new AuthenticationError('auth-request','Invalid verification link.',400);
      await admit(context.env,body.id);
      return reply(await completeAccountEmailProof(context.env,{id:body.id,token:body.token,newPassword:body.newPassword}));
    }
    if (action === 'login' || action === 'reauthenticate') {
      // No legacy fallback here. Accounts without a completed verified migration
      // keep their existing journey until that separate rollout is ready.
      const email = normalizeEmail(body.email);
      if (!email || email.length > 254 || typeof body.password !== 'string' || !body.password || body.password.length > 1024) {
        throw new AuthenticationError('invalid-credentials', 'Check your email and password.', 400);
      }
      await admit(context.env,email);
      let oldSession;
      if (action === 'reauthenticate') {
        oldSession = await readAccountSession(db, context.request);
        requireExpectedAccount(oldSession,body.expectedAccount);
        if (email !== oldSession.email) throw new AuthenticationError('auth-identity-mismatch', 'Sign in using the current account.', 403);
      }
      const proof = await authenticateExistingAccount(context.env, { email, password: body.password, newPassword: body.newPassword },{deferProvisioned:action==='login'});
      if(proof.prepared)return reply({ok:false,code:'auth-account-prepared'},202);
      const issued = await issueAccountSession(db, { email: proof.email, subject: proof.subject, credentialVersion: proof.credentialVersion, persistent: body.remember === true });
      if (!oldSession && hasSessionCookie(context.request)) {
        try { oldSession = await readAccountSession(db, context.request); }
        catch(error) { if(error?.code!=='auth-session-expired')throw error; }
      }
      if (oldSession) await revokeAccountSession(db, oldSession);
      const response=reply({ ok: true, email: issued.email }, 200, issued.setCookie);
      if(authenticationSettings(context.env).passkeysEnabled===true){
        const {passkeyLoginGrant}=await import('../../_lib/account-passkeys.js');
        try { const grant=await passkeyLoginGrant(context.env,context.request,proof,issued);if(grant)response.headers.append('Set-Cookie',grant); } catch { /* Optional enrollment must never fail a valid login. */ }
      }
      return response;
    }
    if (action === 'logout') {
      // An expired/invalid cookie may always be cleared locally, but cannot
      // select an account or revoke anybody else's sessions.
      if (hasSessionCookie(context.request)) {
        try { const current=await readAccountSession(db, context.request); requireExpectedAccount(current,body.expectedAccount); await revokeAccountSession(db,current); }
        catch (error) { if (error.code !== 'auth-session-expired') throw error; }
      }
      return reply({ ok: true }, 200, clearSessionCookie());
    }
    const session = await readAccountSession(db, context.request);
    requireExpectedAccount(session,body.expectedAccount);
    if (action === 'retryAuthenticationDeletion') {
      if(!['creator','owner'].includes(session.role))throw new AuthenticationError('auth-forbidden','Creator access is required.',403);
      requireRecentAuthentication(session);
      const email=normalizeEmail(body.email);if(!email||email.length>254)throw new AuthenticationError('auth-request','Invalid account.',400);
      return reply({ok:true,...await finishAccountAuthenticationDeletion(context.env,email)});
    }
    if (action === 'changePassword') {
      await admit(context.env,session.email);
      validateNewPassword(body.newPassword);
      const proof=await authenticateExistingAccount(context.env,{email:session.email,password:body.currentPassword});
      const link=await loadAuthenticationLink(db,session.email);
      if(link?.credential_version!==proof.credentialVersion)throw new AuthenticationError('auth-session-expired','Please sign in again.',401);
      const operation=await beginCredentialChange(context.env,link,body.newPassword);
      let changed;
      try{changed=await applyCredentialChange(context.env,operation,body.newPassword,
        ()=>changeSupabasePassword(context.env,{email:session.email,subject:proof.subject,password:body.newPassword,accessToken:proof.accessToken}),
        body.currentPassword!==body.newPassword?{differentCurrentProof:proof}:{});}
      catch(error){
        if(error.credentialChangeNotApplied===true){
          // The original password proof remains valid only at this operation's
          // exact cancelled version, never at a later competing change.
          const retained=await issueAccountSession(db,{email:session.email,subject:proof.subject,credentialVersion:operation.credential_version+1,persistent:session.persistent===1});
          return reply({error:error.message,code:error.code},error.status||503,retained.setCookie);
        }
        throw error;
      }
      const issued=await issueAccountSession(db,{email:session.email,subject:changed.subject,credentialVersion:changed.credentialVersion,persistent:session.persistent===1});
      return reply({ok:true,email:session.email},200,issued.setCookie);
    }
    if(action==='migrationAccounts'){
      if(!['creator','owner'].includes(session.role))throw new AuthenticationError('auth-forbidden','Creator access is required.',403);
      const cursor=body.cursor??'';
      if(typeof cursor!=='string'||cursor.length>254)throw new AuthenticationError('auth-request','Invalid account cursor.',400);
      const page=await db.prepare(`SELECT a.email,a.real_name AS name,
        CASE WHEN l.phase='ready' AND a.password_hash='' AND a.password_salt='' THEN 'migrated'
          WHEN l.phase='pending' THEN 'pending' ELSE 'unmigrated' END AS status
        FROM account_profiles a LEFT JOIN account_authentication_links l ON l.email=a.email
        WHERE a.email>? ORDER BY a.email LIMIT 101`).bind(cursor).all();
      const accounts=page.results.slice(0,100);
      return reply({ok:true,accounts,nextCursor:page.results.length>100?accounts.at(-1).email:''});
    }
    if(action==='migrationSummary'){
      if(!['creator','owner'].includes(session.role))throw new AuthenticationError('auth-forbidden','Creator access is required.',403);
      const counts=await db.prepare(`SELECT count(*) AS total,
        coalesce(sum(CASE WHEN a.password_hash='' AND a.password_salt='' AND EXISTS
          (SELECT 1 FROM account_authentication_links l WHERE l.email=a.email AND l.phase='ready') THEN 1 ELSE 0 END),0) AS migrated
        FROM account_profiles a`).first();
      return reply({ok:true,total:Number(counts.total),migrated:Number(counts.migrated)});
    }
    if (action === 'status') return reply({ ok: true, email: session.email, role: session.role });
    if (action === 'activity') {
      const renewed = await renewAccountSession(db, context.request, session);
      return reply({ ok: true }, 200, renewed.setCookie);
    }
    if (action === 'signOutOtherDevices') {
      const revokedSessions=await revokeOtherAccountSessions(db, session);
      return reply({ ok: true, revokedSessions });
    }
    throw new AuthenticationError('auth-request', 'Unknown account action.', 400);
  } catch (error) {
    if (['d1-statement-budget-exceeded', 'd1-request-deadline-exceeded'].includes(error?.code)) throw error;
    if (error instanceof AuthenticationError) return reply({ error: error.message, code: error.code }, error.status);
    // Never reflect provider, database or request error details into auth replies.
    return reply({ error: 'Sign-in is currently unavailable. Please try again later.', code: 'auth-unavailable' }, 503);
  }
}

export function onRequestGet(context) {
  if(new URL(context.request.url).searchParams.get('options')==='1'){
    let config;try{config=authenticationSettings(context.env);}catch{config={};}
    const enabled=accountSessionsEnabled(context.env);return reply({enabled,cohortRollout:enabled&&Array.isArray(config.migrationEmails),passkeys:enabled&&config.passkeysEnabled===true&&new URL(context.request.url).origin==='https://rtc.curiousmind.app',recovery:enabled&&config.emailDeliveryReady===true&&Boolean(context.env.POSTMARK_API_TOKEN)});
  }
  return reply({ error: 'Use POST for account requests.' }, 405);
}
