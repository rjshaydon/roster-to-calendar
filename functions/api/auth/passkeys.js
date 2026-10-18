import { AuthenticationError, providerRequest } from '../../_lib/supabase-auth.js';
import { boundedBody, admit } from '../../_lib/authentication-request.js';
import { readAccountSession, requireSameOrigin, requireExpectedAccount, requireRecentAuthentication, issueAccountSession, hasSessionCookie, revokeAccountSession } from '../../_lib/account-sessions.js';
import { passkeysEnabled, readPasskeyCookie, sealPasskeyCookie, clearPasskeyCookie, passkeyAccount, registrationProof, grantMatches, validatePasskeyOptions, validateCredential, listAccountPasskeys, removeAccountPasskey, verifyPasskeySignIn } from '../../_lib/account-passkeys.js';

function reply(body,status=200,cookie) {
  return Response.json(body,{status,headers:{'Cache-Control':'no-store','Vary':'Cookie',...(cookie?{'Set-Cookie':cookie}:{})}});
}
const invalid = () => new AuthenticationError('auth-passkey-expired','Please start passkey setup or sign-in again.',400);
export async function onRequestPost({env,request}) {
  if(!passkeysEnabled(env,request))return reply({code:'auth-passkey-disabled',error:'Passkeys are not enabled here. Use your password to sign in.'},503);
  try {
    requireSameOrigin(request);
    const body=await boundedBody(request,32768),db=env.ROSTER_DB;
    if(body.action==='startSignIn') {
      // Bound anonymous challenge work per source address using the existing
      // private limiter. No identity directory or roster query is involved.
      await admit(env,String(request.headers.get('CF-Connecting-IP')||'unknown'),'passkey');
      const data=validatePasskeyOptions(await providerRequest(env,'passkeys/authentication/options',{body:{}}));
      return reply({ok:true,...data},200,await sealPasskeyCookie(env,{kind:'authentication',challengeId:data.challengeId,persistent:body.remember===true}));
    }
    if(body.action==='finishSignIn') {
      const ceremony=await readPasskeyCookie(env,request);
      if(ceremony.kind!=='authentication'||body.challengeId!==ceremony.challengeId)throw invalid();
      await admit(env,String(request.headers.get('CF-Connecting-IP')||'unknown'),'passkey');
      const proof=await verifyPasskeySignIn(env,ceremony.challengeId,body.credential);
      const issued=await issueAccountSession(db,{...proof,persistent:ceremony.persistent});
      if(hasSessionCookie(request)){
        let previous;try{previous=await readAccountSession(db,request);}catch(error){if(error.code!=='auth-session-expired')throw error;}
        if(previous)await revokeAccountSession(db,previous);
      }
      const response=reply({ok:true,email:issued.email},200,issued.setCookie);
      response.headers.append('Set-Cookie',clearPasskeyCookie());return response;
    }
    const session=await readAccountSession(db,request);requireExpectedAccount(session,body.expectedAccount);
    const link=await passkeyAccount(db,session);
    if(body.action==='status') {
      const preference=await db.prepare('SELECT passkey_reminder_dismissed FROM account_profiles WHERE email=?').bind(session.email).first();
      // Suppressed reminders need no provider call at all. Account settings
      // explicitly request the list; no background polling or roster reads.
      const suppressed=preference?.passkey_reminder_dismissed===1;
      if(body.prompt===true&&suppressed)return reply({ok:true,remind:false,passkeys:[]});
      const passkeys=await listAccountPasskeys(env,link.provider_subject);
      return reply({ok:true,remind:!suppressed&&passkeys.length===0,remindersEnabled:!suppressed,passkeys});
    }
    if(body.action==='reminders') {
      if(typeof body.enabled!=='boolean')throw invalid();
      const result=await db.prepare(`UPDATE account_profiles SET passkey_reminder_dismissed=? WHERE email=? AND passkey_reminder_dismissed<>?`)
        .bind(body.enabled?0:1,session.email,body.enabled?0:1).run();
      return reply({ok:true,remindersEnabled:body.enabled,changed:Number(result.meta?.changes)>0});
    }
    if(body.action==='startRegistration') {
      await admit(env,session.email);
      const accessToken=await registrationProof(env,request,session,link,body.password);
      const data=validatePasskeyOptions(await providerRequest(env,'passkeys/registration/options',{accessToken,body:{}}),true);
      return reply({ok:true,...data},200,await sealPasskeyCookie(env,{kind:'registration',challengeId:data.challengeId,
        email:session.email,subject:link.provider_subject,version:link.credential_version,session:session.token_hash,accessToken}));
    }
    if(body.action==='finishRegistration') {
      const ceremony=await readPasskeyCookie(env,request);
      if(ceremony.kind!=='registration'||body.challengeId!==ceremony.challengeId||!grantMatches(ceremony,session,link))throw invalid();
      validateCredential(body.credential);
      const data=await providerRequest(env,'passkeys/registration/verify',{accessToken:ceremony.accessToken,body:{challenge_id:ceremony.challengeId,credential:body.credential}});
      if(typeof data?.id!=='string')throw invalid();
      // A logout/password reset while the provider was verifying cannot leave
      // a credential enrolled by a session which has just been revoked.
      try { const current=await readAccountSession(db,request);const currentLink=await passkeyAccount(db,current);
        if(!grantMatches(ceremony,current,currentLink))throw invalid();
      }catch(error){await removeAccountPasskey(env,link.provider_subject,data.id);throw error;}
      return reply({ok:true},200,clearPasskeyCookie());
    }
    if(body.action==='remove') {
      requireRecentAuthentication(session);
      const passkeys=await listAccountPasskeys(env,link.provider_subject);
      if(!passkeys.some(item=>item.id===body.id))throw invalid();
      await removeAccountPasskey(env,link.provider_subject,body.id);
      return reply({ok:true});
    }
    throw invalid();
  }catch(error){
    if(['d1-statement-budget-exceeded','d1-request-deadline-exceeded'].includes(error?.code))throw error;
    if(error instanceof AuthenticationError)return reply({error:error.message,code:error.code},error.status);
    return reply({code:'auth-passkey-unavailable',error:'Passkeys are temporarily unavailable. You can still use your password.'},503);
  }
}
