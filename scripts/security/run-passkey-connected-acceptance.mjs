// One disposable synthetic account through the public app APIs. No staff
// credentials, privileged keys, real-device keys or emails. Never print proofs.
import {generateKeyPairSync,randomBytes,createHash,sign} from 'node:crypto';
import {mkdir,writeFile,unlink} from 'node:fs/promises';
const base='https://rtc.curiousmind.app',directory=process.argv[2];
if(!directory?.startsWith('/tmp/'))throw Error('Provide a private /tmp evidence directory.');
const options=await(await fetch(base+'/api/auth/session?options=1')).json();
if(options.passkeys!==true)throw Error('Production passkeys are disabled; no account was created.');
await mkdir(directory,{recursive:true,mode:0o700});
const email='passkey-acceptance-'+randomBytes(6).toString('hex')+'@example.test',password=randomBytes(32).toString('base64url');
const privateFile=directory+'/fixture.json',cookies=new Map(),checks=[];
await writeFile(privateFile,JSON.stringify({email,password}),{mode:0o600});
const bytes=value=>Buffer.from(value,'base64url'),encoded=value=>Buffer.from(value).toString('base64url');
function length(major,n){if(n<24)return Buffer.from([(major<<5)|n]);if(n<256)return Buffer.from([(major<<5)|24,n]);const out=Buffer.alloc(3);out[0]=(major<<5)|25;out.writeUInt16BE(n,1);return out;}
function cbor(value){
 if(Buffer.isBuffer(value))return Buffer.concat([length(2,value.length),value]);
 if(typeof value==='string'){const b=Buffer.from(value);return Buffer.concat([length(3,b.length),b]);}
 if(typeof value==='number')return length(value>=0?0:1,value>=0?value:-1-value);
 if(value instanceof Map)return Buffer.concat([length(5,value.size),...[...value].flatMap(([k,v])=>[cbor(k),cbor(v)])]);
 throw Error('Unsupported synthetic CBOR value');
}
const hash=value=>createHash('sha256').update(value).digest();
function check(name,passed,extra={}){checks.push({name,passed:Boolean(passed),...extra});if(!passed)throw Error(name);}
async function post(path,body){
 const start=Date.now(),response=await fetch(base+path,{method:'POST',redirect:'error',signal:AbortSignal.timeout(30000),
  headers:{Origin:base,'Content-Type':'application/json',Cookie:[...cookies].map(([k,v])=>k+'='+v).join('; ')},body:JSON.stringify(body)});
 const set=response.headers.getSetCookie();
 for(const cookie of set){const pair=cookie.split(';')[0],at=pair.indexOf('='),name=pair.slice(0,at),value=pair.slice(at+1);if(!value||/Max-Age=0(?:;|$)/i.test(cookie))cookies.delete(name);else cookies.set(name,value);}
 return {response,data:await response.json().catch(()=>({})),wallMs:Date.now()-start,set};
}
const auth=body=>post('/api/auth/session',body),passkey=body=>post('/api/auth/passkeys',{...body,expectedAccount:email});
let attempted=false,cleaned=false;
try{
 attempted=true;let result=await auth({action:'signup',email,password,realName:'Passkey acceptance synthetic test'});
 check('password-signup',result.response.status===202&&result.data.code==='auth-account-prepared');
 result=await auth({action:'login',email,password,remember:true});check('password-login',result.response.status===200&&result.data.ok===true);
 check('opaque-session-and-encrypted-setup-grant',cookies.has('__Host-roster_session')&&cookies.has('__Host-roster_passkey')&&!JSON.stringify(result.data).includes(password));
 result=await passkey({action:'status',prompt:true});check('first-login-reminder',result.response.ok&&result.data.remind===true);
 result=await passkey({action:'reminders',enabled:false});check('disable-reminders',result.response.ok);
 result=await passkey({action:'status',prompt:true});check('suppressed-reminder',result.response.ok&&result.data.remind===false);
 result=await passkey({action:'reminders',enabled:true});check('re-enable-reminders',result.response.ok);
 result=await passkey({action:'startRegistration'});check('registration-challenge-without-repeat-password',result.response.ok&&Boolean(result.data.options));
 const creation=result.data,publicKey=creation.options;
 check('exact-production-rp',publicKey.rp?.id==='rtc.curiousmind.app');
 const {privateKey,publicKey:key}=generateKeyPairSync('ec',{namedCurve:'prime256v1'}),jwk=key.export({format:'jwk'}),credentialId=randomBytes(32);
 const cose=cbor(new Map([[1,2],[3,-7],[-1,1],[-2,bytes(jwk.x)],[-3,bytes(jwk.y)]]));
 const idLength=Buffer.alloc(2);idLength.writeUInt16BE(credentialId.length);
 const registrationAuth=Buffer.concat([hash(publicKey.rp.id),Buffer.from([0x45]),Buffer.alloc(4),Buffer.alloc(16),idLength,credentialId,cose]);
 const client=Buffer.from(JSON.stringify({type:'webauthn.create',challenge:publicKey.challenge,origin:base,crossOrigin:false}));
 const credential={id:encoded(credentialId),rawId:encoded(credentialId),type:'public-key',authenticatorAttachment:'platform',clientExtensionResults:{},response:{clientDataJSON:encoded(client),attestationObject:encoded(cbor(new Map([['fmt','none'],['attStmt',new Map()],['authData',registrationAuth]])))}};
 result=await passkey({action:'finishRegistration',challengeId:creation.challengeId,credential});check('provider-verified-enrollment',result.response.ok&&result.data.ok===true);
 result=await passkey({action:'status',prompt:true});check('registered-key-and-no-reminder',result.response.ok&&result.data.passkeys?.length===1&&result.data.remind===false);
 const keyId=result.data.passkeys[0].id;
 result=await auth({action:'logout',expectedAccount:email});check('password-session-logout',result.response.ok);
 result=await passkey({action:'startSignIn',remember:true});check('discoverable-signin-challenge',result.response.ok&&Boolean(result.data.options));
 const challenge=result.data,assertionClient=Buffer.from(JSON.stringify({type:'webauthn.get',challenge:challenge.options.challenge,origin:base,crossOrigin:false}));
 const counter=Buffer.alloc(4);counter.writeUInt32BE(1);const authData=Buffer.concat([hash('rtc.curiousmind.app'),Buffer.from([5]),counter]);
 const assertion={id:encoded(credentialId),rawId:encoded(credentialId),type:'public-key',authenticatorAttachment:'platform',clientExtensionResults:{},response:{clientDataJSON:encoded(assertionClient),authenticatorData:encoded(authData),signature:encoded(sign('sha256',Buffer.concat([authData,hash(assertionClient)]),privateKey)),userHandle:publicKey.user.id}};
 result=await passkey({action:'finishSignIn',challengeId:challenge.challengeId,credential:assertion});check('provider-verified-passkey-signin',result.response.ok&&result.data.email===email&&cookies.has('__Host-roster_session'));
 result=await post('/api/state',{action:'loadAccountContext',expectedAccount:email,responseMode:'fast'});check('ordinary-workspace-after-passkey',result.response.ok&&result.data.ok===true);
 result=await passkey({action:'remove',id:keyId});check('remove-own-passkey',result.response.ok);
 result=await passkey({action:'status'});check('provider-key-removed',result.response.ok&&result.data.passkeys?.length===0);
 result=await auth({action:'logout',expectedAccount:email});check('passkey-session-logout',result.response.ok);
 result=await auth({action:'login',email,password});check('password-fallback-unchanged',result.response.ok&&result.data.email===email);
}catch(error){checks.push({name:'stopped',passed:false,step:checks.at(-1)?.name||'initial-request'});}
finally{
 if(attempted){
  try{
   if(!cookies.has('__Host-roster_session'))await auth({action:'login',email,password});
   const result=await post('/api/state',{action:'deleteAccount',email,expectedAccount:email});
   cleaned=result.response.ok&&result.data.deletedEmail===email&&result.data.authenticationCleanupPending===false;
   checks.push({name:'exact-synthetic-account-and-provider-cleanup',passed:cleaned});if(cleaned)await unlink(privateFile);
  }catch{checks.push({name:'exact-synthetic-account-and-provider-cleanup',passed:false});}
 }
}
const report={at:new Date().toISOString(),syntheticOnly:true,staffAccountsTouched:false,realDeviceKeysUsed:false,emailsSent:false,cleaned,checks};
await writeFile(directory+'/result.json',JSON.stringify(report,null,2)+'\n',{mode:0o600});
console.log(JSON.stringify(report));if(!checks.every(x=>x.passed)||!cleaned)process.exitCode=1;
