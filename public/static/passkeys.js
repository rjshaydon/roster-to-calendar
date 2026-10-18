// Browser WebAuthn transport only. Supabase verifies the proof server-side;
// neither provider tokens nor private keys are stored by this module.
function bytes(value) {
  return Uint8Array.from(atob(value.replace(/-/g,'+').replace(/_/g,'/')),c=>c.charCodeAt(0));
}
function encoded(value) {
  return btoa(String.fromCharCode(...new Uint8Array(value))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
}
export function passkeySupported(scope=globalThis) {
  return scope.isSecureContext===true&&Boolean(scope.PublicKeyCredential&&scope.navigator?.credentials?.create&&scope.navigator?.credentials?.get);
}
export function passkeyOptions(options,registration) {
  if(registration&&typeof globalThis.PublicKeyCredential?.parseCreationOptionsFromJSON==='function')return PublicKeyCredential.parseCreationOptionsFromJSON(options);
  if(!registration&&typeof globalThis.PublicKeyCredential?.parseRequestOptionsFromJSON==='function')return PublicKeyCredential.parseRequestOptionsFromJSON(options);
  const result={...options,challenge:bytes(options.challenge)};
  if(registration){result.user={...options.user,id:bytes(options.user.id)};result.excludeCredentials=(options.excludeCredentials||[]).map(x=>({...x,id:bytes(x.id)}));}
  else result.allowCredentials=(options.allowCredentials||[]).map(x=>({...x,id:bytes(x.id)}));
  return result;
}
export function passkeyResponse(credential,registration) {
  if(!credential)throw Error('Passkey request cancelled.');
  if(typeof credential.toJSON==='function')return credential.toJSON();
  const response={clientDataJSON:encoded(credential.response.clientDataJSON)};
  if(registration)response.attestationObject=encoded(credential.response.attestationObject);
  else Object.assign(response,{authenticatorData:encoded(credential.response.authenticatorData),signature:encoded(credential.response.signature),userHandle:credential.response.userHandle?encoded(credential.response.userHandle):null});
  return {id:credential.id,rawId:encoded(credential.rawId),type:'public-key',authenticatorAttachment:credential.authenticatorAttachment,
    clientExtensionResults:credential.getClientExtensionResults?.()||{},response};
}
export function passkeyError(error) {
  if(['NotAllowedError','AbortError'].includes(error?.name))return 'Passkey request cancelled. You can still use your password.';
  if(error?.name==='InvalidStateError')return 'This passkey is already registered. You can use it to sign in.';
  return error?.message||'Could not complete the passkey request. You can still use your password.';
}
export function createPasskeyCeremony(credentials=globalThis.navigator?.credentials) {
  let controller;
  return {
    cancel(){controller?.abort();},
    async run(options,registration=false){
      controller?.abort();const active=new AbortController();controller=active;
      try{
        const publicKey=passkeyOptions(options,registration);
        const credential=await credentials[registration?'create':'get']({publicKey,signal:active.signal});
        return passkeyResponse(credential,registration);
      }finally{if(controller===active)controller=null;}
    },
  };
}
