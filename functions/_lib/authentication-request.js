import {AuthenticationError} from './supabase-auth.js';

export async function boundedBody(request, maximumBytes = 8192) {
  if (!request.headers.get('Content-Type')?.toLowerCase().startsWith('application/json')) {
    throw new AuthenticationError('auth-request', 'A JSON request is required.', 400);
  }
  const reader = request.body?.getReader();
  if (!reader) throw new AuthenticationError('auth-request', 'Invalid request.', 400);
  let length = 0; const chunks = [];
  try {
    while (true) {
      const next = await reader.read(); if (next.done) break;
      length += next.value.byteLength;
      if (length > maximumBytes) { await reader.cancel(); throw new AuthenticationError('auth-request', 'Request is too large.', 413); }
      chunks.push(next.value);
    }
    const bytes = new Uint8Array(length); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
    try { const body = JSON.parse(new TextDecoder().decode(bytes)); if (!body || Array.isArray(body) || typeof body !== 'object') throw Error(); return body; }
    catch { throw new AuthenticationError('auth-request', 'Invalid request.', 400); }
  } finally { reader.releaseLock(); }
}


export async function admit(env,identity,mode='login') {
  if(!['login','passkey'].includes(mode))throw new AuthenticationError('auth-request','Invalid sign-in request.',400);
  if (!env.AUTH_LOGIN_LIMITER?.limit && !env.AUTH_LOGIN_LIMITER?.fetch) throw new AuthenticationError('auth-unavailable','Sign-in is currently unavailable. Please try again later.');
  const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(identity));
  const key=Array.from(new Uint8Array(digest),x=>x.toString(16).padStart(2,'0')).join('');
  // A service binding synthesizes truthy RPC methods for arbitrary names.
  // Prefer its explicit fetch API; probing .limit first invokes a nonexistent
  // remote method instead of the service's native-limiter HTTP adapter.
  if(env.AUTH_LOGIN_LIMITER.fetch){
    const response=await env.AUTH_LOGIN_LIMITER.fetch('https://admission.internal/check',{method:'POST',body:mode+':'+key});
    if(response.status===204)return;
    if(response.status!==429)throw new AuthenticationError('auth-unavailable','Sign-in is currently unavailable. Please try again later.');
  }else if((await env.AUTH_LOGIN_LIMITER.limit({key:mode+':'+key}))?.success===true)return;
  throw new AuthenticationError('auth-rate-limited','Too many attempts. Please wait before trying again.',429);
}
