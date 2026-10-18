// Binding-only service: no public route, credentials, passwords or database.
export default {async fetch(request,env){
  if(request.method!=='POST'||new URL(request.url).pathname!=='/check')return new Response(null,{status:404});
  const declared=Number(request.headers.get('content-length')||0);
  if(declared>100)return new Response(null,{status:413});
  const reader=request.body?.getReader();let text='';
  if(!reader)return new Response(null,{status:400});
  try{let bytes=0;const decoder=new TextDecoder();while(true){const n=await reader.read();if(n.done)break;bytes+=n.value.byteLength;if(bytes>100){await reader.cancel();return new Response(null,{status:413});}text+=decoder.decode(n.value,{stream:true});}text+=decoder.decode();}finally{reader.releaseLock();}
  if(!/^(login|passkey|support):[a-f0-9]{64}$/.test(text))return new Response(null,{status:400});
  // Password admission stays per account at five/minute. Anonymous passkey
  // ceremonies share a hospital's public IP, so use a separate bounded bucket.
  const limiter=text.startsWith('passkey:')?env.PASSKEY_RATE:text.startsWith('support:')?env.SUPPORT_RATE:env.LOGIN_RATE;
  if(!limiter?.limit)return new Response(null,{status:503});
  const {success}=await limiter.limit({key:text});
  return new Response(null,{status:success?204:429,headers:{'Cache-Control':'no-store'}});
}};
