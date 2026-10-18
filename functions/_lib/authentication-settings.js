// Credentials remain in the encrypted configuration. Rollout changes can use a
// separate, non-secret JSON binding without copying privileged keys again.
const booleans=new Set(['enabled','sessionsEnabled','migrationEnabled','migrationPaused','signupEnabled','emailDeliveryReady','passkeysEnabled','requestsEnabled','requestTransfersEnabled','requestsAdminOnly']);
const dates=new Set(['legacyStateUntil','legacyMigrationUntil']);
export function authenticationSettings(env={}){
 const config=JSON.parse(env.ROSTER_AUTH_CONFIG||'{}');
 if(!config||typeof config!=='object'||Array.isArray(config))throw Error('Invalid authentication configuration.');
 if(env.ROSTER_AUTH_ROLLOUT===undefined)return config;
 const rollout=JSON.parse(env.ROSTER_AUTH_ROLLOUT);
 if(!rollout||typeof rollout!=='object'||Array.isArray(rollout))throw Error('Invalid authentication rollout.');
 for(const [key,value]of Object.entries(rollout)){
  if(booleans.has(key)){if(typeof value!=='boolean')throw Error('Invalid authentication rollout.');}
  else if(dates.has(key)){if(typeof value!=='string'||!/^\d{4}-\d{2}-\d{2}T/.test(value)||!Number.isFinite(Date.parse(value)))throw Error('Invalid authentication rollout.');}
  else if(key==='migrationEmails'){if(!Array.isArray(value)||value.length>100||value.some(email=>typeof email!=='string'||!email.includes('@')||email.length>254||/[\s\x00-\x1f\x7f]/.test(email)))throw Error('Invalid authentication rollout.');}
  else throw Error('Invalid authentication rollout.');
 }
 return {...config,...rollout};
}
