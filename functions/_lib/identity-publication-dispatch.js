import {guardedFetch} from './outbound-network.js';
export function requestIdentityPublication(context,operationId){
 return guardedFetch(context.env,new URL('/api/automation/identity-maintenance',context.request.url),{method:'POST',headers:{Authorization:`Bearer ${String(context.env.IDENTITY_MAINTENANCE_TOKEN||context.env.ROSTER_AUTOMATION_TOKEN||'')}`,'Content-Type':'application/json'},body:JSON.stringify({mode:'publish',operationId})},{label:'Identity publication maintenance'});
}
