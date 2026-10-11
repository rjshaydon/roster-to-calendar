const defaultFindMyShiftUrl = "https://roster-to-calendar.pages.dev/api/automation/findmyshift-check";

export default {
  async scheduled(_controller, env, ctx) {
    if (automationPaused(env)) {
      console.warn(JSON.stringify({ event: "roster-watchdog", status: "paused", reason: "d1-write-quota-protection" }));
      return;
    }
    ctx.waitUntil(checkFindMyShift(env));
    if(env.IDENTITY_REGISTRY_ENABLED==='true') ctx.waitUntil(checkIdentityAudit(env,'register'));
    // The endpoint checks its R2 checkpoint before D1: completed weeks are
    // cheap no-ops; an interrupted Sunday pass can resume on later days.
    if(env.IDENTITY_SCHEDULED_AUDIT_ENABLED==='true') ctx.waitUntil(checkIdentityAudit(env));
  },

  async fetch(request, env) {
    if (new URL(request.url).pathname === '/check' && request.method === 'POST') {
      const token = String(env.ROSTER_WATCHDOG_TOKEN || '');
      if (!token || request.headers.get('authorization') !== `Bearer ${token}`) return new Response('Unauthorized', { status: 401 });
      if (automationPaused(env)) return Response.json({ ok: true, status: 'paused' });
      return Response.json(await checkFindMyShift(env));
    }
    if (new URL(request.url).pathname !== "/health") return new Response("Not found", { status: 404 });
    return Response.json({
      ok: true,
      service: "roster-queue-watchdog",
      configured: Boolean(String(env.ROSTER_WATCHDOG_TOKEN || "").trim()),
      paused: automationPaused(env),
      intervalMinutes: 5,
      sourceId: 'dandenong-findmyshift',
    });
  },
};

async function checkIdentityAudit(env,mode='audit') {
  const url=new URL('/api/automation/identity-maintenance',env.FINDMYSHIFT_CHECK_URL||defaultFindMyShiftUrl);
  const response=await fetch(url,{method:'POST',headers:{authorization:`Bearer ${String(env.ROSTER_WATCHDOG_TOKEN||'')}`,'content-type':'application/json'},body:JSON.stringify({mode})});
  if(!response.ok) console.warn(JSON.stringify({event:'identity-audit',status:'deferred',http:response.status}));
}

function automationPaused(env = {}) {
  return !["1", "true", "yes", "on"].includes(String(env.ROSTER_AUTOMATION_ENABLED || "").trim().toLowerCase());
}

async function checkFindMyShift(env) {
  const token = String(env.ROSTER_WATCHDOG_TOKEN || "").trim();
  if (!token) return { ok: false, error: "ROSTER_WATCHDOG_TOKEN is not configured." };
  const response = await fetch(String(env.FINDMYSHIFT_CHECK_URL || defaultFindMyShiftUrl), {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', "User-Agent": "roster-queue-watchdog" },
    body: JSON.stringify({ poll: true }),
    signal: AbortSignal.timeout(60000),
  });
  const result = await response.json().catch(() => ({}));
  console.log(JSON.stringify({ event: "findmyshift-check", ok: response.ok, status: result.status || "" }));
  return { ok: response.ok, ...result };
}
