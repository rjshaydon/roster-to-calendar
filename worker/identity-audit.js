import { runScheduledIdentityAudit } from "../functions/_lib/d1-calendar.js";

export default {
  async scheduled(controller, env, ctx) {
    const scheduledAt = new Date(controller.scheduledTime || Date.now());
    ctx.waitUntil(runScheduledIdentityAudit(env.ROSTER_DB, {
      now: scheduledAt.toISOString(),
      owner: `scheduled:${controller.scheduledTime || Date.now()}`,
      allowNew: isMelbourneWeeklyAuditWindow(scheduledAt),
    }).then((result) => console.log(JSON.stringify({ event: "identity-audit", ...result }))));
  },

  async fetch(request) {
    if (new URL(request.url).pathname !== "/health") return new Response("Not found", { status: 404 });
    return Response.json({ ok: true, service: "identity-audit" });
  },
};

export function isMelbourneWeeklyAuditWindow(date) {
  const parts = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Melbourne", weekday: "short", hour: "2-digit", minute: "2-digit", hourCycle: "h23",
  }).formatToParts(date);
  const value = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return value.weekday === "Sun" && value.hour === "03" && value.minute === "30";
}
