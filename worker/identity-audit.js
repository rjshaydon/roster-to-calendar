export default {
  async scheduled(controller) {
    // Deliberately inert: the first implementation resumed paused scans every
    // hour and could exhaust the Free-tier D1 read allowance. Scheduled audits
    // stay off until separately approved with measured row-read budgets.
    console.log(JSON.stringify({ event: "identity-audit", status: "disabled", scheduledAt: controller.scheduledTime || Date.now() }));
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
