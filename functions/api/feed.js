import {
  applyEventOverrides,
  buildRosterViewFromStoredImports,
  customEventsToEvents,
  exportIcs,
} from "../_lib/roster.js";
import {
  loadAccountBySubscriptionToken,
  prepareAccountResponse,
} from "./state.js";

export async function onRequestGet(context) {
  try {
    if (!context.env.ROSTER_STORE) {
      return new Response("Cloud storage is not configured.", { status: 503 });
    }

    const url = new URL(context.request.url);
    const token = String(url.searchParams.get("token") || "").trim();
    if (!token) {
      return new Response("Subscription token is required.", { status: 400 });
    }

    const record = await loadAccountBySubscriptionToken(context.env.ROSTER_STORE, token);
    if (!record) {
      return new Response("Subscription calendar was not found.", { status: 404 });
    }

    const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, record);
    if (!prepared.subscription?.enabled || !prepared.claims?.length) {
      return new Response("No claimed roster names are linked to this subscription.", { status: 404 });
    }

    const aliases = prepared.claims.map((claim) => ({
      key: claim.key,
      displayName: claim.displayName,
      sourceType: claim.sourceType,
    }));
    const doctorKey = aliases[0]?.key || "";
    if (!doctorKey) {
      return new Response("No claimed doctor selection is available for this subscription.", { status: 404 });
    }

    const session = prepared.state?.session && typeof prepared.state.session === "object" ? prepared.state.session : {};
    const settings = session.settings && typeof session.settings === "object" ? session.settings : {};
    const overrides = session.overrides && typeof session.overrides === "object" ? session.overrides : {};
    const conflictSelections = session.conflictSelections && typeof session.conflictSelections === "object" ? session.conflictSelections : {};
    const customEvents = Array.isArray(session.customEvents) ? session.customEvents : [];

    const rosterView = await buildRosterViewFromStoredImports(
      prepared.state?.imports || [],
      doctorKey,
      settings,
      overrides,
      conflictSelections,
      aliases,
    );
    const rosterEvents = applyEventOverrides(rosterView.events, overrides);
    const events = [...rosterEvents, ...customEventsToEvents(customEvents, settings)].sort((left, right) => {
      const leftDate = left.start.slice(0, 10);
      const rightDate = right.start.slice(0, 10);
      if (leftDate !== rightDate) return leftDate.localeCompare(rightDate);
      if (left.allDay !== right.allDay) return left.allDay ? -1 : 1;
      return left.title.localeCompare(right.title);
    });

    const displayName = prepared.realName || aliases[0]?.displayName || record.email;
    const ics = exportIcs(events, displayName);
    return new Response(ics, {
      status: 200,
      headers: {
        "Content-Type": "text/calendar; charset=utf-8",
        "Cache-Control": "private, max-age=300",
        "Content-Disposition": `inline; filename="${displayName.replace(/\//g, "-")} subscription.ics"`,
      },
    });
  } catch (error) {
    return new Response(error.message || "Subscription feed failed.", { status: 400 });
  }
}
