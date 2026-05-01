import { loadAccountBySubscriptionToken, prepareAccountResponse } from "./state.js";

export async function onRequestGet(context) {
  try {
    if (!context.env.ROSTER_STORE) {
      return new Response("Cloud storage is not configured.", { status: 503 });
    }

    const url = new URL(context.request.url);
    const token = String(url.searchParams.get("token") || "").trim();
    const view = String(url.searchParams.get("view") || "full").trim() === "filtered" ? "filtered" : "full";
    if (!token) {
      return new Response("Subscription token is required.", { status: 400 });
    }

    const record = await loadAccountBySubscriptionToken(context.env.ROSTER_STORE, token);
    if (!record) {
      return new Response("Subscription calendar was not found.", { status: 404 });
    }

    const prepared = await prepareAccountResponse(context.env.ROSTER_STORE, record);
    const artifact = prepared.state?.subscriptionFeeds?.[view] || null;
    if (!prepared.subscription?.enabled || !artifact?.ics) {
      return new Response("No stored subscription calendar is available for this view.", { status: 404 });
    }

    const displayName = artifact.doctorDisplay || prepared.realName || record.email;
    return new Response(artifact.ics, {
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
