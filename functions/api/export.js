import { applyEventOverrides, buildRosterView, customEventsToEvents, doctorOptions, exportIcs, parseUploadForm } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources, doctorKey, doctorDisplay, settings, overrides, customEvents, conflictSelections, doctorAliases } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const doctors = doctorOptions(sources.mmc, sources.ddh);
    const requestedKeys = new Set([doctorKey, ...(doctorAliases || []).map((alias) => alias.key)].filter(Boolean));
    const selectedDoctor = doctors.find((doctor) => requestedKeys.has(doctor.key));
    if (!selectedDoctor) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const rosterEvents = applyEventOverrides(
      buildRosterView(sources.mmc, sources.ddh, doctorKey, settings, overrides, conflictSelections, doctorAliases).events,
      overrides,
    );
    const events = [...rosterEvents, ...customEventsToEvents(customEvents, settings)].sort((left, right) => {
      const leftDate = left.start.slice(0, 10);
      const rightDate = right.start.slice(0, 10);
      if (leftDate !== rightDate) return leftDate.localeCompare(rightDate);
      if (left.allDay !== right.allDay) return left.allDay ? -1 : 1;
      return left.title.localeCompare(right.title);
    });
    if (!events.length) {
      throw new Error("No calendar events were found for the selected doctor.");
    }

    const displayName = doctorDisplay || selectedDoctor.displayName;
    const ics = exportIcs(events, displayName);
    return new Response(ics, {
      status: 200,
      headers: {
        "Content-Type": "text/calendar; charset=utf-8",
        "Content-Disposition": `attachment; filename="${displayName.replace(/\//g, "-")} roster.ics"`,
      },
    });
  } catch (error) {
    return Response.json({ error: error.message || "Unexpected server error." }, { status: 400 });
  }
}
