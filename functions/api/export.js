import { buildRosterView, customEventsToEvents, doctorOptions, exportIcs, parseUploadForm } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources, doctorKey, doctorDisplay, settings, overrides, customEvents } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const doctors = doctorOptions(sources.mmc?.workbook, sources.ddh?.workbook);
    const selectedDoctor = doctors.find((doctor) => doctor.key === doctorKey);
    if (!selectedDoctor) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const rosterEvents = buildRosterView(sources.mmc?.workbook, sources.ddh?.workbook, doctorKey, settings, overrides).events;
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
