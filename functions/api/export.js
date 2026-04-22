import { buildRosterView, doctorOptions, exportIcs, parseUploadForm } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources, doctorKey, doctorDisplay, settings, overrides } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const doctors = doctorOptions(sources.mmc?.workbook, sources.ddh?.workbook);
    const selectedDoctor = doctors.find((doctor) => doctor.key === doctorKey);
    if (!selectedDoctor) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const events = buildRosterView(sources.mmc?.workbook, sources.ddh?.workbook, doctorKey, settings, overrides).events;
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
