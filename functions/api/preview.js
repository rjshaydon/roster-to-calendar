import {
  buildRosterView,
  doctorOptions,
  parseUploadForm,
  previewSummary,
  serializeEvent,
  serializeReviewItem,
  sourceNames,
} from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources, doctorKey, settings, overrides } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const validDoctors = doctorOptions(sources.mmc?.workbook, sources.ddh?.workbook).map((doctor) => doctor.key);
    if (!validDoctors.includes(doctorKey)) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const view = buildRosterView(sources.mmc?.workbook, sources.ddh?.workbook, doctorKey, settings, overrides);
    const events = view.events;
    return Response.json({
      ...previewSummary(events),
      events: events.map(serializeEvent),
      review: view.reviewItems.map(serializeReviewItem),
      issues: view.issues,
      sources: sourceNames(sources),
      lastParsed: new Date().toISOString(),
    });
  } catch (error) {
    return Response.json({ error: error.message || "Unexpected server error." }, { status: 400 });
  }
}
