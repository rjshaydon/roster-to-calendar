import {
  buildRosterView,
  doctorOptions,
  parseUploadForm,
  previewSummary,
  serializeConflict,
  serializeEvent,
  serializeReviewItem,
  sourceNames,
} from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources, doctorKey, settings, overrides, conflictSelections } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const validDoctors = doctorOptions(sources.mmc, sources.ddh).map((doctor) => doctor.key);
    if (!validDoctors.includes(doctorKey)) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const view = buildRosterView(sources.mmc, sources.ddh, doctorKey, settings, overrides, conflictSelections);
    const events = view.events;
    return Response.json({
      ...previewSummary(events),
      events: events.map(serializeEvent),
      review: view.reviewItems.map(serializeReviewItem),
      issues: view.issues,
      conflicts: view.conflicts.map(serializeConflict),
      imports: view.imports,
      sources: sourceNames(sources),
      lastParsed: new Date().toISOString(),
    });
  } catch (error) {
    return Response.json({ error: error.message || "Unexpected server error." }, { status: 400 });
  }
}
