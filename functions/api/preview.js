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
    const { sources, doctorKey, settings, overrides, conflictSelections, doctorAliases } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const validDoctors = new Set(doctorOptions(sources.mmc, sources.ddh).map((doctor) => doctor.key));
    const requestedKeys = new Set([doctorKey, ...(doctorAliases || []).map((alias) => alias.key)].filter(Boolean));
    if (![...requestedKeys].some((key) => validDoctors.has(key))) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const view = buildRosterView(sources.mmc, sources.ddh, doctorKey, settings, overrides, conflictSelections, doctorAliases);
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
