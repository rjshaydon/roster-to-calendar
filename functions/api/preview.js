import { doctorOptions, generateEvents, parseUploadForm, previewSummary, serializeEvent } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources, doctorKey } = await parseUploadForm(context.request);
    if (!doctorKey) {
      throw new Error("A doctor selection is required.");
    }
    const validDoctors = doctorOptions(sources.mmc?.workbook, sources.ddh?.workbook).map((doctor) => doctor.key);
    if (!validDoctors.includes(doctorKey)) {
      throw new Error("The selected doctor was not found in the uploaded roster files.");
    }
    const events = generateEvents(sources.mmc?.workbook, sources.ddh?.workbook, doctorKey);
    return Response.json({
      ...previewSummary(events),
      events: events.map(serializeEvent),
    });
  } catch (error) {
    return Response.json({ error: error.message || "Unexpected server error." }, { status: 400 });
  }
}
