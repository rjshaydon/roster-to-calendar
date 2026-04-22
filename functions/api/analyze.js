import { defaultSettings, doctorOptions, parseUploadForm, sourceNames } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources } = await parseUploadForm(context.request);
    const doctors = doctorOptions(sources.mmc?.workbook, sources.ddh?.workbook);
    return Response.json({
      sources: sourceNames(sources),
      doctors,
      settings: defaultSettings(),
    });
  } catch (error) {
    return Response.json({ error: error.message || "Unexpected server error." }, { status: 400 });
  }
}
