import { defaultSettings, doctorOptions, parseUploadForm, sourceNames } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources } = await parseUploadForm(context.request);
    const doctors = doctorOptions(sources.mmc, sources.ddh);
    return Response.json({
      sources: sourceNames(sources),
      imports: [...sources.mmc, ...sources.ddh].map((entry) => ({
        id: entry.id,
        name: entry.file.name,
        sourceType: sources.mmc.includes(entry) ? "mmc" : "ddh",
        addedAt: entry.addedAt || "",
        size: entry.file.size,
        lastModified: entry.file.lastModified,
      })),
      doctors,
      settings: defaultSettings(),
    });
  } catch (error) {
    return Response.json({ error: error.message || "Unexpected server error." }, { status: 400 });
  }
}
