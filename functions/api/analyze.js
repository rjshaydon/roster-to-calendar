import { defaultSettings, doctorOptions, parseUploadForm, sourceNames } from "../_lib/roster.js";

export async function onRequestPost(context) {
  try {
    const { sources } = await parseUploadForm(context.request);
    const doctors = doctorOptions(sources.mmc, sources.ddh, sources.casey, sources.mch);
    return Response.json({
      sources: sourceNames(sources),
      imports: [
        ...sources.mmc.map((entry) => ({ entry, sourceType: "mmc" })),
        ...sources.ddh.map((entry) => ({ entry, sourceType: "ddh" })),
        ...sources.casey.map((entry) => ({ entry, sourceType: "casey" })),
        ...sources.mch.map((entry) => ({ entry, sourceType: "mch" })),
      ].map(({ entry, sourceType }) => ({
        id: entry.id,
        name: entry.file.name,
        sourceType,
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
