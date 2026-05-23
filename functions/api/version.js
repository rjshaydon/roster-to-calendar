export async function onRequestGet() {
  let branch = process.env.CF_PAGES_BRANCH || "";
  let commit = (process.env.CF_PAGES_COMMIT_SHA || "").slice(0, 7) || "";
  if (!branch || !commit) {
    try {
      const { readFile } = await import("node:fs/promises");
      const { join } = await import("node:path");
      const root = process.cwd();
      const head = (await readFile(join(root, ".git", "HEAD"), "utf8")).trim();
      const refMatch = head.match(/^ref:\s+(.+)$/m);
      if (refMatch) {
        branch = refMatch[1].replace("refs/heads/", "");
        commit = (await readFile(join(root, ".git", refMatch[1].trim()), "utf8")).trim().slice(0, 7);
      } else {
        commit = head.slice(0, 7);
      }
    } catch {
      /* local dev without .git access */
    }
  }
  return Response.json({ branch: branch || "unknown", commit: commit || "unknown" });
}
