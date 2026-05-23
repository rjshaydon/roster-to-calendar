export async function onRequestGet() {
  let branch = "";
  let commit = "";
  try {
    const head = await readGitFile("HEAD");
    if (head) {
      const refMatch = head.match(/^ref:\s+(.+)$/m);
      if (refMatch) {
        branch = refMatch[1].replace("refs/heads/", "");
        commit = (await readGitFile(refMatch[1].trim())) || "";
      } else {
        commit = head.trim();
      }
    }
    if (!commit) commit = process.env.CF_PAGES_COMMIT_SHA || "";
    if (!branch) branch = process.env.CF_PAGES_BRANCH || "";
  } catch {
    branch = process.env.CF_PAGES_BRANCH || "";
    commit = process.env.CF_PAGES_COMMIT_SHA || "";
  }
  return Response.json({ branch: branch || "unknown", commit: (commit || "").slice(0, 7) || "unknown" });
}

async function readGitFile(relativePath) {
  try {
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    const gitDir = path.default.resolve(process.cwd(), ".git");
    const content = await fs.default.readFile(path.default.join(gitDir, relativePath), "utf8");
    return content.trim();
  } catch {
    return null;
  }
}
