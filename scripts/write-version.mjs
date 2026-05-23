import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(fileURLToPath(import.meta.url), "..", "..");
const gitDir = join(root, ".git");

let branch = "";
let commit = "";

try {
  const head = (await readFile(join(gitDir, "HEAD"), "utf8")).trim();
  const refMatch = head.match(/^ref:\s+(.+)$/m);
  if (refMatch) {
    branch = refMatch[1].replace("refs/heads/", "");
    commit = (await readFile(join(gitDir, refMatch[1].trim()), "utf8")).trim();
  } else {
    commit = head;
  }
} catch {
  // silently fall back to unknown
}

await writeFile(
  join(root, "public", "version.json"),
  JSON.stringify({ branch: branch || "unknown", commit: (commit || "").slice(0, 7) || "unknown" }),
  "utf8",
);
