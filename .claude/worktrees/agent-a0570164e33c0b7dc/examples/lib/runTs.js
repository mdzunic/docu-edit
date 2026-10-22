// Every example is a plain, flag-free `node examples/<name>.js` — but the
// generation API these examples call is TypeScript, like the rest of
// src/lib/ (see CLAUDE.md), and plain Node cannot import a .ts file without
// an experimental flag this command intentionally never passes. This
// bridges the two: it shells out to `vite-node` (already a real
// dependency — vitest's own execution engine, not something added for
// this feature) to run the example's actual logic (examples/_impl/*.ts),
// which imports src/lib/generate directly.
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

/** Runs `implRelativePath` (a .ts file, resolved relative to the calling
 * example's own URL) via vite-node, and exits with its exit code. */
export function runExample(importerUrl, implRelativePath) {
  const implPath = fileURLToPath(new URL(implRelativePath, importerUrl));
  const viteNodeBin = fileURLToPath(new URL("../node_modules/.bin/vite-node", importerUrl));
  const result = spawnSync(viteNodeBin, [implPath], { stdio: "inherit" });
  if (result.error) {
    throw result.error;
  }
  process.exit(result.status ?? 1);
}
