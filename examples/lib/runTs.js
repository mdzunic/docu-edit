// Every example is a plain, flag-free `node examples/<name>.js` — but the
// generation API these examples call is TypeScript, like the rest of
// src/lib/ (see CLAUDE.md), and plain Node cannot import a .ts file (or its
// extensionless relative imports) without flags this command intentionally
// never passes. This bridges the two: it loads the example's actual logic
// (examples/_impl/*.ts), which imports src/lib/generate directly, through
// Vite's own module runner — `runnerImport` from the `vite` package, already
// a real dependency, not something added for this feature. (It used to
// spawn `vite-node`, which vitest no longer ships since v4.)
import { fileURLToPath } from "node:url";
import { runnerImport } from "vite";

/** Runs `implRelativePath` (a .ts file, resolved relative to the calling
 * example's own URL) to completion, and exits 0 — or 1, printing the error,
 * if it throws. */
export async function runExample(importerUrl, implRelativePath) {
  const implPath = fileURLToPath(new URL(implRelativePath, importerUrl));
  try {
    await runnerImport(implPath, { configFile: false, logLevel: "error" });
  } catch (error) {
    console.error(error);
    process.exit(1);
  }
  // The runner keeps its module graph alive; the example itself is done.
  process.exit(0);
}
