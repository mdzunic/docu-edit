#!/usr/bin/env -S npx vite-node
// One-off generator for the SPEC-077 theme fixtures
// `e2e/fixtures/word-theme.docx` and `e2e/fixtures/lo-theme.odt` (run
// manually: `npx vite-node scripts/build-spec077-fixtures.mjs`). Like the
// SPEC-074 fixtures these go through this project's OWN writer: foreign
// producers' theme-reference shapes are covered by hand-authored unit tests
// (src/lib/docx/theme-related specs and the odf reader's); what these
// exercise is AC-52/53's open -> theme change -> save -> reopen pipeline,
// which needs real package bytes.
//
// Generation is a script and not a `npm test` side effect on purpose: these
// fixtures are committed and e2e/SPEC-077.spec.ts uploads them, so a unit
// run that rewrote them in place could be mid-write while Playwright read
// them. src/lib/fixtures.theme.spec.ts instead reads the committed bytes and
// fails with a pointer back here when the model has moved on.
//
// The model itself lives in src/lib/fixtures.theme.ts so this script and
// that spec cannot drift apart. It pins `modified`, so re-running with an
// unchanged model rewrites the same content rather than churning the stamp.
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { buildThemeFixtures } from "../src/lib/fixtures.theme.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(e2eFixturesDir, { recursive: true });

async function main() {
  for (const { name, bytes } of await buildThemeFixtures()) {
    const target = path.join(e2eFixturesDir, name);
    writeFileSync(target, bytes);
    console.log(`wrote ${target} (${bytes.length} bytes)`);
  }
}

await main();
// writeDocx/writeOdt's chunked save path yields through idleScheduler's
// shared MessageChannel, whose started port keeps Node's event loop alive
// after the files are fully written — without an explicit exit this script
// hangs forever (same reason build-spec074-fixtures.mjs ends this way).
process.exit(0);
