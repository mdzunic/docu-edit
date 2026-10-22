#!/usr/bin/env -S npx vite-node
// One-off generator for the SPEC-053 field fixtures
// `e2e/fixtures/word-fields.docx` and `e2e/fixtures/lo-fields.odt` (run
// manually: `npx vite-node scripts/build-spec053-fixtures.mjs`). Like the
// SPEC-029/031 fixtures these go through this project's OWN writer: foreign
// producers' field XML shapes are covered by hand-authored unit tests
// (src/lib/docx/fields.spec.ts, src/lib/odf/fields.spec.ts); what these
// exercise is AC-51..AC-53's end-to-end open -> render -> update pipeline,
// which needs real package bytes.
//
// Generation is a script and not a `npm test` side effect on purpose: these
// fixtures are committed and e2e/SPEC-053.spec.ts uploads word-fields.docx,
// so a unit run that rewrote them in place could be mid-write while
// Playwright read them — a mystery e2e failure unrelated to the code under
// test. src/lib/fixtures.fields.spec.ts instead reads the committed bytes
// and fails with a pointer back here when the model has moved on.
//
// The model itself lives in src/lib/fixtures.fields.ts so this script and
// that spec cannot drift apart. It pins `modified`, so re-running with an
// unchanged model rewrites the same content rather than churning the stamp.
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { buildFieldFixtures } from "../src/lib/fixtures.fields.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(e2eFixturesDir, { recursive: true });

async function main() {
  for (const { name, bytes } of await buildFieldFixtures()) {
    const target = path.join(e2eFixturesDir, name);
    writeFileSync(target, bytes);
    console.log(`wrote ${target} (${bytes.length} bytes)`);
  }
}

await main();
// writeDocx/writeOdt's chunked save path yields through idleScheduler's
// shared MessageChannel, whose started port keeps Node's event loop alive
// after the files are fully written — without an explicit exit this script
// hangs forever (same reason build-spec047-fixture.mjs ends this way).
process.exit(0);
