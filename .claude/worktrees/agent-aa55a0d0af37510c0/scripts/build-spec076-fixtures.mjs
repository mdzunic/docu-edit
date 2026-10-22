#!/usr/bin/env -S npx vite-node
// One-off generator for the SPEC-076 table fixtures under `e2e/fixtures/`
// (run manually: `npx vite-node scripts/build-spec076-fixtures.mjs`). Eight
// packages — four documents in both .docx and .odt: a four-page table with a
// repeating header, a banded table, a table sorted by two columns, and a
// tab-separated list (AC-75).
//
// Generation is a script and not a `npm test` side effect on purpose: these
// fixtures are committed and e2e/SPEC-076.spec.ts uploads them, so a unit run
// that rewrote them in place could be mid-write while Playwright read them.
// src/lib/fixtures.tables.spec.ts instead reads the committed bytes and fails
// with a pointer back here when the model has moved on.
//
// The models live in src/lib/fixtures.tables.ts so this script and that spec
// cannot drift apart. It pins `modified`, so re-running with an unchanged
// model rewrites the same content rather than churning the stamp.
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { buildTableFixtures } from "../src/lib/fixtures.tables.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(e2eFixturesDir, { recursive: true });

async function main() {
  for (const { name, bytes } of await buildTableFixtures()) {
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
