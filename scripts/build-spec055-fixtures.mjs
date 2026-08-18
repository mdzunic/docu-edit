#!/usr/bin/env -S npx vite-node
// One-off generator for the SPEC-055 caption fixtures
// `e2e/fixtures/word-captions.docx` and `e2e/fixtures/lo-captions.odt` (run
// manually: `npx vite-node scripts/build-spec055-fixtures.mjs`). Like the
// SPEC-053 field fixtures these go through this project's OWN writer, so what
// they exercise is the end-to-end open -> render -> update pipeline that needs
// real package bytes. The model lives in src/lib/fixtures.captions.ts so this
// script and src/lib/fixtures.captions.spec.ts cannot drift apart; it pins
// `modified` so re-running an unchanged model rewrites the same content.
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { buildCaptionFixtures } from "../src/lib/fixtures.captions.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(e2eFixturesDir, { recursive: true });

async function main() {
  for (const { name, bytes } of await buildCaptionFixtures()) {
    const target = path.join(e2eFixturesDir, name);
    writeFileSync(target, bytes);
    console.log(`wrote ${target} (${bytes.length} bytes)`);
  }
}

await main();
// The chunked save path keeps Node's event loop alive (idleScheduler's
// MessageChannel); exit explicitly so this script does not hang.
process.exit(0);
