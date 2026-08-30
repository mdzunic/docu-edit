// Regenerates the checked-in SPEC-107 media e2e fixtures from the same
// DOM-free builders the unit specs use (src/lib/media/fixtures/build.ts), so
// e2e assertions and unit assertions can never drift apart:
//   npx vite-node scripts/build-media-fixtures.ts
//
// Production notes for each fixture live in the builder module's header
// comment — the normative record of how every file was produced.
import { writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  buildEncryptedPdfBytes,
  buildGifFixtureBytes,
  buildJpegFixtureBytes,
  buildPagesPdfBytes,
  buildPngFixtureBytes,
  buildSvgFixtureBytes,
  buildWebpFixtureBytes,
  buildXrefStreamPdfBytes,
} from "../src/lib/media/fixtures/build";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "e2e", "fixtures");
mkdirSync(fixturesDir, { recursive: true });

const files: Array<[string, Uint8Array]> = [
  ["spec107-photo.jpg", buildJpegFixtureBytes()],
  ["spec107-shot.png", await buildPngFixtureBytes()],
  ["spec107-mark.svg", buildSvgFixtureBytes()],
  ["spec107-anim.gif", buildGifFixtureBytes()],
  ["spec107-shot.webp", buildWebpFixtureBytes()],
  ["spec107-pages.pdf", buildPagesPdfBytes()],
  ["spec107-xrefstream.pdf", await buildXrefStreamPdfBytes()],
  ["spec107-encrypted.pdf", buildEncryptedPdfBytes()],
];

for (const [name, bytes] of files) {
  writeFileSync(join(fixturesDir, name), bytes);
}
console.log(files.map(([name, bytes]) => `${name} (${bytes.byteLength}B)`).join(", "));
// The core idle scheduler's keep-alive timer would otherwise hold the
// process open after the work is done (a known trap for Node callers).
process.exit(0);
