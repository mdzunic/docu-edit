// Deterministic generator for SPEC-137's workbook performance-benchmark
// fixtures: e2e/fixtures/spec137-large.xlsx (500,000 cells, 25,000 rows x 20
// columns — e2e/SPEC-137.perf.spec.ts's primary fixture) and
// spec137-library.xlsx (100,000 cells — its library-landing rows, and
// e2e/SPEC-137.spec.ts's functional fixture). Neither is committed to git
// (.gitignore); both benchmarks build them on demand when absent.
//
// Run manually with:  npx vite-node scripts/build-spec137-fixtures.ts
//
// Deterministic the same way scripts/build-spec047-fixture.mjs is: every
// randomised value comes from a seeded mulberry32 stream (see
// src/lib/spreadsheet/fixtures/build.ts's SPEC-137 section), and the
// surrounding package (docProps dates, zip entry timestamps) was already
// fixed by zipWriter.ts and coreXml() for every fixture this module builds —
// so the same script produces the same bytes on every run.
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { buildSpec137LargeXlsxBytes, buildSpec137LibraryXlsxBytes } from "../src/lib/spreadsheet/fixtures/build";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "e2e", "fixtures");
mkdirSync(fixturesDir, { recursive: true });

async function main() {
  const large = await buildSpec137LargeXlsxBytes();
  writeFileSync(join(fixturesDir, "spec137-large.xlsx"), large);
  console.log(`wrote e2e/fixtures/spec137-large.xlsx (${large.length} bytes)`);

  const library = await buildSpec137LibraryXlsxBytes();
  writeFileSync(join(fixturesDir, "spec137-library.xlsx"), library);
  console.log(`wrote e2e/fixtures/spec137-library.xlsx (${library.length} bytes)`);
}

await main();
// The idle scheduler's shared MessageChannel (used by the save/zip path's
// yields) keeps a started port alive after the files are fully written —
// without an explicit exit this script (and the benchmark's
// execFileSync waiting on it) hangs forever, as SPEC-047's own fixture
// script notes.
process.exit(0);
