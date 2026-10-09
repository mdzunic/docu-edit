// Regenerates the checked-in SPEC-116 workbook e2e fixtures from the same
// DOM-free builders the unit specs use, so e2e and unit assertions can never
// drift apart:
//   npx vite-node scripts/build-spreadsheet-fixtures.ts
//
// How each file is produced (src/lib/spreadsheet/fixtures/build.ts records
// the producer shapes each one emulates):
// - mini-book.xlsx          — Excel-shaped: three sheets (Summary, Data
//                             written without `r` attributes, a hidden sheet)
// - mini-book.ods           — LibreOffice-shaped: three tables (a hidden one
//                             first, the active one second), trailing
//                             1048553-row / 16380-column runs
// - spec116-1904.xlsx       — the 1904 date system
// - spec116-macros.xlsm     — a dummy xl/vbaProject.bin
// - spec116-template.xltx   — the SpreadsheetML template content type
// - spec116-wide.ods        — one cell at column 16,000 of row 900,000
// - spec117-styles.xlsx     — SPEC-117, Excel-shaped: every style, size,
//                             merge, freeze, hidden row/column and view
//                             setting the reader models
// - spec117-styles.ods      — SPEC-117, LibreOffice-shaped: the same
//                             constructs in Calc's own shapes
// spec116-numbers.xlsx is NOT generated here: Apple Numbers exported it
// (scripts/make-numbers-fixture.applescript, e2e/fixtures/spec116-numbers.md).
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  buildMiniBookOdsBytes,
  buildMiniBookXlsxBytes,
  buildSpec1904XlsxBytes,
  buildSpecMacrosXlsmBytes,
  buildSpecTemplateXltxBytes,
  buildSpecWideOdsBytes,
  buildSpec117StylesOdsBytes,
  buildSpec117StylesXlsxBytes,
} from "../src/lib/spreadsheet/fixtures/build";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "e2e", "fixtures");
mkdirSync(fixturesDir, { recursive: true });

writeFileSync(join(fixturesDir, "mini-book.xlsx"), await buildMiniBookXlsxBytes());
writeFileSync(join(fixturesDir, "mini-book.ods"), await buildMiniBookOdsBytes());
writeFileSync(join(fixturesDir, "spec116-1904.xlsx"), await buildSpec1904XlsxBytes());
writeFileSync(join(fixturesDir, "spec116-macros.xlsm"), await buildSpecMacrosXlsmBytes());
writeFileSync(join(fixturesDir, "spec116-template.xltx"), await buildSpecTemplateXltxBytes());
writeFileSync(join(fixturesDir, "spec116-wide.ods"), await buildSpecWideOdsBytes());
writeFileSync(join(fixturesDir, "spec117-styles.xlsx"), await buildSpec117StylesXlsxBytes());
writeFileSync(join(fixturesDir, "spec117-styles.ods"), await buildSpec117StylesOdsBytes());

// The idle scheduler keeps Node's event loop alive; the files are written.
process.exit(0);
