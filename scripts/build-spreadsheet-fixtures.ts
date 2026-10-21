// Regenerates the checked-in workbook e2e fixtures (SPEC-116 on) from the same
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
// - spec118-formats.xlsx    — SPEC-118, Excel-shaped: custom and built-in
//                             number formats, a narrow column and a
//                             too-narrow date
// - spec118-generated.xlsx  — SPEC-118, openpyxl-shaped: t="d" dates under
//                             a General style
// - spec118-formats.ods     — SPEC-118, LibreOffice-shaped data styles and
//                             unstyled dates, times and percentages
// - spec119-formulas.xlsx   — SPEC-119, Excel-shaped: normal, shared (both
//                             orders, an off-grid shift), array, dynamic-
//                             array, data-table and external formulas, the
//                             `_xlfn.`/`_xlpm.` wire forms, an unparseable
//                             LAMBDA call and a sheet with showFormulas="1"
// - spec119-formulas.ods    — SPEC-119, LibreOffice-shaped: OpenFormula,
//                             oooc: and msoxl: formulas, unions,
//                             intersections, an inline array, a matrix and a
//                             repeated formula cell
// - spec120-calc.xlsx       — SPEC-120, Excel-shaped: a wrong cached SUM,
//                             TODAY(), an add-in call, the conformance rows
//                             and a two-cell cycle on "Loop"
// - spec120-fullcalc.xlsx   — SPEC-120, XlsxWriter-shaped: fullCalcOnLoad="1"
//                             and every formula cached 0
// - spec120-manual.xlsx     — SPEC-120, Excel-shaped: calcMode="manual"
// - spec120-calc.ods        — SPEC-120, LibreOffice-shaped: case-sensitive,
//                             wildcards, AutoCalculate, uncached formulas
// - spec121-links.xlsx      — SPEC-121, Excel-shaped: cross-sheet, 3-D and
//                             VLOOKUP formulas, workbook- and sheet-scoped
//                             names, a chart, legacy + threaded comments,
//                             a table, conditional formatting, a validation
//                             and a hyperlink naming another sheet, a tab
//                             colour, a hidden and a very hidden sheet,
//                             calcChain and TitlesOfParts
// - spec121-links.ods       — SPEC-121, LibreOffice-shaped twin: named
//                             range/expression, database range, validation,
//                             conditional format, annotation, internal link,
//                             an embedded chart object, settings.xml
//                             per-table entries with TabColor
// - spec122-shapes.xlsx     — SPEC-122, Excel-shaped: "Edit" (shared, array,
//                             dynamic-array and data-table formulas, rich,
//                             phonetic and _x000D_ strings, a vm picture,
//                             outlines, a merge, a freeze, x14 extensions),
//                             "Locked" (sheetProtection, one unlocked cell),
//                             "Inline" (inline strings)
// - spec122-shapes.ods      — SPEC-122, LibreOffice-shaped: the tails, repeated
//                             runs, a span, an annotation, a link, header rows
//                             and a row group, local names and conditional
//                             formats, a validation, de-ce1, a protected table
// - spec123-styles.xlsx     — SPEC-123, Excel-shaped: named cell styles
//                             (cellStyleXfs, cellStyles with builtinIds, a
//                             hidden and a custom one, cellXfs xfIds), tinted
//                             theme colours, numFmt 164, a styled column, a
//                             styled row, a dxf and the x14 extLst
// - spec123-styles.ods      — SPEC-123, LibreOffice-shaped: common styles
//                             (Default, Heading, Heading_20_1, Good,
//                             Invoice_20_total), parented ce styles, N1xx data
//                             styles with style:map, a column default cell
//                             style and a row formatted to XFD
// - spec124-structure.xlsx  — SPEC-124, Excel-shaped: a frozen first column,
//                             a shared-formula total, a plain and an array
//                             formula, a table with its own autoFilter, two
//                             merges, a workbook autoFilter, a conditional
//                             format, a validation, a hyperlink, a drawing
//                             (a picture and a chart), a legacy comment and
//                             its VML note, an x14 list validation,
//                             calcChain and TitlesOfParts
// - spec124-structure.ods   — SPEC-124, LibreOffice-shaped twin: a content
//                             validation, a formula column, a chart frame
//                             embedded directly in a cell, a merge, a named
//                             range, a named expression, a database range
//                             and a table-local conditional format
// - spec125-fill.xlsx       — SPEC-125, Excel-shaped: the copy block, a merge,
//                             the cut cells, a "Mon" seed, a 20-row guide
//                             column, a fill-down seed, cross-sheet formulas
//                             and a "Series" sheet with a seed column per
//                             series kind
// - spec125-fill.ods        — SPEC-125, LibreOffice-shaped twin
// - spec126-*.csv / .tsv    — SPEC-126, literal bytes per producer: Excel
//                             "CSV UTF-8", a German Excel export
//                             (Windows-1252, `;`), a database dump (all
//                             quoted, ragged, LF), Excel's `sep=;` line,
//                             Excel "Unicode Text" (UTF-16 LE) and a classic
//                             Mac pipe file (lone CR)
// - spec128-filter.xlsx     — SPEC-128, Excel-shaped: "Sales" under an
//                             autoFilter (values with blanks, a custom pair,
//                             its sortState), filterMode, four filter-hidden
//                             rows and one user-hidden row, rows without `r`;
//                             "Colours" a colorFilter and its dxf; "Log" a
//                             cell-colour worksheet sortState;
//                             _xlnm._FilterDatabase names
// - spec128-filter.ods      — SPEC-128, LibreOffice-shaped twin of "Sales":
//                             __Anonymous_Sheet_DB__0 with filter-set-items,
//                             visibility "filter" and "collapse" rows
// spec116-numbers.xlsx and spec120-numbers.xlsx are NOT generated here: Apple
// Numbers exported them (scripts/make-numbers-fixture.applescript,
// scripts/make-numbers-calc-fixture.applescript and their notes in
// e2e/fixtures/).
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
  buildSpec118FormatsOdsBytes,
  buildSpec118FormatsXlsxBytes,
  buildSpec118GeneratedXlsxBytes,
  buildSpec119FormulasOdsBytes,
  buildSpec119FormulasXlsxBytes,
  buildSpec120CalcOdsBytes,
  buildSpec120CalcXlsxBytes,
  buildSpec120FullCalcXlsxBytes,
  buildSpec120ManualXlsxBytes,
  buildSpec121LinksOdsBytes,
  buildSpec121LinksXlsxBytes,
  buildSpec122ShapesOdsBytes,
  buildSpec122ShapesXlsxBytes,
  buildSpec123StylesOdsBytes,
  buildSpec123StylesXlsxBytes,
  buildSpec124StructureOdsBytes,
  buildSpec124StructureXlsxBytes,
  buildSpec125FillOdsBytes,
  buildSpec125FillXlsxBytes,
  buildSpec126ExcelUtf8CsvBytes,
  buildSpec126MacPipeCsvBytes,
  buildSpec126QuotedCsvBytes,
  buildSpec126Semicolon1252CsvBytes,
  buildSpec126SepCsvBytes,
  buildSpec126UnicodeTsvBytes,
  buildSpec128FilterOdsBytes,
  buildSpec128FilterXlsxBytes,
} from "../src/lib/spreadsheet/fixtures/build";
// SPEC-129 (xlsx): spec129-cf.xlsx — Excel-shaped conditional formatting
// (every kind, an x14 data bar twin, an x14-only 3Stars, a custom x14 icon
// set, a sparkline ext kept verbatim, tinted theme colours, bgColor-only dxfs).
import { buildSpec129CfXlsxBytes } from "../src/lib/spreadsheet/fixtures/build";

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
writeFileSync(join(fixturesDir, "spec118-formats.xlsx"), await buildSpec118FormatsXlsxBytes());
writeFileSync(join(fixturesDir, "spec118-generated.xlsx"), await buildSpec118GeneratedXlsxBytes());
writeFileSync(join(fixturesDir, "spec118-formats.ods"), await buildSpec118FormatsOdsBytes());
writeFileSync(join(fixturesDir, "spec119-formulas.xlsx"), await buildSpec119FormulasXlsxBytes());
writeFileSync(join(fixturesDir, "spec119-formulas.ods"), await buildSpec119FormulasOdsBytes());
writeFileSync(join(fixturesDir, "spec120-calc.xlsx"), await buildSpec120CalcXlsxBytes());
writeFileSync(join(fixturesDir, "spec120-fullcalc.xlsx"), await buildSpec120FullCalcXlsxBytes());
writeFileSync(join(fixturesDir, "spec120-manual.xlsx"), await buildSpec120ManualXlsxBytes());
writeFileSync(join(fixturesDir, "spec120-calc.ods"), await buildSpec120CalcOdsBytes());
writeFileSync(join(fixturesDir, "spec121-links.xlsx"), await buildSpec121LinksXlsxBytes());
writeFileSync(join(fixturesDir, "spec121-links.ods"), await buildSpec121LinksOdsBytes());
writeFileSync(join(fixturesDir, "spec122-shapes.xlsx"), await buildSpec122ShapesXlsxBytes());
writeFileSync(join(fixturesDir, "spec122-shapes.ods"), await buildSpec122ShapesOdsBytes());
writeFileSync(join(fixturesDir, "spec123-styles.xlsx"), await buildSpec123StylesXlsxBytes());
writeFileSync(join(fixturesDir, "spec123-styles.ods"), await buildSpec123StylesOdsBytes());
writeFileSync(join(fixturesDir, "spec124-structure.xlsx"), await buildSpec124StructureXlsxBytes());
writeFileSync(join(fixturesDir, "spec124-structure.ods"), await buildSpec124StructureOdsBytes());
writeFileSync(join(fixturesDir, "spec125-fill.xlsx"), await buildSpec125FillXlsxBytes());
writeFileSync(join(fixturesDir, "spec125-fill.ods"), await buildSpec125FillOdsBytes());
writeFileSync(join(fixturesDir, "spec126-excel-utf8.csv"), buildSpec126ExcelUtf8CsvBytes());
writeFileSync(join(fixturesDir, "spec126-semicolon-1252.csv"), buildSpec126Semicolon1252CsvBytes());
writeFileSync(join(fixturesDir, "spec126-quoted.csv"), buildSpec126QuotedCsvBytes());
writeFileSync(join(fixturesDir, "spec126-sep.csv"), buildSpec126SepCsvBytes());
writeFileSync(join(fixturesDir, "spec126-unicode.tsv"), buildSpec126UnicodeTsvBytes());
writeFileSync(join(fixturesDir, "spec126-mac-pipe.csv"), buildSpec126MacPipeCsvBytes());
writeFileSync(join(fixturesDir, "spec128-filter.xlsx"), await buildSpec128FilterXlsxBytes());
writeFileSync(join(fixturesDir, "spec128-filter.ods"), await buildSpec128FilterOdsBytes());
// --- SPEC-129 (xlsx) ---
writeFileSync(join(fixturesDir, "spec129-cf.xlsx"), await buildSpec129CfXlsxBytes());
// --- end SPEC-129 (xlsx) ---

// --- SPEC-129 (ODS) ------------------------------------------------------------
// - spec129-cf.ods     — LibreOffice-shaped `calcext:conditional-formats` over
//                        the named styles Bad, Good and Accent 3 (Accent_20_3)
// - spec129-legacy.ods — `style:map` conditions only, on automatic cell styles
//                        over a repeated-cell run and a column default
import { buildSpec129CfOdsBytes, buildSpec129LegacyOdsBytes } from "../src/lib/spreadsheet/fixtures/build";

writeFileSync(join(fixturesDir, "spec129-cf.ods"), await buildSpec129CfOdsBytes());
writeFileSync(join(fixturesDir, "spec129-legacy.ods"), await buildSpec129LegacyOdsBytes());
// --- end SPEC-129 (ODS) --------------------------------------------------------

// --- SPEC-130 (xlsx) -----------------------------------------------------------
// - spec130-validation.xlsx — Excel-shaped data validation: "Form" with a
//                             literal list (input message, stop alert), whole
//                             (warning), decimal (information, imeMode),
//                             date, text-length and custom rules, a range
//                             list with showDropDown="1", a cross-sheet list
//                             in the x14 ext and a defined-name list; xr:uids;
//                             "Lists" the sources; the "Colours" definedName
import { buildSpec130ValidationXlsxBytes } from "../src/lib/spreadsheet/fixtures/build";

writeFileSync(join(fixturesDir, "spec130-validation.xlsx"), await buildSpec130ValidationXlsxBytes());
// --- end SPEC-130 (xlsx) -------------------------------------------------------

// The idle scheduler keeps Node's event loop alive; the files are written.
process.exit(0);
