---
id: SPEC-140
title: Workbook generation API
status: draft
area: platform
complexity: normal
story_points: 8
depends_on: [SPEC-045, SPEC-106, SPEC-121, SPEC-122, SPEC-123, SPEC-124, SPEC-126, SPEC-127]
---
## Why
SPEC-045 gave systems a way to generate documents and SPEC-106 gave them
decks. The third question every buyer of that API asks is the one with the
most volume behind it: *and the spreadsheet?* The monthly sales export, the
reconciliation a finance system mails out, the CSV somebody wants as a
formatted workbook with totals that actually compute — these are generated
from data far more often than typed. Today `src/lib/generate/` cannot produce
a single cell, and a `SpreadsheetModel` can be reached only by opening a file
or pressing "New spreadsheet".

The writers exist (SPEC-121's `writeXlsx`/`writeOds`, SPEC-122's sheet
serialisation), so does the engine that makes a formula's cached value
correct (SPEC-120), the style vocabulary and its patch command (SPEC-117,
SPEC-123), merges and frozen panes (SPEC-117, SPEC-124), and the CSV parser
(SPEC-126). This spec puts a builder and a JSON description in front of
them. It follows SPEC-106 to the letter: validate on entry, never emit an
invalid file, and **generation adds no serialisation code of its own** — a
construct the writers cannot emit is fixed in the writer. **It is the
normative source for `WorkbookBuilder`, the JSON workbook-description
format and its schema, `readWorkbook`, and the rule that a generated
workbook's cached formula values are computed before it is written.**

What a workbook generator must not do is write formulas with no cached
values and hope: Excel can be told to recalculate on open, but LibreOffice
by default does not, Numbers shows what it is given, and every preview pane
and search indexer reads the cached `<v>`. So the builder evaluates every formula with
SPEC-120's engine, and refuses one it cannot evaluate.

## Acceptance criteria
- [ ] `src/lib/generate/workbook.ts` exports `WorkbookBuilder` per the Reference: `new WorkbookBuilder(options?: { title?; format?: "xlsx" | "ods"; dateSystem?: "1900" | "1904"; locale? })`, starting from SPEC-121's `createBlankWorkbook()` with its sheets replaced by the ones the builder adds (none added → one empty "Sheet1"), every method returning `this`, every argument validated on entry through `generate/validate.ts`'s `expect*`/`rejectUnknownKeys` so a `DocumentGenerationError` naming the offending path (`"sheets[0].cells.B5.formula: unknown function SUMM"`) always precedes any byte
- [ ] Sheets are fluent: `addSheet(name, options?)` appends a sheet and makes it current (`tabColor` `#rrggbb`, `hidden`, `showGridLines`), `sheet(name)` makes an existing one current, and every cell, range, size, merge and freeze method acts on the current sheet and throws "no sheet — call addSheet first" when there is none; names pass SPEC-121's `validateSheetName(model, name)` (1–31 characters, none of `: \ / ? * [ ]`, no leading or trailing `'`, not "History", unique ignoring case), its message becoming the error text at `sheets[i].name`
- [ ] Cell values per "Cell input": strings (≤ 32,767 characters; a leading "=" stays text), finite numbers, booleans, `null`, `{ error: "#N/A" }` from SPEC-116's `ErrorCode` list, and dates — a JS `Date` is read in **UTC** and an ISO `{ date: "2026-10-08" }` / `"2026-10-08T13:45:00"` without offset by its calendar fields, both converted with SPEC-116's `serial.ts` in the workbook's date system and given "yyyy-mm-dd" (or "yyyy-mm-dd hh:mm" with a time) unless a number format is supplied; an invalid `Date`, `NaN` or an offset-bearing ISO string is a validation error
- [ ] `setCell(ref, input)`, `setRow(ref, values)` (left to right from `ref`) and `setRows(ref, rows, options?)` (a 2-D block) write values; `setRows(…, { parseText: true })` routes each string through SPEC-126's `inferCsvField(field, locale, dateSystem)` so "1,234.50", "45%" and "10/08/2026" become typed values with their number formats in the builder's locale — and **never** a formula: SPEC-122's `parseInput` is deliberately not used here, because it turns a leading "=" into a formula and CSV formula injection is exactly the attack a data-import option must not open
- [ ] Formulas are SPEC-119 canonical text (a leading "=" is accepted and stripped): `setFormula(ref, text)` or `{ formula }` validates syntax with `parseFormula` (a `FormulaSyntaxError` becomes `"…formula: unexpected ')' at 7"`), rejects a function absent from SPEC-120/127's registry by name, and at output time rejects a reference to a sheet the workbook does not have and any circular reference (`"sheets[0].cells.B2.formula: circular reference (Sales!B2)"`); every formula cell's cached value is then computed by SPEC-120's `recalculate`, so `toWorkbook()` holds the values Excel would show
- [ ] Styles are SPEC-117's `CellStyle` through a validated subset (`CellStyleSpec`: font family/size/bold/italic/underline single|double/strike/colour, a solid fill, borders per edge plus `all`/`outline` shorthands with SPEC-117's border styles, alignment horizontal/vertical/wrap/indent/rotation, and `numberFormat`), colours as `#rrggbb` only; `style(range, spec | presetName)`, `numberFormat(range, code)` and a cell's own `style`/`numberFormat` are applied in call order through SPEC-123's `applyStylePatch(model, sheetId, ranges, patch: StylePatch)`, so a builder style and an editor style mean the same thing (a `{ ok: false, message }` result becomes the error at that call's path); `defineStyle(name, spec)` registers a preset; number format codes are validated by SPEC-118's `parseFormatCode`
- [ ] Geometry per "Sizes, merges and panes": `columnWidth("B:D", { chars } | { px })` stores `widthPx` by SPEC-117's maximum-digit-width formula and keeps `sourceWidthChars` so the `.xlsx` carries the exact width (0–255 characters), `rowHeight(3, { pt } | { px })` sets `heightPt` + `customHeight` (0–409 pt), `merge("A1:D1")` goes through SPEC-124's `mergeCells(model, sheetId, [range], "cells")` and is refused — stricter than the editor, which asks or extends — when it would discard a value outside the top-left cell (`"sheets[0].merges[2]: A1:D1 would discard the value in C1"`) or overlap another merge, and `freeze({ rows, cols })` sets `Sheet.frozen`
- [ ] The builder accumulates cells in private state and materialises the model **once**, at `toWorkbook()`/`toBytes()`, in the order pinned in "Materialisation" — never one copy-on-write command per cell, which is quadratic on a 100,000-cell sheet; the result satisfies SPEC-116's invariants (sorted, sparse, plain JSON — `JSON.parse(JSON.stringify(model))` `toEqual`s it), has an empty `opaque` and no `partName` or `sourceXml` anywhere, and a workbook past `MAX_WORKBOOK_CELLS` is refused with the product's own `WORKBOOK_TOO_LARGE_MESSAGE` rather than written
- [ ] Output mirrors SPEC-106's deck builder: `toWorkbook(): SpreadsheetModel`, `toBytes(format?, options?): Promise<Uint8Array>`, `toBlob(…)`, `toBuffer(…)` (a `DocumentGenerationError` where `Buffer` is not global) and `fileName(format?)` (the same naming rule as `spreadsheetDownload`), `format` defaulting to the builder's; bytes come from SPEC-121's `writeXlsx(model, { modified })`/`writeOds(model, { modified })` on their from-scratch path and nothing else — `src/lib/generate/` gains no XML, ZIP or format code — and `options.modified` is passed straight to those writers' own `modified` option, so the same input with the same `modified` yields `equalBytes`-identical packages, part by part after `unzip`
- [ ] `src/lib/generate/workbookDescription.ts` exports `WorkbookDescription` and `workbookFromDescription(description): SpreadsheetModel`, a plain-JSON format mapping **1:1** onto the builder per "The JSON workbook description", published as `schema/workbook-description.json` beside `schema/document-description.json`; `WorkbookBuilder.fromDescription(description)` is the static factory; unknown keys at any level are validation errors with JSON paths, in `rejectUnknownKeys`'s own wording (`"sheets[1].cells.C4.fomula: unknown property"`), and `{}` describes a one-sheet empty workbook
- [ ] `src/lib/generate/readWorkbook.ts` exports `readWorkbook(bytes, title?): Promise<SpreadsheetModel>`: `unzip`, `detectPackage` and SPEC-116's `spreadsheetFormatOf` decide the format from the package's own declarations, `readSpreadsheetBytes` reads it behind `nodeDom.ts`'s `ensureNodeXmlGlobals()`, and anything that is not an `.xlsx`/`.ods` workbook is a `DocumentGenerationError` ("bytes: not an .xlsx or .ods workbook") — no `File`, no disk, no browser
- [ ] Round trip: build → write → `readWorkbook` reproduces every construct this API emits in both formats under "Round-trip relation" — values and formulas exactly, effective styles per cell, column widths within 1 px, row heights, merges, frozen panes, sheet names, order, visibility and tab colours; number formats by code in `.xlsx` and by rendered text (SPEC-118 `formatValue`, `en-US`) in `.ods`
- [ ] Open-in-editor matches SPEC-045's: a generated `Blob` wrapped in a `File` goes through `openAnyDocumentFile` and `storeOpenedFile` (the funnel SPEC-116 extends to workbooks) and lands at `editorPathFor("spreadsheet", id)` with no download round trip, tested in `src/lib/generate/openWorkbookInEditor.spec.ts` under jsdom as `openInEditor.spec.ts` tests documents; `storeSpreadsheet(builder.toWorkbook())` is the documented model-level alternative
- [ ] The boundary holds: `workbook.ts`, `workbookDescription.ts`, `workbookStyles.ts` and `readWorkbook.ts` import only `src/lib/spreadsheet/`, `src/lib/core/` and the product-neutral `generate/` modules (`errors.ts`, `validate.ts`, `nodeDom.ts`) — never `runs.ts`, `blocks.ts`, `builder.ts`, `description.ts`, `read.ts`, the deck modules, `documentModel.ts` or `openDocument.ts` — asserted by a new `src/lib/generate/workbookBoundary.spec.ts`; nothing under `src/lib/spreadsheet/` imports `generate/` (SPEC-116's `spreadsheet/boundary.spec.ts`, unchanged, already forbids it); `generate/index.ts` re-exports the workbook surface beside both builders
- [ ] Generation is DOM-free and fast: the build and write path touches no `document`, `window` or `DOMParser` (asserted by running it under plain Node with no polyfill — a from-scratch writer path that needs one is fixed in the writer), and a 10,000-row × 10-column workbook with a formula column, a styled header row and two number-formatted columns builds and writes `.xlsx` in under 2 seconds and `.ods` in under 2 seconds in Node (`src/lib/generate/workbookPerformance.spec.ts`, the shape `performance.spec.ts` has)
- [ ] Documentation: `docs/generate-workbook-api.md` (new — `docs/generate-api.md` is shared by two builders already) covers every method, every cell-input form, `CellStyleSpec`, the JSON description, validation and paths, `readWorkbook`, open-in-editor and the Node notes; `docs/generate-api.md` gains a short "Workbooks" section linking it and nothing else
- [ ] Three runnable examples under SPEC-045's `examples/<name>.js` + `examples/_impl/<name>.ts` + `lib/runTs.js` pattern, each writing to `examples/output/`: `sales-report.js` (monthly sales by region — SUM/AVERAGE totals, a growth-percentage formula column, currency and percent formats, a bold filled header, frozen header row, column widths → `.xlsx`), `csv-to-xlsx.js` (reads the committed `examples/data/orders.csv` with SPEC-126's `parseCsvText`, `setRows` with `parseText`, widths from content → `.xlsx`) and `budget.js` (Income, Expenses and Summary sheets with cross-sheet formulas, a merged title, `.xlsx` and `.ods` from one description) — each run by `src/lib/generate/examples.spec.ts`, which re-reads the output with `readWorkbook` and checks one computed total
- [ ] e2e (`e2e/SPEC-140.spec.ts`): the Playwright file builds the Reference's "Q3 sales" workbook plus a second sheet "Notes" with `WorkbookBuilder` in Node (as `e2e/SPEC-045.spec.ts` does), writes `.xlsx` and `.ods` to temp files, and opens each through `open-document-input` → `/sheets/:id` → `sheet-cell[data-ref="A1"]` reads "Region", `E2` reads "$3,960.00" (the computed `SUM(B2:D2)`, formatted), the formula bar shows `=SUM(B2:D2)` with `E2` active, a "Notes" `sheet-tab` is present, row 1 stays visible after scrolling down (frozen) → download → `readWorkbook` on the downloaded bytes gives `E2` the value 3960

## Out of scope
Defined names, Excel tables, charts, images, data validation, conditional
formatting, notes/comments, hyperlinks, page setup, rich text runs within a
cell, array formulas, row/column outline grouping, sheet protection and
pivot tables in the builder — each a one-method follow-up over its own
spec's model once landed (SPEC-129–136, SPEC-139); until then the
description rejects their keys as unknown properties rather than ignoring
them. Theme colours in
`CellStyleSpec` (`#rrggbb` only). Reading an existing workbook *into* the
description format. Streaming output for workbooks past `MAX_WORKBOOK_CELLS`.
CSV or PDF as an output of this API (SPEC-126 and SPEC-136 export those from
the editor). A hosted HTTP service (SPEC-045's exclusion carries over).

## Reference

### Dependencies restated
- `specs/045-document-generation-api.md` / `specs/106-…` — `src/lib/generate/`
  holds `DocumentBuilder` (`builder.ts`), `PresentationBuilder`
  (`presentation.ts`), `DocumentGenerationError` (`errors.ts`, imports
  nothing), the `expect*` and `rejectUnknownKeys` helpers (`validate.ts`,
  imports only `errors.ts`), `ensureNodeXmlGlobals()` (`nodeDom.ts`), and
  `index.ts` re-exporting the public surface. Rules carried over verbatim:
  validate on entry; never emit an invalid file; add no serialisation code;
  the same input with the same `modified` gives `equalBytes`-identical
  parts; paths name the offending field. SPEC-106's output set (`toDeck`,
  `toBytes(format?)`, `toBlob`, `toBuffer`, `fileName`) is the shape copied.
  `generate/` is a **composition point** (it imports the text product
  already, and the deck product after SPEC-106), which is why it may import
  `lib/spreadsheet` and why `core/boundary.spec.ts`'s product blocks do not
  list it.
- `specs/116-workbooks-in-the-library.md` — `SpreadsheetModel`, `Sheet`,
  `Row`, `Cell`, `CellValue`, `ErrorCode`, `MAX_ROWS`, `MAX_COLS`,
  `MAX_WORKBOOK_CELLS`, `WORKBOOK_TOO_LARGE_MESSAGE`; invariants (sorted,
  sparse, copy-on-write, plain JSON); `serial.ts`'s `dateToSerial(y, m, d,
  system)`; `address.ts`'s `parseA1`, `parseRange`, `colToLetters`;
  `readSpreadsheetBytes(bytes, format, title, parts?)`;
  `spreadsheetFormatOf(info)` in `core/packageDetect.ts`;
  `spreadsheetDownload(model, targetFormat?)` → `Promise<{ bytes; fileName;
  mimeType; notices }>` (as SPEC-121 extends it) — `fileName` follows its
  naming rule;
  `openAnyDocumentFile` + `storeOpenedFile` + `storeSpreadsheet` +
  `editorPathFor("spreadsheet", id)` = `/sheets/:id`.
- `specs/117-…` — `CellStyle` (font, fill, borders, alignment,
  `numberFormat`), `SheetBorderStyle`; `ColumnSpan { min; max; widthPx?;
  sourceWidthChars? }`, `Row.heightPt`/`customHeight`, `Sheet.merges`,
  `Sheet.frozen?: { rows; cols }`, `Sheet.showGridLines?`, `Sheet.tabColor?`;
  column width px = trunc(((256·w + trunc(128/MDW))/256)·MDW), MDW 7 for
  the default Calibri 11.
- `specs/118-…`/`specs/119-…`/`specs/120-…` — `parseFormatCode`,
  `formatValue`, the six `SpreadsheetLocale`s; canonical formula text,
  `parseFormula` throwing `FormulaSyntaxError` with an offset; the function
  registry (`formula/registry.ts`) and `recalculate(model, changed)`, which
  reports cycles; formulas are always en-US syntax regardless of locale.
- `specs/121-create-workbook-and-sheet-management.md` —
  `createBlankWorkbook()` (`spreadsheet/sheets.ts`: one "Sheet1", `format:
  "xlsx"`, the default style, `DEFAULT_THEME`, `opaque: []`),
  `validateSheetName(model, name, ownIndex?): string | null`,
  `writeXlsx(model, options?: { modified?: Date; fullCalcOnLoad?: boolean })`
  and `writeOds(model, options?: { modified?: Date })` with a from-scratch
  path for a workbook with no retained package.
- `specs/122-cell-entry-and-the-sheet-writer.md` — the sheet writer both
  formats use, `internStyle(model, style)`, and `parseInput(text, locale,
  dateSystem)`, which this API does **not** use for data (it makes formulas).
- `specs/123-cell-formatting.md` — `StylePatch`, `patchStyle(style, patch)`
  and `applyStylePatch(model, sheetId, ranges, patch): FormatResult` in
  `spreadsheet/formatting.ts` (`null` removes a field). `specs/124-…` —
  `mergeCells(model, sheetId, ranges, "centre" | "across" | "cells")`.
  `specs/126-csv-and-tsv.md` — `parseCsvText(text, delimiter, onRecord)` and
  `inferCsvField(field, locale, dateSystem): { value; numberFormat? }`.
  `specs/127-…` — the full function registry (SPEC-120's core functions plus
  the library), so a builder formula is rejected by name only when neither
  registers it.
- Equality: `equalBytes` (`src/lib/core/equality.ts`), part-level after
  `unzip`. Model comparisons are field-level under "Round-trip relation";
  the text side's `equalSemantic` is `DocumentModel`-typed and unused.

### `WorkbookBuilder`
```ts
const wb = new WorkbookBuilder({ title: "Q3 sales" });
wb.defineStyle("header", { font: { bold: true }, fill: "#dce6f1",
                           borders: { bottom: "thin" } })
  .addSheet("Sales")
  .setRow("A1", ["Region", "July", "August", "September", "Total"])
  .setRows("A2", [["East", 1200, 1350, 1410], ["West", 980, 1010, 1130]])
  .setFormula("E2", "SUM(B2:D2)").setFormula("E3", "SUM(B3:D3)")
  .style("A1:E1", "header")
  .numberFormat("B2:E3", "$#,##0.00")
  .columnWidth("A", { chars: 14 })
  .freeze({ rows: 1 });
const bytes = await wb.toBytes();                       // .xlsx
```
```ts
export type WorkbookOutputFormat = "xlsx" | "ods";
export type CellInput = string | number | boolean | Date | null | {
  value?: string | number | boolean | null; date?: string; error?: ErrorCode;
  formula?: string; numberFormat?: string; style?: CellStyleSpec | string };
export interface CellStyleSpec {
  font?: { family?: string; sizePt?: number; bold?: boolean; italic?: boolean;
           underline?: "single" | "double"; strike?: boolean; color?: string };
  fill?: string;                                   // solid "#rrggbb"
  borders?: Partial<Record<"top" | "right" | "bottom" | "left" | "all" | "outline",
            SheetBorderStyle | { style: SheetBorderStyle; color?: string }>>;
  alignment?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top"
    | "center" | "bottom"; wrap?: boolean; indent?: number; rotation?: number };
  numberFormat?: string;
}
```
Methods: `addSheet`, `sheet`, `setCell`, `setRow`, `setRows`, `setFormula`,
`style`, `numberFormat`, `defineStyle`, `columnWidth`, `rowHeight`,
`merge`, `freeze`, then `toWorkbook`, `toBytes`, `toBlob`, `toBuffer`,
`fileName` and the static `fromDescription`. Addresses are A1 text
(`"B3"`, `"A1:D4"`, `"B:D"`, `"3:5"`), rows in `rowHeight` 1-based as the
user sees them; `all` sets every edge of every cell in the range, `outline`
the range's outer edges only. `CellStyleSpec` maps onto `CellStyle` field
for field in `workbookStyles.ts` (`fill` → `{ kind: "solid", color }`,
`rotation` → `textRotation`, booleans → `true` or absent); `false` in a later
call becomes SPEC-123's `null` (remove) in the `StylePatch`, so `bold: false`
undoes an earlier `bold: true`.

### Cell input
| Input | Stored `value` | Number format when none given |
| --- | --- | --- |
| `"text"` (≤ 32,767 chars; "=…" is text) | the string | — |
| finite number (`-0` → `0`) | the number | — |
| `true` / `false` | the boolean | — |
| `null` | no cell (a style may still store one) | — |
| `{ error: "#N/A" }` | `CellError` | — |
| `Date` (UTC fields) / `{ date: "YYYY-MM-DD[THH:MM[:SS[.sss]]]" }` | serial in `dateSystem` | "yyyy-mm-dd" or "yyyy-mm-dd hh:mm" |
| `{ formula: "SUM(B2:B4)" }` | the computed result | — |
| string with `parseText` | `inferCsvField`'s value | `inferCsvField`'s format |

A 1900-system date before 1900-03-01 follows SPEC-116's fictitious
1900-02-29 rule exactly as the reader does, so a generated date re-reads as
the same calendar day.

### Sizes, merges and panes
`{ chars: w }` → `widthPx` by the MDW formula above with `sourceWidthChars:
w`; `{ px }` → `widthPx` only; adjacent columns with equal settings share
one `ColumnSpan`. `{ px }` row heights convert at 0.75 pt/px. Merges are
applied after values and styles; a value in any non-top-left cell is an
error, never silently dropped (SPEC-124's editor command would ask; a
generator has no one to ask). `freeze({ rows: 1 })` freezes the first row;
`{ rows: 0, cols: 0 }` clears.

### Materialisation
`toWorkbook()` (and every `to*` output, which calls it) builds a fresh
model each time, so the builder stays reusable: (1) `createBlankWorkbook()`
with title, format, date system and the builder's sheets; (2) per sheet,
the accumulated cells sorted into rows — the last write to a cell wins —
with dates converted and `parseText` applied; (3) style calls, in call
order, through `applyStylePatch` (a cell's own `style`/`numberFormat`
counts as a style call made when the cell was set); (4) merges; (5) column
spans, row heights, frozen panes, sheet options; (6) validation that needs
the whole workbook — sheet references, the cell ceiling; (7)
`recalculate(model, keys)` (SPEC-120) with the `cellKey` of every formula
cell, a non-empty `RecalcResult.circular` becoming the circular-reference
error at its first cell; (8) a structural self-check against SPEC-116's
invariants. Volatile functions (`NOW`, `TODAY`, `RAND`) are evaluated at
generation time through `src/lib/core/clock.ts`'s `now()`, so tests pin them
with `setClock`.

### The JSON workbook description
1:1 with the methods, validated by `workbookFromDescription`, published as
`schema/workbook-description.json` (draft-07, `additionalProperties: false`
throughout, as `document-description.json` is):
```json
{ "title": "Q3 sales", "format": "xlsx", "dateSystem": "1900", "locale": "en-US",
  "styles": { "header": { "font": { "bold": true }, "fill": "#dce6f1" } },
  "sheets": [ {
    "name": "Sales", "tabColor": "#4472c4", "hidden": false, "showGridLines": true,
    "cells": { "A1": "Region", "B5": { "formula": "SUM(B2:B4)", "numberFormat": "$#,##0.00" },
               "F2": { "date": "2026-10-08" } },
    "data": [ { "at": "A2", "rows": [["East", 1200], ["West", 980]], "parseText": false } ],
    "styles": [ { "range": "A1:E1", "style": "header" },
                { "range": "B2:E5", "numberFormat": "$#,##0.00" } ],
    "columns": [ { "columns": "A", "width": { "chars": 14 } } ],
    "rows": [ { "rows": "1", "height": { "pt": 24 } } ],
    "merges": ["A10:E10"], "freeze": { "rows": 1, "cols": 0 } } ] }
```
| Key | Method |
| --- | --- |
| `title`, `format`, `dateSystem`, `locale` | constructor options |
| `styles` (top level) | `defineStyle` per entry |
| `sheets[i]` + `name`, `tabColor`, `hidden`, `showGridLines` | `addSheet` |
| `cells` (A1 → `CellInput`) | `setCell`, in key order |
| `data[]` (`at`, `rows`, `parseText?`) | `setRows` |
| `styles[]` (`range` + `style` and/or `numberFormat`) | `style` / `numberFormat` |
| `columns[]`, `rows[]` | `columnWidth`, `rowHeight` |
| `merges`, `freeze` | `merge`, `freeze` |

Paths use the JSON shape — `sheets[0].cells.B5.formula`,
`sheets[0].data[1].rows[3][2]`, `styles.header.font.sizePt` — and the
fluent API reports the same path for the same mistake, counting sheets and
calls in order, so a description and the program that would build it fail
identically. A description that names a construct this spec does not build
(`"names"`, `"charts"`, `"validations"`…) fails as an unknown property.

### Round-trip relation
`comparableWorkbook(model)` (test helper in `src/lib/generate/
workbookRoundTrip.spec.ts`) keeps, per sheet in order: name, visibility,
tab colour, frozen panes, merges, column widths (±1 px), row heights, and
per stored cell its value, formula and **effective** `CellStyle` with
`sourceXf`/`sourceStyleName`/`namedStyle` removed — style indices differ
between a generated and a re-read model by design. For `.ods` the
`numberFormat` code is replaced by `formatValue(value, code, en-US)`.
Generated → written → read must be `toEqual` under it, both formats.

### Tests
- `workbook.spec.ts` — every method's model output, each validation error
  with its exact path text, current-sheet rules, date conversion in both
  date systems including serial 60, `parseText` (a field "=SUM(A1:A2)" stays
  text), style mapping, merges' value refusal, the materialisation
  order (a later style patch wins; the last cell write wins), reuse of one
  builder for two outputs.
- `workbookDescription.spec.ts` — every row of the key table, unknown keys
  at each level, `{}`, the fluent/JSON path parity, a schema-shape test
  beside SPEC-045's.
- `workbookRoundTrip.spec.ts` — the relation above in both formats; byte
  determinism with a fixed `modified`; formulas' cached values equal
  SPEC-120's results after re-read.
- `readWorkbook.spec.ts`, `openWorkbookInEditor.spec.ts`,
  `workbookPerformance.spec.ts`, `workbookBoundary.spec.ts`,
  `examples.spec.ts`; `src/lib/core/boundary.spec.ts`,
  `src/lib/spreadsheet/boundary.spec.ts` and every existing
  `src/lib/generate/*.spec.ts` pass unchanged; `e2e/SPEC-140.spec.ts`.

### Specs this changes
- `specs/045-…` — `src/lib/generate/` gains a third product, `index.ts`
  more exports and `docs/generate-api.md` a "Workbooks" link section;
  `DocumentBuilder`, the document description, its schema and examples are
  unchanged.
- `specs/106-…` — nothing in the deck builder changes; its output-method
  shape is reused.
- `specs/121-…` — its from-scratch writers gain a second caller and no new
  semantics; a from-scratch path found touching `DOMParser` under Node is
  fixed there, not worked around here.
- `specs/000-roadmap.md` gains this spec's row.
