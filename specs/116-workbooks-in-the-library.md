---
id: SPEC-116
title: Workbooks in the library — open .xlsx and .ods, the workbook model and a byte-preserving save
status: draft
area: io
complexity: complex
story_points: 13
depends_on: [SPEC-003, SPEC-008, SPEC-009, SPEC-058, SPEC-060, SPEC-067, SPEC-107]
---
## Why
The library holds documents, decks, images and PDFs. It does not hold the
file that sits beside all of them in every shared folder: the budget, the
timesheet, the price list, the export somebody's system produced. Drop an
`.xlsx` on the landing page today and it falls through the open funnel to the
`.docx` path, which fails with a Word error about a file that was never a
Word document; drop an `.ods` and it is refused as "isn't a text document or
presentation we can open". Either way the user leaves the product.

The pieces are already here. `src/lib/core` reads and writes ZIP packages,
parses OPC relationships and Content-Types, detects ODF and OPC packages by
their own declarations (`core/packageDetect.ts`), reads `docProps/core.xml`
and `meta.xml` (`core/opcProperties.ts`, `core/odfMeta.ts`), reads a
DrawingML theme (`core/themeDrawingMl.ts`), and resolves theme colours
(`core/theme.ts`). `documentKinds.ts` is a registry whose own header says
adding a kind "(e.g. spreadsheets)" is an edit there plus a new product layer;
SPEC-107 proved that by adding `"media"`. The store already splits heavy
payloads out of the record the landing page lists.

What is missing is a **fourth kind**: `"spreadsheet"`, a product layer at
`src/lib/spreadsheet/`, and a workbook view at `/sheets/:id`. This spec adds
the kind, both readers, a byte-preserving download and a read-only grid that
shows every sheet's values. **It is the normative source for the workbook
model (`SpreadsheetModel`, `Sheet`, `Row`, `Cell`, `CellValue`), its
invariants, A1 addressing (`spreadsheet/address.ts`), date serials
(`spreadsheet/serial.ts`), the "General" display of a value
(`spreadsheet/general.ts`), the reader-version upgrade path, workbook
detection, and the store's per-sheet row split.** Every later spreadsheet
spec (SPEC-117–140) edits this model.

Like SPEC-107 and the presentations preparation milestone, it edits nothing.
Cell styles, column widths, merges and selection are SPEC-117; number formats
SPEC-118; formulas SPEC-119/120; the first edit SPEC-121. A downloaded
workbook is the file that came in, part for part.

## Acceptance criteria
- [ ] `documentKinds.ts` gains the kind per the Reference: `DocumentKind = "text" | "presentation" | "media" | "spreadsheet"`, `SPREADSHEET_FORMATS = ["xlsx", "ods"]`, `kindOfFormat` returns `"spreadsheet"` for each, `editorPathFor("spreadsheet", id)` is `/sheets/${id}`, `LibraryFormat` gains `SpreadsheetFormat`, and `OPENABLE_EXTENSIONS` gains `xlsx`, `xlsm`, `xltx`, `xltm`, `ods`, `ots` — every kind-dispatched call site (router guards, `HomeView`, `OpenFromUrlView`, `useLibrary.downloadDocuments`, the pickers' `accept`) keeps deriving from this module and gains no local knowledge of what a workbook is
- [ ] The router's editor guards stop being pairwise: `/doc/:id`, `/slides/:id`, `/slides/:id/present/:n`, `/media/:id` and the new `/sheets/:id` each redirect **any** other kind's id to `editorPathFor(kind, id)` through one shared `guardKind(expected)` helper in `src/router.ts`, so a fifth kind never needs N edits — a spreadsheet id at `/doc/:id` lands on `/sheets/:id`, and the reverse holds for every pair (tested as a table over all kinds × all routes)
- [ ] `src/lib/spreadsheet/` is the fourth product layer, importing only its own siblings and `src/lib/core` — never `documentModel.ts`, the text commands, `docx/`, `odf/`, `md/`, `presentation/`, `media/`, the store or Vue — enforced by a new `src/lib/spreadsheet/boundary.spec.ts` modelled line-for-line on `presentation/boundary.spec.ts`, and by a new block in `core/boundary.spec.ts` asserting that no core, text, presentation or media module imports `spreadsheet/`
- [ ] `src/lib/spreadsheet/model.ts` exports the model per "The workbook model": `SpreadsheetModel` (`kind: "spreadsheet"`, `id`, `title`, `format`, optional `packageVariant`, `readerVersion`, `sheets`, `activeSheetIndex`, `dateSystem`, optional `odsNullDate`/`odsContentEnvelope`, `opaque`, optional `preview`/`properties`/`media`), `Sheet`, `Row`, `Cell`, `CellValue`, `CellError`, `ErrorCode`, `TextRun`, and the constants `MAX_ROWS = 1_048_576`, `MAX_COLS = 16_384`, `MAX_WORKBOOK_CELLS = 2_000_000`; the model is plain JSON — `JSON.parse(JSON.stringify(model))` `toEqual`s it — and its file header states the four invariants (copy-on-write rows, sparseness, the source is the splice base, dates are serials) as normative for every later spec
- [ ] `src/lib/spreadsheet/address.ts` exports the addressing helpers per the Reference (`colToLetters`, `lettersToCol`, `formatA1`, `parseA1`, `formatRange`, `parseRange`, `parseSqref`, `getCell`, `findRow`, `usedRange`); `lettersToCol(colToLetters(c)) === c` for every `c` in `0..MAX_COLS-1`, `parseA1` accepts optional `$` markers and returns `null` (never throws) for anything outside `A1..XFD1048576`, and `getCell` is a binary search over the sparse rows and cells
- [ ] `src/lib/spreadsheet/serial.ts` converts between calendar dates and serials per "Dates are serials": the 1900 system honours Excel's fictitious 1900-02-29 (serial 60 — `serialToParts(60)` is `{ year: 1900, month: 2, day: 29 }`), the 1904 system starts at serial 0 = 1904-01-01, times are day fractions, ODS `office:date-value`/`office:time-value` (including `PT36H05M00S` durations and fractional seconds) convert exactly, and a `table:null-date` other than 1899-12-30 or 1904-01-01 is recorded in `odsNullDate` and honoured both ways
- [ ] Workbook detection runs inside the existing ZIP funnel from the package's own declarations, never the extension: an OPC package whose main part's content type is the SpreadsheetML workbook, template, macro-enabled workbook or macro-enabled template type opens as `"xlsx"` (recording `packageVariant`); an ODF package whose `mimetype` is `application/vnd.oasis.opendocument.spreadsheet` or `…spreadsheet-template` opens as `"ods"`; `core/packageDetect.ts` gains the content-type constants and `spreadsheetFormatOf(info)` beside the existing presentation one
- [ ] Every workbook this product cannot read is refused with a message that names the way out, never misparsed: an `.xlsb` package (`application/vnd.ms-excel.sheet.binary.macroEnabled.main`) → `XLSB_UNSUPPORTED_MESSAGE`; an OLE compound file (`D0 CF 11 E0 A1 B1 1A E1`) whose name ends `.xls`, `.xlsx` or `.xlsm` → `LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE` (the docx path's own `.doc` message is unchanged for every other name); an iWork package that is a Numbers document → `NUMBERS_UNSUPPORTED_MESSAGE`, checked **before** the Keynote dispatch so SPEC-087 never tries to read a spreadsheet as a deck (today it reaches the Keynote reader and is refused as "This iWork file isn't a Keynote presentation we can open", which is true but names no way out — that message stays for Pages); a flat `.fods` keeps the existing `FLAT_ODF_MESSAGE` — the strings are pinned in the Reference
- [ ] Behavioural string changes, with every existing assertion updated in the same change: `UNSUPPORTED_FILE_TYPE_MESSAGE` becomes "We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .pdf and image files" (`e2e/SPEC-009.spec.ts`, `e2e/SPEC-050.spec.ts`, `AppShell.spec.ts`), `UNSUPPORTED_ODF_MESSAGE` becomes "This OpenDocument file isn't a text document, presentation or spreadsheet we can open", `titleFromFileName` strips the six new extensions, and `openDocumentFile`'s text-only contract rejects a workbook with `SPREADSHEET_NOT_TEXT_MESSAGE` ("This is a spreadsheet file — open it from the library instead") rather than misparsing it
- [ ] The `.xlsx` reader (`spreadsheet/xlsx/reader.ts`) produces the model per "Reading .xlsx": sheets in `<sheets>` order with name, `state` → `visibility`, and worksheet vs chartsheet/dialogsheet/macrosheet → `kind: "grid" | "foreign"`; shared strings (plain and rich `<r>` runs, `_xHHHH_` escapes decoded, phonetic runs ignored); every cell type (`n`, `s`, `str`, `inlineStr`, `b`, `e`, `d`); rows and cells without an `r` attribute placed sequentially; formula cells contribute their cached `<v>` as `value`; `workbookPr@date1904` → `dateSystem`; `bookViews/workbookView@activeTab` → `activeSheetIndex`; `docProps/core.xml` → `properties`
- [ ] The `.ods` reader (`spreadsheet/ods/reader.ts`) produces the model per "Reading .ods": `table:table` elements in order; `table:number-rows-repeated` / `table:number-columns-repeated` honoured without ever materialising repeated **empty** rows or cells (LibreOffice's trailing `1048553`-row and `16380`-column runs cost nothing); a repeated non-empty cell or row is expanded; `table:covered-table-cell` occupies its column but stores nothing; `office:value-type` float/percentage/currency → number, date/time → serial, boolean → boolean, string → the `text:p` text (paragraphs joined with `"\n"`, `text:s`/`text:tab`/`text:line-break` expanded), `calcext:value-type="error"` → `CellError`; a table whose style has `table:display="false"` is `hidden`; `settings.xml`'s `ActiveTable` → `activeSheetIndex`; `meta.xml` → `properties`
- [ ] Worksheet parts, `xl/sharedStrings.xml` and `.ods` `content.xml` are read with a forward-only XML scanner, `src/lib/core/xmlScanner.ts` per the Reference — never `DOMParser` — because a million-cell sheet as a DOM costs an order of magnitude more memory than its text; the scanner reports exact source offsets, which is how the `.ods` reader slices each table's `sourceXml`, and it refuses a part containing a DTD rather than expanding entities
- [ ] A workbook whose stored cells would exceed `MAX_WORKBOOK_CELLS` is refused with `WORKBOOK_TOO_LARGE_MESSAGE` as soon as the reader counts past the ceiling — it stops scanning, it never materialises the remainder — and a cell or row outside `MAX_ROWS × MAX_COLS` is a parse error naming its address, not a silent drop; `MAX_FILE_SIZE_BYTES` stays 50 MB for every kind
- [ ] Downloading an untouched workbook re-emits every source part byte-identical (`equalBytes` per part, in source order; the ZIP container itself is rebuilt, exactly as the presentation suites assert), `.ods` keeps `mimetype` first and STOREd, a macro-enabled workbook keeps `xl/vbaProject.bin` and its content types, and the file name's extension follows `format` + `packageVariant` per the Reference table — asserted for every fixture in `src/lib/spreadsheet/roundTrip.spec.ts`
- [ ] `.ods` splice equality holds at read: `odsContentEnvelope.prefix + sheets.map(s => s.sourceXml).join("") + odsContentEnvelope.suffix` equals the source `content.xml` text exactly, asserted for every `.ods` fixture so the writers SPEC-121/122 build on a verified base
- [ ] `src/lib/spreadsheet/open.ts` exports `SPREADSHEET_READER_VERSION = 1`, `READER_UPGRADES` (empty) and `upgradeSpreadsheetModel(model)` per "Reader versions"; the store calls it on every read path (initial load, refresh), a record with an older `readerVersion` and no dirty sheet is re-parsed from `opaque` keeping its `id`, `title`, starred state and sheet ids, and a unit test drives the path with a synthetic version-0 record
- [ ] The store holds the fourth kind per "The store": `StoredModel` gains `SpreadsheetModel`, `isSpreadsheetModel`/`getSpreadsheet(id)`/`storeSpreadsheet(model)` join their siblings, `getOrCreateDocument` and `saveDraft` throw on a spreadsheet id, `useAutosave.beginSession` bails, and each sheet's `rows` is written to its own `SHEET_ROWS_KEY_PREFIX` record only when that array's reference changed — so a rename, a star or a library listing never structured-clones a cell — rejoined on every read path and removed by `deleteDocument`, with `libraryStorageUsage` counting them
- [ ] The library treats workbooks as first-class rows and cards: `formatBadgeLabel` covers `XLSX`, `XLSM`, `XLTX`, `XLTM`, `ODS`, `OTS` (by `format` + `packageVariant`), `FORMAT_CHIPS` gains an `xlsx` and an `ods` chip after `odp` (variants filter with their format), `FILTERABLE_FORMATS` gains both, the card shows a `SheetMiniature` built from `model.preview` and an "N sheets" stat, and search, sort, starring, multi-select, trash, restore and bulk download work on workbook entries with no new code path
- [ ] `/sheets/:id` (`SpreadsheetView.vue`) is a wiring hub with the SPEC-058 app bar (title, format label — "Excel workbook (.xlsx)", "Excel macro-enabled workbook (.xlsm)", "OpenDocument Spreadsheet (.ods)" and the template variants — `spreadsheet-download`, and SPEC-060's Close as `close-document`), a not-found state `spreadsheet-not-found`, and `SpreadsheetView.size.spec.ts` opening its ratchet ledger at the line count it lands on
- [ ] The grid (`SheetGrid.vue`, `spreadsheet-grid`) is DOM-virtualised per "The read-only grid": lettered column headers, numbered row headers, the default 64 px × 20 px cells, only the viewport's cells plus overscan in the DOM, `role="grid"` with `aria-readonly="true"`, `aria-rowcount`/`aria-colcount` and per-cell `aria-rowindex`/`aria-colindex`, each rendered cell a `sheet-cell` with `data-ref="B3"`; every value is shown through `generalText` (numbers right-aligned, text left, booleans and errors centred) and clipped to its cell; scrolling a 1,000,000-row sheet to its last used row keeps fewer than 2,000 cell elements mounted
- [ ] Sheet tabs (`sheet-tabs`, one `sheet-tab` per visible sheet, `role="tablist"`/`role="tab"`, `aria-selected`) open on the file's active sheet and switch the grid on click, through the ARIA tablist keyboard pattern (Left/Right/Home/End move between tabs inside the focused tablist, roving `tabindex`), and on Ctrl+PageUp/Ctrl+PageDown from the grid — Excel's binding, delivered only where the browser does not keep it for its own tab switching (Chrome and Firefox do in an ordinary window), so it is a convenience and never the only keyboard path; hidden and very-hidden sheets have no tab; a foreign sheet's tab shows a `sheet-foreign` panel reading "This chart sheet can't be shown yet" (or "This sheet type can't be shown yet") in place of the grid
- [ ] `src/lib/spreadsheet/general.ts` exports `generalText(value: CellValue): string` per "The General display" — the 11-character rule, scientific notation above it, `TRUE`/`FALSE`, the error code, the string itself — and SPEC-118 later subsumes it without changing any output it already produces for a full-width cell
- [ ] Fixtures: `src/lib/spreadsheet/fixtures/build.ts` builds `e2e/fixtures/mini-book.xlsx` and `mini-book.ods` (regenerated by `npx vite-node scripts/build-spreadsheet-fixtures.ts`) plus `spec116-1904.xlsx`, `spec116-macros.xlsm`, `spec116-template.xltx` and `spec116-wide.ods`, each emulating its producer's real shapes per "Fixtures"; and the human-generated `e2e/fixtures/spec116-numbers.xlsx`, exported by Apple Numbers through `scripts/make-numbers-fixture.applescript`, is committed **before this spec is labelled ready** (the factory cannot run Numbers), with a note recording the Numbers version
- [ ] Tests: `spreadsheet/address.spec.ts`, `serial.spec.ts`, `general.spec.ts`, `model.spec.ts` (JSON round trip, invariants), `xlsx/reader.spec.ts`, `ods/reader.spec.ts`, `roundTrip.spec.ts` (byte-identical download for every fixture, splice equality, the Numbers fixture reads), `open.spec.ts` (upgrade path), `core/xmlScanner.spec.ts`, `boundary.spec.ts`, the router kind-guard table, `HomeView.spreadsheet.spec.ts` (badges, chips, miniature, bulk actions over mixed kinds) and `views/SpreadsheetView.spec.ts`
- [ ] e2e (`e2e/SPEC-116.spec.ts`): drop `mini-book.xlsx` on the library → a row with an `XLSX` badge and "3 sheets" → open it → `B2`'s cell reads "Revenue", a number cell reads right-aligned → switch to the second tab → its values show → Close → drop `mini-book.ods` → open → the first tab is the file's active sheet → download → the downloaded bytes' parts are `equalBytes` with the fixture's → visit `/doc/<the workbook id>` → lands on `/sheets/<id>` → filter by the `ODS` chip shows one row

## Out of scope
Any edit (SPEC-121 onward): no typing, no "New spreadsheet", no undo. Cell
styles, fonts, fills, borders, column widths and row heights from the file,
merged cells, frozen panes, hidden rows and columns, gridline settings, zoom,
text overflow into neighbouring cells, selection, the formula bar and the
status bar (SPEC-117). Number formats — a date shows as its serial and a
percentage as its fraction until SPEC-118, which is the honest state of a
viewer that does not yet read formats. Formulas — their cached values are
shown; reading, showing and evaluating the formulas themselves is
SPEC-119/120. Charts, images and shapes on sheets (SPEC-131/132 — preserved
verbatim here). CSV and TSV (SPEC-126 — they have no package to detect).
`.xls`, `.xlsb`, `.numbers` and flat `.fods` import (refused with messages;
backlog). Decrypting a password-protected workbook. Running or inspecting
macros: a macro-enabled workbook's VBA project is bytes this product carries,
never code it reads. External workbook links (their cached values show like
any other). Spreadsheet small-screen layout beyond the grid scrolling
natively.

## Reference

### Dependencies restated
- `specs/003-local-document-library.md` / `specs/067-…` — the store is
  `src/lib/documentStore.ts`: an in-memory map written through to IndexedDB;
  `listDocuments()` → `LibraryEntry[]` (`id`, `title`, `format`, `kind`,
  `lastModified`, `starred`, `trashedAt?`); `trashDocument`/
  `restoreDocument`/`deleteDocument`/`setStarred`/`libraryStorageUsage`; the
  search/sort/filter/multi-select surface in `src/lib/libraryView.ts`
  (`FILTERABLE_FORMATS`, `FORMAT_CHIPS`), `src/composables/useLibrary.ts` and
  `src/views/library/*`; badges in `src/lib/format.ts` (`formatBadgeLabel`).
  SPEC-047's media split: `MEDIA_KEY_PREFIX` records written by
  `persistMedia` only when `model.media`'s reference changed, stripped from
  the shell by `stripMedia`, rejoined by `rejoinMedia` on every read path —
  the pattern the row split copies.
- `specs/008-round-trip-preservation.md` — never discard content you did not
  parse; preserve unknown parts verbatim. Part-level `equalBytes`
  (`src/lib/core/equality.ts`) is the round-trip relation for packages whose
  ZIP container is rebuilt (presentations and media assert the same way).
- `specs/009-drag-drop-and-open-from-url.md` — `openAnyDocumentFile` in
  `src/lib/openDocument.ts` is the one funnel (picker, drop, `/open?url=`):
  media is sniffed first (`detectMedia`), then one `unzip`, then
  `detectPackage(parts)` → presentation dispatch (`presentationFormatOf`) →
  Keynote (`isLegacyKeynote`/`isIwaPackage`) → ODF text → the `.docx` path.
  `MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024`; `hasSupportedExtension` gates a
  drop by `OPENABLE_EXTENSIONS`. The `.docx` reader's CFB check
  (`DOC_MAGIC_BYTES`, `LEGACY_DOC_MESSAGE`) lives in `src/lib/docx/reader.ts`.
- `specs/058-…` / `specs/060-…` — the app bar comes from `AppShell.vue`'s
  slots; Close is `data-testid="close-document"` (as `MediaView.vue` uses it).
- `specs/107-images-and-pdfs-in-the-library.md` — the precedent this spec
  follows step for step: a kind in `documentKinds.ts`, a product layer with a
  boundary spec, a model with `kind` as the store's discriminant, kind-guarded
  route, library badges/chips, `getMedia`-style narrowing, `saveDraft` and
  `getOrCreateDocument` throwing on the wrong kind.
- `docs/presentations-architecture.md` — the layering: core is
  product-neutral and flat, each product imports core and its own siblings
  only, products meet in the app layer (`documentKinds.ts`,
  `openDocument.ts`, `documentStore.ts`, the router, the views). This spec
  adds a fourth box, not a new rule.

### The workbook model
```ts
// src/lib/spreadsheet/model.ts
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import type { OpaquePart } from "../core/packageParts";

export type SpreadsheetFormat = "xlsx" | "ods";        // SPEC-126 widens with "csv" | "tsv"
export type PackageVariant = "macroEnabled" | "template" | "macroEnabledTemplate";

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;                        // A..XFD
export const MAX_WORKBOOK_CELLS = 2_000_000;           // stored cells, all sheets (SPEC-137 revisits)

export type ErrorCode =
  | "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A"
  | "#GETTING_DATA" | "#SPILL!" | "#CALC!" | "#FIELD!" | "#BLOCKED!" | "#CONNECT!" | "#BUSY!" | "#UNKNOWN!";
export interface CellError { error: ErrorCode }

/** Dates and times are numbers — serials in the workbook's date system. */
export type CellValue = number | string | boolean | CellError | null;

/** One run of a rich-text string. SPEC-117 adds `font?: CellFont`; until
 *  then a run carries its text only (the formatting stays in the source). */
export interface TextRun { text: string }

export interface Cell {
  col: number;           // 0-based
  value: CellValue;      // the literal, or a formula cell's cached result; null = no value
  runs?: TextRun[];      // rich text; `value` is then the runs' concatenated text
  // Later specs add optional fields only: formula (119), styleId (117), …
}

export interface Row {
  index: number;         // 0-based
  cells: Cell[];         // ascending by col, unique
  // SPEC-117 adds heightPt?, customHeight?, hidden?, styleId?
}

export interface Sheet {
  id: string;            // minted at read; persisted in the store; never written to a file
  name: string;
  visibility: "visible" | "hidden" | "veryHidden";
  kind: "grid" | "foreign";   // foreign: chartsheet/dialogsheet/macrosheet — placeholder + verbatim
  rows: Row[];           // sparse, ascending by index
  partName?: string;     // xlsx: the sheet part ("xl/worksheets/sheet1.xml"); its bytes are in `opaque`
  sourceXml?: string;    // ods: this sheet's <table:table> text, with the whitespace that preceded it
}

export interface SheetPreview {
  sheetName: string;
  /** Up to PREVIEW_ROWS (12) × PREVIEW_COLS (6) display strings from the
   *  active sheet's top-left corner; "" for empty cells. */
  cells: string[][];
}

export interface SpreadsheetModel {
  kind: "spreadsheet";
  id: string;
  title: string;
  format: SpreadsheetFormat;
  packageVariant?: PackageVariant;
  readerVersion: number;
  sheets: Sheet[];
  activeSheetIndex: number;
  dateSystem: "1900" | "1904";
  odsNullDate?: string;  // ISO date, only when neither 1899-12-30 nor 1904-01-01
  odsContentEnvelope?: { prefix: string; suffix: string };
  opaque: OpaquePart[];  // the whole source package, verbatim, in order
  preview?: SheetPreview;
  properties?: DocumentProperties;
  media?: MediaPart[];   // unused until SPEC-131; present so the store's media split applies unchanged
}
```

**Invariants** (stated in `model.ts`'s header; every later spec keeps them):
1. **Copy-on-write rows.** No command mutates a `Row` or `Cell` in place, nor
   a sheet's `rows` array: an edit builds new objects for what it touches and
   a new (shallow-copied) `rows` array, and every untouched row keeps its
   object identity. The undo stack (SPEC-121), recalculation's change
   detection (SPEC-120) and this spec's persistence split all find what
   changed by reference identity.
2. **Sparse.** A cell with `value: null` and nothing else is not stored; a row
   with no cells and no properties is not stored. Nothing is ever allocated
   per column or per row of the 1,048,576 × 16,384 grid.
3. **The source is the splice base.** A sheet's source — the xlsx part bytes
   in `opaque`, the ods `sourceXml` — is kept for the life of the sheet, even
   after it is edited. Writers (SPEC-121/122) re-serialise only what the model
   models and splice it into the source; every unmodelled element is
   re-emitted verbatim. This is what lets a later spec model a new element
   and upgrade already-stored workbooks by reading it from the retained
   source (see "Reader versions").
4. **Dates are serials.** There is no date type: a date is a number whose
   meaning comes from its number format (SPEC-118). Times are day fractions.

The model never holds a `Map`, `Set` or class instance, and never reaches the
store as a Vue reactive proxy (a proxy fails IndexedDB's structured clone).

### Addressing (`src/lib/spreadsheet/address.ts`)
```ts
export interface CellRef { row: number; col: number }                        // 0-based
export interface GridRange { top: number; left: number; bottom: number; right: number } // inclusive, normalised
export function colToLetters(col: number): string;          // 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD"
export function lettersToCol(letters: string): number | null;
export function formatA1(ref: CellRef): string;             // { row: 2, col: 1 } → "B3"
export function parseA1(text: string): (CellRef & { rowAbs: boolean; colAbs: boolean }) | null;
export function formatRange(range: GridRange): string;      // single cell → "B3", else "A1:C3"
export function parseRange(text: string): GridRange | null; // "A1:C3", "C3:A1" (normalised), "B3"
export function parseSqref(text: string): GridRange[];      // space-separated list; invalid entries dropped
export function findRow(sheet: Sheet, row: number): Row | undefined;           // binary search
export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined;
export function usedRange(sheet: Sheet): GridRange | null;  // bounding box of stored cells
```
`core/tables.ts` already exports an unrelated `CellRange` for text tables;
spreadsheet code never uses that name. Letters are case-insensitive on input
and upper-case on output.

### Dates are serials (`src/lib/spreadsheet/serial.ts`)
```ts
export function dateToSerial(y: number, m: number, d: number, system: "1900" | "1904"): number;
export function serialToParts(serial: number, system: "1900" | "1904"):
  { year: number; month: number; day: number; hour: number; minute: number; second: number; ms: number };
export function odsDateValueToSerial(iso: string, system: "1900" | "1904", nullDate?: string): number;
export function odsTimeValueToSerial(duration: string): number;      // "PT12H30M00S" → 0.520833…
export function serialToOdsDateValue(serial: number, system: "1900" | "1904", nullDate?: string): string;
export function serialToOdsTimeValue(serial: number): string;
```
- **1900 system.** Serial 1 = 1900-01-01 … 59 = 1900-02-28, **60 = the
  fictitious 1900-02-29** Excel inherited from Lotus, 61 = 1900-03-01, and
  from there serial = days since 1899-12-30. Dates before 1900-01-01 (only
  an `.ods` can hold them) count backwards from 1899-12-31 = 0, so the scale
  stays monotonic: 1899-12-30 = −1.
- **1904 system.** Serial 0 = 1904-01-01; serial = days since then.
- **ODS.** `table:null-date` absent or 1899-12-30 → `"1900"`; 1904-01-01 →
  `"1904"`; anything else → `"1900"` plus `odsNullDate`, and conversions go
  through the stored null date. `office:date-value` may carry a time
  (`2026-10-08T13:45:00.5`); `office:time-value` is an ISO 8601 duration
  that may exceed 24 hours (`PT36H` → 1.5).
- Times are fractions of a day; conversions round to the nearest millisecond.

### The General display (`src/lib/spreadsheet/general.ts`)
`generalText(value)` at full column width, matching Excel's "General":
- `null` → `""`; a string → itself; `true`/`false` → `"TRUE"`/`"FALSE"`; an
  error → its code.
- A number: `0` → `"0"`. If `1e-9 ≤ |n| < 1e11`, the fixed-notation text of
  `n` rounded to at most 15 significant digits, with trailing zeros and a
  trailing point removed; if that is longer than 11 characters (sign and
  point included), round off decimals until it fits. Otherwise scientific:
  mantissa rounded so the whole text is at most 11 characters, upper-case
  `E`, signed exponent of at least two digits (`123456789012` →
  `"1.23457E+11"`, `0.0000000001234` → `"1.234E-10"`).
- SPEC-118 adds width-aware shrinking and the locale's decimal separator;
  at full width in `en-US` its output equals this function's.

### Reading .xlsx
- **Package.** Main part via the root `officeDocument` relationship
  (`core/packageDetect.ts`). Main content types:
  `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` (workbook),
  `…spreadsheetml.template.main+xml` (`packageVariant: "template"`),
  `application/vnd.ms-excel.sheet.macroEnabled.main+xml` (`"macroEnabled"`),
  `application/vnd.ms-excel.template.macroEnabled.main+xml` (`"macroEnabledTemplate"`).
  `application/vnd.ms-excel.sheet.binary.macroEnabled.main` is `.xlsb` → refused.
- **Workbook** (`xl/workbook.xml`, small → `core/xml.ts` `parseXml`):
  `sheets/sheet@name`, `@state` (`hidden`/`veryHidden`), `@r:id` → the
  workbook rels; relationship type `…/worksheet` → `kind: "grid"`,
  `…/chartsheet`, `…/dialogsheet`, `…/xlMacrosheet` → `kind: "foreign"`.
  `workbookPr@date1904` (`1`/`true`) → `"1904"`.
  `bookViews/workbookView[0]@activeTab` → `activeSheetIndex` (clamped; if it
  points at a hidden sheet, the first visible one).
- **Shared strings** (scanner): each `<si>` is either `<t>` or a sequence of
  `<r>` runs (`<rPr>` ignored here, SPEC-117 reads it) — `rPh`/`phoneticPr`
  are skipped. `_xHHHH_` escapes decode to the code point (`_x000D_` → CR;
  `_x005F_` is a literal underscore, so `_x005F_x0041_` stays `_x0041_`).
- **Cells** (scanner over `sheetData`): `row@r` 1-based, or previous + 1 when
  absent; `c@r` A1, or previous column + 1 when absent. `c@t`: absent/`n` →
  `<v>` as a number (empty `<v>` → no value); `s` → shared string by index
  (out of range → parse error); `str` → `<v>` text (a formula's string
  result); `inlineStr` → `<is>` text/runs; `b` → `<v>` `1`/`0`; `e` → the
  error code (unknown codes → `#VALUE!`); `d` → ISO 8601 → serial. A `<c>`
  with only a style and no value is skipped in this spec (SPEC-117 keeps it
  for its `styleId`). `<f>` is skipped in this spec — its cached `<v>` is the
  value (SPEC-119 reads the formula).
- Everything else in the package — styles, theme, drawings, comments,
  tables, calcChain, VBA, custom XML — is in `opaque` and untouched.

### Reading .ods
- **Package.** `mimetype` `application/vnd.oasis.opendocument.spreadsheet`
  or `…spreadsheet-template` (`packageVariant: "template"`).
- **content.xml** (scanner): `office:spreadsheet` holds, in order, optional
  declarations (`table:calculation-settings` — whose `table:null-date`
  sets the date system — `table:content-validations`, `table:label-ranges`),
  the `table:table` run, then `table:named-expressions`,
  `table:database-ranges`, `table:data-pilot-tables` and friends.
  `odsContentEnvelope.prefix` is everything up to the first `table:table`
  start tag's preceding whitespace; each sheet's `sourceXml` is that
  whitespace plus its `<table:table …>…</table:table>`; `suffix` is the rest.
- **Rows and columns.** `table:table-row` may be nested in
  `table:table-header-rows` / `table:table-row-group` /
  `table:table-rows`; all count in document order.
  `table:number-rows-repeated="N"`: an empty row advances the row index by N
  and stores nothing; a non-empty row is materialised N times (counted
  against `MAX_WORKBOOK_CELLS`). The same for
  `table:number-columns-repeated` on cells. `table:covered-table-cell`
  advances the column and stores nothing.
- **Values.** `office:value-type` → `float`, `percentage`, `currency`:
  `office:value`; `date`: `office:date-value`; `time`: `office:time-value`;
  `boolean`: `office:boolean-value`; `string`: `office:string-value` if
  present, else the `text:p` text. `calcext:value-type="error"` → the error
  code read from the `text:p` (unknown → `#VALUE!`). Cells with
  `table:formula` contribute their cached value.
- **Visibility.** A `table:table@table:style-name` whose automatic style has
  `style:table-properties@table:display="false"` → `"hidden"`.
- **Active sheet.** `settings.xml`'s `config:config-item
  config:name="ActiveTable"` (by name); absent → 0.

### The forward-only scanner (`src/lib/core/xmlScanner.ts`)
```ts
export type XmlScanEvent =
  | { type: "open"; name: string; local: string; ns: string | undefined;
      attrs: { name: string; local: string; ns: string | undefined; value: string }[];
      selfClosing: boolean; start: number; end: number }
  | { type: "close"; name: string; local: string; ns: string | undefined; start: number; end: number }
  | { type: "text"; text: string; start: number; end: number };   // entities decoded, CDATA unwrapped
export function scanXml(text: string, onEvent: (event: XmlScanEvent) => void | "stop"): void;
export class XmlScanError extends Error { readonly offset: number }
```
Namespaces are resolved from in-scope `xmlns` declarations; offsets are
UTF-16 indices into `text`; comments and processing instructions are
skipped; the five predefined entities and numeric character references are
decoded; a `<!DOCTYPE` is an `XmlScanError` (no producer of these formats
writes one, and refusing it closes the entity-expansion door). Returning
`"stop"` ends the scan — the cell ceiling uses it. Product-neutral, so it
lives in core; nothing here knows about spreadsheets.

### Detection and messages
`core/packageDetect.ts` gains `SPREADSHEETML_MAIN_CONTENT_TYPES` (the four
above), `XLSB_MAIN_CONTENT_TYPE`, `ODS_MIMETYPES` (the two above) and
`spreadsheetFormatOf(info): { format: SpreadsheetFormat; packageVariant?:
PackageVariant } | "xlsb" | null`. `openFile` dispatches it right after
`presentationFormatOf`. Numbers is told apart from Keynote inside the IWA
branch: an iWork package containing an `Index/CalculationEngine.iwa` entry
and no `Index/Slide` entries is Numbers (confirm both markers against
`spec116-numbers`'s Numbers source document on the operator's Mac before
implementing; a `.numbers`-named ZIP with neither marker also gets the
Numbers message, the one place the extension steers — and only toward a
better message, as `.key` already does).

| Constant | String |
| --- | --- |
| `XLSB_UNSUPPORTED_MESSAGE` | "Excel binary workbooks (.xlsb) aren't supported — save as .xlsx first" |
| `LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE` | "Old .xls files and password-protected workbooks aren't supported — save as an unprotected .xlsx first" |
| `NUMBERS_UNSUPPORTED_MESSAGE` | "Apple Numbers files can't be opened yet — in Numbers, choose File ▸ Export To ▸ Excel, then open the .xlsx" |
| `WORKBOOK_TOO_LARGE_MESSAGE` | "This workbook has more than 2,000,000 filled cells — too large to open here" |
| `SPREADSHEET_NOT_TEXT_MESSAGE` | "This is a spreadsheet file — open it from the library instead" |

### File names and labels
| `format` | `packageVariant` | Extension | Badge | Format label |
| --- | --- | --- | --- | --- |
| xlsx | — | `.xlsx` | XLSX | Excel workbook (.xlsx) |
| xlsx | macroEnabled | `.xlsm` | XLSM | Excel macro-enabled workbook (.xlsm) |
| xlsx | template | `.xltx` | XLTX | Excel template (.xltx) |
| xlsx | macroEnabledTemplate | `.xltm` | XLTM | Excel macro-enabled template (.xltm) |
| ods | — | `.ods` | ODS | OpenDocument Spreadsheet (.ods) |
| ods | template | `.ots` | OTS | OpenDocument Spreadsheet template (.ots) |

A template opens as an ordinary workbook (no gallery, unlike SPEC-046's
`.dotx`) and, being byte-preserved, downloads as the template it is.
`spreadsheetFormatLabel(format, variant)` joins `presentationFormatLabel` in
`src/lib/format.ts`; `spreadsheetDownload(model)` in
`spreadsheet/download.ts` returns `{ bytes, fileName, mimeType }`.

### Reader versions (`src/lib/spreadsheet/open.ts`)
```ts
export const SPREADSHEET_READER_VERSION = 1;
export interface ReaderUpgrade {
  from: number;                                   // upgrades a record at `from` to `from + 1`
  /** Brings ONE dirty sheet up to date by reading the newly modelled element
   *  from the sheet's retained source and merging it into the model. */
  upgradeDirtySheet(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>): void;
}
export const READER_UPGRADES: ReaderUpgrade[] = [];
export function readSpreadsheetBytes(bytes: Uint8Array, format: SpreadsheetFormat, title: string,
  parts?: Map<string, Uint8Array>): Promise<SpreadsheetModel>;
export function upgradeSpreadsheetModel(model: SpreadsheetModel): Promise<SpreadsheetModel>;
```
A spec that teaches the readers to extract something new bumps
`SPREADSHEET_READER_VERSION` and appends its `ReaderUpgrade`. On load, a
record behind the current version is upgraded: if no sheet is dirty
(SPEC-121 adds `Sheet.dirty`) it is simply re-read from `opaque`, keeping
`id`, `title`, the library's starred/trash state and each sheet's `id`
(matched by position); otherwise clean sheets are re-read and each dirty
sheet runs every pending step in order. The upgraded record is written back
once. This is the same move SPEC-079 made with `upgradePresentationModel`,
made general up front because this queue grows the readers fourteen times.

### The store
- `StoredModel = DocumentModel | PresentationModel | MediaModel |
  SpreadsheetModel`; `isSpreadsheetModel(m)` tests `kind === "spreadsheet"`;
  `getSpreadsheet(id)` narrows like `getPresentation`/`getMedia`;
  `storeSpreadsheet(model)` mirrors `storePresentation`; `documentKindOf`
  answers `"spreadsheet"`.
- **Row split.** `SHEET_ROWS_KEY_PREFIX = "sheetrows:"`; record id
  `sheetrows:<docId>:<sheetId>`, value `{ id, docId, sheetId, rows }`. The
  shell record stores each sheet with `rows: []`; `persist` writes a sheet's
  rows record only when `sheet.rows !== lastPersistedRows.get(key)` (the
  copy-on-write invariant makes that sound), on the same latch discipline as
  `persistMedia`. Every read path rejoins them; `deleteDocument` (and
  emptying the trash) deletes them; `libraryStorageUsage` counts them. A
  deleted sheet's record is deleted on the next persist.
- `preview` is recomputed by `buildPreview(model)` (`spreadsheet/preview.ts`,
  `generalText` of the active sheet's top-left 12 × 6) whenever a sheet's
  rows reference changes, so the landing page renders a card from the shell
  record alone.

### The read-only grid
`src/views/SpreadsheetView.vue` is a wiring hub under the same rule as the
other views (`SpreadsheetView.size.spec.ts`). Its parts:
`src/views/spreadsheet/useWorkbookLoad.ts` (load, upgrade, kind guard,
not-found), `useSheetViewport.ts` (scroll → visible row/column window), and
the components `src/components/spreadsheet/SpreadsheetAppBar.vue`,
`SheetGrid.vue`, `SheetTabs.vue`, plus `src/components/library/SheetMiniature.vue`.

- **DOM, not canvas.** A visible cell needs real text layout (fonts, wrap,
  rich runs), an accessibility node and a selector a test can find; the
  viewport bounds how many exist. Rendered: the cells in the viewport plus
  `GRID_OVERSCAN_ROWS = 8` and `GRID_OVERSCAN_COLS = 4`, positioned
  absolutely inside a spacer sized to the scroll extent; headers are sticky.
  SPEC-117 replaces the uniform 64 × 20 px geometry with `GridMetrics`.
- **Scroll extent.** Rows: `max(usedRange.bottom + 100, 1000)`; columns:
  `max(usedRange.right + 10, 26)`. Scrolling within one screen of the end
  grows the extent by the same margins, up to `MAX_ROWS`/`MAX_COLS`. Above
  `MAX_SCROLL_EXTENT_PX = 15_000_000` (under every target browser's maximum
  element size) the spacer stops growing and scroll offsets map onto rows
  proportionally, so a used range at row 1,000,000 is reachable everywhere.
- **ARIA.** `role="grid"`, `aria-readonly="true"`, `aria-rowcount` =
  `MAX_ROWS + 1` and `aria-colcount` = `MAX_COLS + 1` (headers included);
  header cells `role="columnheader"` / `role="rowheader"`; data cells
  `role="gridcell"` with 1-based `aria-rowindex`/`aria-colindex` offset by
  the header row/column; the grid's accessible name is the sheet name.
- Values are text nodes (`generalText`), clipped by `overflow: hidden`.

### Fixtures
Built by `spreadsheet/fixtures/build.ts`, shared by the unit suite and
`scripts/build-spreadsheet-fixtures.ts` (which ends in `process.exit(0)` —
the idle scheduler keeps Node alive). They emulate real producers, because a
fixture shaped like our own writer's output proves nothing:
- `mini-book.xlsx` (Excel-shaped): three sheets — "Summary" (strings incl.
  rich text and an `_x000D_` escape, numbers, a date serial, booleans, a
  `#DIV/0!` error, formulas with cached values including a shared formula
  and a `t="str"` result, `B2` = "Revenue"), "Data" (rows and cells **without
  `r`**, 500 rows), "Hidden" (`state="hidden"`); `mc:Ignorable="x14ac xr xr2
  xr3"`, `xr:uid` attributes, `x14ac:dyDescent` and `spans` on rows,
  `xl/calcChain.xml`, `docProps/app.xml` with `TitlesOfParts`, a theme and a
  styles part (untouched here, read by SPEC-117).
- `mini-book.ods` (LibreOffice-shaped): three tables, the second active in
  `settings.xml`; trailing `table:number-rows-repeated="1048553"` and
  `table:number-columns-repeated="16380"` runs; a repeated non-empty row;
  covered cells under a span; every `office:value-type`; `calcext:value-type`;
  `of:=` formulas with cached values; an error cell; a hidden table; no
  `table:null-date`.
- `spec116-1904.xlsx` (`date1904="1"`), `spec116-macros.xlsm` (a dummy
  `xl/vbaProject.bin` and its relationship/content type), `spec116-template.xltx`,
  `spec116-wide.ods` (a cell at column 16,000 and row 900,000).
- `spec116-numbers.xlsx` — human-generated, see the AC. The AppleScript
  (`scripts/make-numbers-fixture.applescript`) builds a two-table document
  with text, numbers, a date, a formula and a hidden-free second sheet, and
  exports it with Numbers' "Export To ▸ Excel"; like the Keynote fixture,
  `open -a Numbers` first or the script's `activate` fails with -600.

### Tests
- `core/xmlScanner.spec.ts` — events, offsets, namespaces (default and
  prefixed, rebinding in scope), entities, CDATA, `"stop"`, the DTD refusal,
  malformed input with offsets.
- `spreadsheet/xlsx/reader.spec.ts` / `ods/reader.spec.ts` — every rule in
  "Reading .xlsx" / "Reading .ods", the cell ceiling (a synthetic 2,000,001-cell
  sheet is refused without materialising it), out-of-grid addresses.
- `spreadsheet/roundTrip.spec.ts` — byte-identical parts for every fixture,
  `.ods` splice equality, extensions and MIME types per the table.
- `address.spec.ts`, `serial.spec.ts` (the 1900 boundary days 59/60/61, 1904,
  ODS durations, custom null date), `general.spec.ts`, `model.spec.ts`,
  `open.spec.ts`, `boundary.spec.ts`, the router guard table,
  `documentStore` row-split tests (a rename writes no rows record; editing one
  sheet's rows array writes exactly one), `HomeView.spreadsheet.spec.ts`,
  `views/SpreadsheetView.spec.ts`, `e2e/SPEC-116.spec.ts`.

### Specs this changes
- `specs/003-…`/`specs/067-…` — the library holds a fourth kind; every
  listed behaviour applies to it unchanged.
- `specs/009-…`/`specs/050-…` — the accepted extensions and the drop
  rejection string grow; their e2e assertions are updated with them.
- `specs/087-import-keynote.md` — an iWork package is checked for Numbers
  before the Keynote dispatch; its iWork rejection ("This iWork file isn't a
  Keynote presentation we can open") still covers Pages and undecodable
  graphs, and its e2e/unit assertions for a Numbers input (if any) move to
  the Numbers message.
- `specs/107-…` — the router's pairwise kind guards become one generic guard.
- `specs/000-roadmap.md` gains Phase 17.
