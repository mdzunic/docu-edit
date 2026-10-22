---
id: SPEC-125
title: Copy, paste, fill and move
status: draft
area: editing
complexity: complex
story_points: 13
depends_on: [SPEC-016, SPEC-122, SPEC-123, SPEC-124]
---
## Why
Most spreadsheet data arrives by paste — a range from another workbook, a
table from a web page, a block out of Excel or Google Sheets, a list someone
dropped into a chat. And most spreadsheet *structure* is made by the fill
handle: a column of dates, a row of month names, one formula dragged down a
thousand rows. After SPEC-124 the grid can be typed into, formatted and
reshaped, but Ctrl/Cmd+C copies nothing useful and a range pasted from
Excel goes nowhere. A user who has to retype a column leaves.

SPEC-016 built the text product's clipboard, and its lessons carry over: all
flavours are written in the synchronous `copy` handler; the internal payload
is untrusted input and is rebuilt field by field by a validating parser;
external HTML is imported through a whitelist, never spliced. What differs
is the cargo. A cell range is a rectangle of values, formulas, styles and
merges, and pasting it means shifting relative references, tiling over a
larger selection, refusing to cut a merge in half, and — for a cut —
re-pointing every formula in the workbook that looked at the moved cells.

**This file is the normative source for the internal cell clipboard payload
(`SheetClipboardPayload`, its wire form inside `text/html`, and
`parseClipboardPayload`), TSV and HTML interop in both directions, paste and
paste-special semantics, cut-move reference following (`RangeMove`,
`repointFormula`, `MOVE_ADJUSTERS`), the fill handle and series detection
(`detectSeries`), fill down/right, and drag-move.** Ctrl/Cmd+Enter (fill the
selection with the typed entry) is SPEC-122's. Every paste, fill and move is
one undo step.

## Acceptance criteria
- [ ] Copy (Ctrl/Cmd+C, the cell menu's "Copy", which triggers the same `copy` event through `document.execCommand("copy")` as the text product does) of a single-rectangle selection writes `text/plain` (TSV) and `text/html` (a table carrying the payload) via `ClipboardEvent.clipboardData.setData`, per "Flavours"; the copied rectangle is clipped to the sheet's used range (copying column A copies A1 to its last used row) and marked by a marching-ants outline `SheetCopyMarquee.vue` (`sheet-copy-marquee`, `data-mode="copy|cut"`, static dashes under `prefers-reduced-motion`) until Escape, another copy or cut, any edit, or the paste of a cut
- [ ] Copy refuses a multi-range selection with "That doesn't work on multiple selections" and an area over `MAX_COPY_CELLS = 1_000_000` with "That's too many cells to copy — select fewer than 1,000,000 cells"; the same cap bounds every paste, fill and move destination, with "That's too many cells to change at once"
- [ ] TSV is each cell's SPEC-118 display text at full width (no "########"), cells joined by `\t`, rows by `\r\n`, no trailing newline; a field containing a tab, newline or `"` is quoted with `"` doubled; the HTML is `<meta charset="utf-8"><table data-docu-edit-cells="…">` with one `<td>` per position (merges as `colspan`/`rowspan`), escaped display text and inline CSS from the whitelist in "HTML out" only — no `class`, no script, nothing a reader must execute
- [ ] The payload is `SheetClipboardPayload` version 1 per "The payload", JSON encoded as UTF-8 then base64 in the attribute; it carries values, canonical formulas, whole-array formulas, styles (without `sourceXf`/`sourceStyleName`), rich-text runs, relative merges, whole-row/column properties, the source workbook id, sheet name and date system, and a cut token on a cut
- [ ] `parseClipboardPayload(attribute): SheetClipboardPayload | null` treats the attribute as hostile, mirroring `src/lib/clipboard.ts`'s `parseFragmentPayload`: it checks size before decoding, parses JSON in a `try`, checks `kind`/`version`, and **rebuilds every field from a whitelist** under the bounds of "Validation" (shape, ordering, value types, error codes, string and formula lengths, formulas that `parseFormula` accepts, style fields, colours, number-format codes `parseFormatCode` accepts, merges inside the shape and disjoint); any violation returns `null` and the paste falls through to the HTML-table import — a security test feeds oversize, malformed, prototype-polluting (`__proto__`), out-of-shape and script-bearing payloads and asserts none reaches the model
- [ ] Paste (Ctrl/Cmd+V while the grid has focus; the cell editor's paste stays SPEC-122's plain-text paste) reads `text/html` first — a `table[data-docu-edit-cells]` whose payload validates, else an external table per "External sources" — then `text/plain`; the cell menu's "Paste" uses `navigator.clipboard.read()` (with `unsanitized: ["text/html"]` where supported) and, when the read HTML lacks the attribute but its `text/plain` equals the view's last copy's TSV, uses that copy's in-memory payload (SPEC-016's snapshot rule); an unavailable clipboard shows SPEC-016's toast "Clipboard unavailable — paste with the keyboard instead"
- [ ] Paste geometry per "Paste geometry": a block pastes at the active cell; one cell into a multi-cell selection tiles; an N×M block into a selection whose height and width are exact multiples tiles; any other selection pastes once at its top-left; a paste reaching past the grid edge is refused with "The paste area goes past the edge of the sheet"; a destination that partially overlaps a merge (or whose pasted merges would) is refused with "Can't paste over part of a merged cell"; a partial array is refused with SPEC-124's `PARTIAL_ARRAY_MESSAGE`; a protected destination with SPEC-122's message; a multi-range destination with "That doesn't work on multiple selections"
- [ ] Paste (All) replaces values, formulas, styles and merges in the destination rectangle per "Paste semantics": relative references shift by the paste offset through SPEC-119's `shiftFormula` (absolute ones stay; off-grid → `#REF!`), a blank source position clears its destination (value and style), a whole-row/column copy replaces whole rows/columns including heights, widths and row/column styles, same-workbook pastes keep style ids by identity while other workbooks' styles go through `internStyle`, date-formatted values convert between 1900 and 1904 date systems (±1462 days), and a formula from another workbook that names a sheet this workbook lacks is pasted as its cached value with the toast "N formulas referring to other sheets were pasted as values"
- [ ] Cut (Ctrl/Cmd+X, cell menu "Cut") writes the same flavours with a cut token and leaves the model untouched; pasting it in the same workbook view runs `moveRange` per "Cut and move" — the destination is overwritten, the source cleared to unstyled blanks, merges and whole-array formulas move, every formula in the workbook (moved ones included) that referenced cells entirely inside the source is re-pointed to the destination, references to overwritten destination cells become `#REF!`, and moved formulas are otherwise **not** shifted (Excel); a cut pastes once, any edit before the paste cancels it, and a cut pasted into another workbook (or after cancellation) pastes as a copy
- [ ] `move.ts` exports `RangeMove`, `repointFormula(text, hostSheetId, move, sheetNameOf)`, `moveMapper(model, move)` (a SPEC-124 `RetainedReferenceMapper`) and the `MOVE_ADJUSTERS` registry per "Cut and move"; this spec registers `cells`, `formulas`, `merges` and `retained` (unmodelled references re-pointed through SPEC-124's `rewriteRetainedReferences`), and `move.spec.ts` holds the contract test later adjusters join
- [ ] Paste special (`PasteSpecialDialog.vue`, `sheet-paste-special-dialog`; Ctrl/Cmd+Shift+V, alternate Ctrl+Alt+V — Control+Option+V on macOS — both by `KeyboardEvent.code`, plus the Edit menu and cell menu) offers All, Formulas, Values, Formats, Values and number formats (`sheet-paste-special-all|formulas|values|formats|values-number-formats`) and the checkboxes Skip blanks and Transpose (`sheet-paste-special-skip-blanks`, `-transpose`) with the meanings of "Paste special"; options the clipboard's source cannot supply are disabled; it reads through `navigator.clipboard.read()`, and a refused read shows "Clipboard access was blocked — allow it for this site to use Paste special"
- [ ] External HTML tables import per "External sources": Excel (`<style>` class rules read as data for `mso-number-format` and the CSS whitelist, `x:num`, `x:bool`, `x:err`, `x:str`), Google Sheets (`data-sheets-value`, `data-sheets-numberformat`, `data-sheets-formula` converted from R1C1 to A1 relative to each destination cell), LibreOffice (`sdval`, `sdnum`) and any other `<table>` (inline-CSS whitelist, text through `parseInput`), `colspan`/`rowspan` becoming merges; parsing is `DOMParser("text/html")` on a detached document, so nothing executes or loads, and the security test of SPEC-016's sanitisation list (script, style with `url()`, `on*`, `javascript:`, remote `<img>`, `iframe`) passes for this importer too
- [ ] Plain text pastes per "TSV in": quoted fields with embedded tabs, newlines and doubled quotes, `\r\n`/`\n`/`\r` rows, a trailing empty line ignored, each field through SPEC-122's `parseInput(text, locale, dateSystem)` (so "1,234.50" in en-US is the number 1234.5 and "=A1+1" a formula), destination styles kept except the number format `parseInput` implies; a single value pastes into the active cell (or tiles over the selection)
- [ ] The fill handle (`SheetFillHandle.vue`, `sheet-fill-handle`, a 7×7px square on the selection's bottom-right corner, `cursor: crosshair`) drags along the axis the pointer moved further on, outlines the target (`sheet-fill-preview`) and on release runs `fillRange(model, sheetId, seed, target, mode)`; holding Ctrl (Windows/Linux) or Alt/Option (every platform) at release toggles copy ↔ series; dragging back inside the selection clears the contents (not styles) of the cells left behind; Escape cancels
- [ ] `detectSeries` implements every row of "Series detection": per seed line (each column of a vertical fill, each row of a horizontal one), numbers (one → copy, two or more → least-squares linear trend, results rounded to 15 significant digits), dates by day, by month (EDATE, month-ends clamped) or by year, times by hour, weekday and month names (full and short, English plus the user's SPEC-118 locale, case kept), `Q1`–`Q4`/`Qtr 1`/`Quarter 1` with an optional year, text with a trailing integer (zero padding kept) and ordinals, formulas shifted by their distance from the seed cell, and everything else repeating the seed pattern; filling up or left extends the series backwards; every example in the table is a unit test
- [ ] After a drag fill, the fill options chip (`FillOptionsChip.vue`, `sheet-fill-options`, `role="toolbar"`) offers "Copy cells", "Fill series", "Fill formatting only" and "Fill without formatting" (`sheet-fill-copy|series|formats|no-formats`) with the applied one pressed; choosing one re-applies the fill as the **same** undo step; Escape or any other command dismisses it
- [ ] Double-clicking the fill handle fills down to the extent rule of "Double-click fill" (the adjacent column's contiguous data, stopping before any non-empty cell below the selection) and does nothing when there is no adjacent data
- [ ] Ctrl/Cmd+D fills down and Ctrl/Cmd+R fills right (`SHEET_SHORTCUTS` `sheet.fillDown`, `sheet.fillRight`; Edit ▸ Fill ▸ Down/Right): each column's top cell (each row's left cell) is copied into the rest of the selection with formulas shifted and styles copied, never a series; a one-row (one-column) selection copies from the row above (column to the left)
- [ ] Drag-move (`useRangeMove.ts`): hovering within 4px of the selection's outline (not the fill handle) shows `cursor: move`; dragging outlines the block at its snapped destination (`sheet-move-preview`, tooltip with the destination address); dropping runs `moveRange` (cut semantics), or copy-paste semantics when Alt/Option (or Ctrl on Windows/Linux) is held at drop; dropping onto non-empty cells outside the source asks "There's already data here. Do you want to replace it?" (`sheet-move-overwrite`, "Replace"/"Cancel"); a drop partially over a merge is refused like a paste; Escape or a drop outside the grid cancels
- [ ] Every paste, paste special, cut-paste, fill, fill-chip re-application, fill down/right and drag-move is one undo step through SPEC-121's `record(label, scope)` ("Paste", "Paste special", "Move cells", "Fill", "Fill down", "Fill right"; a cut across sheets records both), runs SPEC-122's row refit (`refitRows`, which SPEC-124 extends to pastes and fills), marks every touched sheet dirty and hands changed formula cells to recalculation; the marching ants are view state and never enter the undo stack
- [ ] `SpreadsheetView.vue` gains only `useSheetClipboard({ … })`, `useFillHandle({ … })` and `useRangeMove({ … })` wiring plus hookup; the composables live in `src/views/spreadsheet/` (each ≤ 400 lines), the components `SheetCopyMarquee.vue`, `SheetFillHandle.vue`, `FillOptionsChip.vue`, `PasteSpecialDialog.vue` and `SheetMoveOverwriteDialog.vue` in `src/components/spreadsheet/`, and `SpreadsheetView.size.spec.ts`'s budget is not raised
- [ ] Fixtures: real clipboard HTML captured from each producer is committed under `e2e/fixtures/spec125-clipboard/` — `excel.html` (Excel desktop, or Excel for the web), `sheets.html` (Google Sheets), `libreoffice.html` (LibreOffice Calc), `numbers.html` (Apple Numbers) — each a range holding text, a number with a thousands format, a currency, a percentage, a date, a boolean, a formula and a merged pair, captured on the operator's Mac by `scripts/capture-clipboard-html.sh` (`osascript -e 'the clipboard as «class HTML»'`, hex-decoded) with a `NOTES.txt` recording producer, version and date; all four are committed **before this spec is labelled ready** (the factory can produce none of them; the operator's Mac has Numbers and a browser for Sheets and Excel for the web, and LibreOffice is the operator's to install), and `spreadsheet/fixtures/build.ts` adds `spec125-fill.xlsx` and `.ods` (seed columns for every series kind, cross-sheet formulas, a merge)
- [ ] Tests: `spreadsheet/clipboard.spec.ts` (payload build/parse round trip, every validation bound, the security cases, TSV both ways, HTML out, each external fixture to the expected cells), `clipboard.paste.spec.ts` (geometry, semantics, paste special incl. transpose and skip blanks, date-system conversion, cross-workbook rules), `move.spec.ts` (re-point table, `#REF!` on overwrite, overlapping source/destination, cross-sheet moves, the registry contract), `fill.spec.ts` (series table, reverse fills, chip modes, double-click extent, fill down/right), `useSheetClipboard.spec.ts`, `useFillHandle.spec.ts`, `useRangeMove.spec.ts`
- [ ] e2e (`e2e/SPEC-125.spec.ts`, Chromium with `clipboard-read`/`clipboard-write` granted): open `spec125-fill.xlsx` → copy B2:C3 (marching ants show) → paste at E5 → `E6`'s formula bar shows the shifted formula → Escape clears the ants → cut A10:A12 → paste at H10 → the Sheet2 formula that summed A10:A12 now reads `H10:H12` → write `excel.html` to the clipboard with `navigator.clipboard.write` → Ctrl/Cmd+V → the currency cell is the number 1234.5 with its format → Ctrl/Cmd+Shift+V → Values + Transpose → drag the fill handle from a cell reading "Mon" down six rows → "Tue"…"Sun" → choose "Copy cells" in the chip → every cell reads "Mon" → double-click the fill handle beside a 20-row column → it fills 20 rows → Ctrl/Cmd+D → Alt/Option-drag a block to copy it → undo three times → download → reopen → the cells as left

## Out of scope
Paste operations (add, subtract, multiply, divide), paste link, "Column
widths", "Comments" and "Validation" paste-special options, a paste options
chip after an ordinary paste — backlog. Insert cut/copied cells (shifting
the destination) and Shift-drag insert (backlog, with SPEC-124's shifted
inserts). Copying multi-range selections. Images, shapes and charts on the
clipboard (SPEC-131/132). Clipboard history. Fill series dialog (step, stop
value, growth trends) and Flash Fill. Drag-move across sheets. Pasting into
filtered ranges and copying only visible cells (SPEC-128 adds that rule).
Ctrl/Cmd+Enter fill (SPEC-122).

## Reference

### Dependencies restated
- `specs/016-cut-copy-paste.md` / `src/lib/clipboard.ts` — flavours are
  written in the synchronous `copy`/`cut` handler with
  `clipboardData.setData` and read with `getData` in `paste`; the internal
  payload is validated (`parseFragmentPayload`: `JSON.parse` in a `try`,
  `mimeType`/`version` checked, every block rebuilt from whitelisted fields,
  `null` on any doubt → fall through to `text/html`); the context menu
  pastes through the async API and recognises its own copy by comparing the
  clipboard with the last copy's snapshot; sanitisation drops active content,
  `on*`, `url()`/`javascript:`/`data:` values and remote images, enforced by
  a security test; `src/lib/htmlPaste.ts` parses with a detached
  `DOMParser`.
- `specs/116-…` — `SpreadsheetModel { id; sheets; dateSystem; … }`, `Cell {
  col; value; formula?; styleId?; runs? }`, `CellValue`, `ErrorCode` (the
  fifteen codes), `MAX_ROWS`, `MAX_COLS`, the invariants (copy-on-write rows,
  sparse, plain JSON), `GridRange`, `CellRef`, `usedRange`.
- `specs/117-…` — `CellStyle`, `TextRun { text; font?: CellFont }`,
  `ColumnSpan`, `Row.heightPt|customHeight|hidden|styleId`, `Sheet.merges`,
  `SheetSelection`. `specs/118-…` — `formatValue(value, code, ctx)`,
  `parseFormatCode`, the locale's month and day names, date/time code
  classification. `specs/119-…` — `parseFormula` (throws
  `FormulaSyntaxError`), `shiftFormula(text, dRow, dCol)` (relative refs
  only, off-grid → `#REF!`), `Cell.arrayRange`, `formulaUnparsed`, and the
  reference-rewrite primitive `rewriteReferences` (as SPEC-124 restates it).
- `specs/121-…` — `record(label, scope)`, `Sheet.dirty`. `specs/122-…` —
  `internStyle`, `parseInput(text, locale, dateSystem, options?): { value;
  formula?; numberFormat?; builtinFormatId?; quotePrefix? }`, protection
  honour, the cell editor's own paste, Ctrl/Cmd+Enter, and the automatic row
  refit `refitRows(model, sheetId, ranges, measurer, ctx)` in `autofit.ts`.
  `specs/123-…` — `patchStyle`, `applyStylePatch`,
  `SHEET_SHORTCUTS` with `code`/`ctrl` bindings, `SheetCellContextMenu.vue`.
  `specs/124-…` — `adjustRange`, `PARTIAL_ARRAY_MESSAGE`,
  `RetainedReferenceMapper { range; index; formula }`,
  `rewriteRetainedReferences(model, mapper)`, the refit's paste/fill trigger.

### Flavours
| Flavour | Content | Why |
| --- | --- | --- |
| `text/plain` | TSV of display text | every app reads it |
| `text/html` | a `<table data-docu-edit-cells="<base64 JSON>">` with display text and inline CSS | other apps get a formatted table; this app gets the payload from the same flavour |
The payload rides inside `text/html` rather than a private MIME type (as
SPEC-016's does) because the async Clipboard API — the cell menu's paste,
paste special — exposes only `text/plain`, `text/html` and images (custom
formats need Chromium-only `web ` types); one standard flavour carries it
on every path. **HTML out** (whitelist): `font-family`, `font-size` (pt),
`font-weight:700`, `font-style:italic`, `text-decoration` (underline,
line-through), `color`, `background-color`, `text-align`,
`vertical-align`, `white-space:pre-wrap` for wrapped cells, and
`border-top|right|bottom|left` (`1px`/`2px`/`3px` `solid|dashed|dotted|double`
+ colour) — theme colours resolved to hex.

### The payload
```ts
// src/lib/spreadsheet/clipboard.ts
export interface SheetClipboardPayload {
  kind: "docu-edit/cells";
  version: 1;
  workbookId: string;                 // source SpreadsheetModel.id
  sheetName: string;
  dateSystem: "1900" | "1904";
  origin: CellRef;                    // source top-left
  rows: number; cols: number;         // the copied rectangle after used-range clipping
  wholeRows?: true; wholeCols?: true;
  cells: ClipboardCell[];             // sparse, row-major ascending, unique (r, c)
  styles: CellStyle[];                // distinct styles referenced; no sourceXf/sourceStyleName
  merges: GridRange[];                // relative to origin
  rowProps?: { r: number; heightPt?: number; customHeight?: true; hidden?: true; s?: number }[];  // wholeRows
  colProps?: { c: number; widthPx?: number; hidden?: true; s?: number }[];                        // wholeCols
  cut?: string;                       // the cut token
}
export interface ClipboardCell {
  r: number; c: number;               // relative to origin
  v: CellValue;                       // the value, or a formula's cached value
  f?: string;                         // canonical formula
  a?: GridRange;                      // relative arrayRange — only when the whole array was copied
  s?: number;                         // index into styles
  t?: TextRun[];                      // rich text
}
export const MAX_COPY_CELLS = 1_000_000;
export function buildClipboardPayload(model, sheetId, range, cutToken?): SheetClipboardPayload;
export function clipboardFlavours(model, sheetId, range, cutToken?): { text: string; html: string };
export function parseClipboardPayload(attribute: string): SheetClipboardPayload | null;
```
An array member copied without its anchor contributes its cached value only.

### Validation
| Field | Rule |
| --- | --- |
| attribute | ≤ 96 MiB of base64 (≈ 64 MiB of JSON), checked before `atob`; UTF-8 decoded with `fatal: true` |
| top level | a plain object; `kind === "docu-edit/cells"`, `version === 1`; unknown keys ignored; nothing is copied by spread — every field is read and re-created |
| `rows`, `cols` | integers ≥ 1, `rows × cols ≤ MAX_COPY_CELLS`; `origin` inside the grid |
| `cells` | ≤ `rows × cols` entries, integer `r < rows`, `c < cols`, strictly increasing row-major |
| `v` | finite number, string ≤ 32,767 characters, boolean, `null`, or `{ error }` with a known `ErrorCode` |
| `f` | string ≤ 8,192 characters that `parseFormula` accepts — otherwise dropped and `v` kept |
| `a` | inside the shape and anchored at its own cell, else dropped |
| `s` | integer `< styles.length` |
| `styles` | ≤ 64,000; each rebuilt field by field: font family ≤ 255 printable characters, `sizePt` 1–409, booleans only `true`, colours `#rrggbb` or `{ theme: <slot> }` with `lumMod`/`lumOff`/`shade`/`tint` in [0, 1], border styles from `SheetBorderStyle`, alignment enums, `indent` 0–250, `textRotation` 0–180 or 255, `numberFormat` ≤ 255 characters that `parseFormatCode` accepts, `namedStyle` ≤ 255 characters |
| `t` | ≤ 32,767 characters in total, fonts as above |
| `merges` | inside the shape, at least two cells, pairwise disjoint |
| `rowProps`/`colProps` | only with the matching flag; indices inside the shape; sizes within SPEC-124's dialog bounds |
| `workbookId`, `sheetName`, `cut` | strings ≤ 255 (`sheetName` 1–31) |

### Paste geometry
| Clipboard | Destination selection | Result |
| --- | --- | --- |
| N×M | one cell | pasted at it |
| 1×1 | any single range | tiled over it |
| N×M | a range whose height is a multiple of N and width of M | tiled |
| N×M | any other single range | once, at its top-left |
| whole rows/columns | any | onto whole rows/columns starting at the active cell's row/column; Transpose disabled |
| any | several ranges | refused |
The destination rectangle is the union of the tiles; it is checked against
the grid edge, merges (a merge entirely inside is replaced; one crossing
the edge refuses), arrays and protection before anything changes.

### Paste semantics
| Option | Values / formulas | Styles | Merges |
| --- | --- | --- | --- |
| All (default paste) | replaced; formulas `shiftFormula`'d by (destination − origin) | replaced (blank source position → default) | replaced |
| Formulas | formulas shifted, constants as values | destination kept | kept |
| Values | cached values, formulas and runs dropped | kept | kept |
| Formats | kept | replaced (+ row/column properties for whole rows/columns) | replaced |
| Values and number formats | as Values | only `numberFormat`, merged with `patchStyle` | kept |
| Skip blanks | a source position with no value and no formula leaves its destination untouched under every option |
| Transpose | rows ↔ columns; in each formula (through `rewriteReferences`) a reference relative on both axes has its offset from the host transposed — (dRow, dCol) → (dCol, dRow) — so a block that referred to itself still does; absolute and mixed references shift by the host's displacement as in an ordinary paste; merges transposed |
External sources supply values (and Excel/Sheets/LibreOffice formats,
Sheets formulas); options they cannot supply are disabled.

### Cut and move
```ts
// src/lib/spreadsheet/move.ts
export interface RangeMove { fromSheetId: string; from: GridRange; toSheetId: string; toTop: number; toLeft: number }
export interface MoveAdjuster { id: string; move(model: SpreadsheetModel, move: RangeMove): void }
export const MOVE_ADJUSTERS: MoveAdjuster[];
export type MoveResult = { ok: true } | { ok: false; message: string };
export function moveRange(model: SpreadsheetModel, move: RangeMove): MoveResult;
export function repointFormula(text: string, hostSheetId: string, move: RangeMove, sheetNameOf: (id: string) => string): string;
export function moveMapper(model: SpreadsheetModel, move: RangeMove): RetainedReferenceMapper;
```
| A reference anywhere in the workbook, moved formulas included | becomes |
| --- | --- |
| entirely inside the source (cell or range) | moved by the offset, onto the destination sheet (prefix added, changed or dropped to match the host) |
| range partly overlapping the source | unchanged |
| entirely inside the destination and not in the source | `#REF!` (its cells were overwritten) |
| whole-row/column, 3-D, external, anything else | unchanged |
Source and destination may overlap (moving a block down one row); the
move reads the source before writing. Registered: `cells` (values,
formulas, styles, runs, array anchors; the source left as unstyled blanks),
`formulas` (`repointFormula` over every sheet), `merges` (source merges
move, destination merges replaced), `retained` (SPEC-124's
`rewriteRetainedReferences` with `moveMapper`: an `sqref` entry entirely
inside the source moves, one entirely inside the destination is dropped,
partial overlaps stay). Later specs that register a SPEC-124
`ReferenceAdjuster` for something that travels with cells or points at them
— the autofilter (128), conditional formats (129), validations (130), chart
references (132), defined names and tables (133), notes and hyperlinks
(135), pivots (139) — register a `MoveAdjuster` with the same id; drawings
(131) and page breaks (136) deliberately do not.

### External sources
| Producer (detected by) | Value | Number format | Style | Formula |
| --- | --- | --- | --- | --- |
| Excel (`xmlns:x="urn:schemas-microsoft-com:office:excel"`) | `x:num="…"` → number (bare `x:num` → `parseInput` of the text); `x:bool`; `x:err`; `x:str` → string; else `parseInput` | `mso-number-format` from the cell's class rule or inline style, CSS escapes decoded (`\0022` → `"`, `\X` → `X`) | class rules (`.xl65 {…}`, `td {…}`) from the `<style>` text, read as data, plus inline CSS, through the whitelist below | not read (`x:fmla` positions are unknown) |
| Google Sheets (`data-sheets-root` / `google-sheets-html-origin`) | `data-sheets-value` JSON: `"1"` type (2 string, 3 number, 4 boolean), `"2"`/`"3"`/`"4"` the value | `data-sheets-numberformat` `"2"` | inline CSS | `data-sheets-formula` R1C1 (`R[-1]C[0]`, `R2C3`) → A1 relative to each destination cell; unconvertible → value |
| LibreOffice (`sdval`/`sdnum` attributes) | `sdval` → number; else `parseInput` | `sdnum="lang;index;code"` → `code` when non-empty | inline CSS and `<font>`/`<b>`/`<i>`/`<u>` | not available |
| Any other `<table>` (Numbers, web pages) | `parseInput` of the text | from `parseInput` | inline CSS, `<b>`/`<i>`/`<u>`/`<s>` | not available |
CSS whitelist in: `font-weight` (≥ 600 or `bold`), `font-style:italic`,
`text-decoration` (underline, line-through), `text-underline-style:double`,
`color`, `background`/`background-color`, `font-family` (first family,
quotes stripped), `font-size` (pt, px × 0.75), `text-align`,
`vertical-align` (top/middle/bottom), `white-space:normal` (Excel's wrap),
`mso-rotate`, `border*` (`.5pt` → thin, `1.0pt` → medium, ≥ `1.5pt` →
thick, `double`, `dotted`, `dashed`, `hairline` → hair). Colours through
`core/colors.ts` `parseColorInput` plus the CSS basic names, `windowtext` →
automatic. Text is the cell's `textContent` with `<br>` as `\n`. Input over
64 MiB or over `MAX_COPY_CELLS` cells is refused before parsing.

### Series detection
`detectSeries(seed: { value: CellValue; formula?: string; code?: string }[], locale): Series`
| Seed line (in fill order) | Continuation |
| --- | --- |
| `5` | `5, 5, …` (toggle: `6, 7, …`) |
| `1, 3` | `5, 7, …` |
| `1, 2, 4` | least squares: `5.33333333333333, 6.83333333333333` |
| `0.1, 0.2` | `0.3, 0.4` (15 significant digits — never `0.30000000000000004`) |
| `2026-01-31` (date) | `2026-02-01, …` (toggle: copy) |
| `2026-01-31, 2026-02-28` | month-end: `2026-03-31, 2026-04-30` |
| `2026-01-15, 2026-02-15` | `2026-03-15, …` (EDATE from the first seed) |
| `2024-03-01, 2025-03-01` | `2026-03-01, …` |
| `2026-01-01, 2026-01-08` | weekly: `2026-01-15, …` |
| `9:00` (time-only format) | `10:00, 11:00` |
| `Mon` / `MONDAY` / `monday` | `Tue` / `TUESDAY` / `tuesday` …, wrapping |
| `Jan, Apr` | `Jul, Oct, Jan` |
| `Montag` (de-DE locale) | `Dienstag` |
| `Q3 2026` / `Qtr 4` | `Q4 2026, Q1 2027` / `Qtr 1` |
| `Item 9` / `Item 009` / `Item 1, Item 3` | `Item 10` / `Item 010` / `Item 5` |
| `1st` | `2nd, 3rd, 4th, …, 11th, 12th, 13th, 21st` |
| `=A1*2` at B1 | `=A2*2` at B2 (shifted by the distance) |
| `TRUE`, `#N/A`, `abc`, `Item 1, Thing 2`, mixed kinds | the seed pattern repeats |
Dates (a seed whose code SPEC-118 classifies as a date) try, in order:
equal year steps on the same month and day; equal month steps on the same
day of month, or all on month-ends; equal whole-day steps; else the linear
trend on serials. Name cycles need a constant integer step between seeds,
else the pattern repeats.
Each filled cell takes the style of its seed position (k mod n). Chip
modes: Copy cells (pattern repeat), Fill series (the detected series; for a
non-series seed, copy), Fill formatting only, Fill without formatting
(values/formulas, destination styles kept).

### Double-click fill
With selection rows `top..bottom` and columns `left..right`: the guide
column is `left−1` if the cell at (`bottom+1`, `left−1`) is non-empty, else
`right+1` if (`bottom+1`, `right+1`) is; none → nothing. The fill ends at
the last row of the guide column's contiguous non-empty run from
`bottom+1`, or the row before the first non-empty cell below the selection
in any of its own columns, whichever comes first.

### Shortcuts
| Id (`SHEET_SHORTCUTS`) | Windows / Linux | macOS |
| --- | --- | --- |
| `sheet.pasteSpecial` | Ctrl+Shift+V; Ctrl+Alt+V | Cmd+Shift+V; Control+Option+V |
| `sheet.fillDown` / `sheet.fillRight` | Ctrl+D / Ctrl+R | Cmd+D / Cmd+R |
Ctrl/Cmd+D and Ctrl/Cmd+R are the browser's bookmark and reload keys; they
are `preventDefault`ed only while the grid has focus (as Google Sheets does),
and Edit ▸ Fill ▸ Down/Right is the menu path. Copy, cut and paste are the
native `copy`/`cut`/`paste` events on the
grid's focus element, not registry entries. Fill-handle and drag-move
modifiers: Ctrl (Windows/Linux) or Alt/Option. Escape clears the ants and
cancels a pending cut.

### Tests
Unit and e2e files per the ACs. Relations: field-level `toEqual` on the
model after paste/fill/move and on re-read sheets after download; identity
checks for untouched rows (copy-on-write) and after undo; part-level
`equalBytes` for parts a paste did not touch.

### Specs this changes
- SPEC-016's approach (synchronous flavours, validating parser, snapshot
  recognition, whitelist import) gains a spreadsheet sibling; the text
  product's clipboard is unchanged.
- SPEC-124's retained-reference machinery gains a second mapper
  (`moveMapper`); its later `ReferenceAdjuster` owners 128, 129, 130, 132,
  133, 135 and 139 also register `MoveAdjuster`s (131 and 136 do not).
- SPEC-122's `refitRows` runs after pastes and fills.
- SPEC-123's cell context menu gains Cut, Copy, Paste and Paste special…;
  the Edit menu gains Paste special… and Fill ▸ Down/Right.
- SPEC-128 adds "copy visible cells only" for filtered ranges.
