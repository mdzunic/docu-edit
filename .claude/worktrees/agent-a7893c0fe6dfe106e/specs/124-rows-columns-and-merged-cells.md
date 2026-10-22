---
id: SPEC-124
title: Rows, columns and merged cells
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-117, SPEC-119, SPEC-122, SPEC-123]
---
## Why
After SPEC-123 a sheet can be typed into and formatted but not reshaped: no
row for the forgotten month, no deleting an obsolete column, no widening the
column whose totals read "########", no title merged across a table. These
are the most frequent structural edits a spreadsheet gets, and the most
dangerous ones. Every reference in a workbook — formulas on every sheet,
merges, frozen panes, conditional formats, validations, defined names,
charts, tables, pivots, comments — points at row and column numbers, and an
insert that moves the cells but not the references corrupts the workbook
silently, with every number still looking plausible.

Half of those references live in XML this product does not model yet. The
splice-base invariant (SPEC-116) re-emits unmodelled elements verbatim —
which, after a row insert, means at their old positions: a conditional
format would colour the wrong row, a print area would cut the wrong page, a
chart would plot the wrong series. So a structural edit must also rewrite
the reference-bearing text in the retained source, element by element,
until the spec that models each element takes it over.

**This file is the normative source for `StructuralEdit`, `adjustIndex`/
`adjustRange`, `adjustFormula` (structural adjustment of formula text,
`formula/adjust.ts`), the `REFERENCE_ADJUSTERS` registry every
reference-bearing structure joins, and `rewriteRetainedReferences` — the
textual adjustment of unmodelled reference-bearing XML in both formats —
plus the insert/delete/resize/autofit/hide commands, merges and the
freeze-panes UI.** Inserting or deleting cells with a shift (rather than
whole rows or columns) is backlog. Commands mark sheets dirty; SPEC-122
serialises them.

## Acceptance criteria
- [ ] `src/lib/spreadsheet/structure.ts` exports `StructuralEdit`, `adjustIndex(index, edit)` and `adjustRange(range, edit)` per "Geometry of an edit" — pure functions, every row of the table a unit test, `null` meaning "no longer exists"
- [ ] `src/lib/spreadsheet/formula/adjust.ts` exports `adjustFormula(text, hostSheetId, edit, sheetNameOf): string` per "Formula adjustment": token-level over SPEC-119's reference-rewrite primitive (trivia, `$` markers, sheet-name quoting and every non-reference token survive byte-for-byte), absolute and relative references both move, a reference into a deleted band becomes `#REF!`, a range straddling a deletion shrinks, whole-row and whole-column references follow their axis, references on **other** sheets that name the edited sheet are adjusted (sheet names compared case-insensitively), 3-D and external references are left alone, `formulaUnparsed` text is never touched, and an unaffected formula returns the identical string — every example in the table is a unit test
- [ ] `structure.ts` exports `ReferenceAdjuster { id; adjust(model, edit): void }`, `REFERENCE_ADJUSTERS` and `applyStructuralEdit(model, edit): StructuralResult` per "The adjuster registry": refusals are checked before anything changes, then the adjusters run in registry order under copy-on-write (adjusters replace arrays and objects, never mutate one a snapshot may hold); this spec registers `rows`, `formulas`, `merges`, `frozen`, `selection`, `columns` and `retained`, and `structure.spec.ts` holds the contract test every later adjuster joins (no structure of its kind → model untouched by identity; a row inserted above its structure moves it)
- [ ] Inserting rows/columns shifts everything at or after `at`, and the new rows take the row above's `styleId`, `heightPt` and `customHeight` and its non-zero cell styles as styled empty cells (columns: the column to the left; at index 0: the old first row/column), bounded by SPEC-123's `MAX_FORMAT_MATERIALISED` (beyond it only row/column properties are copied)
- [ ] An insert that would push a cell holding a value or formula past `MAX_ROWS`/`MAX_COLS` is refused with `INSERT_PUSHES_DATA_MESSAGE` ("Can't insert — that would push filled cells off the end of the sheet"); formatting-only rows, cells and merges that fall off the end are dropped; an edit that would split an array formula's range (an insert strictly inside it, a delete band partly inside it) is refused with `PARTIAL_ARRAY_MESSAGE` ("You can't change part of an array"); a protected sheet refuses unless its protection permits the operation (`sheetProtection@insertRows|insertColumns|deleteRows|deleteColumns="0"`), with SPEC-122's protected message
- [ ] Deleting rows/columns removes the band's cells, row/column properties and merges inside it, shrinks merges straddling it (a merge reduced to one cell disappears), and turns every reference into it `#REF!`; deleting every row of a sheet leaves an empty sheet
- [ ] Right-clicking a row or column header (`sheet-row-header`/`sheet-col-header`; Shift+F10 or the Menu key when the selection is whole rows/columns — there it opens this menu instead of SPEC-123's cell context menu) opens `SheetHeaderContextMenu.vue` (`sheet-header-menu`) with the entries and labels of "Header menu"; right-clicking a header outside the selection first selects that row or column, and a selection holding several row bands deletes or hides every band as one step
- [ ] Ctrl+Alt+= / Cmd+Option+= (`sheet.insertCells`) inserts and Ctrl+Alt+- / Cmd+Option+- (`sheet.deleteCells`) deletes whole rows/columns when the selection is whole rows or whole columns (inserting as many as are selected, above/left); on any other selection they open `SheetInsertDeleteDialog.vue` (`sheet-insert-dialog`/`sheet-delete-dialog`) offering "Entire row" and "Entire column" (`sheet-insert-delete-entire-row|entire-column`); both bind by `KeyboardEvent.code` (`Equal`, `Minus`, so Option's dead characters on macOS do not matter) through SPEC-123's `SHEET_SHORTCUTS` — Google Sheets' binding, because Ctrl/Cmd+=, Ctrl/Cmd+- and Ctrl/Cmd+Shift+= (`+`) are SPEC-117's sheet zoom (SPEC-063's `view.zoomIn`/`zoomOut`), which keeps them
- [ ] Dragging a header boundary resizes per "Sizes": a `col-resize`/`row-resize` cursor within 3px of a boundary, a guide line `sheet-resize-guide` and the tooltip "Width: 8.43 (64 pixels)" / "Height: 15.00 (20 pixels)" while dragging, Escape cancels, release commits one undo step, a release under 4px hides the column/row (keeping its size for unhide), and when the dragged header belongs to a whole-column/row selection every selected column/row takes the new size
- [ ] Double-clicking a header boundary autofits per "Autofit" (every selected column/row when the boundary belongs to the selection), as do the header menu's "Autofit column width"/"Autofit row height" and the Format menu; measurement goes through SPEC-122's injectable `CellTextMeasurer`, so unit tests are deterministic (`fixedMeasurer(charPx, linePx)`) and the view supplies SPEC-122's `createCanvasCellMeasurer()`
- [ ] SPEC-122's automatic row refit (`refitRows`, which already runs on cell entry, Fill and Clear contents) also runs when a SPEC-123 formatting command changes a font family/size, wrap or rotation and when a later command pastes or fills values (SPEC-125), as part of that command's undo step and under SPEC-122's limits (rows without `customHeight` holding a stored non-empty cell in the touched ranges, the first `AUTOFIT_MAX_ROWS_PER_COMMAND` of them)
- [ ] "Row height…" (`sheet-row-height-dialog`, `sheet-row-height-input`) edits points (0–409, rounded to the nearest 0.75 pt, 0 hides, setting `customHeight`) and "Column width…" (`sheet-column-width-dialog`, `sheet-column-width-input`) edits Excel characters (0–255, two decimals, converted through SPEC-117's MDW formula, the typed value kept in `sourceWidthChars`, 0 hides), each showing the pixel equivalent ("= 64 pixels") and opening blank when the selected rows/columns differ
- [ ] Hide and unhide set `Row.hidden` (a property-only `Row` where none exists) and `ColumnSpan.hidden` without touching sizes; "Unhide" is enabled when the selection spans a hidden row/column (select-all reaches a hidden first row); hiding more than `MAX_FORMAT_MATERIALISED` rows that do not exist yet is refused with "That's too many rows to change at once"
- [ ] Merges per "Merges": `mergeCells(model, sheetId, ranges, "centre" | "across" | "cells")` and `unmergeCells(model, sheetId, ranges)`; a range partially overlapping an existing merge **extends** to cover it (never refused); when a cell other than the kept upper-left one holds a value or formula, `SheetMergeWarningDialog.vue` (`sheet-merge-warning`) says "Merging keeps only the upper-left value" with "Merge" (`sheet-merge-confirm`) and "Cancel" (`sheet-merge-cancel`); the split button `sheet-merge` (pressed when the active cell is merged, its primary action then Unmerge) and `sheet-merge-menu` (`sheet-merge-centre`, `sheet-merge-across`, `sheet-merge-cells`, `sheet-unmerge`) fill SPEC-123's band-6 slot
- [ ] View ▸ Freeze offers "Top row" (`sheet-menu-freeze-top-row`), "First column" (`sheet-menu-freeze-first-column`), "Freeze at current cell" (`sheet-menu-freeze-at-cell`, disabled at A1) and "Unfreeze" (`sheet-menu-unfreeze`, enabled only when frozen); freezing scrolls the grid to its top-left first so `Sheet.frozen` counts from row and column 0, marks the sheet dirty and is one undo step ("Freeze panes")
- [ ] Unmodelled references are rewritten per "Retained references": every row of the xlsx and ods tables is adjusted by `rewriteRetainedReferences(model, mapper)` through the `retained` adjuster, by `core/xmlScanner.ts` offsets, splicing only the changed attribute values and text nodes — an adjusted part differs from its source only inside those spans (asserted by a span-diff helper), a part with no change keeps its object identity, and an untouched workbook still downloads byte-identical
- [ ] `opaque`, `odsContentEnvelope` and `sourceXml` stay the splice base but hold the **adjusted** source after a structural edit (copy-on-write: a new `OpaquePart`, a new `opaque` array, new envelope/strings), so a later spec that models an element and upgrades a dirty sheet reads current positions; SPEC-121's undo snapshot is widened to keep `opaque`, `odsContentEnvelope` and each sheet's `sourceXml` by reference, and undo restores them exactly
- [ ] Every command is one undo step through `record(label, scope)` with the labels of "Header menu", "Merge cells", "Unmerge cells", "Resize column(s)", "Resize row(s)", "Autofit", "Freeze panes"; a structural edit drops `xl/calcChain.xml` (SPEC-121), marks the edited sheet and every sheet whose formulas or retained text changed dirty, and hands the changed formula cells to recalculation (SPEC-120); the view's live `SheetSelection` is adjusted with `adjustRange` after the command
- [ ] Round trip: the e2e edits on `spec124-structure.xlsx`/`.ods`, downloaded and reopened, reproduce the adjusted formulas, merges, frozen panes, sizes and hidden flags (field-level `toEqual` on the re-read sheets), every retained reference in the tables at its adjusted position, and every part the edit did not touch `equalBytes` with the source
- [ ] `SpreadsheetView.vue` gains only `useStructure({ … })`, `useHeaderResize({ … })` and `useMerges({ … })` wiring plus hookup; logic lives in those composables under `src/views/spreadsheet/` (each ≤ 400 lines) and the components `SheetHeaderContextMenu.vue`, `SheetInsertDeleteDialog.vue`, `SheetSizeDialog.vue`, `SheetMergeButton.vue`, `SheetMergeWarningDialog.vue`; `SpreadsheetView.size.spec.ts`'s budget is not raised
- [ ] Fixtures per "Fixtures" built by `spreadsheet/fixtures/build.ts` into `e2e/fixtures/spec124-structure.xlsx` and `spec124-structure.ods`
- [ ] Tests: `structure.spec.ts` (geometry table, every command, refusals, inheritance, the registry contract), `formula/adjust.spec.ts` (the example table plus a randomised check that insert-then-delete of the same band restores every formula not turned into `#REF!`), `xlsx/retainedReferences.spec.ts` and `ods/retainedReferences.spec.ts` (every table row, element removal, span-only diffs), `structureRoundTrip.spec.ts`, `useStructure.spec.ts`, `useHeaderResize.spec.ts`, `useMerges.spec.ts`
- [ ] e2e (`e2e/SPEC-124.spec.ts`): open `spec124-structure.xlsx` → right-click row header 3 → "Insert 2 rows above" → on Sheet2 the cell that read `=Sheet1!A5` reads `=Sheet1!A7` in `formula-bar-content` → delete column B from its header menu → a formula that referenced B5 shows `#REF!` → drag column C's boundary wider → double-click column D's boundary (it grows to its longest text) → hide row 10, unhide it → select A1:C1 with a value in B1 → Merge & Centre → the warning → Merge → View ▸ Freeze ▸ Top row → undo, redo → download → reopen → merges, freeze and formulas as left, and the downloaded package's conditional-format `sqref`, defined name, table `ref` and chart `c:f` are shifted; then on `spec124-structure.ods` insert a column → after download and reopen its named range and conditional-format target have moved

## Out of scope
Insert/delete cells with shift right/down/left/up, "Insert cut cells" and
Shift-drag insert (backlog). Outline grouping of rows and columns
(whole-queue backlog). Hiding all rows below a point by default
(`sheetFormatPr@zeroHeight`) — read and preserved, not authored. Structural
edits on several grouped sheets at once, and therefore adjusting 3-D
references. The insert-options chip (format same as above/below/clear).
Moving rows or columns by dragging headers. Modelling any of the retained
elements — each table row names the spec that will. Cut, copy, paste and
moving ranges (SPEC-125, which reuses this spec's retained-reference
machinery with its own mapper).

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — the model (`Sheet { id; name;
  rows; partName?; sourceXml? }`, `Row { index; cells }`, `Cell { col; value;
  formula?; styleId?; runs? }`, `SpreadsheetModel { sheets; opaque:
  OpaquePart[]; odsContentEnvelope?: { prefix; suffix } }`), `MAX_ROWS`,
  `MAX_COLS`, the four invariants (copy-on-write rows, sparse, the source is
  the splice base, dates are serials), `GridRange`/`CellRef`, `findRow`,
  `getCell`, `usedRange`; `core/xmlScanner.ts` (`scanXml(text, onEvent)`,
  events carry `start`/`end` UTF-16 offsets); `OpaquePart { name; data /*
  base64 */ }` and `decodedPartBytes` in `core/packageParts.ts`.
- `specs/117-sheet-grid-styles-sizes-and-selection.md` — `Sheet.columns:
  ColumnSpan[] { min; max; widthPx?; hidden?; styleId?; sourceWidthChars? }`,
  `Row.heightPt?`, `Row.customHeight?`, `Row.hidden?`, `Row.styleId?`,
  `Sheet.defaultColWidthPx` (64), `Sheet.defaultRowHeightPt` (15),
  `Sheet.merges: GridRange[]`, `Sheet.frozen?: { rows; cols }`,
  `Sheet.selection?: SheetSelection`; column widths are px at 100% zoom,
  chars↔px through the default font's max digit width (MDW): px =
  trunc(((256·w + trunc(128/MDW))/256)·MDW); `GridMetrics`; the header
  testids `sheet-row-header`, `sheet-col-header`.
- `specs/119-the-formula-language.md` — `Cell.formula` canonical text
  (no "="), `Cell.formulaUnparsed?: true`, `Cell.arrayRange?: GridRange`;
  `formula/tokens.ts` `tokenize`/`joinTokens` (`joinTokens(tokenize(t)) ===
  t`); `formula/references.ts` `RefPart { row?; col?; rowAbs; colAbs }`,
  `ReferenceNode { sheet?; sheetEnd?; start: RefPart; end?: RefPart }`,
  `shiftFormula`, `renameSheetInFormula`, and the token-level
  reference-rewrite primitive both are built on,
  `rewriteReferences(text, rewrite: RefRewrite)`, where `rewrite(ref)`
  returns a replacement `ReferenceNode`, `null` for `#REF!` or `undefined`
  to keep the token, and every other token and all trivia are re-emitted
  untouched; `formula/dialect.ts` `fromXlsxFormula`, `toXlsxFormula`,
  `fromOdsFormula`, `toOdsFormula`.
- `specs/121-…` — `record(label, scope)`, by-reference snapshots, `Sheet.dirty`,
  `structureDirty`, calcChain dropped on any structural change; sheet names
  inside retained text stay **source-time** (renames are applied at write
  time by `sourceSheetRemap`), so any code interpreting one resolves it with
  `resolveRetainedSheetName(model, parts, name, hostSheetId?)` in
  `formula/sheetRefs.ts`; ODS address lists parse and print through
  `odsRangeAddressToCanonical`/`canonicalToOdsRangeAddress` in
  `ods/rangeAddress.ts`.
  `specs/122-…` — the xlsx sheet splice (CT_Worksheet order; `<dimension>`,
  `<sheetViews>`, `<cols>`, `<sheetData>`, `<mergeCells>` rebuilt),
  ods table re-serialisation with `Cell.preservedXml`, protection honour;
  `autofit.ts` (`autofitRowHeightPt(model, sheet, row, measurer, ctx)`,
  `refitRows`, `CellTextMeasurer`, `MeasuredFont`, `fixedMeasurer`,
  `AUTOFIT_MAX_ROWS_PER_COMMAND = 10_000`) and the view's
  `createCanvasCellMeasurer()` (`src/views/spreadsheet/cellMeasurer.ts`).
- `specs/123-cell-formatting.md` — `applyStylePatch`, SPEC-117's
  `effectiveStyleId` for a position's current style,
  `MAX_FORMAT_MATERIALISED = 1_000_000`, the toolbar's band 6 slot,
  `SHEET_SHORTCUTS` with `code` bindings, the painter's `StylePattern`,
  `clearRange`, `SheetCellContextMenu.vue`.

### Geometry of an edit
```ts
// src/lib/spreadsheet/structure.ts
export type StructuralEdit = {
  kind: "insertRows" | "deleteRows" | "insertCols" | "deleteCols";
  sheetId: string;
  at: number;      // 0-based first row/column inserted or deleted
  count: number;   // ≥ 1; at + count ≤ MAX_ROWS / MAX_COLS
};
export function adjustIndex(index: number, edit: StructuralEdit): number | null;
export function adjustRange(range: GridRange, edit: StructuralEdit): GridRange | null;
```
Along the edit's axis (`last` = `MAX_ROWS−1` or `MAX_COLS−1`):

| Case | Insert `n` at `a` | Delete `[a, a+n)` |
| --- | --- | --- |
| index `i < a` | `i` | `i` |
| index `i ≥ a` | `i+n`; `> last` → `null` | `i` in band → `null`; `i ≥ a+n` → `i−n` |
| range start `s`, end `e` | `s ≥ a` → `s+n`; `e ≥ a` → `e+n`, clamped to `last` (a start past `last` → `null`) | both in band → `null`; `s` in band → `a`; `e` in band → `a−1`; `e ≥ a+n` → `e−n` |
So an insert *at* a range's first row moves the whole range, an insert
strictly inside expands it, and an insert just after its last row leaves it
alone (Excel). The other axis is never touched.

### Formula adjustment
`adjustFormula(text, hostSheetId, edit, sheetNameOf: (sheetId) => string)`
walks `rewriteReferences`: a reference is on the edited sheet when its
`sheet` (or the host's name when absent) equals `sheetNameOf(edit.sheetId)`
case-insensitively; its parts move per the geometry table regardless of
`$`; a whole-column reference ignores row edits and vice versa; a
`sheetEnd` (3-D) or external (`[1]Book!A1`) reference is left alone; a
reference whose area is gone prints as `#REF!` (`Sheet2!#REF!` when it
named a sheet). Strings, names and structured references are not
references to this function. Host Sheet1, edit on Sheet1 unless noted:

| Formula | Edit | Result |
| --- | --- | --- |
| `A5+1` | insertRows a=2 n=2 | `A7+1` |
| `$A$5*2` | insertRows a=2 n=2 | `$A$7*2` |
| `SUM(A1:A10)` | insertRows a=4 n=1 | `SUM(A1:A11)` |
| `SUM(A5:A10)` | insertRows a=4 n=1 | `SUM(A6:A11)` |
| `SUM(A1:A10)` | insertRows a=10 n=1 | unchanged |
| `A5 +  B5` | insertRows a=0 n=1 | `A6 +  B6` (trivia kept) |
| `A5` | deleteRows a=4 n=1 | `#REF!` |
| `SUM(A1:A10)` | deleteRows a=2 n=3 | `SUM(A1:A7)` |
| `SUM(A4:A10)` | deleteRows a=2 n=3 | `SUM(A3:A7)` |
| `SUM(A1:A4)` | deleteRows a=2 n=3 | `SUM(A1:A2)` |
| `SUM(A3:A5)` | deleteRows a=2 n=3 | `SUM(#REF!)` |
| `5:10` / `A:A` | insertRows a=0 n=1 | `6:11` / unchanged |
| `SUM(B:D)` | deleteCols a=2 n=1 | `SUM(B:C)` |
| `A1048576` / `A1:A1048576` | insertRows a=0 n=1 | `#REF!` / `A2:A1048576` |
| `Sheet1!C1` (host Sheet2) | insertCols a=1 n=1 | `Sheet1!D1` |
| `'my sheet'!A5` (host Sheet2) | insertRows a=0 n=1 on "My Sheet" | `'my sheet'!A6` |
| `Sheet2!A5` (host Sheet1) | deleteRows a=4 n=1 on Sheet2 | `Sheet2!#REF!` |
| `Sheet1:Sheet3!A5` / `[1]Sheet1!A5` | any | unchanged |
| `INDIRECT("A5")` | insertRows a=0 n=1 | unchanged |

### The adjuster registry
```ts
export interface ReferenceAdjuster { id: string; adjust(model: SpreadsheetModel, edit: StructuralEdit): void }
export const REFERENCE_ADJUSTERS: ReferenceAdjuster[];
export type StructuralResult = { ok: true } | { ok: false; message: string };
export function applyStructuralEdit(model: SpreadsheetModel, edit: StructuralEdit): StructuralResult;
export const INSERT_PUSHES_DATA_MESSAGE = "Can't insert — that would push filled cells off the end of the sheet";
export const PARTIAL_ARRAY_MESSAGE = "You can't change part of an array";
```
| Order | id | Owner | Moves |
| --- | --- | --- | --- |
| 1 | `rows` | 124 | row edits: `Row` objects re-indexed, band removed, inserted rows' inherited properties and styled cells; column edits: `Cell.col` shifted per row, inherited cell styles |
| 2 | `formulas` | 124 | `adjustFormula` over every sheet's `Cell.formula`, and `Cell.arrayRange` via `adjustRange` |
| 3 | `merges` | 124 | `Sheet.merges` via `adjustRange`; single-cell results dropped |
| 4 | `frozen` | 124 | the band `[0, rows−1]` / `[0, cols−1]` via `adjustRange` |
| 5 | `selection` | 124 | `Sheet.selection` ranges and active cell (clamped into the sheet) |
| 6 | `columns` | 124 | `ColumnSpan`s shifted, split, removed; inserted columns copy the left span's width and style (not `hidden`); adjacent equal spans re-joined |
| 7 | `retained` | 124 | `rewriteRetainedReferences(model, structuralMapper(model, edit))` |
| 8+ | registered by later specs, in landing order: `autofilter` (SPEC-128), `conditionalFormats` (SPEC-129), `validations` (SPEC-130), `drawings` (SPEC-131), `charts` (SPEC-132), `names` and `tables` (SPEC-133), `notes` and `hyperlinks` (SPEC-135), `pageBreaks` (SPEC-136), `pivots` (SPEC-139) | — | each appends its adjuster with the contract test, and deletes its rows from "Retained references" in the same change; print areas and print titles get **no** adjuster of their own — they are SPEC-133's reserved names, moved by `names` (a second adjuster would shift them twice) |

### Retained references
```ts
export interface RetainedReferenceMapper {
  range(sheetName: string, range: GridRange): GridRange | null;           // null: area gone
  index(sheetName: string, axis: "row" | "col", index: number): number | null;
  formula(canonical: string, hostSheetId: string): string;
}
export function structuralMapper(model: SpreadsheetModel, edit: StructuralEdit): RetainedReferenceMapper;
export function rewriteRetainedReferences(model: SpreadsheetModel, mapper: RetainedReferenceMapper): void;
```
Format halves live in `xlsx/retainedReferences.ts` and
`ods/retainedReferences.ts`. A part belongs to a sheet through the sheet
part's relationships (tables, drawings, comments, VML, pivot tables);
charts, pivot caches and defined names carry sheet names in their own
references. Those names are source-time text (SPEC-121 renames them only at
write time), so every sheet name read out of retained text — a formula's
sheet prefix, an ODS address's sheet, `worksheetSource@sheet` — is resolved
to its model sheet with SPEC-121's `resolveRetainedSheetName(model, parts,
name, hostSheetId)` (the host selecting a duplicated sheet's override)
before `mapper.range`/`mapper.formula` compare it with the edited sheet, and
the rewrite writes the name back exactly as it found it. Range-valued text is a space-separated list (`sqref`); a list
entry that maps to `null` is dropped, and an element whose list empties —
or whose single-cell `ref` is deleted — is removed (for a legacy comment,
with its VML shape matched by `x:Row`/`x:Column`). Formula-valued xlsx text
goes `fromXlsxFormula` → `mapper.formula` → `toXlsxFormula` and is written
back only when the canonical text changed. Index-valued text moves through
`mapper.index`; an index in a deleted band becomes `at`.

| xlsx part | Element / attribute | Kind | Retired by |
| --- | --- | --- | --- |
| sheet | `conditionalFormatting@sqref`; `cfRule/formula` | sqref; formula | 129 |
| sheet | `dataValidation@sqref`; `formula1`, `formula2` | sqref; formula | 130 |
| sheet | `autoFilter@ref`, `sortState@ref`, `sortCondition@ref` | range | 128 |
| sheet | `hyperlink@ref` | range | 135 |
| sheet | `rowBreaks/brk@id`, `colBreaks/brk@id` | index | 136 |
| sheet | `protectedRange@sqref`, `ignoredError@sqref`, `dataRef@ref`, `inputCells@r`, `cellWatch@r` | sqref | — |
| sheet `extLst` | `x14:conditionalFormatting/xm:sqref`, its `xm:f`; `x14:dataValidation/xm:sqref`, `x14:formula1|2/xm:f`; `x14:sparkline/xm:f`, `xm:sqref` | sqref; formula | 129 / 130 / — |
| `xl/workbook.xml` | `definedName` text | formula | 133 (`_xlnm._FilterDatabase` 128; `_xlnm.Print_Area`/`Print_Titles` with every other name, 133) |
| `xl/tables/tableN.xml` | `table@ref`, `autoFilter@ref` | range | 133 |
| `xl/charts/chartN.xml` | every `c:f` | formula | 132 |
| `xl/pivotCache/pivotCacheDefinitionN.xml` | `worksheetSource@ref` (+`@sheet`) | range | 139 |
| `xl/pivotTables/pivotTableN.xml` | `location@ref` | range | 139 |
| `xl/drawings/drawingN.xml` | `xdr:from`/`xdr:to` `xdr:row`, `xdr:col` (0-based; `editAs="absolute"` anchors untouched) | index | 131 |
| `xl/commentsN.xml`, `xl/threadedComments/*.xml` | `comment@ref`, `threadedComment@ref` | range | 135 |
| `xl/drawings/vmlDrawingN.vml` | `x:ClientData` `x:Row`, `x:Column`, `x:Anchor` (its 1st, 3rd, 5th and 7th numbers — the anchor columns and rows) | index | 135 |

| ods location | Element@attribute | Kind | Retired by |
| --- | --- | --- | --- |
| envelope prefix | `table:content-validation@table:base-cell-address`, `@table:condition` | range; formula | 130 |
| envelope prefix | `table:label-range@table:label-cell-range-address`, `@table:data-cell-range-address` | range | — |
| suffix and inside tables | `table:named-range@table:cell-range-address`, `@table:base-cell-address`; `table:named-expression@table:expression`, `@table:base-cell-address` | range; formula | 133 |
| suffix | `table:database-range@table:target-range-address`; `table:filter@table:condition-source-range-address` | range | 128 / 133 |
| suffix | `table:data-pilot-table@table:target-range-address`; `table:source-cell-range@table:cell-range-address` | range | 139 |
| suffix | `table:consolidation@table:source-cell-range-addresses`, `@table:target-cell-address` | range | — |
| table start tag | `table:table@table:print-ranges` | range | 136 |
| inside tables | `calcext:conditional-format@calcext:target-range-address`; `calcext:condition@calcext:value`, `@calcext:base-cell-address`; `calcext:formatting-entry@calcext:value` | range; formula | 129 |
| inside tables, `Cell.preservedXml` | `draw:*@table:end-cell-address`; `draw:object@draw:notify-on-update-of-ranges` | range | 131 / 132 |
| inside tables | `table:scenario@table:scenario-ranges` | range | — |
| `Object N/content.xml` | `chart:plot-area@table:cell-range-address`, `chart:series@chart:values-cell-range-address`, `@chart:label-cell-address`, `chart:domain@table:cell-range-address`, `chart:categories@table:cell-range-address` | range | 132 |
ODS range text is a list of `[$]Sheet.[$]A[$]1[:[[$]Sheet].[$]B[$]2]` (sheet
names may be `'quoted'`; the second sheet is optional), parsed and printed
entry by entry through SPEC-121's `odsRangeAddressToCanonical`/
`canonicalToOdsRangeAddress` (`ods/rangeAddress.ts`) — an entry the edit does
not move keeps its bytes. In a formula-valued
ODS attribute **every bracketed `[…]` reference is adjusted on its own and
nothing else is touched** — OpenFormula always brackets references, whatever
wrapper surrounds them (`of:=…`, `cell-content-is-in-list(…)`,
`formula-is(…)`). `table:tracked-changes` is history and never adjusted. A
modelled element's row is retired by its spec in the same change that adds
its adjuster, so nothing is adjusted twice.

### Sizes
Resize writes `widthPx` (dropping `sourceWidthChars`) or `heightPt` +
`customHeight`. **Autofit**: `autofitColumnWidthPx(model, sheet, col,
measurer, ctx)`, added to SPEC-122's `autofit.ts`, = ceil(the widest
SPEC-118 full-width display text among the column's stored cells) + 5 px
(the per-column padding in SPEC-117's chars→px formula), skipping cells
inside merges and wrapped cells; no text → `defaultColWidthPx`; capped at
the width of 255 characters. Rows use SPEC-122's
`autofitRowHeightPt(model, sheet, row, measurer, ctx)` (restated: the
tallest cell — one line of its font unwrapped, greedy word-wrap at (column
width − 5 px) when wrapped, the rotated bounding box for 1–180, characters ×
line height for 255 — converted px → pt (× 0.75), rounded up to a 0.75 pt
step, never below `defaultRowHeightPt`); the manual AutoFit writes that
height and clears `customHeight`. Both take SPEC-122's `CellTextMeasurer`
(`fixedMeasurer` in tests, `createCanvasCellMeasurer()` from
`src/views/spreadsheet/cellMeasurer.ts` in the view).

### Merges
1. Each range is expanded until it contains every merge it intersects.
2. Merges inside the expanded range are removed.
3. **Kept value**: the upper-left cell (per row for "across"); any other cell
   with a value or formula triggers the warning; on confirm their values
   and formulas are cleared, their styles kept (Excel draws merge borders
   from them).
4. "centre" adds `{ alignment: { horizontal: "center" } }` to the upper-left
   cell through SPEC-123's `applyStylePatch`; "across" merges each row of a
   multi-column range separately; a single cell is a no-op.
5. A merge partly over an array range → `PARTIAL_ARRAY_MESSAGE`.
`unmergeCells` removes every merge intersecting the ranges; content and
styles stay. SPEC-123's Clear formats/Clear all also unmerge merges inside
the cleared ranges, and its painter reproduces the source's merges in the
target (warning rules apply).

### Header menu
| Entry | testid | Label (rows; columns analogous) | Undo label |
| --- | --- | --- | --- |
| Insert above / left | `sheet-insert-before` | "Insert 1 row above", "Insert 3 rows above" / "Insert 1 column left" | "Insert rows" / "Insert columns" |
| Insert below / right | `sheet-insert-after` | "Insert 1 row below" / "Insert 2 columns right" | same |
| Delete | `sheet-delete-band` | "Delete row", "Delete rows 4–6" / "Delete columns B–D" | "Delete rows" / "Delete columns" |
| Hide | `sheet-hide-band` | "Hide row", "Hide rows 4–6" | "Hide rows" |
| Unhide | `sheet-unhide-band` | "Unhide rows" / "Unhide columns" | "Unhide rows" |
| Size | `sheet-size-band` | "Row height…" / "Column width…" | "Resize rows" |
| Autofit | `sheet-autofit-band` | "Autofit row height" / "Autofit column width" | "Autofit" |
`N` is the number of selected rows/columns in the band containing the
clicked header. SPEC-123's cell context menu gains "Insert…" and
"Delete…" (opening the dialogs).

### Fixtures
`spec124-structure.xlsx`, Excel-shaped (`mc:Ignorable`, `xr:uid`,
`x14ac:dyDescent`, `spans`, `xl/calcChain.xml`): Sheet1 with 30 data rows, a
shared formula, an array formula, merges, a frozen `pane`; Sheet2 with
`Sheet1!A5`, `'Sheet1'!B2:B20`, `Sheet1:Sheet2!A1` and `B5`-referencing
formulas; a `conditionalFormatting` with `<formula>`, a list
`dataValidation` with `formula1`, an x14 `extLst` validation with
`xm:sqref`, a hyperlink, an `autoFilter`, defined names including
`_xlnm.Print_Area`, a table part, a drawing with a `twoCellAnchor` picture and
a chart whose `c:f` reads Sheet1, a legacy comment (comments + VML).
`spec124-structure.ods`, LibreOffice-shaped: content validations, named
ranges and expressions, a database range, `calcext:conditional-formats`,
`table:print-ranges`, a chart object `Object 1/content.xml`, a `draw:frame`
with `table:end-cell-address` inside a cell, covered cells for a merge, and
`settings.xml` frozen panes (`HorizontalSplitMode=2`).

### Tests
Unit and e2e files per the Tests and e2e ACs. Relations: part-level
`equalBytes` (`src/lib/core/equality.ts`) for every part an edit did not
touch and for the no-edit download; field-level `toEqual` on re-read sheets
(ignoring minted sheet ids) for the edited ones; the span-diff helper for
adjusted parts. Undo tests compare `rows`, `columns`, `merges`, `frozen`,
`opaque` and `sourceXml` by identity after an undo.

### Specs this changes
- SPEC-121's undo snapshot also keeps `opaque`, `odsContentEnvelope` and each
  sheet's `sourceXml` by reference (already in its
  `WORKBOOK_SNAPSHOT_FIELDS`/`SHEET_BY_REFERENCE_FIELDS`).
- SPEC-116's "the source is the splice base" now reads: the *adjusted*
  source — unmodelled reference-bearing text is moved by structural edits.
- SPEC-123: band 6 gains Merge; Clear formats unmerges; the painter carries
  merges; the cell context menu gains Insert…/Delete… and gives Shift+F10 /
  the Menu key to the header menu when the selection is whole rows or
  columns; font family/size, wrap and rotation patches run SPEC-122's
  `refitRows`; `SHEET_SHORTCUTS` gains `sheet.insertCells` (Ctrl+Alt+= /
  Cmd+Option+=) and `sheet.deleteCells` (Ctrl+Alt+- / Cmd+Option+-).
- SPEC-122: `autofit.ts` gains `autofitColumnWidthPx`; `refitRows` gains
  the formatting, paste and fill triggers.
- SPEC-121: its `resolveRetainedSheetName` is how `RetainedReferenceMapper`
  reads sheet names in retained text; its `ods/rangeAddress.ts` parses the
  ODS address lists.
- SPEC-128–139: each spec that models a reference-bearing structure
  registers a `ReferenceAdjuster` and retires its retained-reference rows.
