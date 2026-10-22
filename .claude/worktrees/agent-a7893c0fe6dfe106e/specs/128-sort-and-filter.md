---
id: SPEC-128
title: Sort and filter
status: draft
area: editing
complexity: normal
story_points: 8
depends_on: [SPEC-076, SPEC-122, SPEC-124, SPEC-125, SPEC-127]
---
## Why
After SPEC-122 a workbook can be typed into, and after SPEC-124 its rows and
columns can be inserted, deleted and hidden — but the two things people do to
a list first are still missing. There is no way to put the rows in order, and
no way to look at only the rows that matter. Worse, the files already carry
both: Excel writes an `<autoFilter>` with its criteria and a `<sortState>`
into every sheet somebody has filtered, LibreOffice writes a
`table:database-range` with `table:display-filter-buttons="true"`, and both
hide the rows the filter rejected. Today those rows open hidden with no
dropdown on the header that hid them and nothing to say a filter is on; the
user sees gaps in the row numbers and cannot get the rows back except by
unhiding them by hand — which, as far as the file is concerned, destroys the
filter.

SPEC-076 already sorts table rows in the word processor, and its dialog
decisions carry over: a "My data has headers" toggle, ties broken by the
original order, merged cells refused with a stated reason in the dialog
rather than a failed action. Its comparator does not: it reads every cell as
text and parses numbers and dates out of it, with numeric collation
(`"A2" < "A10"`). A sheet's cells are already typed, and Excel's order is
defined over the types — numbers, then text, then logicals, then errors,
blanks last in both directions — with plain collation (`"A10" < "A2"`). So
this spec writes its own comparator in `src/lib/spreadsheet/sort.ts` and
**moves nothing out of `src/lib/tableSort.ts`**: the only shareable line is an
`Intl.Collator` construction with different options.

**This file is the normative source for spreadsheet sort semantics (type
order, collation, blanks, stability, what moves and what refuses), the
autofilter model and its criteria semantics, the distinction between
filter-hidden and user-hidden rows, and the xlsx/ODS wire mappings of
`autoFilter`, `sortState`, `_xlnm._FilterDatabase` and the ODS database range
that carries an autofilter.**

## Acceptance criteria
- [ ] `src/lib/spreadsheet/sort.ts` exports `compareSortValues`, `sortCollator`, `guessHasHeader`, `sortRange` and `MAX_SORT_KEYS = 64` per "Sort semantics": ascending order is numbers < text < `FALSE` < `TRUE` < errors (all errors equal), descending reverses every class, and blank cells sort **last in both directions**; ties fall through to the next key and then to the original row order, so applying the same keys twice changes nothing (`toEqual` on the sheet's rows)
- [ ] Text compares with `sortCollator(locale.tag, caseSensitive)` — `Intl.Collator` over the user's `SpreadsheetLocale` (SPEC-118) with `usage: "sort"`, `numeric: false` and sensitivity pinned to `"accent"` (case-insensitive, the default) or `"variant"` with `caseFirst: "lower"` (case-sensitive) — so `"A10"` sorts before `"A2"`, `"apple"` and `"Apple"` tie by default and `"apple"` precedes `"Apple"` when case-sensitive; a formula returning `""` is text, not blank
- [ ] `sortRange` permutes, among the **visible** rows of the range's data rows only, the cells in the range's columns as whole records: value, formula, `styleId`, `runs` and every other `Cell` field travel together; row properties (`heightPt`, `hidden`, `filterHidden`, row `styleId`) stay with the row index; rows hidden by the user or by a filter are neither moved nor moved into; cells outside the range's columns never move
- [ ] A moved formula is shifted like a copy — `shiftFormula(formula, toRow - fromRow, 0)` — so relative references move with the record and absolute ones do not (Excel's semantics, including its known surprise for references to other rows); formulas **outside** the range that point into it are not adjusted; a cell with `formulaUnparsed` moves verbatim; `sortRange` returns the row permutation and the changed cell keys, runs every `ROW_PERMUTATION_ADJUSTERS` entry with that permutation per "Sort semantics" (the hook later per-cell structures — SPEC-135's notes, threads and links — register into; this spec registers none), and the view hands the keys to SPEC-120's `recalculate`
- [ ] `sortRange` refuses, changing nothing, with the pinned messages: a merge that spans more than one data row or crosses the range's left/right edge (a single-row merge wholly inside the range moves with its row), part of a multi-row array formula, a multi-area selection, fewer than two data rows, a duplicated key column, and a sheet protected per SPEC-122 whose retained `sheetProtection` does not carry `sort="0"`
- [ ] `guessHasHeader(model, sheet, range)` implements "The header guess" exactly — the first row is a header when every non-blank cell in it is a string and either some column holds a string above a non-string value, or the whole first row is bold and the second is not, or the range is an autofilter's range — and the guess only ever seeds the "My data has headers" checkbox, which the user overrides
- [ ] Data ▸ "Sort A to Z" (`sheet-menu-sort-asc`) and "Sort Z to A" (`sheet-menu-sort-desc`) sort by the active cell's column: a single-cell selection expands to its `currentRegion` (SPEC-117), a multi-cell selection sorts exactly itself, a selection inside an autofilter's range sorts that range's data rows; a refusal opens the `sheet-alert` dialog with the message and changes nothing; each sort is one undo entry labelled "Sort"
- [ ] Data ▸ "Sort range…" (`sheet-menu-sort-range`) opens `sort-dialog`: one `sort-level` row per key (`sort-level-column`, `sort-level-order`), `sort-add-level` (disabled at 64 levels with the reason "Up to 64 levels"), `sort-delete-level`, `sort-copy-level`, `sort-move-up`, `sort-move-down`, `sort-has-headers` ("My data has headers"), `sort-case-sensitive` ("Case sensitive"), `sort-ok`, `sort-cancel`; column choices read the header text when headers are on (blank header → "Column B") and "Column B" otherwise; the order choices read "A to Z"/"Z to A", or "Smallest to Largest"/"Largest to Smallest" for a column whose non-blank data cells are mostly numbers; refusals and a duplicated column show inline in `sort-dialog-error`
- [ ] Every sort records its `SortState` per "The sort record": into `autoFilter.sort` when the range is an autofilter's data range, otherwise into `Sheet.sortState`; it is a record of the last sort, never re-applied on open or on edit
- [ ] `src/lib/spreadsheet/filter.ts` exports the autofilter model per "The autofilter model" — `Sheet.autoFilter?: AutoFilter` (`range` including its header row, `columns: FilterColumn[]`, optional `sort`), `FilterCriteria` (`values` with `(Blanks)` and date groups, `custom` with one or two conditions joined by and/or, `top10`, `dynamic`, and `preserved` for colour and icon filters), plus `rowPasses`, `applyAutoFilter` and `filterValueList` — every criterion's matching rule pinned in "Criteria semantics", including Excel's encodings (`notEqual " "` is "non-blank", wildcards `*` `?` `~` in text conditions)
- [ ] `Row` gains `filterHidden?: true`, distinct from SPEC-117's user-set `hidden?: true`; a row is displayed hidden when either is set; `GridMetrics` gives such a row zero height; SPEC-124's Unhide clears `hidden` only, so a filtered-out row stays hidden until the filter changes; turning the filter off or clearing its criteria clears `filterHidden` and never touches `hidden` — asserted on a sheet holding both kinds
- [ ] Reading never changes what is visible: a row the file shows stays shown even if today's values fail the criteria (Excel shows stale results until Reapply), and a row the file hides inside the autofilter's data rows is classified per "Filter-hidden or user-hidden" — `filterHidden` when a modelled criterion rejects it or the filter carries a preserved criterion, `hidden` otherwise; an ODS row's `table:visibility="filter"` is `filterHidden` and `"collapse"` is `hidden`
- [ ] Data ▸ "Filter" (`sheet-menu-filter`, a checkbox item) and Ctrl/Cmd+Shift+L toggle the active sheet's autofilter: on, it covers the selection per "Choosing the range" and shows a `filter-button` in every header cell of the range; off, it removes `autoFilter` and every row's `filterHidden`; "Clear filters" (`sheet-menu-clear-filters`) removes all criteria and keeps the buttons; "Reapply filter" (`sheet-menu-reapply-filter`, Ctrl/Cmd+Alt+L — best effort on macOS, where Chrome keeps Cmd+Option+L for its Downloads page; the menu item is the guaranteed path) re-evaluates the criteria against current values — each is one undo entry
- [ ] A `filter-button` (`aria-haspopup="dialog"`, `aria-expanded`, `aria-label` `Filter "Region"`, `data-state` `none`/`filtered`/`sorted-asc`/`sorted-desc`) opens `filter-popover` on click or Alt+Down on its header cell; the popover holds `filter-sort-asc`/`filter-sort-desc` (labels by column type, as the Sort dialog), `filter-clear` (`Clear filter from "Region"`), a conditions submenu `filter-conditions` ("Text Filters", "Number Filters" or "Date Filters" by the column's majority type, items per the Reference), `filter-search`, `filter-select-all` ("(Select All)", tri-state), one `filter-value` checkbox per distinct display text sorted by the sort comparator with "(Blanks)" last, `filter-not-all` ("Not all items showing") past `FILTER_VALUE_LIST_MAX = 10_000`, `filter-ok` and `filter-cancel`; Escape closes it and returns focus to the header cell
- [ ] The custom-filter dialog (`filter-custom-dialog`: `filter-custom-op-1`, `filter-custom-value-1`, `filter-custom-join` And/Or, `filter-custom-op-2`, `filter-custom-value-2`) and the Top 10 dialog (`filter-top10-dialog`: `filter-top10-direction` Top/Bottom, `filter-top10-value` 1–500 items or 1–100 percent, `filter-top10-unit` Items/Percent) write `custom` and `top10` criteria; a column holding a `preserved` criterion shows "Filtered by colour or icon — this filter can't be edited here" with `filter-clear` as its only action
- [ ] Applying criteria sets `filterHidden` exactly on the data rows `rowPasses` rejects (AND across columns), evaluating rows up to `usedRange(sheet).bottom` only, and the range's bottom first grows to its header row's `currentRegion` bottom when that is lower (Excel extends a filter over appended data; it never shrinks); a preserved criterion keeps each row's existing `filterHidden` state rather than guessing
- [ ] While the active sheet's autofilter has at least one criterion, `sheet-status-filter` reads "N of M records found" (M = data rows of the range up to the used range, N = those not `filterHidden`), visible row headers inside the range carry `data-filtered="true"` and render in the accent colour, and SPEC-117's status-bar aggregates skip filter-hidden rows
- [ ] SUBTOTAL honours the distinction: SPEC-127's `EvalContext.rowVisibility(sheetId, row): "visible" | "filtered" | "hidden"` is widened here to answer `"filtered"` for `filterHidden` (`"hidden"` wins when both flags are set) — so SPEC-127's `SUBTOTAL` 1–11 skips filter-hidden rows only and 101–111 both kinds; every filter change (apply, clear, reapply, toggle off) hands `visibilitySensitiveCells(model)` to `recalculate`; copy, clear contents and fill down/right over a selection that spans filter-hidden rows act on the visible rows only (Excel), through one `visibleRowSpans(sheet, range)` helper
- [ ] xlsx read and write per "xlsx wire mapping": `<autoFilter ref>` with `filterColumn` (`filters`/`filter`/`dateGroupItem`/`@blank`, `customFilters`/`customFilter`, `top10`, `dynamicFilter`; `colorFilter` and `iconFilter` kept verbatim as `preserved`; `hiddenButton`/`showButton`), `sortState`/`sortCondition` inside it and at worksheet level, `sheetPr@filterMode` set exactly when some row is `filterHidden`, `hidden="1"` written for either kind of hidden row, and the hidden `_xlnm._FilterDatabase` defined name kept in step with every autofilter (added, moved, removed) in `xl/workbook.xml`
- [ ] ODS read and write per "ODS wire mapping": the sheet's anonymous `table:database-range` (`__Anonymous_Sheet_DB__<index>`) with `table:display-filter-buttons="true"`, its `table:filter` tree (`table:filter-and`/`table:filter-or`/`table:filter-condition`/`table:filter-set-item`, every `table:operator` in the table) and `table:sort`/`table:sort-by`; a filter tree that does not decompose into per-column criteria is kept verbatim and offers only "Clear filter"; rows written `table:visibility="filter"` for filter-hidden and `"collapse"` for user-hidden, and visibility joins SPEC-122's row repeat-compression key
- [ ] An untouched filtered workbook downloads byte-identical part for part (`equalBytes`); a modelled element whose `sourceXml` survives (unchanged since read) is re-emitted verbatim inside a rebuilt sheet; filter → save → reopen reproduces the model's `autoFilter`, `sortState` and every row's `hidden`/`filterHidden` under `toEqual` minus minted ids, in both formats
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`: a clean sheet is re-read; a dirty sheet gets `autoFilter`/`sortState` parsed from its retained source — the retained source has carried SPEC-124's and SPEC-125's retained-reference rewrites since it was read, so the parsed positions already match the current model — and the hidden rows inside the filter range reclassified by the read rule; tested with a synthetic previous-version record holding one dirty and one clean filtered sheet
- [ ] Registers `{ id: "autofilter", adjust }` in `REFERENCE_ADJUSTERS` per "Structural edits" and, in the same change, retires its rows from SPEC-124's "Retained references" tables (`autoFilter@ref`, `sortState@ref`, `sortCondition@ref`, the autofilter `table:database-range@table:target-range-address` and its `table:filter@table:condition-source-range-address`, and the `_xlnm._FilterDatabase` defined-name text), so nothing is adjusted twice: inserting two rows above an autofilter at `A5:D20` moves it to `A7:D22` (and its `sortState` with it), a column inserted inside it shifts the `colId` of every criterion to its right, deleting a criterion's column drops that criterion, deleting the header row removes the autofilter — the test joins `structure.spec.ts`'s registry contract test
- [ ] Registers `{ id: "autofilter", move }` in SPEC-125's `MOVE_ADJUSTERS` per "Moves": an autofilter (with its sort record) or a worksheet sort record lying entirely inside a cut-paste or drag-move source moves by the offset — onto the destination sheet, replacing an autofilter there — and is otherwise unchanged; the test joins `move.spec.ts`'s registry contract test
- [ ] Fixtures from `spreadsheet/fixtures/build.ts`: `e2e/fixtures/spec128-filter.xlsx` (Excel-shaped per the Reference: `xr:uid`, `filterMode`, `_xlnm._FilterDatabase`, a values filter with `(Blanks)`, a `customFilters` pair, a `colorFilter` with its dxf, a worksheet `sortState`, filter-hidden rows plus one user-hidden row inside the range, `x14ac:dyDescent` rows) and `spec128-filter.ods` (LibreOffice-shaped: `__Anonymous_Sheet_DB__0`, `table:filter-set-item`, `table:visibility="filter"` and one `"collapse"` row, trailing repeated empty rows)
- [ ] Tests: `spreadsheet/sort.spec.ts` (type order both directions, collation both sensitivities across all six locales, stability, the permutation over hidden rows, formula shifting, every refusal, the header guess table), `filter.spec.ts` (every criterion and Excel encoding, date dynamics under `setClock`, reapply with a preserved criterion, the read classification), `xlsx/autoFilter.spec.ts`, `ods/databaseRanges.spec.ts`, the adjuster test in `structure.spec.ts`, the move adjuster test in `move.spec.ts`, the upgrade step in `open.spec.ts`, `views/spreadsheet/useSortFilter.spec.ts`, and `roundTrip.spec.ts` additions
- [ ] e2e (`e2e/SPEC-128.spec.ts`): open `spec128-filter.xlsx` → the header row shows `filter-button`s and `sheet-status-filter` reads "8 of 12 records found" → open Region's popover → uncheck "West" → OK → "6 of 12 records found" and the West rows are gone → Data ▸ Clear filters → the user-hidden row is still hidden → click an Amount cell → Data ▸ Sort Z to A → the visible Amount column reads descending → Ctrl/Cmd+Z restores the order → Ctrl/Cmd+Shift+L → no `filter-button` remains → download → reopen → no autofilter; then open `spec128-filter.ods` → the `table:visibility="filter"` rows are hidden and "Reapply filter" leaves the same rows hidden

## Out of scope
Sorting by cell colour, font colour or icon, and filtering by them (read,
honoured and preserved as `preserved` criteria / `sortState.sourceXml`; their
editing is Backlog). Custom sort lists ("Jan, Feb, …"). Sorting left to right
(columns rather than rows) — a worksheet `sortState@columnSort` is preserved,
never produced. The year ▸ month ▸ day tree in the filter popover: dates
appear as their display text in the flat list (`dateGroupItem` criteria read
from a file are evaluated, written back and shown as one "Date groups" entry
until edited) — Backlog. Advanced Filter (criteria ranges, `_xlnm.Criteria`,
`_xlnm.Extract`, LibreOffice's standard filter with a copy-output range),
preserved verbatim. Excel tables' own autofilters and named database ranges
(SPEC-133). Right-click ▸ Sort/Filter entries and ribbon buttons (the Data
menu and the keyboard are the surfaces here). Emulating Excel's "word sort"
that ignores hyphens and apostrophes — the collator's behaviour is pinned
instead. Re-filtering automatically as values change (Excel does not either;
Reapply is explicit). Printing a filtered sheet (SPEC-136 prints what is
visible).

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — the model is plain JSON with
  copy-on-write rows (an edit builds new `Row`/`Cell` objects and a new
  `rows` array; untouched rows keep identity), sparse storage, the retained
  source as the splice base (xlsx part bytes in `opaque`, ODS `sourceXml`,
  `odsContentEnvelope { prefix; suffix }` around the tables).
  `address.ts`: 0-based `CellRef { row; col }`, inclusive normalised
  `GridRange { top; left; bottom; right }`, `parseRange`, `formatRange`,
  `parseSqref`, `getCell`, `findRow`, `usedRange`. `open.ts`:
  `SPREADSHEET_READER_VERSION`, `ReaderUpgrade { from; upgradeDirtySheet(model,
  sheet, sourceParts) }`, `READER_UPGRADES`, `upgradeSpreadsheetModel`.
- SPEC-117 — `Row.hidden?: true` (read from xlsx `hidden` and, until this
  spec, from both ODS `table:visibility` values), `Row.heightPt?`,
  `Row.styleId?`, `Row.sourceAttrs?`, `Sheet.merges: GridRange[]`,
  `GridMetrics` (`rowTop`, `rowHeight`, `rowAt`), `effectiveStyleId(sheet,
  row, col)` and `cellStyle(model, id)` in `styles.ts`, `CellStyle.font.bold?:
  true`, `SheetSelection { active; ranges; activeRangeIndex }`, `NavContext {
  …; isRowHidden(r) }` and `currentRegion(at, nav)` in `selection.ts`;
  `selectionAggregates` in `aggregates.ts` (hidden rows excluded) behind
  `sheet-status-sum`/`-average`/`-count` in `SpreadsheetStatusBar.vue`;
  primary modifier from `core/platform.ts` (`isMacPlatform`).
- SPEC-118 — `formatValue(value, code, ctx: FormatContext): FormattedValue`
  with `FormatContext { locale: SpreadsheetLocale; dateSystem; widthPx? }`;
  `SpreadsheetLocale.tag` (en-US, en-GB, de-DE, fr-FR, es-ES, cs-CZ), a
  per-user preference; `plainText(fv)`; `parseFormatCode(code):
  ParsedFormat` whose `sections[i].kind` is `"date"` for a date/time
  section.
- SPEC-119/120 — `Cell.formula` canonical text; `shiftFormula(text, dRow,
  dCol)` (out-of-grid → `#REF!`); `Cell.arrayRange?: GridRange`;
  `Cell.formulaUnparsed?: true`; `recalculate(model, changed: CellKey[])`;
  `EvalContext`; `FunctionSpec`; SPEC-120's Data menu (`menu-title-data`,
  "Recalculate"). SPEC-127 — `SUBTOTAL`/`AGGREGATE` read
  `EvalContext.rowVisibility(sheetId, row)` and are `visibilitySensitive`;
  `visibilitySensitiveCells(model)` in `recalc.ts`.
- SPEC-121 — `useSpreadsheetUndo` `{ record(label, scope) /* before */,
  commit(alsoRecalc?) /* after */, undo, redo }`: each sheet snapshotted as
  `{ ...sheet }` with object fields `structuredClone`d except
  `SHEET_BY_REFERENCE_FIELDS` (`rows`, `sourceXml`); `Sheet.dirty?: true`,
  `SpreadsheetModel.structureDirty?: true`; writers `writeXlsx`/`writeOds`;
  `stripRetention(model)` and the `UNCONVERTIBLE_CONTENT` table behind the
  "Not included in the .ods copy: …" toast.
- SPEC-122 — the dirty-sheet splice (CT_Worksheet order: … `sheetProtection`,
  `protectedRanges`, `scenarios`, **`autoFilter`, `sortState`**,
  `dataConsolidate`, `customSheetViews`, `mergeCells` …), ODS table
  re-serialisation with row/column repeat compression; `parseInput(text,
  locale, dateSystem)`; protection is honoured with its own message.
- SPEC-124 — `StructuralEdit { kind: "insertRows" | "deleteRows" |
  "insertCols" | "deleteCols"; sheetId; at; count }`, `REFERENCE_ADJUSTERS:
  { id; adjust(model, edit): void }[]` in `structure.ts` (ids `rows`,
  `formulas`, `merges`, `frozen`, `selection`, `columns`, `retained`, then
  `autofilter` reserved for this spec), `adjustIndex(index, edit)` and
  `adjustRange(range, edit): GridRange | null` (insert at or before a range's
  start → shift; strictly inside → grow; delete covering → `null`; delete
  overlapping → shrink), and `rewriteRetainedReferences`, whose rows for the
  elements this spec models are retired here.
- SPEC-125 — `src/lib/spreadsheet/move.ts`: `RangeMove { fromSheetId; from:
  GridRange; toSheetId; toTop; toLeft }`, `MoveAdjuster { id; move(model,
  move): void }`, `MOVE_ADJUSTERS` (registered: `cells`, `formulas`,
  `merges`, `retained`), `moveRange`, `repointFormula(text, hostSheetId,
  move, sheetNameOf)` (a reference entirely inside the source follows the
  move; one entirely inside the destination becomes `#REF!`; anything else
  is unchanged) and `moveMapper`; `move.spec.ts` holds the contract test
  later move adjusters join.
- SPEC-076 — `src/lib/tableSort.ts` (`sortTableRows`, `SortKey`): the
  precedent for the header toggle, stability by original index and refusing
  merges in the dialog. Text product; not imported, not moved (see "Why").
- `src/lib/core/clock.ts` — `now()` / `setClock(instant)`, the one seam for
  "today" (date filters; tests never hard-code a date).

### Sort semantics (`src/lib/spreadsheet/sort.ts`)
```ts
export const MAX_SORT_KEYS = 64;
export interface SheetSortKey { col: number; descending: boolean }   // absolute 0-based column
export interface SortOptions { keys: SheetSortKey[]; hasHeader: boolean; caseSensitive: boolean; localeTag: string }
export type SortOutcome =
  | { ok: true; permutation: { from: number; to: number }[]; changed: CellKey[] }
  | { ok: false; reason: SortRefusal; message: string };
export type SortRefusal = "merges" | "array" | "multiArea" | "tooFewRows" | "duplicateKey" | "protected";
export function sortCollator(localeTag: string, caseSensitive: boolean): Intl.Collator;
export function compareSortValues(a: CellValue, b: CellValue, collator: Intl.Collator): number; // ascending order
export function guessHasHeader(model: SpreadsheetModel, sheet: Sheet, range: GridRange): boolean;
export function sortRange(model: SpreadsheetModel, sheetId: string, range: GridRange, options: SortOptions): SortOutcome;
/** Per-cell structures that travel with a sorted record register here. */
export interface RowPermutationAdjuster {
  id: string;
  permute(model: SpreadsheetModel, sheetId: string, range: GridRange,
    permutation: readonly { from: number; to: number }[]): void;   // absolute row indices
}
export const ROW_PERMUTATION_ADJUSTERS: RowPermutationAdjuster[];   // empty here
```
- **Classes.** A blank (`value === null`, formula or not) is outside the
  order. Ascending: number < string < `false` < `true` < error; within
  numbers, numeric order; within strings, the collator; errors all equal.
  `descending` negates the class and in-class comparison of its key; blanks
  are then placed after every non-blank, so they are last both ways. Dates
  are numbers.
- **Collator.** `new Intl.Collator(tag, { usage: "sort", numeric: false,
  ignorePunctuation: false, sensitivity: cs ? "variant" : "accent",
  caseFirst: "lower" })`, built once per sort.
- **Stability.** Decorate with the record's original index; the comparator
  falls through keys, then to the index (no reliance on engine stability).
- **Records.** Data rows = the range's rows minus the header row when
  `hasHeader`. Visible positions `P` = data rows that are neither `hidden`
  nor `filterHidden`. Each position contributes the record of its cells in
  `[left, right]` (absent row or cell = blank). Sorted records are written
  back into `P` in order; a record moving from row `r` to `r′` shifts each
  formula by `r′ − r`; a single-row `arrayRange` wholly inside the range moves
  by the same delta. New `Row` objects are built only for rows whose cells
  changed; row properties are copied from the destination's old row.
- **Permutation hook.** After the records are written back, `sortRange`
  runs each `ROW_PERMUTATION_ADJUSTERS` entry in registry order with the
  sheet, the sorted range and the permutation (every moved data row's
  `from` → `to`), under the same copy-on-write rule as SPEC-124's
  adjusters. It exists for structures keyed to single cells that must
  travel with their record (SPEC-135 registers `annotations`); range-based
  structures — conditional formats, validations, drawings, charts, tables,
  pivots — never register and stay where they are, as in Excel.
  `sort.spec.ts` asserts a registered test adjuster receives exactly the
  permutation `SortOutcome` returns.
- **Refusals** (`message` strings, shown verbatim):

| `reason` | Message |
| --- | --- |
| `merges` | "Sorting isn't possible here — the range contains merged cells that span rows. Unmerge them first." |
| `array` | "Sorting would split an array formula — the range contains part of one." |
| `multiArea` | "Sorting works on one range at a time — select a single range." |
| `tooFewRows` | "There's nothing to sort — select a cell inside a range of at least two rows." |
| `duplicateKey` | "Column B is used by more than one level — remove the duplicate." |
| `protected` | SPEC-122's protected-sheet message |

### The header guess
`guessHasHeader` is `false` for a range of fewer than two rows, and `true`
when the range's top row is the sheet's `autoFilter.range.top` (a filter's
first row is its header by definition). Otherwise it is `true` iff the first
row has ≥ 1 non-blank cell, **every** non-blank first-row cell holds a
string, and at least one of: (a) some column has a string in row 1 and a
number, boolean or error in row 2; (b) every non-blank first-row cell
resolves bold (`cellStyle(model, effectiveStyleId(sheet, row, col)).font
?.bold`, SPEC-117) and no non-blank second-row cell does.
Excel's own heuristic is undocumented; this one is pinned and tested as a
table of cases.

### The sort record
```ts
export interface SortState {
  ref: GridRange;            // the sorted data rows (header excluded)
  keys: SheetSortKey[];      // level order
  caseSensitive?: true;
  sourceXml?: string;        // verbatim source while unchanged (keeps sortBy="cellColor", customList, columnSort)
}
// Sheet gains: sortState?: SortState
```
xlsx `<sortState ref="A2:D20" [caseSensitive="1"]><sortCondition
[descending="1"] ref="B2:B20"/>…` ↔ `SortState`; a `sortCondition` with
`sortBy` other than `value` or a `customList` is not a key — the element is
kept as `sourceXml` until a new sort replaces it.

### The autofilter model (`src/lib/spreadsheet/filter.ts`)
```ts
export interface AutoFilter {
  range: GridRange;               // header row = range.top; data rows = top+1..bottom
  columns: FilterColumn[];        // ascending colId, at most one per colId
  sort?: SortState;
  sourceXml?: string;             // xlsx <autoFilter> / ods <table:database-range>, verbatim while unchanged
  odsName?: string;               // "__Anonymous_Sheet_DB__0"
  odsUnmappedFilter?: string;     // ods <table:filter> that does not decompose per column (see ODS)
}
export interface FilterColumn { colId: number; criteria?: FilterCriteria; hideButton?: true }
export type FilterCriteria =
  | { kind: "values"; values: string[]; blanks: boolean; dateGroups: DateGroupItem[] }
  | { kind: "custom"; join: "and" | "or"; conditions: [FilterCondition] | [FilterCondition, FilterCondition] }
  | { kind: "top10"; top: boolean; percent: boolean; value: number }
  | { kind: "dynamic"; type: DynamicFilterType }
  | { kind: "preserved"; xml: string; label: "colour" | "icon" };
export interface FilterCondition {
  operator: "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual";
  value: string;                  // wire text: invariant number text, or text with wildcards * ? ~
}
export interface DateGroupItem { grouping: "year" | "month" | "day" | "hour" | "minute" | "second";
  year: number; month?: number; day?: number; hour?: number; minute?: number; second?: number }
export type DynamicFilterType = "aboveAverage" | "belowAverage" | "today" | "yesterday" | "tomorrow"
  | "thisWeek" | "lastWeek" | "nextWeek" | "thisMonth" | "lastMonth" | "nextMonth" | "thisQuarter"
  | "lastQuarter" | "nextQuarter" | "thisYear" | "lastYear" | "nextYear" | "yearToDate"
  | "Q1" | "Q2" | "Q3" | "Q4" | "M1" | "M2" | "M3" | "M4" | "M5" | "M6" | "M7" | "M8" | "M9" | "M10" | "M11" | "M12";
// Sheet gains: autoFilter?: AutoFilter      Row gains: filterHidden?: true
export const FILTER_VALUE_LIST_MAX = 10_000;
export function rowPasses(model: SpreadsheetModel, sheet: Sheet, row: number, ctx: FormatContext): boolean;
export function applyAutoFilter(model: SpreadsheetModel, sheetId: string, ctx: FormatContext): { changedRows: number[] };
export function filterValueList(model: SpreadsheetModel, sheet: Sheet, colId: number, ctx: FormatContext):
  { items: { text: string; checked: boolean }[]; blanks: boolean | null; truncated: boolean };
export function visibleRowSpans(sheet: Sheet, range: GridRange): GridRange[];
```
Commands (one undo entry each, each sets `Sheet.dirty` and replaces — never
mutates — `autoFilter`; `autoFilter` and `sortState` therefore join SPEC-121's
`SHEET_BY_REFERENCE_FIELDS`, held by reference in the undo snapshot):
`toggleAutoFilter`, `setFilterCriteria(model,
sheetId, colId, criteria | undefined)`, `clearFilters`, `reapplyFilter`.
Creating, moving or removing an autofilter also sets `structureDirty` (the
workbook part carries `_xlnm._FilterDatabase`).

### Criteria semantics
A cell's **display text** is `plainText(formatValue(value, its number
format, ctx))` with no `widthPx` (never "####"), compared
case-insensitively. A **blank** is a cell whose display text is `""`. A
**date cell** is a number whose format's first section has `kind: "date"`;
a column's **type** for labels is the majority of its non-blank data cells
— date cells, other numbers, or everything else (ties → text).
- `values` — passes when the display text equals an entry, or `blanks` and
  the cell is blank, or the cell is a date cell whose serial's parts
  (`serialToParts`) equal a `DateGroupItem` down to its `grouping`. Empty `values`, no
  `blanks`, no groups hides every data row (Excel allows it).
- `custom` — each condition: `notEqual` with value `" "` means
  **non-blank** (Excel's encoding); a value with an unescaped `*` or `?`
  under `equal`/`notEqual` is a wildcard match on the display text (`~`
  escapes); otherwise, if the value parses as an invariant number and the
  cell is numeric, a numeric comparison; otherwise a text comparison of the
  display text with `sortCollator(tag, false)`. A blank fails every
  operator except `notEqual`. Numeric cells never match a text `greaterThan`
  et al. `join` combines the two results.
- `top10` — over the numeric cells of all data rows (visible or not):
  threshold = the `value`-th largest (smallest) value, or the value at rank
  `max(1, floor(count × value / 100))` for `percent`; passes when `≥`
  (`≤`) the threshold.
- `dynamic` — `aboveAverage`/`belowAverage` strictly against the mean of
  the numeric data cells; the date types compare the serial's date part with
  "today" from `now()` (local time) converted by `dateToSerial` in the
  workbook's date system: weeks start on Sunday, quarters are calendar
  quarters, `yearToDate` is 1 January ≤ d ≤ today, `Qn`/`Mn` match any
  year. Non-numeric cells fail.
- `preserved` — not evaluated: `applyAutoFilter` sets `filterHidden` to
  `rejectedByModelled || (hasPreserved && row.filterHidden)`.
Rows are evaluated from `range.top + 1` to `min(range.bottom,
usedRange.bottom)`; trailing rows past the used range are never hidden (no
million-row materialisation). A hidden-by-filter blank row is stored as a
`Row` with no cells and `filterHidden: true`.

### Filter-hidden or user-hidden
xlsx writes both as `<row hidden="1">`. On read, for a hidden row inside the
autofilter's data rows: `filterHidden` when `rowPasses` is false or the
filter carries any `preserved` criterion, else `hidden`. A hidden row outside
the range, or on a sheet without an autofilter, is `hidden`. Visible rows are
never hidden by reading. ODS states it on the wire: `table:visibility="filter"`
→ `filterHidden`, `"collapse"` → `hidden`. Writers: xlsx `hidden="1"` when
either; ODS `"collapse"` when `hidden` (whatever `filterHidden` is — the read
rule recovers the filter half), `"filter"` when only `filterHidden`.

### Choosing the range
Filter on: a single-cell selection → its `currentRegion`; a one-row
selection → those columns down to the `currentRegion` bottom; a multi-row
selection → itself. An empty region refuses in `sheet-alert`: "Select a cell
inside the data you want to filter". One autofilter per sheet.

### xlsx wire mapping
| Wire | Model |
| --- | --- |
| `<autoFilter ref="A1:D20" xr:uid>` | `range`; whole element → `sourceXml` |
| `filterColumn@colId` / `@hiddenButton="1"` or `@showButton="0"` | `colId` / `hideButton` |
| `filters@blank="1"`, `filter@val`, `dateGroupItem@year/month/day/hour/minute/second/dateTimeGrouping` | `values` |
| `customFilters@and="1"`, `customFilter@operator` (default `equal`) `@val` | `custom` |
| `top10@top` (default 1) `@percent` `@val` (`@filterVal` written, ignored on read) | `top10` |
| `dynamicFilter@type` (`@val`/`@maxVal`/`@valIso` ignored on read; `@val` = the mean written for the average types) | `dynamic` |
| `colorFilter`, `iconFilter` | `preserved` (`"colour"`/`"icon"`) |
| `autoFilter/sortState`, worksheet `<sortState>` | `autoFilter.sort`, `Sheet.sortState` |
| `sheetPr@filterMode="1"` | written iff some row `filterHidden` (attribute spliced into the retained `sheetPr`) |
| `<definedName name="_xlnm._FilterDatabase" localSheetId="i" hidden="1">'Sheet'!$A$1:$D$20</definedName>` | derived; reconciled by string splice in `<definedNames>` (created after `<sheets>` if absent) whenever `structureDirty` |

`autoFilter` and worksheet `sortState` join SPEC-122's rebuilt set at their
CT_Worksheet positions. Table autofilters (`xl/tables/*.xml`) and
`customSheetViews` stay verbatim. (SPEC-133, later, models defined names and
moves this reconciliation onto its `setReservedName`; nothing here waits for
it.)

### ODS wire mapping
`table:database-ranges` sits in `odsContentEnvelope.suffix`. A
`table:database-range` with `table:display-filter-buttons="true"` whose
`table:target-range-address` (`Sheet1.A1:Sheet1.D20`, quoted sheet names,
`$` ignored) lies on sheet S is S's autofilter — the anonymous
`__Anonymous_Sheet_DB__<n>` first, a named one only when S has none. The
writer regenerates the element when any sheet is dirty or `structureDirty`:
entries this spec does not own re-emitted verbatim (SPEC-133 later
serialises its own), each sheet's autofilter from `sourceXml` while
unchanged and still at its index, else serialised as
`__Anonymous_Sheet_DB__<sheet index>`; a new element goes after
`table:named-expressions`, before `table:data-pilot-tables`. Addresses go
through `odsRangeAddressToCanonical`/`canonicalToOdsRangeAddress` in
`src/lib/spreadsheet/ods/rangeAddress.ts` (SPEC-121's module: `"Sheet1.$A$1:.$B$3"`
↔ `"Sheet1!$A$1:$B$3"`, quoted names, `null` when unparseable).

`table:filter`: the top-level `table:filter-and` (or a lone child) is split
per `table:field-number` (offset from `range.left`):

| `table:operator` | Model |
| --- | --- |
| `=` with `table:filter-set-item`s, or a `table:filter-or` of `=` on one field | `values` (an item `table:value=""` → `blanks`) |
| `=` `!=` `<` `>` `<=` `>=` (`table:data-type="number"` → numeric) | `custom` operator |
| `begins-with` / `ends-with` / `contains` | `equal "x*"` / `"*x"` / `"*x*"` (wildcards in x escaped with `~`) |
| `does-not-begin-with` / `does-not-end-with` / `does-not-contain` | `notEqual` with the same patterns |
| `empty` / `!empty` | `values` blanks only / `notEqual " "` |
| `top values` / `bottom values` / `top percent` / `bottom percent` | `top10` |
| `match`, `!match`, colour conditions, an OR across fields | whole tree → `odsUnmappedFilter` |

Writes use the `table:filter-set-item` form for `values` and a
`table:filter-and` over columns. `table:sort` (`table:case-sensitive`) with
`table:sort-by table:field-number table:order="ascending|descending"
table:data-type="automatic"` ↔ `autoFilter.sort`.

### Structural edits
`autoFilter.range`, `autoFilter.sort.ref`, `Sheet.sortState.ref` and every
`SheetSortKey.col` follow `adjustRange`/`adjustIndex`; `colId`s at or right of an
inserted column shift by `count`; a deleted column drops its `FilterColumn`
and its sort keys (a sort record left with no keys is removed); deleting the
header row removes `autoFilter`; inserted rows are visible. The adjuster
replaces the objects (copy-on-write) and clears their `sourceXml`.

### Moves
The `autofilter` move adjuster: when `autoFilter.range` lies entirely inside
`move.from`, the autofilter (its `colId`s are relative and need no change,
its `sort.ref` and keys shift with it) moves by `(toTop − from.top, toLeft −
from.left)` onto `toSheetId`, replacing that sheet's autofilter; the old
data rows' `filterHidden` flags are cleared (row properties never move with
cells) and the filter is reapplied at its new place. `Sheet.sortState` moves
by the same rule. A range that overlaps the source or the destination only
partly is left alone — a filter does not break because some of its cells
were moved. Both objects are replaced, never mutated.

### UI
- `src/views/spreadsheet/useSortFilter.ts` (≤ 400 lines per
  `composableSize.spec.ts`; split a `useFilterPopover.ts` if it grows) owns the
  commands, the popover state and the keyboard shortcuts (ignored from inputs
  and the cell editor, the SPEC-021 convention); `SpreadsheetView.vue` gains
  one wiring call and the template hookup, and `SpreadsheetView.size.spec.ts`'s
  budget does not rise.
- Components in `src/components/spreadsheet/`: `SortDialog.vue`,
  `FilterPopover.vue`, `FilterCustomDialog.vue`, `FilterTop10Dialog.vue` (the
  dialogs on `BaseDialog.vue`), and `SheetAlertDialog.vue` (`sheet-alert`,
  `role="alertdialog"`, one `sheet-alert-ok`), which this spec creates and
  later specs (SPEC-139) reuse for their refusals. The `filter-button` is drawn by SPEC-117's
  `SheetCellLayer.vue` inside each of the range's header-row `sheet-cell`s
  (right-aligned, 16 px), so it virtualises with the grid.
- SPEC-120's Data menu gains, after "Recalculate" and a separator: Sort A to
  Z, Sort Z to A, Sort range…, separator, Filter, Clear filters, Reapply
  filter (SPEC-130 appends "Data validation…" after a separator). Menu items
  use `MenuItem.vue`'s `testid` prop; a disabled item states its reason
  through `menu-item-reason` ("Turn on Filter first" for Clear/Reapply
  without an autofilter).
- Conditions submenu items (Excel's): **Text Filters** — Equals…, Does Not
  Equal…, Begins With…, Ends With…, Contains…, Does Not Contain…, Custom
  Filter…; **Number Filters** — Equals…, Does Not Equal…, Greater Than…,
  Greater Than Or Equal To…, Less Than…, Less Than Or Equal To…, Between…,
  Top 10…, Above Average, Below Average, Custom Filter…; **Date Filters** —
  Equals…, Before…, After…, Between…, Tomorrow, Today, Yesterday, Next Week,
  This Week, Last Week, Next Month, This Month, Last Month, Next Year, This
  Year, Last Year, Year to Date, Custom Filter…. Each "…" opens the custom
  dialog pre-set; the others apply a `dynamic` criterion at once. Values typed
  into the custom dialog go through `parseInput`: a number or date is stored
  as invariant number text, anything else as text with its operator's
  wildcards; `*` and `?` typed by the user stay wildcards (Excel).
- The search box narrows the list case-insensitively; while it is non-empty
  "(Select All)" reads "(Select All Search Results)" and OK keeps only the
  checked matching items. Icons `filter`, `sort-ascending` and
  `sort-descending` join `src/components/icons/paths.ts`.

### Fixtures
Built by `spreadsheet/fixtures/build.ts` like SPEC-116's: `spec128-filter.xlsx`
has sheet "Sales" (Region, Rep, Amount, Date; 12 data rows, half of them
`<row>`s with no `r`) with an `autoFilter` over `A1:D13`, a `filters` column
on Region (`East`, `North`, `West`, `blank="1"` — the two South rows fail), a
`customFilters and="1"` pair on Amount (two more rows fail), those four rows
`hidden="1"`, one further `hidden="1"` row that every criterion passes (the
user-hidden one), `sheetPr filterMode="1"`, its `_xlnm._FilterDatabase` name
and an `autoFilter/sortState`; sheet "Colours" carries an autoFilter with a
`colorFilter dxfId="0"` (and its dxf in `styles.xml`) over two hidden rows,
and sheet "Log" a worksheet-level `sortState` with a `sortBy="cellColor"`
condition. `spec128-filter.ods` mirrors "Sales" with LibreOffice's shapes.

### Tests
- `sort.spec.ts`, `filter.spec.ts` — per the AC; collation over every
  `SpreadsheetLocale` tag; date dynamics with `setClock` (never real today).
- `xlsx/autoFilter.spec.ts`, `ods/databaseRanges.spec.ts` — every row of both
  wire tables both ways, `sourceXml` verbatim reuse, `_FilterDatabase`
  reconciliation, `filterMode`.
- `structure.spec.ts` — the `autoFilter` adjuster cases; `open.spec.ts` — the
  upgrade step; `roundTrip.spec.ts` — byte identity and model `toEqual`.
- `views/spreadsheet/useSortFilter.spec.ts`, the `SortDialog`/`FilterPopover`
  component specs, `e2e/SPEC-128.spec.ts`.

### Specs this changes
- SPEC-117 — `GridMetrics`, `NavContext.isRowHidden` and
  `selectionAggregates` treat `filterHidden` like `hidden`; reading ODS
  `table:visibility="filter"` yields `filterHidden` (only `"collapse"` stays
  `hidden`).
- SPEC-127 — its `EvalContext.rowVisibility` answers `"filtered"`.
- SPEC-121 — `SHEET_BY_REFERENCE_FIELDS` gains `autoFilter` and
  `sortState` (commands replace them); duplicate-sheet copies them (the
  copy's `odsName` dropped); `UNCONVERTIBLE_CONTENT` loses its "filters" row
  (autofilters now cross through the model) and gains "colour and icon
  filters" (marker: a `preserved` criterion or `odsUnmappedFilter`).
- SPEC-122 — xlsx rows write `hidden="1"` for either flag; `autoFilter` and
  `sortState` join the rebuilt set; ODS visibility joins the row-compression
  key; the commit path's refusal for protection covers sort and filter.
- SPEC-124 — Unhide clears `hidden` only; the retained-reference rows listed
  in the adjuster AC are retired (its e2e still holds — the modelled path
  writes the same shifted `ref`).
- SPEC-125 — copy, clear and fill act on `visibleRowSpans` of a filtered
  selection (paste is unchanged: like Excel it writes hidden rows too);
  `MOVE_ADJUSTERS` gains `autofilter`.
- SPEC-135 (later) registers its annotations in `ROW_PERMUTATION_ADJUSTERS`
  so notes, threads and links travel with their cells; nothing here names
  it.
- SPEC-076 — unchanged; its comparator stays in the text product.
- SPEC-133 — `_xlnm._FilterDatabase` is its reserved name whose meaning this
  spec owns; an Excel table's autofilter criteria reuse `FilterColumn` /
  `FilterCriteria` and `rowPasses` when SPEC-133 models them.
- SPEC-139 builds on `compareSortValues` (this spec's collator, not
  SPEC-076's) and `SortState`.
