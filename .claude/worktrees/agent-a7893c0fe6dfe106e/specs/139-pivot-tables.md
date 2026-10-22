---
id: SPEC-139
title: Pivot tables
status: draft
area: structure
complexity: complex
story_points: 21
depends_on: [SPEC-120, SPEC-124, SPEC-125, SPEC-128, SPEC-129, SPEC-133]
---
## Why
The pivot table is why many people open a spreadsheet at all: a thousand
rows of sales become "revenue by region by quarter" in four clicks. Today a
workbook that carries one opens with its cached cells visible — SPEC-116
preserves `xl/pivotTables/*`, `xl/pivotCache/*` and the `.ods`
`table:data-pilot-tables` verbatim — but nothing knows those cells belong to
anything. Type into one and SPEC-122 overwrites it like any cell; insert
rows above it and SPEC-124 moves the cells while the pivot definition's
`location` keeps pointing at the old rows, which is one of Excel's "We found
a problem with some content" repairs. A pivot's source range does not
follow a structural edit either. So pivots need modelling for correctness
before anyone builds one, and once modelled, building one is the feature.

Everything a pivot needs is already in this queue: values and their
comparison (SPEC-120), the sort order and collation (SPEC-128), Excel tables
and defined names as sources (SPEC-133), structural edits and the
`REFERENCE_ADJUSTERS` registry (SPEC-124), number formats (SPEC-118), the
writers' splice discipline (SPEC-121/122). What is missing is the pivot
itself: a model, a pure computation that turns a source range into an output
grid, both wire formats, refresh, and the field list. **This file is the
normative source for the pivot model (`SheetPivot` and its parts), the pivot
computation and output geometry, pivot refresh and overlap rules, both pivot
wire mappings (xlsx `pivotCacheDefinition`/`pivotCacheRecords`/
`pivotTableDefinition`, ODS `table:data-pilot-table`), and the pivot field
list.**

The model covers what most pivots use — rows, columns, values with the
eleven Excel aggregates, report filters, the three layouts, subtotals,
grand totals, item sorting and hidden items. A pivot using anything else
(calculated fields, grouping, "Show values as", value filters, OLAP,
external data) is **foreign**: its cached cells show, its parts survive
byte for byte, and the field list says so rather than recomputing it wrong.

## Acceptance criteria
- [ ] `Sheet.pivots?: SheetPivot[]` per "The pivot model" (types in `src/lib/spreadsheet/model.ts`); a pivot belongs to the sheet its output is on, references its source by sheet id, Excel table name or defined name, references fields by name, and stays plain JSON; pivot edits are copy-on-write (a new `pivots` array and a new pivot object), and `"pivots"` is appended to SPEC-121's `SHEET_BY_REFERENCE_FIELDS` because `itemOrder`, `hiddenItems` and `odsXml` grow with the data — undo snapshots hold the array by reference instead of cloning it
- [ ] `src/lib/spreadsheet/pivot.ts` exports `computePivot(model, pivot, ctx): PivotOutput | PivotFailure` — pure and deterministic per "Computation": every source row below the header is a record (blank rows and filter-hidden rows included), items are grouped by the item key (numbers at 15 significant digits, text case-insensitively with the first occurrence as caption, `null` as the blank item), ordered numbers < text < booleans < errors with the blank item last in both directions, and aggregated per the aggregation table — `count` counts every non-blank value including errors, `countNums` numbers only, every other aggregate propagates the first error in record order, and totals are computed from records, never from displayed subtotals
- [ ] The output grid follows "Output geometry" exactly for all three layouts, column axes, the values pseudo-field on either axis, subtotals at top or bottom, both grand totals, and report filters above the body; the two worked examples in the Reference are unit-tested cell for cell, as is the empty pivot's 3 × 18 placeholder
- [ ] Captions follow the table in "Captions" by the workbook's `format` — "Row Labels", "Grand Total", "(blank)", "Sum of Sales" for `.xlsx`; "Total Result", "(empty)", "Sum - Sales" for `.ods` — and a file's own captions (`rowHeaderCaption`, `colHeaderCaption`, `dataCaption`, `grandTotalCaption`, `missingCaption`) override them; a duplicate value-field name gains "2", "3"…; renaming a value field to a name already used by a field of the pivot is refused with "A field with that name already exists in this pivot table"
- [ ] Output is written into cells as values: item cells hold the item's value (a date item stays a serial carrying its source number format), header and label cells hold strings, data cells hold the aggregate (number or error) carrying the value field's `numberFormat` through `internStyle`, intersections with no records hold no cell; every other aspect of each in-extent cell's existing style is kept, and the pivot style (`PivotStyleLight16`, the default; "None" = no style) is a render-time layer by cell role per "Rendering", never baked into cell styles
- [ ] Refresh is manual and pinned: at open every pivot shows its cached cells and nothing is recomputed (Excel's behaviour); the field list's "Refresh" (`pivot-refresh`, Alt+F5 — `sheet.refreshPivot` in SPEC-123's `SHEET_SHORTCUTS`) recomputes the active pivot from the current source and Data ▸ "Refresh all" (`sheet-menu-refresh-all`, Ctrl+Alt+F5 — `sheet.refreshAllPivots`, best effort: Linux desktops keep Ctrl+Alt+F-keys for switching consoles, so the menu item is the guaranteed path) recomputes every non-foreign pivot in sheet order, toasting "Refreshed 3 pivot tables" or "Refreshed 2 pivot tables — 1 uses features that can't be refreshed here"; each refresh is one undo step "Refresh pivot table", and its changed rows go to SPEC-120's `recalculate`
- [ ] Refresh re-derives the field list from the source header (new columns appended to `fields`, vanished fields removed from every area); a blank header cell refuses with "Every column in the source needs a heading — fill in the blank heading cells, then try again"; a source deleted by a structural edit or sheet deletion refuses with "The pivot table's source data is no longer valid" and the cached cells stay
- [ ] Any output change (refresh, field-list edit, creation) is refused — SPEC-128's `sheet-alert` dialog shows the message and the model is untouched — when the new extent would cover a stored cell with a value or formula outside the pivot's current extent ("There's already data in Sheet1!E3:F9 — move or clear it, then try again"), another pivot ("A pivot table can't overlap another pivot table"), an Excel table ("A pivot table can't overlap a table"), a merge not wholly inside the current extent, or more than `MAX_PIVOT_OUTPUT_CELLS = 1_000_000` cells or the sheet's edge ("This pivot table won't fit on the sheet"); cells the old extent covered and the new one does not are cleared
- [ ] Cells inside any pivot's extent (foreign included) are protected: cell entry, paste, clear, fill, sort and filter that would change one are refused in `sheet-alert` with "You can't change this part of a pivot table" through `pivotBlocksCellEdit(model, sheetId, ranges)` — run where SPEC-122 runs `firstLockedCell`, before anything is written — and so is a SPEC-125 cut-move whose source block or destination covers part of a pivot (a block containing a pivot's whole extent moves the pivot), while formatting commands (SPEC-123) are allowed; a SPEC-124 insert or delete whose band intersects a pivot's extent rows (or columns) is refused with "This change would affect a pivot table — move the pivot table first, or use the field list" — `pivotBlocksStructuralEdit(model, edit)` joins the refusals `applyStructuralEdit` checks before anything changes, returning its `{ ok: false, message }` — and one wholly above (or left of) it moves it
- [ ] `pivot.ts` registers `{ id: "pivots", move }` in SPEC-125's `MOVE_ADJUSTERS` — a pivot whose whole extent lies in the moved block moves with it (anchor, body, extent, onto the destination sheet), a `{ sheetId, range }` source entirely inside the block re-points by the offset, partial overlaps stay, a source entirely inside an overwritten destination sets `sourceInvalid` — and `{ id: "pivots", adjust }` in SPEC-124's `REFERENCE_ADJUSTERS`: a row insert above a pivot moves its `anchor`, `body` and `extent`; a `{ sheetId, range }` source moves, grows and shrinks through SPEC-124's `adjustRange(range, edit)`, and its `null` (the whole source deleted) sets `sourceInvalid: true`; a sheet rename marks every pivot whose source or output is on that sheet `dirty`; deleting a sheet removes its pivots and invalidates sources on it; a SPEC-133 table rename (`sheet-table-name`) rewrites `{ tableName }` sources — each with a unit test; SPEC-124's four pivot rows in "Retained references" (`worksheetSource@ref`, `location@ref`, `table:data-pilot-table@table:target-range-address`, `table:source-cell-range@table:cell-range-address`) are narrowed to foreign pivots, whose bytes are otherwise verbatim, and modelled pivots move through the adjuster alone
- [ ] Insert ▸ "Pivot table…" (`sheet-menu-pivot-table`), and "Summarize with pivot table" (`sheet-table-summarize-pivot`) on SPEC-133's table design bar, open `pivot-create-dialog`: the source (`pivot-source-input`) defaults to the table when opened from the design bar, else the selection when it spans more than one cell, else the Excel table holding the active cell (by name), else SPEC-117's `currentRegion` around it ("Select a range with headings to make a pivot table" when that is a single empty cell); the destination is "New sheet" (`pivot-destination-new`, default — a SPEC-121 sheet inserted before the source sheet, anchor `A3`) or "Existing sheet" (`pivot-destination-existing` + `pivot-location-input`, a cell such as `Sheet2!B5`); OK (`pivot-create-ok`) creates the empty pivot named "PivotTable1" (the lowest free number in the workbook), selects it and opens the field list — one undo step "Insert pivot table"
- [ ] The field list (`pivot-field-list`, `src/components/spreadsheet/PivotFieldList.vue`) is the `"pivotFields"` pane of SPEC-129's sheet side panel (this spec widens `SheetPaneId` and adds its `SHEET_PANES` entry), opened with `open("pivotFields")` whenever the active cell enters a pivot's extent and closed when it leaves one, unless the user switched the panel to another pane meanwhile: one `pivot-field` (`data-field`) per field with a checkbox (`pivot-field-checkbox`) and the four areas `pivot-area-filters`, `pivot-area-columns`, `pivot-area-rows`, `pivot-area-values`, each listing `pivot-area-chip`s (`data-field`); ticking puts a field in Values (Sum, or Count when the column has a blank or a non-number) when all its non-blank values are numbers and none carries a format whose `parseFormatCode(code).kind` is `"date"`, otherwise in Rows; unticking removes it from every area
- [ ] Fields and chips drag between and within areas with a drop indicator, and every drag has a keyboard equivalent: a field's menu (`pivot-field-menu`; Enter, Space or Shift+F10 on a focused field) offers "Add to Filters", "Add to Columns", "Add to Rows", "Add to Values"; a chip's menu (`pivot-chip-menu`) offers "Move up", "Move down", "Move to beginning", "Move to end", "Move to Filters/Columns/Rows/Values", "Remove field" and, per area, "Value field settings…" or "Choose items…"; each change recomputes immediately as one undo step "Change pivot table", announced through a polite live region ("Region added to Rows")
- [ ] Value field settings (`pivot-value-settings`) edits the name (`pivot-value-name`), the aggregate (`pivot-value-aggregate`: Sum, Count, Average, Max, Min, Product, Count Numbers, StdDev, StdDevp, Var, Varp) and the number format (`pivot-value-number-format`, a format code validated by SPEC-118's `parseFormatCode`); in an `.ods` workbook the name field is disabled with the title "OpenDocument files don't store value field names"
- [ ] "Choose items…" on a Filters chip and the `pivot-filter-button` drawn on a filter's value cell open `pivot-filter-items`: "(All)" plus one `pivot-filter-item` checkbox per item in the field's order, a `pivot-filter-search` box, at most 1,000 rendered items ("N more — search to narrow"), OK `pivot-filter-ok`; the filter cell then reads the item, "(All)" or "(Multiple Items)"
- [ ] The field list's footer sets layout (`pivot-layout`: Compact, Outline, Tabular — Compact disabled in an `.ods` workbook, titled "OpenDocument pivot tables have no compact layout"), subtotals (`pivot-subtotals`: at top, at bottom, none — "at top" disabled for Tabular), grand totals (`pivot-grand-totals-rows` "Grand totals for rows", `pivot-grand-totals-columns` "Grand totals for columns"), the pivot's name (`pivot-name`, unique in the workbook), `pivot-refresh` and `pivot-delete` ("Delete pivot table" — clears its extent and removes it, one undo step)
- [ ] A foreign pivot (any trigger under "Foreign pivots (xlsx)" or "Foreign (ods)") shows its cached cells, keeps every one of its parts `equalBytes` on download, is skipped by "Refresh all", and its field list is read-only: fields and areas as read where they parse, every control disabled, `pivot-foreign-notice` reading "This pivot table uses features that can't be edited here"
- [ ] The xlsx reader and writer follow "xlsx wire mapping": a workbook with untouched pivots re-emits every part byte-identical; a dirty pivot's `pivotTableDefinition` is re-serialised by splice in schema order (modelled children rebuilt, `rowItems`/`colItems` omitted, everything else verbatim), its cache definition is rebuilt with `refreshOnLoad="1"` and consistent `sharedItems` flags, its records part becomes `<pivotCacheRecords count="0"/>`, and new pivots add their parts, relationships, `[Content_Types].xml` overrides and `xl/workbook.xml` `<pivotCaches>` entry; deleting the last pivot of a cache removes the cache, its parts and every reference to them
- [ ] The ods reader and writer follow "ODS wire mapping": the `table:data-pilot-tables` element in `odsContentEnvelope.suffix` is re-emitted verbatim while no pivot changed; a dirty pivot's `table:data-pilot-table` is re-serialised (modelled attributes and children rebuilt, `table:buttons` dropped, unknown attributes and children verbatim), new pivots are appended, deleted ones removed, and the element is inserted at its schema position when the workbook had none
- [ ] This spec bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`, which reads a dirty sheet's pivots from its retained parts (xlsx: the sheet's relationships in `opaque`; ods: the envelope suffix, matching each pivot's target sheet to the `table:name` in the sheet's retained `sourceXml`) — retained text SPEC-124/125's reference rewrite has already kept in step with every structural edit and move made before this spec landed, so the step reads it as it stands; stored workbooks gain their pivots on next load with ids, titles and edits intact
- [ ] Performance: `computePivot` over 100,000 records with two row fields, one column field and one value field completes in under 1,000 ms in Node (`src/lib/spreadsheet/pivot.perf.spec.ts`, the shape `src/lib/generate/performance.spec.ts` already has), and the output it builds is a single copy-on-write pass — one new `rows` array for the sheet, untouched rows keeping their identity
- [ ] Fixtures (built by `spreadsheet/fixtures/build.ts`, emulating their producers per "Fixtures"): `e2e/fixtures/spec139-sales.xlsx` (Excel-shaped, with records, `rowItems`/`colItems`, `x14` `extLst` and cached output), `spec139-foreign.xlsx` (one pivot per foreign trigger family) and `spec139-sales.ods` (LibreOffice-shaped); unit tests `spreadsheet/pivot.spec.ts` (computation, aggregation table, geometry, captions, overlap, guards, adjusters), `xlsx/pivot.spec.ts`, `ods/dataPilot.spec.ts`, `pivotRoundTrip.spec.ts` (byte identity for untouched pivots; rewrite → re-read → same model minus minted ids), `views/spreadsheet/usePivotTables.spec.ts`, `components/spreadsheet/PivotFieldList.spec.ts`
- [ ] e2e (`e2e/SPEC-139.spec.ts`): open `spec139-sales.xlsx` → the cached pivot shows "Grand Total" → click inside it → `pivot-field-list` shows Region in Rows → type into a pivot cell → refused with the message → on "Data" change a Sales value, Refresh → the grand total changes → Insert ▸ Pivot table from `Data!A1:E201` to a new sheet → tick Region and Sales → `A3` reads "Row Labels", `B3` "Sum of Sales" → add Product to Columns through `pivot-field-menu` with the keyboard only → set Sales to Average in `pivot-value-settings` → move Product to Filters through its chip menu and choose one item in `pivot-filter-items` → type a value two rows below the grand total, then add Units to Rows so the pivot would grow over it → refused with the "There's already data in …" message and the pivot unchanged → undo the filter change → `B1` reads "(All)" → download, reopen → both pivots and their field lists survive; then open `spec139-sales.ods`, refresh, download → `content.xml`'s data pilot names the same fields; `spec139-foreign.xlsx` shows `pivot-foreign-notice` and downloads its pivot parts `equalBytes`

## Out of scope
Calculated fields and items, grouping (dates, numbers, manual groups), "Show
values as" (% of total, running total, difference from — backlog beyond
this spec), label/value/top-10 filters, sort by value, collapsing and
expanding items, "Show items with no data", repeating item labels, blank
lines after items, a "Defer layout update" mode — each makes a pivot
foreign when read, so nothing is ever computed wrong. Editing which items
of a row or column field are hidden (read, honoured and preserved; only
filter fields are edited here). Pivot styles other than PivotStyleLight16
and None (any other name is kept and drawn as Light16), and a style
gallery. A "Move pivot table" dialog (cutting and pasting a whole pivot moves it), changing its source, "Show details" drill-down,
GETPIVOTDATA (SPEC-127 owns functions), pivot charts and slicers/timelines
(their pivots are foreign; the parts are preserved), OLAP and the Data
Model, external and consolidation sources, refresh on open, multiple pivots
sharing one new cache. Recommended pivot tables.

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — `SpreadsheetModel { format:
  "xlsx" | "ods"; sheets; opaque; odsContentEnvelope?: { prefix; suffix };
  readerVersion; … }`, `Sheet { id; name; rows: Row[]; partName?;
  sourceXml?; … }`, `Row { index; cells }`, `Cell { col; value: CellValue;
  formula?; styleId? }`, `CellValue = number | string | boolean | CellError |
  null`, `ErrorCode` order `#NULL!` … `#UNKNOWN!`. Invariants: copy-on-write
  rows, sparse, **the source is the splice base** (unmodelled XML re-emitted
  verbatim), dates are serials. Addressing in `address.ts`: `CellRef`,
  `GridRange` (inclusive, normalised), `parseA1`, `formatRange`,
  `getCell`, `usedRange`. Reader versions: `SPREADSHEET_READER_VERSION`,
  `ReaderUpgrade { from; upgradeDirtySheet(model, sheet, sourceParts) }`,
  `READER_UPGRADES`. The `.ods` reader scans `content.xml` once with
  `core/xmlScanner.ts` (`scanXml`, exact offsets); `table:data-pilot-tables`
  sits in the envelope suffix after the tables.
- `specs/117-…` — `CellStyle`, `selection.ts`'s `currentRegion`; `specs/118-…`
  — `formatValue(value, code, ctx)`, `FormatContext { locale; dateSystem }`,
  `parseFormatCode`, `BUILTIN_NUMBER_FORMATS`, and the date/time
  classification of a format code; `generalText` (116).
- `specs/120-calculation-engine-and-core-functions.md` — `Cell.value` is the
  cached or last computed result; `recalculate(model, changed)`; numeric
  comparison at 15 significant digits.
- `specs/121-…`/`specs/122-…` — `useSpreadsheetUndo.record(label, scope)`,
  snapshots per sheet with `SHEET_BY_REFERENCE_FIELDS` held by reference;
  SPEC-122's `firstLockedCell` check before the editor, Delete and Fill; `Sheet.dirty`,
  `structureDirty`; `writeXlsx`/`writeOds` substitution path;
  `internStyle(model, style): number` (append-only `cellXfs`/`numFmts`);
  CT_Worksheet splice; dropping `calcChain`.
- `specs/124-rows-columns-and-merged-cells.md` — `StructuralEdit = { kind:
  "insertRows" | "deleteRows" | "insertCols" | "deleteCols"; sheetId; at;
  count }`, `adjustFormula(text, hostSheetId, edit, sheetNameOf)`,
  `adjustRange(range, edit): GridRange | null`, `ReferenceAdjuster { id;
  adjust(model, edit): void }` in `REFERENCE_ADJUSTERS` (`structure.ts`),
  each with a "row insert above moves it" test, and `applyStructuralEdit(
  model, edit): { ok: true } | { ok: false; message }`, whose refusals run
  before anything changes; its retained-reference rewrite covers pivot
  `worksheetSource@ref`/`location@ref` and the two data-pilot addresses
  until this spec models them.
- `specs/125-copy-paste-fill-and-move.md` — `RangeMove { fromSheetId; from;
  toSheetId; toTop; toLeft }`, `MoveAdjuster { id; move(model, move) }` in
  `MOVE_ADJUSTERS` (`move.ts`), `moveRange`, `moveMapper`; a structure that
  registers a `ReferenceAdjuster` and travels with cells registers a
  `MoveAdjuster` with the same id.
- `specs/128-sort-and-filter.md` — `sortCollator(localeTag,
  caseSensitive)` and `compareSortValues(a, b, collator)` in
  `spreadsheet/sort.ts` (number < string < `false` < `true` < error, errors
  equal, blanks outside the order); the Data menu, whose items use
  `MenuItem.vue`'s `testid` prop (`sheet-menu-sort-asc` …); the `sheet-alert`
  dialog (`SheetAlertDialog.vue`) for refusals; filter-hidden rows.
- `specs/133-named-ranges-and-tables.md` — `Sheet.tables?: SheetTable[]`
  (`SheetTable { id; name; ref /* header + data + totals */; headerRow;
  totalsRow; … }`), `SpreadsheetModel.names: DefinedName[]` and
  `resolveDefinedName(model, name, qualifier, hostSheetId)`; the table
  design bar `sheet-table-design` and its rename `sheet-table-name`; table
  styles painted beneath a cell's own properties (the layer pivots copy);
  `.ods` tables written as `table:database-range`s, and an `.ods` read never
  creating a table. 133 assigns "Summarize with PivotTable" to this spec.
- `specs/129-…` — `useSheetSidePanel()` (`src/views/spreadsheet/
  useSheetSidePanel.ts`): `SheetPaneId` (widened per pane), `SHEET_PANES`,
  `{ activePane; open(id); close(); toggle(id); side; widthPx; collapsed;
  tabs }`, rendered by `SidePanel.vue` with the active pane's component in
  its default slot. `specs/123-…` — `SHEET_SHORTCUTS`.
- Equality: `equalBytes` (`src/lib/core/equality.ts`), part-level after
  `unzip`; model comparisons are field-level `toEqual` with minted ids
  stripped.

### The pivot model
```ts
// src/lib/spreadsheet/model.ts — Sheet gains `pivots?: SheetPivot[]`
export type PivotAggregate = "sum" | "count" | "average" | "max" | "min" | "product"
  | "countNums" | "stdDev" | "stdDevp" | "var" | "varp";     // = xlsx dataField@subtotal
export type PivotSource =
  | { sheetId: string; range: GridRange }   // header row included
  | { tableName: string }                   // SPEC-133 table / ODS database range
  | { definedName: string };                // a name resolving to one range
export interface PivotField {
  name: string;                             // header text, unique in the pivot
  sort: "ascending" | "descending" | "manual";
  itemOrder?: CellValue[];                  // manual order from the file; absent = first occurrence
  hiddenItems?: CellValue[];                // unticked items of a row/column field
  subtotals: boolean;                       // automatic subtotal when not the innermost axis field
}
export interface PivotValueField { field: string; aggregate: PivotAggregate;
  displayName: string; numberFormat?: string }
export interface PivotFilterField { field: string; selected: CellValue[] | null }   // null = all
export interface SheetPivot {
  id: string;                               // minted at read, never written
  name: string;
  source: PivotSource;
  sourceInvalid?: true;
  location: { anchor: CellRef; body: GridRange; extent: GridRange };
  fields: PivotField[];                     // every source column, in source order
  rowFields: string[]; columnFields: string[];
  valueFields: PivotValueField[];
  valuesAxis?: { axis: "rows" | "columns"; index: number };   // iff valueFields.length >= 2
  filterFields: PivotFilterField[];
  layout: "compact" | "outline" | "tabular";
  subtotalsAt: "top" | "bottom";            // tabular is always bottom
  grandTotals: { rows: boolean; columns: boolean };
  captions?: Partial<Record<"rowHeader" | "columnHeader" | "data" | "grandTotal" | "missing", string>>;
  styleName?: string;                       // absent = None
  foreign?: true;
  dirty?: true;                             // definition changed: its parts are rebuilt
  xlsx?: { partName: string; cacheId: number; cachePartName: string };
  odsXml?: string;                          // the table:data-pilot-table element as read
}
```
`anchor` is the body's top-left cell; `body` is the xlsx `location@ref`
area; `extent` adds the filter rows above. `grandTotals` uses Excel's
naming: `rows` is "grand totals for rows" — the **right-hand** total column
— and `columns` the **bottom** total row. Fields are named, not indexed, so
a refresh that adds a source column disturbs nothing.

### Computation
`computePivot(model, pivot, ctx: FormatContext)` resolves the source
(`{ sheetId, range }`; a table → its header and data rows, no totals row; a
defined name → its single range, anything else invalid), takes the top row
as headers (`formatValue` text, trimmed; duplicates suffixed "2", "3"… in
source order) and every row below as a record — a whole-column source ends
at the sheet's `usedRange`. Values are `Cell.value`. A record survives when
every filter field with `selected !== null` holds one of its selected items
and no row/column field holds a hidden item.

**Items.** Key: a number rounded to 15 significant digits; text
`normalize("NFC").toLowerCase()` (first occurrence's text is the caption;
`""` is a text item with an empty caption); a boolean; an error code;
`null` → the blank item. Ascending is SPEC-128's `compareSortValues(a, b,
sortCollator(locale, false))` — numbers, text, FALSE, TRUE, errors — with
errors (equal there) tie-broken in `ErrorCode` order and the blank item
after everything; descending reverses all but blank, which stays last; `manual` follows
`itemOrder` and appends unseen items in first-occurrence order.

| Aggregate | Over | No qualifying values | Errors among them |
| --- | --- | --- | --- |
| `sum`, `product` | numbers | 0 | first error, record order |
| `count` | every non-blank value | 0 | counted |
| `countNums` | numbers | 0 | ignored |
| `average` | numbers | `#DIV/0!` | first error |
| `max`, `min` | numbers | 0 | first error |
| `stdDev`, `var` (sample) | numbers | `#DIV/0!` when n < 2 | first error |
| `stdDevp`, `varp` | numbers | `#DIV/0!` when n = 0 | first error |

Text and booleans are never numbers here. Sums accumulate in record order;
variance is two-pass. An intersection with no surviving records has no
value at all (an empty cell, or `captions.missing`).

### Output geometry
Axis entries are the row (column) fields, plus `Σ` — the values pseudo-field
— at `valuesAxis.index` on `valuesAxis.axis`. The row tree nests each entry's
items under the previous entry's (only items with surviving records; `Σ`
expands to one child per value field). Let L = label columns (compact 1;
outline/tabular max(1, row entries)) and H = header rows: 1 + column
entries when the column axis holds a field, else 1.
- **Rows.** Compact: one line per node, caption in column 0 at indent =
  depth. Outline: the same, caption in column *depth*. Tabular: a node's
  caption shares the line of its first leaf, in column *depth*. A
  non-innermost node with `subtotals` carries its subtotal on its own line
  (`subtotalsAt: "top"`, compact/outline) or on a following "<item> Total"
  line (bottom; always for tabular); without subtotals compact/outline keep
  a value-less header line. Above `Σ`, subtotal and grand-total lines split
  into one line per value field ("East Sum of Sales", "Total Sum of Sales").
  `grandTotals.columns` adds "Grand Total" as the last line.
- **Columns** always nest tabularly whatever the layout: leaves left to
  right, a "<item> Total" column after a node with `subtotals`, and with
  `grandTotals.rows` (and a column field) "Grand Total" — or, with `Σ` on
  columns, "Total Sum of Sales" per value field — last.
- **Header rows.** With a column field: row 0 holds the single value
  field's name (empty with `Σ`) in column 0 and "Column Labels" (compact) or
  the first column field's name in column L; the next rows hold column
  captions per level; the last header row holds the row header in the label
  columns ("Row Labels", or the row field names). Otherwise one header row:
  row header(s), then each value field's name over its data column.
- `firstDataRow = H`, `firstDataCol = L`, `firstHeaderRow = H − 1` when the
  column axis is non-empty, else 1 (Excel's convention). **Filters**: F
  filter fields occupy the F rows ending two above the anchor — [field name
  | caption] — so `extent.top = anchor.row − F − 1` (the anchor moves down
  when there is no room). With no fields at all the pivot is a 3 × 18
  placeholder (Excel's `A3:C20`), no cells written, a `pivot-placeholder`
  overlay reading "Choose fields from the field list to build the pivot
  table".

Worked examples — source `Data!A1:C6`: Region/Product/Sales; records East
Apples 100, East Pears 200, West Apples 50, East apples 40, (blank) Pears 30.
Compact, rows Region, columns Product, Sum of Sales, anchor `A3`
(`location ref="A3:D8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"`):

| | A | B | C | D |
| --- | --- | --- | --- | --- |
| 3 | Sum of Sales | Column Labels | | |
| 4 | Row Labels | Apples | Pears | Grand Total |
| 5 | East | 140 | 200 | 340 |
| 6 | West | 50 | | 50 |
| 7 | (blank) | | 30 | 30 |
| 8 | Grand Total | 190 | 230 | 420 |

Tabular, rows Region then Product, Sum of Sales, subtotals on (`A3:C11`,
`firstHeaderRow="1" firstDataRow="1" firstDataCol="2"`):

| | A | B | C |
| --- | --- | --- | --- |
| 3 | Region | Product | Sum of Sales |
| 4 | East | Apples | 140 |
| 5 | | Pears | 200 |
| 6 | East Total | | 340 |
| 7 | West | Apples | 50 |
| 8 | West Total | | 50 |
| 9 | (blank) | Pears | 30 |
| 10 | (blank) Total | | 30 |
| 11 | Grand Total | | 420 |

### Captions
| Use | `.xlsx` workbooks | `.ods` workbooks |
| --- | --- | --- |
| Row / column header (compact) | Row Labels / Column Labels | — (no compact) |
| Values pseudo-field | Values | Data |
| Grand total / subtotal | Grand Total / "<item> Total" | Total Result / "<item> Result" |
| Blank item | (blank) | (empty) |
| Filter: all / several | (All) / (Multiple Items) | - all - / - multiple - |
| Default value name | "Sum of Sales", "Count of", "Average of", "Max of", "Min of", "Product of", "Count of" (countNums), "StdDev of", "StdDevp of", "Var of", "Varp of" | "Sum - Sales" with Sum, Count, Average, Max, Min, Product, Count, StDev, StDevP, Var, VarP |

### Rendering
Roles come from the last `computePivot` output, cached per pivot object
in the view (`filterLabel`, `filterValue`, `header`, `rowItem` + depth,
`columnItem`, `data`, `subtotal`, `grandTotal`); a pivot never computed this
session (cached from the file, or foreign) gets geometric roles from
`location` alone — filter rows, header rows, label columns, data, and a
grand-total last row/column per `grandTotals`. The grid resolves a pivot
cell's style as its own `CellStyle` over the role's pivot-style element.
PivotStyleLight16 as drawn here: header rows and the grand-total row bold
with fill accent1 at `lumMod 0.2, lumOff 0.8` and a thin accent1 rule (below
the header, above the grand total); compact/outline depth-0 item lines and
subtotal lines bold; depth indents of one level each. Excel draws pivot
styles the same way — from the definition, not the cells — so baking them
in would turn them into user formatting that survives a style change.

### xlsx wire mapping
| Part | Relationship from | Content type |
| --- | --- | --- |
| `xl/pivotTables/pivotTableN.xml` | the sheet's rels, type `…/relationships/pivotTable` | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml` |
| `xl/pivotCache/pivotCacheDefinitionN.xml` | workbook rels and the pivot table's rels, `…/pivotCacheDefinition` | `….spreadsheetml.pivotCacheDefinition+xml` |
| `xl/pivotCache/pivotCacheRecordsN.xml` | the cache definition's rels, `…/pivotCacheRecords` | `….spreadsheetml.pivotCacheRecords+xml` |

`xl/workbook.xml` `<pivotCaches><pivotCache cacheId r:id/></pivotCaches>`
sits after `calcPr`/`oleSize`/`customWorkbookViews` (a change sets
`structureDirty`); new parts take the lowest free N, a new cache the
largest `cacheId` + 1.

| Model | xlsx |
| --- | --- |
| `source` | `cacheSource type="worksheet"` + `worksheetSource@ref`/`@sheet` (absent sheet = the pivot's own), or `@name` → a table if SPEC-133 has one, else a defined name |
| `fields[i].name` | `cacheFields/cacheField[i]@name`; `pivotFields/pivotField[i]` |
| items | `sharedItems` children `s`, `n`, `b`, `d` (ISO → serial), `e`, `m` (blank) |
| `sort`, `itemOrder`, `hiddenItems`, `subtotals` | `pivotField@sortType` (absent = manual, order = `items/item@x`), `item@h="1"`, `@defaultSubtotal` |
| `rowFields`, `columnFields`, `valuesAxis` | `rowFields/field@x`, `colFields/field@x`; `x="-2"` is `Σ` |
| `valueFields` | `dataFields/dataField@fld`, `@subtotal` (absent = sum), `@name`, `@numFmtId` |
| `filterFields` | `pageFields/pageField@fld` + `@item` (one item), or `pivotField@multipleItemSelectionAllowed="1"` + hidden items |
| `layout`, `subtotalsAt` | row fields' `pivotField@compact`/`@outline` (default 1/1): 1/1 compact, 0/1 outline, 0/0 tabular; `@subtotalTop` |
| `grandTotals`, `captions`, `styleName` | `@rowGrandTotals`/`@colGrandTotals`; `@rowHeaderCaption`, `@colHeaderCaption`, `@dataCaption`, `@grandTotalCaption`, `@missingCaption`; `pivotTableStyleInfo@name` |
| `location` | `location@ref`, `@firstHeaderRow`, `@firstDataRow`, `@firstDataCol`, `@rowPageCount` = F, `@colPageCount` = 1 |

**Writing a dirty pivot.** `pivotTableDefinition` children in schema order
— `location, pivotFields, rowFields, rowItems, colFields, colItems,
pageFields, dataFields, formats, conditionalFormats, chartFormats,
pivotHierarchies, pivotTableStyleInfo, filters, rowHierarchiesUsage,
colHierarchiesUsage, extLst` — the modelled ones rebuilt, `rowItems`/
`colItems` omitted (optional; Excel rebuilds them on refresh), the rest and
unmodelled root attributes verbatim; a refresh that changed the field list
drops `formats`/`conditionalFormats`/`chartFormats` (they address fields by
index). The cache definition keeps `cacheSource` and its `extLst`, rebuilds
`cacheFields` from the source, sets `recordCount="0"` and
`refreshOnLoad="1"`, drops `refreshedDate`/`refreshedBy`; the records part
becomes `<pivotCacheRecords count="0"/>` — the shape non-Excel producers
that Excel accepts write: Excel recomputes the pivot from the source at
open, so a record copy of the source would double the file for nothing.
`sharedItems` lists every item of a field used on an axis or filter or
holding anything but numbers; an all-number data-only field lists none.
Flags: `containsBlank`, `containsDate`, `containsNumber`, `containsInteger`
(all numbers integral), `containsMixedTypes` (two or more of text, number,
date, boolean, error) set when true; `containsSemiMixedTypes="0"` when no
text and no blank, `containsString="0"` when no text, `containsNonDate="0"`
when every non-blank is a date; `minValue`/`maxValue` (or `minDate`/
`maxDate`) on unmixed numeric (date) fields; `count` on listed items. Number
items are written as `d` when the field's first non-blank source cell has a
format of `kind: "date"` (SPEC-118), the field's `numFmtId` that format's
id.

**Foreign pivots (xlsx)** — any of: `cacheSource@type` other than
`worksheet`, `worksheetSource@r:id` or `@connectionId`; `cacheField@formula`
or `@databaseField="0"`, `fieldGroup`, `calculatedItems`, `cacheHierarchies`
/`pivotHierarchies` or a Data Model `extLst` entry; `dataField@showDataAs`
other than `normal`; `filters`; an axis or page `pivotField` with `showAll`
absent or `1`, `insertBlankRow`, custom subtotal flags, an `item@sd="0"`, or
`x14:pivotField@fillDownLabels="1"`; row fields disagreeing on compact/
outline; `pageWrap` or `pageOverThenDown`; a slicer cache or a chart's
`c:pivotSource` naming the pivot.

### ODS wire mapping
| Model | `table:data-pilot-table` |
| --- | --- |
| `name`, `extent` | `@table:name`, `@table:target-range-address` (the whole extent) |
| `source` | `table:source-cell-range@table:cell-range-address` (`Data.A1:Data.C6`; `$` and quoted names accepted); a `{ tableName }` source is written as `table:database-source-range@table:database-range-name` naming the database range SPEC-133 writes for that table |
| `grandTotals` | `@table:grand-total`: `both`, `none`, `row` = `{ rows: true }` (LibreOffice's "Total rows", the right-hand column), `column` = `{ columns: true }` |
| areas | `table:data-pilot-field@table:source-field-name` + `@table:orientation` row/column/data/page; `@table:is-data-layout-field="true"` is `Σ` |
| `valueFields` | `orientation="data"` + `@table:function` (`sum`, `count`, `average`, `max`, `min`, `product`, `countnums`, `stdev`, `stdevp`, `var`, `varp`); no display name |
| filters, hidden items | `@table:selected-page` (one) or `table:data-pilot-member@table:display="false"` |
| `sort` | `table:data-pilot-sort-info@table:sort-mode` `name` + `@table:order`; `none` → manual, first occurrence; `manual` → member order |
| `subtotals`, `layout` | `table:data-pilot-subtotal@table:function="auto"` present; `table:data-pilot-layout-info@table:layout-mode` `tabular-layout`, `outline-subtotals-top`, `outline-subtotals-bottom` |

Member names are written as the item's `generalText` (`""` for blank) and
matched on read against `generalText`, then the formatted caption.
**Foreign (ods)**: a `table:database-source-range` source (an `.ods` read
makes no tables to bind it to, SPEC-133, and rebinding it to a fixed range
would change what LibreOffice refreshes), a `table:database-source-sql`/
`-table`/`-query` or `table:source-service` source; a `table:filter` in the source; a
`table:data-pilot-groups` or `table:data-pilot-field-reference` child; a
function outside the list; a subtotal function other than `auto`;
`table:show-empty="true"`; `table:data-pilot-display-info@table:enabled=
"true"`; `sort-mode="data"`; `table:add-empty-lines="true"`; repeated item
labels. The ods writer locates `table:data-pilot-tables` in the suffix with
`scanXml`; when inserting one it goes after `table:database-ranges`, else
after `table:named-expressions`, else at the start of the suffix. A changed
pivot list makes `content.xml` a rebuilt part even when no table is dirty.

### UI wiring
`src/views/spreadsheet/usePivotTables.ts` (≤ 400 lines) owns the create
dialog state, the active pivot (active cell inside an extent), field-list
commands, refresh, refusals (through `sheet-alert`), refresh results
(through `useToasts`) and the render-role cache; `SpreadsheetView.vue` gains one wiring call and the
template hookup, its size ratchet unchanged or lower. Components:
`PivotCreateDialog.vue`, `PivotFieldList.vue`, `PivotValueSettingsDialog.vue`,
`PivotFilterItemsDialog.vue` in `src/components/spreadsheet/`, typed
transforms in a sibling `pivotFieldListRender.ts`. Lib: `spreadsheet/pivot.ts`
(model operations, guards, adjuster, messages; re-exports `computePivot`),
`pivotCompute.ts` (computation and geometry), `xlsx/pivot.ts`,
`ods/dataPilot.ts`; all import
only spreadsheet siblings and core (`boundary.spec.ts`).

### Fixtures
`spec139-sales.xlsx` as Excel 365 writes it: a 200-record "Data" sheet
(Region, Product, Date with `numFmtId="14"`, Units, Sales; one blank Region,
one "apples"), a "Pivot" sheet whose pivot (rows Region, columns Product,
Sum of Sales, compact, PivotStyleLight16) has full `pivotCacheRecords`,
`rowItems`/`colItems`, `xr:uid`, `mc:Ignorable`, an `x14` pivot `extLst`, a
`refreshedDate`, and the cached output cells; a second pivot sharing its
cache. `spec139-foreign.xlsx`: a calculated field, a date-grouped field, a
"% of Grand Total" value and a top-10 filter, one pivot each.
`spec139-sales.ods` as LibreOffice writes it: the same data, a tabular data
pilot with a page field (`table:selected-page`), the data layout field,
`table:buttons`, `data-pilot-level` with members, display info, sort info
and layout info, and cached output cells with LibreOffice's captions.

### Specs this changes
- `specs/124-…` — `applyStructuralEdit`'s refusals gain
  `pivotBlocksStructuralEdit`; `REFERENCE_ADJUSTERS` gains `pivots`; its
  four pivot rows in "Retained references" narrow to foreign pivots.
- `specs/122-…`, `specs/125-…`, `specs/128-…` — cell entry, paste, clear,
  fill, cut-move, sort and filter call `pivotBlocksCellEdit` first;
  `MOVE_ADJUSTERS` gains `pivots`.
- `specs/121-…` — sheet rename marks affected pivots dirty (its xlsx pass
  already renames `worksheetSource@sheet`; its ods pass gains the two
  data-pilot addresses, for foreign pivots); sheet delete removes its
  pivots; a duplicated sheet's pivots are copied with fresh ids, the name
  uniqued as 121 already does for the part; the workbook rebuild emits
  `<pivotCaches>`; the ods writer treats a changed pivot list as a
  `content.xml` change; "Download a copy as…" writes modelled pivots from
  the model in the target format (compact becomes outline in `.ods`), so
  its "pivot tables" unconvertible-content row now fires for foreign
  pivots only.
- `specs/133-…` — a table rename rewrites `{ tableName }` pivot sources;
  the design bar gains `sheet-table-summarize-pivot`; creating or resizing
  a table over a pivot's extent is refused with "A pivot table can't
  overlap a table".
- `specs/116-…` — `SPREADSHEET_READER_VERSION` bumps; pivot parts stop
  being purely opaque for pivots that are modelled.
- `specs/129-…` — `SheetPaneId` gains `"pivotFields"`. `specs/123-…` —
  `SHEET_SHORTCUTS` gains `sheet.refreshPivot` and `sheet.refreshAllPivots`.
- `specs/000-roadmap.md` gains this spec's row.
