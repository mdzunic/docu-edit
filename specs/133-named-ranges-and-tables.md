---
id: SPEC-133
title: Named ranges and tables
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-119, SPEC-120, SPEC-123, SPEC-124, SPEC-125, SPEC-127]
---
## Why
Two features carry most of the structure in a modern workbook, and this
product reads neither. A **defined name** (`TaxRate`, `Q1_Sales`,
`Sheet2!Lookup`) lets a formula say what it means; an **Excel table**
(`xl/tables/table1.xml`) turns a block of cells into a typed list that grows
as rows are typed under it, formats itself in bands, totals itself, and is
referenced as `=SUM(Sales[Amount])` or `=[@Price]*[@Qty]`. Today the names sit
unread in `xl/workbook.xml` and the tables in parts nobody opens, so SPEC-120
shows every such formula's cached value until something it depends on
changes — and then computes `#NAME?` or `#REF!` for a formula that is
perfectly good in every other spreadsheet.

The pieces this needs are in place. SPEC-119 parses formulas into an AST that
already has a `name` node and a `structured` node for table references;
SPEC-120 evaluates the AST and calls out to an `EvalContext`; SPEC-124 gives
every reference-bearing structure a place in `REFERENCE_ADJUSTERS`; SPEC-122's
point mode lets a dialog field pick a range off the grid; SPEC-123 interns
styles. What is missing is the model, the two wire formats each way, the
resolver the engine calls, and the surfaces — the Name Box, the Name Manager,
Insert ▸ Table and the table's design controls.

**This file is the normative source for defined names (`DefinedName`,
scopes, validity, resolution, the reserved `_xlnm.*` representation), Excel
tables (`SheetTable`), the meaning of every structured-reference form and its
translation to each format, and how a table style renders.** SPEC-128 and
SPEC-136 own what `_xlnm._FilterDatabase` and `_xlnm.Print_Area` mean; this
file owns how they are stored.

ODS has names but no tables. A table written to `.ods` becomes what
LibreOffice can hold — a named database range with the table's look baked
into the cells and its structured references spelled as plain ranges — and
the spec says so in the UI rather than pretending otherwise.

## Acceptance criteria
- [ ] `model.ts` gains the model per "The model": `SpreadsheetModel.names: DefinedName[]` and `namesDirty?: true`, `DefinedName { name; scopeSheetId?; formula; hidden?; comment?; formulaUnparsed?; sourceAttrs? }`, `Sheet.tables?: SheetTable[]`, `SheetTable { id; name; ref; headerRow; totalsRow; columns; style; autoFilter?; locked?; partName?; xlsxId?; dirty? }`, `SheetTableColumn`, `SheetTableStyle`, `TotalsFunction`; SPEC-121's `createBlankWorkbook` starts with `names: []`; the model stays plain JSON
- [ ] `names.ts` exports `validateName(model, name, scopeSheetId?, except?)` returning a `NameProblem` or `null` per "Name validity" (Excel's rules: 1–255 characters, a letter/underscore/backslash first, then letters, digits, `.`, `_`, `\`, `?`; not an A1 or R1C1 reference; not `_xlnm.`-prefixed; case-insensitively unique within its scope and distinct from every table name) and `NAME_PROBLEM_MESSAGES` with the exact strings pinned there; tables use the same validator plus workbook-wide uniqueness against names of every scope
- [ ] xlsx read: `workbook.xml` `definedNames/definedName` → `names` per "Names on the wire" — `localSheetId` (0-based position in `<sheets>`, hidden and chart sheets counted) → `scopeSheetId`, `hidden="1"`, `comment`, text through `fromXlsxFormula` then `structuredFromXlsx`, unparseable text kept with `formulaUnparsed`, every other attribute verbatim in `sourceAttrs`; xlsx write: whenever `workbook.xml` is rebuilt (`namesDirty` or SPEC-121's `structureDirty`) `<definedNames>` is emitted from the model with `localSheetId` recomputed from the scope sheet's current index — this supersedes SPEC-121's textual renumbering — and a clean workbook re-emits `workbook.xml` byte-identical
- [ ] ods read: `table:named-expressions` in `office:spreadsheet` (workbook scope, in `odsContentEnvelope.suffix`) and inside a `table:table` (that sheet's scope) → `names`, `table:named-range` through `odsRangeAddressToCanonical` and `table:named-expression` through `fromOdsFormula`, both rebased from `table:base-cell-address` to A1 by `rebaseRelative`; ods write emits workbook names into the suffix's `table:named-expressions` (replacing it, or inserting it right after the last table) and local ones as the last child of their dirty table, a reference-only formula as `table:named-range`, anything else as `table:named-expression`
- [ ] Reserved names per "Reserved names": every `_xlnm.*` name (`Print_Area`, `Print_Titles`, `_FilterDatabase`, `Criteria`, `Extract`, others) is a sheet-scoped `DefinedName` spelled as in the file, adjusted like any name, never shown in the Name Box or Name Manager and never creatable or editable there, read and written by their owners through `getReservedName`/`setReservedName`; `_xlfn.`/`_xlpm.` placeholder names are kept verbatim, hidden and never resolved
- [ ] Name resolution per "Resolution": SPEC-120's `EvalContext` gains `resolveName(name, sheetQualifier, host)` backed by `resolveDefinedName(model, name, qualifier, hostSheetId)` — the host's sheet scope first, then the workbook; a qualified `Sheet2!X` that sheet's scope then the workbook — evaluating the name's formula with its relative parts offset from A1 to the host cell modulo the grid (Excel's wrap), names nesting to a depth of 64 before `#NAME?`; an unknown name is `#NAME?`; resolution reports its areas through `ctx.recordDependency`, and defining, editing or deleting a name recalculates exactly the cells whose formulas mention it
- [ ] SPEC-124's `REFERENCE_ADJUSTERS` gains `{ id: "names" }` (each name's formula through `adjustFormula`, host = its scope sheet, else the edited sheet) and `{ id: "tables" }` per "Tables follow cells", and SPEC-125's `MOVE_ADJUSTERS` gains `names` (each formula through `repointFormula`, so a name follows its cells when they are cut-pasted or drag-moved, as in Excel) and `tables` per "Tables and moves"; the same change retires SPEC-124's retained-reference rows for `xl/workbook.xml` `definedName` text, `xl/tables/tableN.xml` `table@ref`/`autoFilter@ref` and ODS `table:named-range`/`table:named-expression` (the `table:database-range` row stays SPEC-128's); SPEC-121's rename and delete apply their `SheetRemap` through `remapSheetReferences` to every name's formula and every table's `calculatedFormula`/`totalsFormula`, delete also removes the sheet's local names and tables, and duplicate copies local names and tables per "Tables on the wire" — this spec adds those steps to `sheets.ts`; `names.spec.ts` and `tables.spec.ts` assert that a row insert above a name's range and above a table, and a `moveRange` of a block containing each, move them
- [ ] The Name Box (`name-box`, SPEC-117) gains names per "The Name Box": typing a reference still navigates; typing an existing visible name or a table name selects its range (switching sheet); typing a valid new name defines a workbook-scope name for the current selection (absolute, sheet-qualified, a union for several areas) as one undo entry "Define name"; an invalid entry shows its `NAME_PROBLEM_MESSAGES` text as a toast and changes nothing; the dropdown `name-box-dropdown` lists visible names and tables as `name-box-option`
- [ ] Formulas ▸ Name Manager… (`sheet-menu-name-manager`; this spec adds the Formulas menu to the spreadsheet menu bar, after Format; Ctrl+F3 / Cmd+F3 as `sheet.nameManager` in SPEC-123's `SHEET_SHORTCUTS` — best effort on macOS, whose Cmd+F3 shows the desktop, so the menu item is the guaranteed path) opens `NameManagerDialog.vue` (`name-manager`): one `name-manager-row` per visible name with its name, current value (`generalText` of its evaluation, "{…}" for an array), refers-to and scope ("Workbook" or the sheet name), and `name-manager-new`/`-edit`/`-delete`/`-close`; New/Edit open `NameEditorDialog.vue` (`name-editor`: `name-editor-name`, `name-editor-scope`, `name-editor-comment`, `name-editor-refers-to` with SPEC-122 point mode, `name-editor-ok`, `name-editor-error`), validated live; changing a name's text rewrites every formula that used the old one; labels "Define name", "Edit name", "Delete name"
- [ ] xlsx tables read per "Tables on the wire": the worksheet's `<tableParts>` → `…/relationships/table` → `xl/tables/tableN.xml` → `SheetTable` (`displayName` → `name`, `ref`, `headerRowCount`, `totalsRowCount`, `tableColumn` `name`/`id`/`totalsRowFunction`/`totalsRowLabel`/`calculatedColumnFormula`/`totalsRowFormula`, `tableStyleInfo`, `autoFilter` kept as `sourceXml`); a `tableType` of `queryTable` or `xml` sets `locked`; clean tables write their part verbatim, `dirty` ones are regenerated keeping every unmodelled attribute and child (matched per column by `xlsxId`), new ones follow the golden `xlsx/tableGolden.ts` (hand-authored against ECMA-376 §18.5 before the serialiser) with a fresh part `xl/tables/table{n}.xml`, a workbook-unique `id`, the sheet relationship, `<tableParts>` in CT_Worksheet order and the `application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml` Override; a deleted table's part, rel and Override leave the package
- [ ] Structured references evaluate per "Structured references": `resolveStructuredReference(model, node, host)` (`tables.ts`) gives every form in the table there — `T`, `T[#All]`, `T[#Data]`, `T[#Headers]`, `T[#Totals]`, `T[Col]`, `T[[#Headers],[Col]]`, `T[[#Headers],[#Data],[Col]]`, `T[[#Data],[#Totals],[Col]]`, `T[[A]:[C]]`, `T[@Col]`, `T[@]`, `T[@[A]:[C]]` — its range or error (`#REF!` for a missing table, column, header or totals row; `#VALUE!` for a this-row reference from outside the data body), with `'` escaping `[`, `]`, `#` and `'` in column names; `EvalContext` gains `resolveStructured`, and changing a table's `ref` or columns recalculates every formula that names the table
- [ ] Canonical form per "Canonical and wire forms": `Cell.formula` always carries the table name and uses `@` for this-row (`Sales[@Amount]`); a formula committed inside a table that omits the name (`[@Amount]`, `[Amount]`) is qualified with the host table by `qualifyStructuredReferences`; `structuredToXlsx`/`structuredFromXlsx` map `@` to and from the file's `[#This Row]` form around SPEC-119's dialect functions; ODS output replaces every structured reference with its A1 equivalent through `expandStructuredReferences` before `toOdsFormula`
- [ ] Insert ▸ Table (`sheet-menu-insert-table`; no keyboard shortcut — Excel's Ctrl+T and Ctrl+L are the browsers' new-tab and address-bar keys, and Ctrl/Cmd+T is in `RESERVED_BROWSER_KEYS`, which SPEC-123's test keeps out of `SHEET_SHORTCUTS`) opens `CreateTableDialog.vue` (`sheet-create-table`: `-range` with point mode defaulting to the selection or the active cell's `currentRegion`, `-headers` defaulting to `detectHeaderRow`, `-ok`, `-cancel`, `-error`); `createTable` per "Creating a table" names it `Table{n}`, normalises the header row ("Column{n}" for blanks, numeric suffixes for duplicates, every header a string), writes a header row above when headers are off, unmerges merged cells inside, refuses an overlap with another table ("Tables can't overlap") or a multi-cell array formula inside ("Tables can't contain multi-cell array formulas"), and applies `TableStyleMedium2` with banded rows and the filter button — one undo entry "Create table"
- [ ] The table grows as data is typed per "Auto-expansion and calculated columns": a committed non-empty entry directly below the last data row (no totals row) or directly right of the last column extends `ref` by one row or column (a new column named from its header cell, else `Column{n}`), fills calculated columns into a new row, and is its own undo entry "Expand table" recorded after the entry's own, so one undo removes only the expansion; a `locked` table never expands
- [ ] Calculated columns per "Auto-expansion and calculated columns": committing a formula into a data cell of a column whose other data cells are all empty, or all hold its current `calculatedFormula`, fills the formula into every data row (A1 parts shifted per row with `shiftFormula`, structured parts unchanged), sets `calculatedFormula`, and records "Fill table column" after the entry
- [ ] Totals row per "Header and totals rows": `sheet-table-option-totals-row` adds the row below the table (into an empty row, else via a full-row insert) with "Total" in the first column and SUBTOTAL `sum` (numeric column) or `count` in the last; the totals-cell menu `sheet-table-totals-menu` offers `sheet-table-totals-<none|average|count|countNums|max|min|sum|stdDev|var>` writing `SUBTOTAL(code,T[Col])` with Excel's 101–110 codes; turning it off clears the row and keeps the functions for next time; SUBTOTAL is SPEC-127's
- [ ] The table design bar (`sheet-table-design`) appears while the active cell is in a table: `sheet-table-name` (renames, validated, rewriting every formula and name that used the old name), the options `sheet-table-option-header-row`/`-totals-row`/`-banded-rows`/`-banded-columns`/`-first-column`/`-last-column`/`-filter-button`, `sheet-table-styles` (the gallery), `sheet-table-resize` (`sheet-resize-table` dialog: the header row must stay in the same row and the new range must overlap the old — "The header row must stay in the same row and the new range must overlap the table"), `sheet-table-convert` (confirm `sheet-table-convert-confirm`; per "Convert to range") — each one undo entry ("Rename table", "Table options", "Table style", "Resize table", "Convert to range"); editing a header cell renames its column the same way
- [ ] Table styles render per "Table styles": `tableStyles.ts` exports `BUILT_IN_TABLE_STYLE_NAMES` (Excel's 60: `TableStyleLight1`–`21`, `TableStyleMedium1`–`28`, `TableStyleDark1`–`11`) and `tableStyleLook(name, theme)`, generated from the family templates there for Light 1–21, Medium 1–28 and Dark 1–7 and the neutral fallback for Dark 8–11, workbook-defined styles and unknown names; the grid paints each table cell's look beneath the cell's own explicitly set properties (and beneath SPEC-129's conditional formats), stripes alternating over visible data rows; the gallery (`sheet-table-style-gallery`, `sheet-table-style-none`, `sheet-table-style-<name>`) previews every built-in with the same function; every name round-trips unchanged
- [ ] ods write per "Tables in ODS": each table becomes a `table:database-range` named after it in the suffix's `table:database-ranges` (`table:target-range-address`, `table:contains-header`, `table:display-filter-buttons`), its look is baked into the written cells as generated automatic styles without touching `model.styles`, and its structured references are written as ranges; reading an `.ods` never creates a table from a database range; inserting a table in an `.ods` workbook shows once per session the notice `sheet-table-ods-notice`: "OpenDocument has no tables — this saves as a named range with the table's formatting applied to its cells"
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade` per "Reader upgrade": its `upgradeDirtySheet` reads a dirty sheet's tables from its retained source and `opaque`, and a new optional `ReaderUpgrade.upgradeWorkbook(model, sourceParts)` — run once when any sheet is dirty — reads the names from the source workbook part or content envelope, mapping scopes and renamed or deleted sheets onto the current model
- [ ] Fixtures from `spreadsheet/fixtures/build.ts`: `e2e/fixtures/spec133-names-tables.xlsx` (Excel-shaped: workbook and sheet-scoped names incl. a relative name and a constant `TaxRate`, a hidden name, `_xlnm.Print_Area`, `_xlnm._FilterDatabase hidden="1"`, a `_xlfn.` placeholder; a table with totals row, a calculated column written `Sales[[#This Row],[Price]]*Sales[[#This Row],[Qty]]`, `xr:uid`/`xr3:uid`, `dataDxfId`; a second table without header row; formulas `SUM(Sales[Amount])`, `Sales[[#Totals],[Amount]]`, `Sales['#Units]`) and `spec133-names.ods` (LibreOffice-shaped: global and table-local `table:named-range`/`table:named-expression` with non-A1 base cells, an anonymous and a named `table:database-range`)
- [ ] Tests: `names.spec.ts` (validity table, resolution order, relative wrap, nesting cap, adjuster incl. row-insert-above, rename/delete/duplicate), `tables.spec.ts` (every structured form, escapes, qualify/expand/wire mapping, create/expand/calculated/totals/rename/resize/convert, the adjuster incl. row-insert-above and the deletion rules), `tableStyles.spec.ts` (all 60 names produce a look; family slots), `xlsx/definedNames.spec.ts`, `xlsx/tableXml.spec.ts` (golden, verbatim-preserving regeneration), `ods/namedExpressions.spec.ts`, `namesTablesRoundTrip.spec.ts` (`equalBytes` untouched, `equalStructural` after save → reopen), `open.spec.ts` (both upgrade hooks), `useNames.spec.ts`, `useSheetTables.spec.ts`
- [ ] e2e (`e2e/SPEC-133.spec.ts`): open `spec133-names-tables.xlsx` → `SUM(Sales[Amount])` shows its value → change an amount → it recalculates → type `Margin` in the Name Box over `B2:B5` → `=SUM(Margin)` in a cell evaluates → Name Manager lists `Margin`, `TaxRate` and not `Print_Area` → type a value under the table → it expands with banding → type `=[@Price]*2` in the new column → the column fills → toggle the totals row → pick Average → rename the table to `Orders` → the formulas read `Orders[...]` → Insert ▸ Table on a fresh block → a `Table2` appears styled → download, reopen → names, both tables and formulas persist; download as `.ods`, reopen → values identical, no tables, the names present

## Out of scope
Custom table styles: creating or editing them (workbook-defined
`<tableStyles>` render with the neutral fallback and round-trip untouched).
Table slicers, refreshing query tables, the resize handle at a table's corner
(the dialog resizes), auto-expansion on paste or fill (typed entry only),
inserting/deleting table rows and columns as table-local operations (SPEC-124's
whole-row/column commands apply), the "Summarize with PivotTable" and
"Remove Duplicates" table commands (SPEC-139 / Backlog), filtering and
sorting inside a table (SPEC-128), LET/LAMBDA names and dynamic arrays (queue
backlog), names referring to other workbooks (kept, never resolved), the
"Create from Selection" and "Apply Names" commands, Go To by name (SPEC-134
calls `resolveDefinedName`), and what print areas and filter databases do
(SPEC-136, SPEC-128).

## Reference

### Dependencies restated
- `specs/116-…` — the model and invariants (copy-on-write, sparse, the
  source is the splice base, plain JSON), `GridRange { top; left; bottom;
  right }` (inclusive), `parseRange`/`formatRange`/`getCell`, reader versions
  (`ReaderUpgrade { from; upgradeDirtySheet(model, sheet, sourceParts) }`),
  `odsContentEnvelope { prefix; suffix }` (the suffix holds
  `table:named-expressions`, `table:database-ranges`, … in that order).
- `specs/119-…` — canonical formula text (no `=`, en-US, `'Sheet'!A1`),
  `parseFormula`, the `name` and `structured` nodes, `ReferenceNode`,
  `shiftFormula(text, dRow, dCol)`, `renameSheetInFormula(text, from, to)`,
  `fromXlsxFormula`/`toXlsxFormula`/`fromOdsFormula`/`toOdsFormula`
  (token-level, whitespace kept), `Cell.formulaUnparsed`,
  `rewriteReferences(text, rewrite)` (the `ref`-token primitive). A table
  reference is one `structured` token and parses to `{ kind: "structured";
  table?: string; text: string }`, kept verbatim and never interpreted by
  SPEC-119: this spec parses `text` (`parseStructuredText`) and gives it its
  meaning, and its own rewrites (rename table or column, `#REF!` on removal,
  qualification, expansion to A1) are token-level over `tokenize`, touching
  `structured` tokens only.
- `specs/120-…` — `evaluate(node, ctx: EvalContext)` with `ctx.host` and
  `ctx.origin` (a relative part resolves to `part + (host − origin)`),
  `recordDependency(area)`, `EvalValue`, `DependencyGraph`,
  `recalculate(model, changed)`; its `EvalContext` leaves
  `resolveName(name, sheetQualifier, host)` and `resolveStructured(node,
  host)` to this spec, which owns their signatures (below); while absent, a
  name or structured node makes a cell uncomputable. `specs/117-…` — `currentRegion`, `name-box`,
  hidden rows, the style resolution order. `specs/122-…` — cell entry commit,
  point mode, CT_Worksheet order (`tableParts` before `extLst`), ODS table
  writer and generated automatic styles. `specs/123-…` — `internStyle`,
  `applyStylePatch`. `specs/124-…` — `StructuralEdit`, `adjustFormula(text,
  hostSheetId, edit, sheetNameOf)`, `REFERENCE_ADJUSTERS`, the row insert and
  merge commands, and its "Retained references" rewrite, which until this
  spec moves `definedName` text, table `ref`s and ODS named ranges in the
  retained source — so the upgrade steps below read already-adjusted text.
  `specs/125-…` — `MOVE_ADJUSTERS`, `MoveAdjuster { id; move(model, move) }`,
  `RangeMove`, `moveRange` (with its refusal pass), `repointFormula`.
  `specs/121-…` — undo (`record`/`commit`, `SHEET_BY_REFERENCE_FIELDS` —
  this spec appends `"tables"` — and `WORKBOOK_SNAPSHOT_FIELDS` — it appends
  `"names"`, `"namesDirty"`), `SheetRemap`/`remapSheetReferences`,
  `sourceSheetRemap(model, parts)`, the duplicate's `<name>_<n>` table
  naming, `createBlankWorkbook`, workbook-part rebuild.
- `src/lib/spreadsheet/ods/rangeAddress.ts` (SPEC-121):
  `odsRangeAddressToCanonical(text): string | null`,
  `canonicalToOdsRangeAddress(canonical): string | null`.
- `specs/127-…` — `SUBTOTAL` (codes 1–11 and 101–111, skipping cells that
  call SUBTOTAL/AGGREGATE) for totals rows. `specs/128-…` — its
  `_xlnm._FilterDatabase` reconciliation (a string splice in
  `<definedNames>`), which this spec moves onto `setReservedName`.
  `specs/132-…` — `resolveChartRange(model, canonical, resolveName?)`'s
  optional name hook.

### The model
```ts
// src/lib/spreadsheet/model.ts (SPEC-133 additions)
export interface DefinedName {
  name: string;               // as written; compared case-insensitively
  scopeSheetId?: string;      // absent = workbook scope
  formula: string;            // canonical; relative parts are relative to A1
  hidden?: true;
  comment?: string;
  formulaUnparsed?: true;     // kept in its source dialect, never evaluated
  sourceAttrs?: string;       // xlsx: unmodelled attributes verbatim; ods: table:range-usable-as
}
export type TotalsFunction = "sum" | "min" | "max" | "average" | "count" | "countNums" | "stdDev" | "var" | "custom";
export interface SheetTableColumn {
  name: string;               // always equals the header cell's string when there is a header row
  xlsxId?: number;            // tableColumn@id (new columns: 1 + max)
  totalsFunction?: TotalsFunction;
  totalsFormula?: string;     // canonical; "custom" only
  totalsLabel?: string;
  calculatedFormula?: string; // canonical
}
export interface SheetTableStyle {
  name?: string;              // absent = no style ("None")
  showRowStripes: boolean; showColumnStripes: boolean; showFirstColumn: boolean; showLastColumn: boolean;
}
export interface SheetTable {
  id: string;                 // minted at read, never written
  name: string;               // displayName — what formulas use
  ref: GridRange;             // header + data + totals
  headerRow: boolean;
  totalsRow: boolean;
  columns: SheetTableColumn[];
  style: SheetTableStyle;
  autoFilter?: { sourceXml?: string };   // present = filter buttons; criteria are SPEC-128's
  locked?: true;              // tableType queryTable/xml: style options only
  partName?: string;          // "xl/tables/table1.xml" — bytes stay in opaque
  xlsxId?: number;            // table@id, unique in the workbook
  dirty?: true;               // regenerate the part
}
// SpreadsheetModel gains: names: DefinedName[]; namesDirty?: true
// Sheet gains:           tables?: SheetTable[]
```
Data body = `ref` minus the header row (if any) and the totals row (if any);
it is never empty — a deletion that would leave no data row deletes the table.

### Name validity

| `NameProblem` | Rule | Message |
| --- | --- | --- |
| `empty` | zero characters | "Enter a name" |
| `tooLong` | over 255 characters | "Names can be at most 255 characters" |
| `invalidCharacters` | not `^[\p{L}_\\][\p{L}\p{N}_.\\?]*$`u | "Names must start with a letter, underscore or backslash and can't contain spaces or symbols" |
| `looksLikeReference` | an A1 cell inside the grid (`B12`, `$xfd$1`), or R1C1 (`R`, `C`, `R2`, `C3`, `R1C1`, any case) | "A name can't look like a cell reference, such as A1 or R1C1" |
| `reserved` | starts `_xlnm.` | "Names starting with _xlnm. are reserved" |
| `duplicate` | same name (case-insensitive) in the same scope | "That name already exists in this scope" |
| `tableConflict` | equals a table name, or (for a table) any name or other table | "That name is already used by a table or a range name" |

`XFE1` (outside the grid) and `Sales2024` are valid. `uniqueTableName` takes
`Table{n}` for a new table (smallest `n` ≥ 1) and `<name>_<n>` for a copy
(smallest `n` ≥ 2, SPEC-121's rule), each passing the validator.

### Resolution
`resolveDefinedName(model, name, qualifier, hostSheetId)`: a qualified name
looks in that sheet's scope, then the workbook; an unqualified one in the
host sheet's scope, then the workbook; reserved names also answer to their
short form (`Print_Area`) within their sheet. The hooks:
```ts
resolveName(name: string, sheetQualifier: string | undefined, host: EvalContext["host"]): EvalValue;
resolveStructured(node: { kind: "structured"; table?: string; text: string }, host: EvalContext["host"]): EvalValue;
```
`resolveName` parses the name's formula once per formula text (a cache keyed
by the text) and evaluates it with `origin` = A1 and the caller's `host`, so
SPEC-120's `part + (host − origin)` rule gives Excel's relative names; for
an evaluation started from a name a result outside the grid wraps modulo
`MAX_ROWS`/`MAX_COLS` (Excel's rule) instead of becoming `#REF!`. An unknown
name returns `#NAME?`; a chain re-entering names 64 deep returns `#NAME?`.
`resolveStructured` returns the range of "Structured references" as a
`RangeValue`, or its error. Both call `ctx.recordDependency` for every area
they resolve to, so the graph learns them the way it learns INDIRECT's; a
change to `names` or to a table's `ref`/columns recalculates every cell whose
formula mentions that name or table (a token scan, cached per formula text),
including cells showing `#NAME?` for a name just defined.
`rebaseRelative(text, dRow, dCol)` performs the same wrap on text (ODS base
cells; Name Box definitions are absolute).

### Names on the wire
- **xlsx** — `<definedName name=… [localSheetId] [hidden="1"] [comment] …>wire</definedName>`.
  Write order: source order, new names appended. An out-of-range
  `localSheetId` keeps the name hidden with the attribute in `sourceAttrs`;
  it never resolves and is written back as read.
- **ods** — `table:named-range` (`table:cell-range-address`,
  `table:base-cell-address`, `table:range-usable-as`) and
  `table:named-expression` (`table:expression` = `of:=…`). A new named range's
  base cell is `$<scope or first sheet>.$A$1`.

### Reserved names

| Name | Owner | Notes |
| --- | --- | --- |
| `_xlnm.Print_Area`, `_xlnm.Print_Titles` | SPEC-136 | sheet-scoped; ODS print ranges are SPEC-136's mapping |
| `_xlnm._FilterDatabase`, `_xlnm.Criteria`, `_xlnm.Extract` | SPEC-128 | `_FilterDatabase` is `hidden="1"` in Excel files |
| any other `_xlnm.*` | — | kept and adjusted only |

Owners read and write them through `names.ts`'s `getReservedName(model,
sheetId, short)` / `setReservedName(model, sheetId, short, formula | null)`.

### The Name Box
Order of interpretation of a typed entry: an A1 reference or range
(optionally sheet-qualified) → navigate (SPEC-117); a visible name (not
hidden, not reserved, in the active sheet's or the workbook's scope) or a
table name → select its range (a table selects its data body; a name whose
formula is not a reference toasts "That name doesn't refer to a range");
otherwise a new name, validated. Typing replaces nothing until Enter;
Escape restores the box.

### Tables on the wire
- **Read.** `table@displayName` (fallback `@name`), `@ref`,
  `@headerRowCount` (`"0"` → no header; absent → 1), `@totalsRowCount`
  (`"1"` → totals), `tableStyleInfo@name`/`show*` (`"1"`/`"true"`), column
  children `calculatedColumnFormula` and `totalsRowFormula` through
  `fromXlsxFormula` + `structuredFromXlsx`.
- **Regenerate.** The root start tag is the source's with `id`, `name`,
  `displayName` (both written as the table name), `ref`, `headerRowCount`,
  `totalsRowCount`, `totalsRowShown` rewritten; `autoFilter` is its
  `sourceXml` with `ref` patched to header + data (omitted with no header
  row or no filter), or a new `<autoFilter ref="…"/>`; `sortState` and
  `extLst` verbatim; each `tableColumn` keeps its source attributes and
  unmodelled children by `xlsxId`. Header cells are always strings equal to
  their column names.
- **Duplicate sheet.** SPEC-121 already copies table parts at write time
  (`id` = max + 1, `<name>_<n>`, structured references in the copy renamed)
  and duplicates sheet-scoped `definedName`s. With tables and names modelled,
  the duplicate command does this on the model instead — the copy's tables
  get fresh `id`s and `xlsxId`s, `uniqueTableName` names, no `partName`,
  `dirty`; the copy's local names are cloned with the copy's scope; the
  copy's formulas are rewritten to the new table names — and the copy is
  marked `dirty` so its `<tableParts>` is rebuilt from `Sheet.tables`;
  SPEC-121's write-time copy of `table` parts and of `definedName`s is
  skipped for content the model holds.

### Structured references

| Form | Resolves to |
| --- | --- |
| `T`, `T[#Data]` | data body, all columns |
| `T[#All]` | `ref` |
| `T[#Headers]` / `T[#Totals]` | header / totals row (`#REF!` when absent) |
| `T[Col]`, `T[[Col]]` | Col's data cells |
| `T[[#Headers],[Col]]`, `T[[#Totals],[Col]]`, `T[[#All],[Col]]` | that part of Col |
| `T[[#Headers],[#Data],[Col]]`, `T[[#Data],[#Totals],[Col]]` | the adjacent parts of Col |
| `T[[A]:[C]]` (either order) | data cells of columns A..C |
| `T[@Col]`, `T[@]`, `T[@[A]:[C]]` | the host row's cells (host outside the data body → `#VALUE!`) |

A column name containing any of `,` `:` `.` `[` `]` `#` `'` `"` `{` `}` `$`
`^` `&` `*` `+` `=` `-` `>` `<` `/`, a tab or a line break (a plain space
does not count) prints in the inner-bracket form `T[[a.b]]`; `[`, `]`, `#`
and `'` are escaped with `'`. Rewrites (rename table, rename column, `#REF!` on
removal) are token-level edits of the structured token only.

### Canonical and wire forms

| Canonical (`Sales` at `A1:D10`, header row) | xlsx file | ods file (`expandStructuredReferences`, host C7) |
| --- | --- | --- |
| `Sales[@Amount]` | `Sales[[#This Row],[Amount]]` | `[.$C7]` |
| `Sales[@]` | `Sales[#This Row]` | `[.$A7:.$D7]` |
| `Sales[Amount]` | `Sales[Amount]` | `[.$C$2:.$C$10]` |
| `SUBTOTAL(109,Sales[Amount])` | same | `SUBTOTAL(109;[.$C$2:.$C$10])` |

SPEC-119's dialect leaves structured tokens verbatim (its out-of-scope
names this mapping as SPEC-133's), so `structuredToXlsx`/`structuredFromXlsx`
do the `@` ↔ `[#This Row]` rewrite; the end-to-end mapping is what the tests
assert.

### Creating a table
`detectHeaderRow(sheet, range)`: true when the range has at least two rows,
every first-row cell is a non-empty string, and either some column's
second-row cell is not a string or every first-row cell is bold while the
second row's are not. Headers off: the row above becomes the header row when
`range.top > 0` and it is empty across the range's columns and outside every
table; otherwise a SPEC-124 row insert at `range.top` makes room. A single
row with headers on gains one empty data row. Defaults: style
`TableStyleMedium2`, `showRowStripes: true`, the rest `false`, `autoFilter: {}`.

### Auto-expansion and calculated columns
- Expansion triggers only on SPEC-122's commit of a typed entry, never on
  paste, fill, undo or a formula result; cells of the new row or column that
  already hold content keep it.
- A calculated column's cells hold ordinary formulas; `calculatedFormula`
  is what fills new rows. Overwriting one cell leaves it an exception, and
  the column keeps its `calculatedFormula`.

### Header and totals rows
- **Totals on** uses the row below the table when it is empty across the
  table's columns and outside every table, else a SPEC-124 row insert below
  makes room. **Off** clears that row's cells, shrinks `ref`, and keeps each
  column's `totalsFunction`/`totalsLabel` for the next time it is turned on.
- **Header off** clears the header cells and moves `ref.top` down one (the
  names stay in `columns`); **on** writes them back into the row above, found
  or made exactly as in "Creating a table".
- SUBTOTAL codes: average 101, countNums 102, count 103, max 104, min 105,
  stdDev 107, sum 109, var 110; `custom` writes `totalsFormula`. SUBTOTAL
  ignores cells holding SUBTOTAL/AGGREGATE and, for 101–111, hidden rows.

### Convert to range
`convertTableToRange` bakes each cell's `tableStyleLook` into its own style
with SPEC-123's `internStyle` (the cell's explicit properties win, as on
screen), rewrites every structured reference to the table across the
workbook into A1 — this-row forms column-absolute and row-relative (`$C7`),
the rest absolute (`$C$2:$C$10`) — removes the table and its part, and is
one undo entry.

### Tables follow cells

| Edit (rows; columns symmetric) | Effect |
| --- | --- |
| insert at `≤ ref.top` | `ref` shifts |
| insert with `top < at ≤ bottom` | `ref` grows (new columns: `Column{n}`, headers written) |
| delete band covering the header / totals row | `headerRow` / `totalsRow` becomes false |
| delete band covering every data row, or every column | the table is deleted |
| other overlapping deletions | `ref` shrinks |

References to deleted tables and columns become `#REF!` in every formula,
name, calculated and totals formula. SPEC-124's merge refuses a range
touching a table ("You can't merge cells in a table"). Clearing every cell of
a table in one command deletes it.

### Tables and moves
The `tables` move adjuster moves a table whose whole `ref` lies inside the
moved block (onto the destination sheet, `dirty`, both sheets `dirty` so
their `<tableParts>` follow); references to it are by name and need nothing.
`moveRange`'s refusal pass (SPEC-125) refuses a move whose source cuts through
a table or whose destination overlaps a table not wholly inside the source:
"You can't move part of a table". A table overwritten by a move's
destination is deleted, its references becoming `#REF!`.

### Table styles
Within each family of seven, member 1 uses `dk1` and members 2–7 `accent1`–
`accent6` as the colour **C**; tints are theme references (`T80` = `lumMod
0.2, lumOff 0.8`; `T60` = 0.4/0.6; `T40` = 0.6/0.4), `W` = `lt1`, `K` = `dk1`.

| Names | Header | Body / stripes | Edges | Totals |
| --- | --- | --- | --- | --- |
| Light 1–7 | bold, bottom thin C | row/column stripes T80 | table top/bottom thin C | bold, top double C |
| Light 8–14 | fill C, bold W | no fill; stripes add top/bottom thin C | outline thin C | bold, top double C |
| Light 15–21 | bold, bottom thin C | stripes T80 | every cell thin C | bold, top double C |
| Medium 1–7, 15–21 | fill C (15–21: K), bold W | stripes T80; inner horizontals thin T40 | outline thin T40 | bold, top double C |
| Medium 8–14, 22–28 | fill C, bold W | body T80 (22–28: T80 + inner thin T40), stripes T60, inner lines W | — | fill C, bold W |
| Dark 1–7 | fill K, bold W | body C shaded 25%, text W, stripes C | — | fill K, bold W |
| Dark 8–11, unknown, workbook-defined | bold, bottom thin K | — | — | bold |

First/last column, when on, are bold. These are this product's renderings of
Excel's built-in looks, pinned so screenshots and tests are deterministic;
Excel and LibreOffice draw their own from the name.

### Tables in ODS
`table:database-range table:name="{name}" table:target-range-address="{ref}"
table:contains-header="{headerRow}" table:display-filter-buttons="{autoFilter}"`.
Cells inside take an automatic style merging `tableStyleLook` beneath the
cell's own style, generated per write by the ODS table writer.

### Reader upgrade
The retained `workbook.xml`, table parts and ODS named ranges that the
steps read have been kept current by SPEC-124's and SPEC-125's `retained`
rewrites, but sheet names in them are still source-time text.
`upgradeWorkbook` maps source sheets to model sheets by `partName` (xlsx) or
by each sheet's retained `sourceXml` `table:name` (ods), resolves
`localSheetId` through the **source** sheet order, applies SPEC-121's
`sourceSheetRemap(model, parts)` to every formula it reads (renamed sheets
renamed, deleted ones `#REF!`), and drops names scoped to sheets no longer
present; `upgradeDirtySheet` applies the same remap to the table formulas it
reads. `ReaderUpgrade` gains the optional member
`upgradeWorkbook?(model: SpreadsheetModel, sourceParts: Map<string, Uint8Array>): void`;
steps without it are unaffected.

### View wiring
`src/views/spreadsheet/useNames.ts` (Name Box, Name Manager) and
`useSheetTables.ts` (create, design bar, totals menu, the entry-commit hook
for expansion and calculated columns), each ≤ 400 lines;
`src/components/spreadsheet/NameManagerDialog.vue`, `NameEditorDialog.vue`,
`CreateTableDialog.vue`, `ResizeTableDialog.vue`, `TableDesignBar.vue`,
`TableStyleGallery.vue`; `SpreadsheetView.vue` gains two wiring calls.

### Tests
Unit files as listed in the ACs, each table above asserted row by row;
`equalBytes` (`src/lib/core/equality.ts`) for untouched `workbook.xml`, table
parts and `content.xml`; `equalStructural` (`src/lib/equality.ts`) for created
names and tables after save → reopen; a second write of a reopened model is
byte-identical. `e2e/SPEC-133.spec.ts` as in the last AC.

### Specs this changes
- SPEC-116: the model gains `names` and `Sheet.tables`; `ReaderUpgrade`
  gains `upgradeWorkbook?`.
- SPEC-117: the read-only `name-box` becomes an editable field (references
  navigate, names select, new names are defined).
- SPEC-119/120: the structured node's semantics and xlsx/ODS mapping are
  defined here; `EvalContext` gains `resolveName` and `resolveStructured`
  (and relative wrap for name evaluation).
- SPEC-121: `createBlankWorkbook` has `names: []`; rename, delete and
  duplicate reach names and tables; `<definedNames>` comes from the model
  and, like regenerated table parts, is excluded from `sourceSheetRemap`;
  the undo snapshot fields gain `tables`, `names`, `namesDirty`.
- SPEC-122: header-cell, calculated-column and expansion hooks on commit;
  `<tableParts>` and ODS named expressions/database ranges written from the
  model. SPEC-123: the table look sits beneath a cell's own style.
- SPEC-124: `names` and `tables` adjusters, their retained rows retired;
  merge refuses tables. SPEC-125: `names` and `tables` move adjusters; the
  move refusal pass gains the partial-table check.
- SPEC-128: its `_xlnm._FilterDatabase` reconciliation goes through
  `setReservedName` on the model instead of splicing `<definedNames>` (now
  rebuilt from `names`). SPEC-132: the view passes `resolveChartRange` a
  `resolveName(name, sheetId)` hook built on `resolveDefinedName` (the
  found name's `formula`, or `null`), so chart references resolve names. (SPEC-134 and SPEC-136, later, consume `resolveDefinedName`,
  `getReservedName`/`setReservedName` and tables.)
- SPEC-123: `SHEET_SHORTCUTS` gains `sheet.nameManager`; the menu bar gains
  the Formulas menu and Insert ▸ Table.
- `specs/000-roadmap.md` gains this spec's row.
