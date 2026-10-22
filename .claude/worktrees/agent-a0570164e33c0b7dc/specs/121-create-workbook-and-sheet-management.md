---
id: SPEC-121
title: Create a workbook, and sheet management
status: draft
area: editing
complexity: complex
story_points: 13
depends_on: [SPEC-116, SPEC-117, SPEC-119, SPEC-120]
---
## Why
After SPEC-116–120 the app opens a workbook, shows it with its styles,
formats and formulas, and recalculates it — but it cannot start one, and
the sheet tabs are a viewer. There is no "New spreadsheet", no way to add,
rename, copy, delete, reorder or hide a sheet, no undo, and the writers can
only re-emit the package that came in. This spec adds the first mutations
and, with them, the machinery every later spreadsheet editing spec rides
on: from-scratch packages for both formats, the substitution save path
that rebuilds only the workbook-level parts a change touched while
re-emitting everything else verbatim, the spreadsheet undo stack and the
edit pipeline behind it (diff → recalculate → mark dirty → persist), and
"Download a copy as…" the other format.

Sheet commands are where a workbook's references live or die. A renamed
sheet must take every formula, defined name, chart series and conditional
format that names it along; a deleted one must turn those into `#REF!`
rather than leave them pointing at nothing; a copied sheet's own drawings,
comments and tables must be copied with names that stay unique, or Excel
answers with its "We found a problem with some content" repair. The
presentation product walked the same arc in SPEC-080; this is its
spreadsheet twin, sized to sheets and to the reference web between them.

**This file is the normative source for the spreadsheet writers'
workbook-level rebuild rules (both formats, from scratch and by
substitution), the retained-XML sheet-reference pass, the spreadsheet undo
stack and edit pipeline, write-through persistence of edits, the sheet
commands, and the cross-format copy (`stripRetention`).** Serialising a
sheet's own cells is SPEC-122's; this spec writes empty sheets only and
states exactly what happens to a sheet whose cells changed before SPEC-122
lands.

## Acceptance criteria
- [ ] The landing page offers "New spreadsheet" (`new-spreadsheet`) beside "New presentation" — a button at the full tier of `src/views/library/LibraryHeaderActions.vue` and an item in its compact overflow menu (emitting a product-neutral `new-spreadsheet` intent), plus a button in `LibraryEmptyState`'s `extra-actions` slot; activating it stores `createBlankWorkbook()` via `storeSpreadsheet` and navigates to `/sheets/:id`. The three create actions move into `createBlankDocument(kind: "text" | "presentation" | "spreadsheet"): string` (stores the blank model, returns `editorPathFor(kind, id)`) in a new app-layer `src/lib/newDocument.ts`, so `HomeView.vue` gains only template hookups and stays inside `HomeView.size.spec.ts`'s budget (never raised); the existing "New document" and "New presentation" tests pass unchanged
- [ ] `createBlankWorkbook()` (in `spreadsheet/sheets.ts`) returns a model per "The blank workbook": one visible grid sheet "Sheet1", `format: "xlsx"` from birth, `dateSystem: "1900"`, `styles` = the single default style (Calibri via the theme's minor font, 11 pt, `dk1`), `theme` = `DEFAULT_THEME`, 64 px default columns and 15 pt default rows, title `DEFAULT_SPREADSHEET_TITLE` ("Untitled spreadsheet"), `opaque: []`, `readerVersion: SPREADSHEET_READER_VERSION`; the library badges it `XLSX` with no new `LibraryFormat`, and `documentKinds.ts`'s comment on the `"new"` format names blank workbooks beside blank presentations
- [ ] Downloading a workbook with no retained package writes a valid from-scratch package in its format per "From scratch": the `.xlsx` holds exactly the parts listed there (a minimal `xl/styles.xml` with the two reserved fills, the Office theme, one `xl/worksheets/sheetN.xml` per sheet), the `.ods` has `mimetype` first and STOREd; each re-opens through `readSpreadsheetBytes` to the same sheets, names, order, visibility, tab colours, active sheet, date system and default sizes (`workbookForCompare` deep equality — minted ids masked), and the golden parts in `xlsx/writer.spec.ts` / `ods/writer.spec.ts` match byte for byte with `modified` injected
- [ ] The sheet tab bar (`sheet-tabs`) gains management: an "Add sheet" button (`sheet-add`, "+") after the last tab; a per-tab context menu (`sheet-tab-menu`, right-click on a tab, or Shift+F10 / the ContextMenu key on a focused tab) with Insert sheet, Rename, Duplicate, Delete, Move left, Move right, Hide, Unhide…, Tab colour (items `sheet-tab-menu-<action>` per "UI"); every item is disabled with a `title` reason rather than hidden when it cannot apply. SPEC-117's `SpreadsheetMenuBar` gains an Edit menu (Undo `sheet-menu-undo`, Redo `sheet-menu-redo`, later specs append) and File ▸ "Download a copy as …" (`sheet-menu-download-copy`), routed through `useSpreadsheetMenu`
- [ ] Insert sheet / `sheet-add` inserts an empty grid sheet immediately after the active sheet, named `Sheet<N>` with N one more than the largest number in any existing `Sheet<number>` name (case-insensitive, hidden sheets included; 1 when none), and makes it active; the new sheet is born `dirty` with no `partName`/`sourceXml`
- [ ] Rename (double-click a tab, F2 on a focused tab, or the menu) edits the name in place in `sheet-tab-rename-input`; Enter commits, Escape cancels, blur commits a valid name and cancels an invalid one; an invalid name keeps the input open with the exact message from "Sheet names" in `sheet-name-error`; a valid rename rewrites every formula in every sheet that references the old name through SPEC-119's `renameSheetInFormula`; a case-only change of the sheet's own name is valid (uniqueness ignores the sheet being renamed), and committing the unchanged name records no undo step
- [ ] Duplicate inserts a copy immediately after its source named `<base> (<n>)` per "Sheet names", active, with the source's cells (rows shared by reference), columns, merges, frozen panes, tab colour and view state; qualified references to the source inside the copy's formulas are rewritten to the copy; the copy receives its own retained source at once, copy-on-write, per "Duplicating a sheet" (xlsx: part and rels closure copied into `opaque` with table/VML/threaded-comment fixups, sheet-scoped defined names duplicated on save; ods: table text with `xml:id`/`draw:id` and embedded objects uniqued), so a later edit of either sheet's retained text never touches the other's; Duplicate is disabled for foreign sheets ("This sheet type can't be copied yet") and for an xlsx sheet with a relationship type outside the safe list ("This sheet has content that can't be copied yet")
- [ ] Delete removes the sheet after a confirmation (`sheet-tab-delete-dialog`, "Delete "<name>"? Its cells and anything on it will be removed." with `sheet-tab-delete-confirm` / `sheet-tab-delete-cancel`) when it has content per "Deleting a sheet", and without one when it is empty; every reference to it in the remaining formulas becomes `#REF!` per "Sheet references in formulas", the workbook recalculates, the next visible sheet to the right (else left) becomes active, and deleting is disabled ("A workbook must keep at least one visible sheet") when it is the last visible sheet
- [ ] Move: dragging a tab reorders with a visible drop indicator (`sheet-tab-drop-indicator`), Escape or a drop outside the tab bar cancels, and Move left / Move right step past the adjacent visible tab; the moved sheet stays active, an `aria-live="polite"` region announces "Sheet "<name>" moved to position <k>" (k counted over visible tabs), and formula cells whose 3-D references span a moved, added, duplicated or deleted sheet are recalculated although their text did not change
- [ ] Hide removes a sheet's tab (disabled with "A workbook must keep at least one visible sheet" for the last visible one; the active sheet moves as on delete); Unhide… opens `sheet-unhide-dialog` listing every `"hidden"` sheet as a `sheet-unhide-option` in a multi-select listbox — `"veryHidden"` sheets are never listed — and unhides the chosen sheets, the last of them becoming active; the menu item is disabled when no sheet is hidden
- [ ] Tab colour opens the core `TextColorPicker` (theme row from `model.theme`, standard palette, recent colours via `core/colors.ts` `loadRecentColors`/`pushRecentColor`) plus a "No colour" item (`sheet-tab-color-none`); the choice sets or clears `Sheet.tabColor` (`ModelColor`, theme references kept as references), and a coloured tab renders a 3 px bottom bar when active and a tinted background otherwise
- [ ] Every command above is one undo step through `useSpreadsheetUndo` per "The undo stack": Cmd/Ctrl+Z, Shift+Cmd/Ctrl+Z and Cmd/Ctrl+Y (the registry's `edit.undo`/`edit.redo`, ignored while focus is in an input or dialog) and the app bar's `spreadsheet-undo`/`spreadsheet-redo` buttons (tooltips "Undo Rename sheet (⌘Z)"); snapshots are by reference — rows, `styles`, `opaque`, `odsContentEnvelope` and each `sourceXml` are held by identity (all copy-on-write), so taking one costs O(sheets), never O(cells) or O(package), asserted by a unit test that a snapshot of a 200,000-cell workbook allocates no `Row`, `Cell` or `OpaquePart`; the stack is per open view, capped at `MAX_UNDO_STEPS = 100`, never persisted; undo re-activates the sheet the step was recorded on and restores its selection
- [ ] Every command and every undo/redo runs the edit pipeline per "The edit pipeline": changed cells found by reference identity (SPEC-120's `changedCellsBetween`, built on `core/arrayWindow.ts` `windowOf`) become recalculation keys (sheet-list changes first call `invalidateDependencyGraph`), every sheet whose `rows` changed — by the command or in the pass's `changedSheetIds` — is replaced with `dirty: true`, and the model is written through (`touchDocument`) — the SPEC-116 row split then writes exactly the changed sheets' rows records, asserted in a store test (a rename of a sheet nobody references writes zero rows records)
- [ ] A workbook opened from a file and saved after zero commands — or after commands that were all undone — re-emits every part byte-identical (`equalBytes` per part, the SPEC-116 round-trip suite unchanged): writers rebuild nothing unless `structureDirty` is set or a sheet is `dirty`, and undo restores both flags with the state they describe
- [ ] Saving an opened `.xlsx` after sheet commands follows "The .xlsx substitution path": `xl/workbook.xml` gets a spliced `<sheets>` (source `sheetId`s kept, new ones above the maximum), renumbered or removed `definedName@localSheetId`, rewritten `workbookView@activeTab` and no `firstSheet`; `xl/_rels/workbook.xml.rels` and `[Content_Types].xml` are spliced, never re-serialised; `xl/calcChain.xml` is dropped with its relationship and override; `docProps/app.xml`'s sheet titles are rebuilt or removed; `tabSelected` agrees with the active sheet in every sheet part; every worksheet part the commands did not touch is `equalBytes` with its source
- [ ] The retained-XML sheet-reference pass per "Sheet references in formulas" rewrites renamed and deleted sheet names in every location in its table, in both formats (cell formulas, defined names, chart ranges, conditional-format and validation formulas, table formulas, internal hyperlinks; ODS named ranges/expressions, database ranges, validations, conditional formats, drawing anchors, chart ranges), simultaneously (a swap of two names round-trips), leaving every other byte of those parts verbatim
- [ ] Deleting or duplicating an xlsx sheet handles its part closure per "Deleting a sheet" / "Duplicating a sheet": on save, a deleted sheet's part, its `.rels` and every part reachable only through it are skipped (with their Content-Types overrides); at duplication, the copy's part and closure are added to `opaque` under the next free names with unique table ids and names, rebased VML shape ids and re-minted threaded-comment ids — the duplicated `.xlsx` opens in this app with every copied part present and no dangling relationship (asserted by a package-integrity helper over every fixture × command)
- [ ] Saving an opened `.ods` after sheet commands follows "The .ods substitution path": `content.xml` = envelope prefix (generated `de-ta` table styles appended to its automatic styles) + each sheet's table text with only its `table:name` / `table:style-name` attributes spliced + suffix (named ranges and database ranges on deleted sheets removed); `settings.xml`'s `ActiveTable` and per-table entries follow renames and deletions; embedded objects and pictures only a deleted table used are dropped with their manifest entries; every untouched table's text is splice-equal to its source
- [ ] Until SPEC-122 lands, a sheet made dirty by this spec (formula rewrites, recalculated values) is written as its retained source plus the sheet-reference pass, and the `.xlsx` writer then sets `calcPr@fullCalcOnLoad="1"` so the consumer recomputes the cached values this path cannot refresh — documented in `xlsx/writer.ts`'s header as the interim SPEC-122 replaces
- [ ] The app bar's Download becomes a split "Download ▾" (`spreadsheet-download` primary, `spreadsheet-download-menu` caret, `spreadsheet-download-copy` item reading "Download a copy as OpenDocument Spreadsheet (.ods)…" or "…as Excel workbook (.xlsx)…"): the copy converts through `stripRetention(model)` (which applies SPEC-119's `formulaForTarget` to every formula cell) and the target's from-scratch writer, never mutates the open model, and when the source holds structures from the `UNCONVERTIBLE_CONTENT` table it shows one toast "Not included in the .ods copy: charts, comments" listing their labels (returned in `spreadsheetDownload`'s `notices`); until SPEC-122 lands the item is disabled ("Not available for this workbook yet") for a workbook with any stored cell
- [ ] `upgradeSpreadsheetModel` (SPEC-116) is amended per "Reader upgrades after edits": a record with `structureDirty` or any dirty sheet is never re-read wholesale — clean sheets (copies included) are re-read individually from their own retained source, keeping every sheet's id, name, visibility, tab colour and position from the model — tested with a synthetic old-version record that has a renamed, a duplicated, a deleted and a new sheet
- [ ] Fixtures: the shared builder adds `e2e/fixtures/spec121-links.xlsx` and `spec121-links.ods` per "Fixtures", shaped like Excel and LibreOffice output respectively (cross-sheet and 3-D formulas, workbook- and sheet-scoped names, a chart, comments, a table, conditional formatting and a validation naming another sheet, `calcChain`, `TitlesOfParts`, a tab colour, a hidden and a very hidden sheet; ODS: named expressions, a database range, `settings.xml` per-table entries with `TabColor`, an embedded chart object)
- [ ] Tests: `spreadsheet/sheets.spec.ts` (every command, name rules and messages, naming, active-sheet results, flags, `createBlankWorkbook`, `stripRetention`, `unconvertibleContent`), `spreadsheet/sheetRefs.spec.ts` (the remap per location, simultaneity, 3-D endpoints, quoting, `#REF!`), `ods/rangeAddress.spec.ts` (both directions, quoted names, unions, untouched entries byte-identical), `xlsx/writer.spec.ts` and `ods/writer.spec.ts` (goldens, substitution rules, closures, integrity), additions to `roundTrip.spec.ts` and `open.spec.ts`, `views/spreadsheet/useSpreadsheetUndo.spec.ts` (by-reference snapshots, cap, flags, pipeline), `views/SpreadsheetView.spec.ts` (menu keyboard path, dialogs, undo wiring), the store row-split test, and the HomeView create test
- [ ] e2e (`e2e/SPEC-121.spec.ts`): New spreadsheet → `sheet-add` twice → tabs read Sheet1, Sheet2, Sheet3 → rename Sheet2 to "Budget" → duplicate it → "Budget (2)" → move it first via the context menu by keyboard → hide Sheet3 and unhide it through the dialog → colour a tab → delete Sheet1 → undo → download, re-open the file → the tab names and order are as left; then open `spec121-links.xlsx` → rename "Data" to "Facts" → the formula bar shows "=Facts!B2*2" for `Summary!C2` → delete "Lookup" after confirming → a cell that referenced it reads "#REF!" → download and re-open → both hold; download a copy of the blank workbook as `.ods` and re-open it → same tab names

## Out of scope
Typing into cells and serialising cells (SPEC-122), formatting (SPEC-123),
row/column structure (SPEC-124), clipboard (SPEC-125). "Move or Copy…" to
another workbook, copying sheets between workbooks or apps, and grouping
several selected sheets (Excel's multi-sheet selection) — backlog. Workbook
structure protection is honoured from SPEC-122; editing protection is
backlog. Renaming the workbook beyond the app bar's existing title edit;
editing `docProps` metadata (substitution saves keep `docProps/core.xml`
verbatim). Persisting the undo stack (a tab switch is view state, persisted
by SPEC-117's `storeSpreadsheetViewState`, never an undo step). Excel's
Shift+F11 accelerator is not bound here (spreadsheet shortcuts get their
table, `SHEET_SHORTCUTS`, in SPEC-123). Templates: a new workbook is always
blank (no gallery). Cross-format conversion of anything the model does not
yet hold — each later spec that models a structure removes its row from
`UNCONVERTIBLE_CONTENT`. Chart sheets cannot be duplicated (SPEC-132).

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — the model (`SpreadsheetModel`,
  `Sheet { id, name, visibility, kind, rows, partName?, sourceXml? }`,
  `Row`, `Cell`, `opaque: OpaquePart[]`, `odsContentEnvelope { prefix,
  suffix }`, `activeSheetIndex`, `dateSystem`, `odsNullDate?`,
  `packageVariant?`); its four invariants (copy-on-write rows, sparseness,
  the source is the splice base, dates are serials) and plain-JSON rule;
  `readSpreadsheetBytes`, `SPREADSHEET_READER_VERSION`, `ReaderUpgrade`,
  `READER_UPGRADES`, `upgradeSpreadsheetModel` in `spreadsheet/open.ts`;
  `spreadsheetDownload(model)` in `spreadsheet/download.ts`;
  `core/xmlScanner.ts` `scanXml` (exact UTF-16 offsets); the store's
  `storeSpreadsheet`/`getSpreadsheet`, the `sheetrows:<docId>:<sheetId>` row
  split written only when a sheet's `rows` reference changed, `buildPreview`;
  `SheetTabs.vue` (`sheet-tabs`, `sheet-tab`, `role="tablist"`/`"tab"`,
  Ctrl+PageUp/PageDown), `SpreadsheetAppBar.vue` (`spreadsheet-download`,
  `close-document`), `SpreadsheetView.size.spec.ts`; the format/extension
  table (`.xlsx`/`.xlsm`/`.xltx`/`.xltm`, `.ods`/`.ots`).
- `specs/117-…` — `SpreadsheetModel.styles: CellStyle[]` (index 0 = the
  default "Normal" style), `defaultStyleId: 0`, `theme?: DocumentTheme`;
  `Sheet.columns`, `defaultColWidthPx` (64 for Calibri 11, MDW 7),
  `defaultRowHeightPt` (15), `merges`, `frozen?`, `showGridLines?`,
  `zoomPercent?`, `tabColor?: ModelColor`, `selection?: SheetSelection`
  (view state, never dirties by itself); `CellStyle.sourceXf?` /
  `sourceStyleName?`; the xlsx colour mapping — `theme="n"` → slot with
  Excel's swap (0→lt1, 1→dk1, 2→lt2, 3→dk2, 4..9→accent1..6, 10→hlink,
  11→folHlink), `tint` t<0 → `lumMod: 1+t`, t>0 → `lumMod: 1−t, lumOff: t`,
  ARGB `FFrrggbb` → `#rrggbb`; ODS tab colour `tableooo:tab-color` on the
  table style. Sheet objects are never mutated in place (a change replaces
  the sheet object; `rows` keeps its identity when untouched); view state
  (`activeSheetIndex`, selection, zoom, gridlines) is persisted by
  `storeSpreadsheetViewState(model)` without touching `lastModified` and
  never dirties anything; `SpreadsheetMenuBar.vue` (items `sheet-menu-*`,
  routed by `src/views/spreadsheet/useSpreadsheetMenu.ts`).
- `specs/119-the-formula-language.md` — canonical formula text in
  `Cell.formula` (no leading "=", en-US syntax, sheet names quoted per its
  quoting predicate); `tokenize`/`joinTokens` (trivia-preserving,
  `joinTokens(tokenize(t)) === t`); `ReferenceNode { sheet?, sheetEnd?,
  start, end? }` (3-D when `sheetEnd`); `rewriteReferences(text, rewrite)`,
  `renameSheetInFormula(text, from, to)` and `shiftFormula` in
  `formula/references.ts`; `fromXlsxFormula`/`toXlsxFormula`/
  `fromOdsFormula`/`toOdsFormula` and `formulaForTarget(cell, target): {
  cell; dropped?: "unparsed" | "dataTable" | "dynamicArray" }` in
  `formula/dialect.ts`; `Cell.formulaUnparsed?: true` (wire text, never
  transformed), `arrayRange`, `dynamicArray`, `cm`, `dataTableXml`.
- `specs/120-calculation-engine-and-core-functions.md` — `CellKey`,
  `cellKey(sheetId, row, col)`, `invalidateDependencyGraph(model)` (sheet
  add/delete/rename/move call it) in `formula/dependencies.ts`;
  `changedCellsBetween(sheetId, before, after)`, `recalculate(model,
  changed, options?)` / `startRecalc` (time-sliced, driven by the view's
  `useRecalc.ts`) and `RecalcResult { valueChanged, changedSheetIds, … }` in
  `formula/recalc.ts`; recalculation replaces a row only when a cached value
  changes; an `"open"` pass marks nothing dirty; `model.forceFullCalcOnLoad`
  (set by a pass that produced a stale cell) and `model.calc?.fullCalcOnLoad`
  (read from the file).
- `specs/124-rows-columns-and-merged-cells.md` (later) — rewrites retained
  text at command time, copy-on-write (`opaque`, `odsContentEnvelope`,
  `sourceXml`), through `xlsx/` and `ods/retainedReferences.ts`; this spec's
  undo snapshot already holds those by reference.
- `specs/080-create-presentation-and-slide-management.md` — the precedent:
  three writer paths (from scratch / verbatim / substitution), `dirty` +
  `structureDirty`, `stripRetention`, the split Download button
  (`PresentationDownloadButton.vue`), `usePresentationUndo.ts`.
- Core: `core/zipWriter.ts` `zip(entries)` (`ZipEntryInput.store` for
  `mimetype`); `core/packageParts.ts` `decodedPartBytes`;
  `core/opcProperties.ts` `buildCorePropertiesXml(input, modified)`;
  `core/odfMeta.ts` `buildMetaXml(input, modified)`; `core/clock.ts` `now()`
  (the clock-coupled-test rule: no test hard-codes today); `core/xml.ts`
  `parseXml`, `parseRelationships` (drops `TargetMode` — so rels are edited
  by splicing, see below), `relsPathFor`, `resolvePartPath`,
  `escapeXmlText`/`escapeXmlAttribute`; `core/theme.ts` `DEFAULT_THEME`,
  `ModelColor`, `resolveColor`; `core/colors.ts`; `core/arrayWindow.ts`
  `windowOf(before, after)`; `src/lib/shortcuts.ts` (`findShortcut`,
  `isOutsideDocumentInput`, `edit.undo`/`edit.redo`); `TextColorPicker.vue`
  (`theme` prop, emits `select` with a hex or a `ThemeColorRef`);
  `BaseDialog.vue`; `UndoRedoControls.vue`; `useToasts`.

### Model additions (`spreadsheet/model.ts`)
```ts
export const DEFAULT_SPREADSHEET_TITLE = "Untitled spreadsheet";
export interface Sheet {
  // …SPEC-116/117 fields…
  /** The sheet's own content changed since open: SPEC-122's writer rebuilds
   *  it. Born true on a sheet with no source (new). Never set by a view
   *  change (selection, scroll, tab switch). */
  dirty?: true;
  /** Set by "Duplicate sheet": the origin's name in the SOURCE file (its
   *  `copiedFrom` when the origin is itself a copy). The copy's own retained
   *  text (its `partName` part in `opaque`, or its `sourceXml`) is a copy of
   *  the origin's source-time text, so the writer maps that name to the
   *  copy's when remapping it. A copy is never an "original". */
  copiedFrom?: string;
}
export interface SpreadsheetModel {
  // …
  /** The sheet list or a sheet's workbook-level properties (name,
   *  visibility, tab colour, order) changed since open. Absent on a freshly
   *  opened workbook; a from-scratch workbook is always fully built. */
  structureDirty?: true;
}
```
A sheet "has a source" when it has `partName` (xlsx) or `sourceXml` (ods);
a sheet with no source is always `dirty`. An **original** is a sheet with a
source and no `copiedFrom`.

### The blank workbook (`createBlankWorkbook()`, `spreadsheet/sheets.ts`)
`{ kind: "spreadsheet", id: crypto.randomUUID(), title:
DEFAULT_SPREADSHEET_TITLE, format: "xlsx", readerVersion:
SPREADSHEET_READER_VERSION, sheets: [{ id, name: "Sheet1", visibility:
"visible", kind: "grid", rows: [], columns: [], merges: [],
defaultColWidthPx: 64, defaultRowHeightPt: 15, dirty: true }],
activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{ font: {
family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } },
namedStyle: "Normal" }], defaultStyleId: 0, theme:
structuredClone(DEFAULT_THEME) }`. No `properties` (the from-scratch
`core.xml` carries the title and `modified` only; `created` is never
fabricated, per `buildCorePropertiesXml`).

### Sheet commands (`src/lib/spreadsheet/sheets.ts`, framework-free)
```ts
export interface SheetCommandResult {
  activeIndex: number;        // the sheet to make active
  /** Formula cells whose RESULT may change though their text did not: 3-D
   *  references spanning a sheet added/duplicated/moved/deleted, and cells
   *  naming a sheet that did not exist before (add/duplicate/rename). The
   *  pipeline adds them to the recalculation keys. */
  alsoRecalc: CellKey[];
}
export function createBlankWorkbook(): SpreadsheetModel;
export function addSheet(model: SpreadsheetModel, afterIndex: number): SheetCommandResult;
export function validateSheetName(model: SpreadsheetModel, name: string, ownIndex?: number): string | null;
export function renameSheet(model: SpreadsheetModel, index: number, name: string): SheetCommandResult; // throws on invalid
export function duplicateSheet(model: SpreadsheetModel, index: number): SheetCommandResult;
export function deleteSheet(model: SpreadsheetModel, index: number): SheetCommandResult;
export function moveSheet(model: SpreadsheetModel, from: number, to: number): SheetCommandResult;
export function setSheetHidden(model: SpreadsheetModel, indices: number[], hidden: boolean): SheetCommandResult;
export function setTabColor(model: SpreadsheetModel, index: number, color: ModelColor | undefined): void;
export function sheetHasContent(model: SpreadsheetModel, index: number): boolean;
export function cannotDuplicateReason(model: SpreadsheetModel, index: number): string | null;
export function stripRetention(model: SpreadsheetModel): SpreadsheetModel;
export function unconvertibleContent(model: SpreadsheetModel, target: SpreadsheetFormat): string[];
```
Each command keeps the 116 invariants and SPEC-117's: formula rewrites build
new `Cell`/`Row` objects and a new `rows` array for the sheets they touch; a
changed sheet is replaced by a new object (`{ ...sheet, name }`), never
mutated; `model.sheets`, `opaque` and `odsContentEnvelope` are replaced, not
spliced in place, when they change. Every command except `setTabColor` on
an unchanged colour sets `structureDirty`. Indices are into
`model.sheets` (hidden sheets included); the UI maps visible tab positions
onto them. The active sheet is always visible.

### Sheet names
Validation (`validateSheetName`; `null` = valid), first failing rule wins;
the message column is the exact text shown in `sheet-name-error`:
| Rule | Message |
| --- | --- |
| empty after trimming | `A sheet name can't be blank` |
| longer than 31 characters | `A sheet name can be at most 31 characters` |
| contains `[ ] : * ? / \` | `A sheet name can't contain any of these characters: [ ] : * ? / \` |
| starts or ends with `'` | `A sheet name can't begin or end with an apostrophe` |
| equals "History" (any case) | `"History" is reserved — choose another name` |
| equals another sheet's name, case-insensitively (`toUpperCase`), hidden sheets included, the sheet itself excluded | `There's already a sheet called "<name>"` |
- The rules apply to names this product creates; a name read from a file
  that breaks them (an ODS name of 40 characters) is kept as read.
- **Duplicate:** base = the source name with a trailing ` (<digits>)`
  removed; the copy is `<base> (<n>)` for the smallest n ≥ 2 that is free;
  if that exceeds 31 characters, base is truncated to fit.
- **Insert:** `Sheet<N>` per the AC (Excel's numbering, not the first gap).

### Sheet references in formulas (`src/lib/spreadsheet/formula/sheetRefs.ts`)
```ts
export interface SheetRemap {
  renamed: Record<string, string>;  // OLD name (matched case-insensitively) → new name
  deleted: string[];                // names now gone
  orderBefore: string[];            // sheet order the text was written against (3-D endpoints)
}
export function remapSheetReferences(canonical: string, remap: SheetRemap): string;
export function remapOdsAddresses(list: string, remap: SheetRemap): string; // "$Sheet1.A1:.B2 $'My S'.C3"
export function sourceSheetRemap(model: SpreadsheetModel, parts: Map<string, Uint8Array>): SheetRemap;
export function resolveRetainedSheetName(model: SpreadsheetModel, parts: Map<string, Uint8Array>,
  name: string, hostSheetId?: string): Sheet | undefined;
```
Token-level over SPEC-119's `tokenize`, like `rewriteReferences` (which
cannot express a prefix-less `#REF!`, so this is its sibling, not a
caller); whitespace and untouched tokens stay byte-identical; every
reference is looked up once, so a swap of two names
is one simultaneous pass — for a single rename the result equals
`renameSheetInFormula(text, from, to)` (property-tested). A reference to a
deleted sheet becomes the error literal `#REF!` as a whole (SPEC-119's
out-of-grid convention; Excel would show `#REF!A1`). A 3-D reference
`Sheet1:Sheet3!A1` whose endpoint is deleted moves that endpoint inward along
`orderBefore`; a span collapsing to one sheet becomes a plain reference;
interior deletions leave the text unchanged. Names are re-quoted as
SPEC-119 prints them. In ODS address lists a deleted sheet becomes
`$#REF!.` (LibreOffice's form). Formulas with `formulaUnparsed`, or text the
tokenizer rejects, are left untouched.

**ODS range addresses** (`src/lib/spreadsheet/ods/rangeAddress.ts`). This
spec creates the module, because `remapOdsAddresses` is the first code in the
queue that parses ODS address lists; SPEC-124's retained-reference rewrite
and SPEC-128–136 reuse it unchanged:
```ts
/** "Sheet1.$A$1:.$B$3", "$'Q1 Data'.A1:'Q1 Data'.B3", "Sheet1.A1:Sheet1.B3 Sheet1.D1:Sheet1.D3"
 *  → "Sheet1!$A$1:$B$3", "'Q1 Data'!A1:B3", "(Sheet1!A1:B3,Sheet1!D1:D3)"; null if unparseable. */
export function odsRangeAddressToCanonical(text: string): string | null;
/** The inverse; writes "$Sheet1.$A$1:.$B$3" (sheet on the start only, `$` as canonical has it),
 *  a union as a space-separated list. */
export function canonicalToOdsRangeAddress(canonical: string): string | null;
```
`remapOdsAddresses` splits a list on top-level spaces and rewrites, through
this pair, only the entries whose sheet the remap renames or deletes; every
other entry keeps its bytes.

**Model side** (the commands): rename/delete rewrite `Cell.formula` in every
sheet; duplicate rewrites qualified references to the source inside the
copy only. **Retained side** (the writers): sheet names inside retained text
(`opaque`, `odsContentEnvelope`, `sourceXml`) stay **source-time** text —
SPEC-124's command-time rewrites adjust rows and columns there but never
rename — so each writer derives one remap from the source workbook part and
the model: `sourceSheetRemap(model, parts)`: an original whose source name
(xlsx: the source `<sheet name>` for its `partName`; ods: its `sourceXml`'s
`table:name`) differs from its model name is renamed; a source sheet no
original holds any more is deleted. A copy's retained text is remapped with
the same map plus `<copiedFrom> → <copy name>`. Any later code that
interprets a sheet name found in retained text (SPEC-124's
`RetainedReferenceMapper`, a reader-upgrade step) resolves it with
`resolveRetainedSheetName` (the same mapping, `hostSheetId` selecting a
copy's override). The pass rewrites, wherever the part is emitted, exactly
these locations (offset splices located with `scanXml`; values re-escaped;
anything not listed stays verbatim):
| Format | Part | Location |
| --- | --- | --- |
| xlsx | worksheets | `c/f` text; `cfRule/formula`; `dataValidation/formula1`, `formula2`; `x14:*/xm:f`; `hyperlink@location` |
| xlsx | `xl/workbook.xml` | `definedName` text |
| xlsx | `xl/charts/*.xml`, chartex | `c:f`, `cx:f` text |
| xlsx | `xl/tables/*.xml` | `calculatedColumnFormula`, `totalsRowFormula` |
| xlsx | `xl/pivotCache/pivotCacheDefinition*.xml` | `worksheetSource@sheet` (renames only) |
| ods | `content.xml` | `table:formula`; `table:named-range@table:cell-range-address`, `@table:base-cell-address`; `table:named-expression@table:expression`, `@table:base-cell-address`; `table:database-range@table:target-range-address`; `table:content-validation@table:condition`, `@table:base-cell-address`; `calcext:condition@calcext:value`, `@calcext:base-cell-address`; `calcext:conditional-format@calcext:target-range-address`; `@table:end-cell-address`; `draw:object@draw:notify-on-update-of-ranges`; `text:a@xlink:href` starting `#` |
| ods | `Object N/content.xml` | `@chart:values-cell-range-address`, `@table:cell-range-address`, `@chart:label-cell-address` |
xlsx formula text goes through `fromXlsxFormula` → remap → `toXlsxFormula`,
ODS through the ODS pair (emitted in `of:` form); a location whose remapped
canonical text equals its original is not touched. An empty remap skips the
pass, and a part containing no case-insensitive occurrence of any remapped
name is emitted without being scanned — so a save after adds, moves or
colour changes alone never decodes a clean worksheet part.

### Deleting a sheet
`sheetHasContent`: `rows.length > 0`, or `kind === "foreign"`, or (xlsx) its
part has a `.rels`, or (ods) its `sourceXml` contains `<table:shapes` or
`<office:annotation`. Model: remap formulas with `deleted: [name]`, splice
the sheet out. xlsx save: remove its workbook relationship; then every part
reachable from `_rels/.rels` through the deleted sheet's part and through no
surviving relationship is skipped with its `.rels` and its
`[Content_Types].xml` `Override` (a deleted copy's own parts included; parts
unreachable even in the source stay, as they came); `definedName`s with its
`localSheetId` are removed and the rest renumbered. ods save: its table text
is gone; `Object N/` directories, `ObjectReplacements/Object N` and
`Pictures/` entries referenced only by deleted tables are skipped with their
manifest entries; `table:named-range` / `table:database-range` elements whose
range lies on it are removed.

### Duplicating a sheet
Model: `{ ...source, id: new, name, rows: source.rows /* shared */, dirty:
source.dirty, copiedFrom: source.copiedFrom ?? <source's source name> }`
(no `copiedFrom`, and no source, when the origin has none) with small
fields cloned and self-references rewritten (which makes it dirty). The
copy gets its own retained source **at command time**, so SPEC-124's later
command-time rewrites of one sheet's retained text never reach the other:
xlsx — a new `opaque` array holding the origin's part
bytes under the next free `xl/worksheets/sheet<N>.xml` (the copy's
`partName`) and its rels closure — each part reachable only from the
origin's part copied under its folder's next free number (`drawing3.xml` →
`drawing<max+1>.xml`; no trailing digits → `_2`) with the copied `.rels`
pointing at the copies; parts also reachable from elsewhere (images, pivot
caches, embedded packages, persons) are shared by reference; the Overrides,
the workbook relationship and the `sheetId` are added on save. Safe
relationship types: worksheet-level `drawing`
(with charts, chart styles/colours, images), `comments`, `vmlDrawing`,
`threadedComment`, `table`, `pivotTable`, `printerSettings`, external
`hyperlink`. Fix-ups: a copied table gets `id` = max table id + 1 and
`name`/`displayName` `<name>_<n>` (smallest free n ≥ 2, unique among tables
and defined names) with structured references to it in the copy renamed; a
copied VML part gets `o:idmap data` = max over all VML parts + 1 and its
`_x0000_s<id>` shape ids rebased by 1024 per block; threaded-comment ids and
the matching `xr:uid`s are re-minted deterministically (a v4-shaped UUID
from FNV-1a of copy sheet id + original id) so writes are reproducible; a
copied pivot table's `name` is uniqued (all fix-ups at command time); on
save, sheet-scoped `definedName`s of the source are duplicated with the
copy's `localSheetId`. ods, at command time: the copy's `sourceXml` is the
origin's with every `xml:id`/`draw:id` value suffixed `-c<k>` consistently
(references to them included) and each `draw:object` copied in `opaque` to
the next free `Object N` (manifest spliced, copy-on-write); on save its
`table:name` is spliced and its `settings.xml` per-table entry is copied
under the new name.

### The undo stack (`src/views/spreadsheet/useSpreadsheetUndo.ts`)
```ts
export const MAX_UNDO_STEPS = 100;
export type UndoScope = "workbook" | { sheetId: string };
export const SHEET_BY_REFERENCE_FIELDS: readonly (keyof Sheet)[] = ["rows", "sourceXml"];
export const WORKBOOK_SNAPSHOT_FIELDS: readonly (keyof SpreadsheetModel)[] =
  ["activeSheetIndex", "structureDirty", "styles", "opaque", "odsContentEnvelope"];
export interface SpreadsheetUndoDeps {
  model: ShallowRef<SpreadsheetModel | null>;
  /** SPEC-120's useRecalc with reason "edit": runs (possibly sliced) and calls
   *  back with the pass's changedSheetIds when it completes. */
  recalc: (changed: CellKey[], done: (changedSheetIds: string[]) => void) => void;
  onApplied: () => void;                    // triggerRef + touchDocument(model.id)
  activateSheet: (sheetId: string) => void;
}
export interface SpreadsheetUndo {
  record(label: string, scope: UndoScope): void;  // BEFORE mutating
  commit(alsoRecalc?: CellKey[]): void;           // AFTER mutating: the edit pipeline
  undo(): void; redo(): void; reset(): void;
  canUndo: Ref<boolean>; canRedo: Ref<boolean>; undoLabel: Ref<string>; redoLabel: Ref<string>;
}
```
A snapshot holds, for **every** sheet, a shallow copy `{ ...sheet }` in which
each object/array field is `structuredClone`d except those in
`SHEET_BY_REFERENCE_FIELDS` (kept by reference — copy-on-write makes that
safe), plus every `WORKBOOK_SNAPSHOT_FIELDS` value **by reference**: each is
a primitive or copy-on-write — `styles` (SPEC-122's `internStyle` replaces
the array when it appends, so an undo past a new style restores the shorter
array), `opaque` (a changed part is a new `OpaquePart` in a new array —
this spec's duplicate, SPEC-124's retained rewrites), `odsContentEnvelope`
(a new object of immutable strings). Cost: O(sheets + small metadata). A
later spec that adds a sheet field able to grow with the cell count makes
it copy-on-write and appends it to `SHEET_BY_REFERENCE_FIELDS`; one that
adds workbook-level state makes it copy-on-write and appends it to
`WORKBOOK_SNAPSHOT_FIELDS`. `forceFullCalcOnLoad` is deliberately absent
(SPEC-120 never clears it). Undo/redo
push the live state as the opposite snapshot, install fresh shallow copies
of the stored one (snapshots stay immutable), run the pipeline, and
activate `scope.sheetId` when it still exists. `commit()` on a command that
changed nothing pops its step. Labels for this spec: "Add sheet", "Rename
sheet", "Duplicate sheet", "Delete sheet", "Move sheet", "Hide sheet",
"Unhide sheet", "Tab colour". Shortcuts are ignored while focus is in an
input, a dialog, or (SPEC-122) the cell editor.

### The edit pipeline (`commit`)
1. For each sheet id present before and after whose `rows` differ:
   SPEC-120's `changedCellsBetween(sheetId, before.rows, after.rows)` (an
   identity diff over `windowOf`) → the keys of cells added, removed or
   replaced. A sheet present on one side only contributes the keys of all
   its stored cells. When the sheet list, a name or the order changed,
   `invalidateDependencyGraph(model)` runs first (SPEC-120's rule for sheet
   commands); otherwise SPEC-120's `updateGraphCells` follows the keys.
2. `deps.recalc(keys ∪ alsoRecalc, done)`; a pass still running when the
   next command, undo or redo starts is completed first, so steps never
   interleave.
3. Every sheet whose `rows` reference differs from the snapshot's — by the
   command, or listed in the pass's `changedSheetIds` when `done` fires — is
   replaced with `dirty: true` (commands set `dirty` themselves for
   non-row content).
4. `deps.onApplied()` — re-render and write through: `touchDocument(id)`
   bumps recency and persists the shell plus the rows records whose
   reference changed (SPEC-116's split; `buildPreview` follows). It runs
   after the command and again when a sliced pass completes. (SPEC-138
   later swaps `touchDocument` for its draft cadence here.)

### From scratch
`writeXlsx(model, options?: { modified?: Date; fullCalcOnLoad?: boolean })`
and `writeOds(model, options?: { modified?: Date })` (`modified` defaults to
core `now()`; `fullCalcOnLoad="1"` is written when the option,
`model.forceFullCalcOnLoad` (SPEC-120) or — from scratch —
`model.calc?.fullCalcOnLoad` asks, or the interim rule below) — paths as SPEC-080: no `opaque` → from scratch; `opaque`,
no `structureDirty`, no dirty sheet → every part verbatim in source order;
otherwise substitution. **.xlsx** parts, in this order:
`[Content_Types].xml` (Defaults `rels`, `xml`; Overrides for workbook,
each worksheet, theme, styles, `sharedStrings` when SPEC-122 emits one,
core, app); `_rels/.rels` (officeDocument, core-properties,
extended-properties); `docProps/core.xml` (`buildCorePropertiesXml({ title },
modified)`); `docProps/app.xml` (`<Application>docu-edit</Application>` only —
no `TitlesOfParts`, so nothing can go stale); `xl/workbook.xml`
(`<workbookPr/>` or `date1904="1"`, `<bookViews><workbookView
activeTab="…"/></bookViews>`, `<sheets>` with `sheetId` 1..n, `state` for
hidden sheets, `<calcPr calcId="191029"/>`); `xl/_rels/workbook.xml.rels`
(sheets `rId1..n`, then theme, styles, sharedStrings); `xl/theme/theme1.xml`
= `OFFICE_THEME_XML`, which this spec **moves** from
`presentation/pptx/writer.ts`'s private `THEME_XML` to a new
`src/lib/core/officeThemeXml.ts` (the pptx writer imports it; its bytes are
unchanged, proven by the pptx goldens); `xl/styles.xml` (`<fonts>` one font
— `sz 11`, `color theme="1"`, `name Calibri`, `family 2`, `scheme minor`;
`<fills>` `none` + `gray125`; one empty `<border>`; one `cellStyleXfs` xf; one
`cellXfs` xf `xfId="0"`; `cellStyle name="Normal" xfId="0" builtinId="0"`;
`dxfs count="0"`; `tableStyles count="0"`) — SPEC-122's style writer later
produces exactly this for a one-style table; `xl/worksheets/sheetN.xml` for
an empty sheet: `<worksheet>` with the main and `r` namespaces,
`<sheetPr><tabColor …/></sheetPr>` when coloured, `<dimension ref="A1"/>`,
`<sheetViews><sheetView workbookViewId="0"` + `tabSelected="1"` on the
active sheet `/></sheetViews>`, `<sheetFormatPr defaultRowHeight="15"/>`,
`<sheetData/>`, `<pageMargins left="0.7" right="0.7" top="0.75"
bottom="0.75" header="0.3" footer="0.3"/>`. **.ods**: `mimetype` (STOREd,
first), `META-INF/manifest.xml` (version 1.2), `meta.xml`
(`buildMetaXml`), `styles.xml` (a `table-cell` default style and `Default`
cell style in Calibri 11 pt, page layout `pm1`, master page `Default`),
`content.xml` (`office:version="1.2"`; `table:calculation-settings` with
`table:null-date` only for 1904 or `odsNullDate`; automatic styles from the
`de-` registry; per sheet `<table:table table:name table:style-name=
"de-ta…"><table:table-column table:style-name="de-co…"
table:default-cell-style-name="Default"/><table:table-row table:style-name=
"de-ro…"><table:table-cell/></table:table-row></table:table>`) and
`settings.xml` holding only `ooo:view-settings` → `Views` → `ActiveTable`.

**Generated ODS styles** (`ods/styles.ts`, `OdsStyleRegistry`, extended by
SPEC-122): automatic styles named `de-ta<n>` (family `table`:
`style:master-page-name` = the source table style's or `Default`,
`table:display`, `style:writing-mode="lr-tb"`, `tableooo:tab-color` for a
tab colour resolved to hex), `de-co<n>` (`style:column-width` in inches from
px at 96 dpi, 4 decimals) and `de-ro<n>` (`style:row-height` in pt,
`style:use-optimal-row-height`), deduplicated by properties, spliced before
`</office:automatic-styles>` (the element is created before `<office:body`
if absent); the counter starts past every `de-` name already in
`content.xml` or `styles.xml`, so no generated name ever collides.

### The .xlsx substitution path
- `xl/workbook.xml` (offset splices only): `<sheets>` replaced — per model
  sheet `<sheet name sheetId [state] r:id/>` using the root's prefix for the
  relationships namespace; an original keeps its source `sheetId` and rId,
  copies (`copiedFrom`) and new sheets get `sheetId` = max (source,
  assigned) + 1 and the next free `rId`; `definedName@localSheetId` renumbered by the moved
  sheet's new index, removed for deleted sheets, duplicated for copies; the
  pass over `definedName` text; every `workbookView` gets `activeTab` = the
  active index and loses `firstSheet`; `calcPr@fullCalcOnLoad="1"` when
  the rule under "From scratch" asks (an absent `calcPr` is created before the
  first of `oleSize`, `customWorkbookViews`, `pivotCaches`, `smartTagPr`,
  `smartTagTypes`, `webPublishing`, `fileRecoveryPr`, `webPublishObjects`,
  `extLst` — CT_Workbook order — else before `</workbook>`). Elements are
  located with `scanXml`.
- `xl/_rels/workbook.xml.rels`: surviving `<Relationship>` elements
  verbatim (so `TargetMode` survives), deleted sheets' and `calcChain`'s
  removed, new ones appended. `[Content_Types].xml`: Overrides spliced for
  added/skipped parts and `calcChain` removed; everything else verbatim.
- Sheet parts: each sheet is emitted at its `partName` (copies have one
  from duplication), a new sheet at the next free
  `xl/worksheets/sheet<N>.xml` in model order. A clean sheet is its retained
  bytes with the pass applied; `<sheetPr>`'s `<tabColor>` is
  spliced (inserted/replaced/removed, `<sheetPr>` created as the first child
  if absent) when the model colour differs from the source's — theme colours
  as `theme="<swap index>"` plus `tint` when the reference is exactly a
  `lumMod` or a `lumMod`+`lumOff` summing to 1, else `rgb` of
  `resolveColor` (SPEC-117's `modelColorToXlsxAttrs(color, theme)` in
  `xlsx/themeColor.ts`, which SPEC-122's style writer calls too); the first
  `<sheetView>`'s `tabSelected` is spliced to
  agree with the active sheet (two selected tabs put Excel into group
  editing).
- `xl/calcChain.xml` is dropped (part, relationship, Override) on any
  structural change or dirty sheet — a stale chain is a known Excel repair
  trigger.
- `docProps/app.xml`: the `HeadingPairs` group whose `TitlesOfParts`
  entries equal the source's grid sheet names in order is rewritten with the
  model's (count updated), likewise chart sheets; if either cannot be
  identified both elements are removed (they are optional, and a deleted
  sheet's name must not linger). `docProps/core.xml` stays verbatim.

### The .ods substitution path
`content.xml` = prefix + tables + suffix. Prefix: generated `de-` styles
appended; the pass over `table:content-validations`. Each table: a clean
original's or copy's `sourceXml` with the pass applied and two start-tag
attribute splices — `table:name` when it differs from the model name, and
`table:style-name` → a `de-ta` style when the model's visibility or tab
colour differs from what the source table style declares (read with the
same helper the reader uses); a new sheet's generated table. Suffix: the
pass, plus the removals in "Deleting a sheet". `settings.xml` (when
present): `ActiveTable` = the active sheet's name; in the `Tables`
`config-item-map-named`, entries are renamed, removed for deleted sheets,
copied for duplicates, and a recoloured table's `TabColor` item is removed
(the style attribute is authoritative). `META-INF/manifest.xml` is spliced
only when entries were added or removed; every other part is verbatim.

### Interim: dirty sheets before SPEC-122
Until SPEC-122 replaces it, a dirty sheet with a source is written exactly
as a clean one (source + pass + splices); its cell formulas are correct
because the pass rewrites `<f>`/`table:formula`, but cached values the
pipeline recalculated are not refreshed, so `writeXlsx` sets
`fullCalcOnLoad="1"` whenever it wrote a dirty sheet this way. LibreOffice
does not recalculate ODS on load by default; a stale cached value there is
the documented cost of the interim. New sheets are always empty here.

### Cross-format copy (`stripRetention`, `unconvertibleContent`)
`spreadsheetDownload(model, targetFormat = model.format): Promise<{ bytes;
fileName; mimeType; notices: string[] }>` — `notices` are whole one-line
messages the view shows as toasts (SPEC-126 adds its CSV notices to the
same list). A cross-format copy writes `stripRetention({ ...model, format:
target })`: no `opaque`, `odsContentEnvelope`, `structureDirty`,
`packageVariant`; per sheet no `partName`/`sourceXml`/`copiedFrom`/`dirty`;
foreign sheets removed (active index re-pointed to a visible sheet); styles
without `sourceXf`/`sourceStyleName`; every formula cell through SPEC-119's
`formulaForTarget(cell, target)` (an `"unparsed"` or `"dataTable"` drop
keeps the cached value and raises the "untranslatable formulas" label); rows
shared by reference except where a cell must change. Pure — the open model
is untouched. `UNCONVERTIBLE_CONTENT: { label: string; xlsx?: Marker; ods?:
Marker }[]` (exported from `sheets.ts`) is the table below;
`unconvertibleContent` returns, in table order, the labels whose marker is
present in the source. Later specs delete their row when their structure
crosses through the model, and may add rows whose marker yields item-level
labels (SPEC-131's "Form control on Sheet1"):
| Label | xlsx marker | ods marker | Row removed by |
| --- | --- | --- | --- |
| chart sheets | a `kind: "foreign"` chartsheet | — | SPEC-132 |
| macros | `packageVariant` macro-enabled | `Basic/` or `Scripts/` entries | — |
| images and shapes | a worksheet `drawing` relationship | `<table:shapes` / `draw:frame` | SPEC-131 |
| charts | a chart part | a `draw:object` chart | SPEC-132 |
| comments | a `comments` relationship | `<office:annotation` | SPEC-135 |
| conditional formatting | `<conditionalFormatting` | `<calcext:conditional-formats` | SPEC-129 |
| data validation | `<dataValidations` | `<table:content-validations` | SPEC-130 |
| named ranges | `<definedName` (non-`_xlnm`) | `<table:named-expressions` | SPEC-133 |
| tables | a table part | `<table:database-ranges` | SPEC-133 |
| filters | `<autoFilter` | `table:display-filter-buttons` | SPEC-128 |
| hyperlinks | `<hyperlinks` | `<text:a` in a table | SPEC-135 |
| pivot tables | a pivotTable part | `<table:data-pilot-tables` | SPEC-139 |
| page setup | `<pageSetup` / `<headerFooter` | a non-default page layout | SPEC-136 |
| untranslatable formulas | a `formulaForTarget` drop (`formulaUnparsed`, `dataTableXml`) | same | — |
| custom date origin | — | `odsNullDate` (to xlsx) | — |
The toast reads "Not included in the .ods copy: " (or ".xlsx") + the labels
joined with ", ". File names follow SPEC-116's table for the target format
with no variant.

### Reader upgrades after edits
`upgradeSpreadsheetModel` on an old record with `structureDirty` or a dirty
sheet: the model's sheet list, order, names, visibility, tab colours and ids
are kept; each clean sheet (copies included) is re-read from its own
`partName`/`sourceXml`, replacing only its cell content and per-sheet
properties, with `sourceSheetRemap` applied to the formulas it reads
(retained text keeps source-time sheet names); dirty sheets run the pending `ReaderUpgrade` steps (a
step on a sheet with no source does nothing). This spec adds no reader
output and does not bump `SPREADSHEET_READER_VERSION`.

### UI
`SheetTabs.vue` grows the `sheet-add` button, roving tab focus (ArrowLeft/
Right between tabs, Enter/Space activates, F2 renames), drag-reorder with
pointer events and the drop indicator, and the live region;
`SheetTabMenu.vue` (core `menuBar` types + the shared menu components; item
testids `sheet-tab-menu-insert`, `-rename`, `-duplicate`, `-delete`,
`-move-left`, `-move-right`, `-hide`, `-unhide`, `-tab-color` — distinct
from SPEC-117's menu-bar `sheet-menu-*` ids and SPEC-124's
`sheet-delete-dialog`), `SheetTabDeleteDialog.vue` and
`SheetUnhideDialog.vue` on `BaseDialog`,
`SpreadsheetDownloadButton.vue` (the split button), all in
`src/components/spreadsheet/`. `UndoRedoControls.vue` gains optional
`undoTestid`/`redoTestid` props (defaults `undo-button`/`redo-button`,
unchanged for the text editor) and renders in `SpreadsheetAppBar` as
`spreadsheet-undo`/`spreadsheet-redo`. Logic lives in
`src/views/spreadsheet/useSpreadsheetUndo.ts`, `useSheetCommands.ts` (menu →
command → `record`/`commit`, dialogs, toasts) and `useWorkbookDownload.ts`;
the Edit menu and File ▸ "Download a copy as …" entries route through
SPEC-117's `useSpreadsheetMenu.ts`; each composable stays within `composableSize.spec.ts`'s 400 lines; `SpreadsheetView.vue`
gains only `use*` wiring and template hookups and ratchets its budget.

### Fixtures
`spreadsheet/fixtures/build.ts` adds `spec121-links.xlsx` (Excel-shaped:
sheets "Summary", "Data", "Lookup", "Archive" (hidden), "Vault"
(`veryHidden`); `Summary!C2` = `Data!B2*2`, `SUM(Data:Lookup!A1)`, a
`VLOOKUP` into `Lookup`; `definedName`s workbook-scoped and
`localSheetId`-scoped incl. `_xlnm.Print_Area`; on "Data" a drawing with a
chart whose `c:f` names `Data`, legacy comments + VML, a threaded comment, a
table part, conditional formatting and a list validation naming `Lookup`, an
internal hyperlink; a tab colour with `theme`+`tint`; `calcChain.xml`;
`docProps/app.xml` with `HeadingPairs`/`TitlesOfParts`; `mc:Ignorable`,
`xr:uid`) and `spec121-links.ods` (LibreOffice-shaped twin: `of:=` formulas
with `[$Data.B2]`, named range and named expression, a database range, a
content validation, `calcext:conditional-formats`, an embedded chart object
with `Object 1/content.xml` ranges, `settings.xml` with `ActiveTable` and
per-table entries carrying `TabColor`, `tableooo:tab-color`, an `xml:id`).

### Tests
- `sheets.spec.ts`, `formula/sheetRefs.spec.ts`, `xlsx/writer.spec.ts`,
  `ods/writer.spec.ts` per the AC; a `packageIntegrity(bytes)` helper
  (every relationship target exists, every part has a content type, no
  Override for a missing part, unique table ids/names, unique `sheetId`s)
  runs over every fixture × every command.
- `roundTrip.spec.ts`: from-scratch write → read (`workbookForCompare`, ids
  masked — compare resolved facts, not minted values); no-op and
  undo-to-origin saves `equalBytes` per part; untouched worksheet parts
  `equalBytes` after structural commands; ods splice equality of untouched
  tables.
- `open.spec.ts`: the amended upgrade path. `useSpreadsheetUndo.spec.ts`:
  by-reference snapshots (no `Row`/`Cell` allocated), cap, labels, flags,
  pipeline keys and dirty marking. The store row-split tests. View and
  HomeView tests. e2e per the AC, downloading with the presentation suites'
  download-capture pattern.

### Specs this changes
- SPEC-116: `spreadsheetDownload` gains `targetFormat` and `notices` and
  returns a `Promise`;
  `upgradeSpreadsheetModel` follows "Reader upgrades after edits"; the
  tabs become editable.
- SPEC-117: `SpreadsheetMenuBar` gains the Edit menu (Undo/Redo) and File ▸
  "Download a copy as …"; the tab-colour splice is the first caller of
  `modelColorToXlsxAttrs`.
- SPEC-124: its `RetainedReferenceMapper` resolves sheet names found in
  retained text through `resolveRetainedSheetName` (they are source-time),
  and its copy-on-write retained rewrites are what this spec's undo holds by
  reference.
- SPEC-080/SPEC-079: the Office theme XML moves to
  `core/officeThemeXml.ts` (two-phase move; pptx bytes unchanged).
- SPEC-068's `LibraryHeaderActions` gains the `new-spreadsheet` intent; the
  text editor's `UndoRedoControls` testids are unchanged.
- SPEC-122 replaces "Interim: dirty sheets before SPEC-122", fills the
  from-scratch sheet bodies and lifts the copy's empty-workbook condition.
