---
id: SPEC-134
title: Find and replace, and Go To
status: draft
area: editing
complexity: normal
story_points: 5
depends_on: [SPEC-017, SPEC-122, SPEC-128, SPEC-133]
---
## Why
Press Ctrl/Cmd+F on a workbook today and the browser's own find bar opens.
On a virtualised grid that is worse than nothing: the DOM holds only the
cells in the viewport plus overscan (SPEC-116), so the browser searches a
couple of hundred cells of a forty-thousand-row sheet, reports "0 results"
for a value three screens down, and the user believes it. Search has to run
over the model, and it has to know what a sheet's text is — the displayed
"$1,200.00", or the `=IF(A2="North",1,0)` behind a cell — which no browser
can.

The other half of moving around a large workbook is Go To: Ctrl+G or F5
with a reference or a defined name, the Name Box accepting `B12` or
`Sheet2!A1:C3`, and Go To Special — select every blank, every formula, every
error — which is how people audit a sheet someone else built.

Most of the find surface exists. SPEC-017's `FindBar.vue` is a controlled
component whose only import is `vue` — no document model, no text commands —
so the spreadsheet view can mount it as SPEC-098 does for decks;
`compileQuery` is product-neutral once it lives in `src/lib/core`. This spec
adds the sheet-shaped parts. **It is the normative source for sheet search
semantics (scope, Look in, match rules, order), the sheet replace rules
(re-entry, refusal, skips, one undo step), the Go To reference grammar and
the Name Box's resolution order, and Go To Special's selection rules.**

## Acceptance criteria
- [ ] `src/components/FindBar.vue` is reused, **not forked**: it gains a `placeholder` prop (default "Find in document"), a `showWholeWords` prop (default `true`) and a named `options` slot rendered inside `.find-options` after the regex toggle; its scope props are SPEC-098's `scopeLabel`/`scopeTitle` — added here with SPEC-098's exact names and defaults if SPEC-098 has not landed; `FindBar.spec.ts`'s existing cases and `DocumentView.vue` are unchanged and green
- [ ] `FindOptions`, `compileQuery`, `QueryError`, `isQueryError`, `escapeRegExp` and `INVALID_REGEX_MESSAGE` are imported from `src/lib/core/findQuery.ts`; if SPEC-098 has not landed, this spec performs SPEC-098's move exactly (same file, same names, `src/lib/findReplace.ts` re-exporting them, `findReplace.spec.ts` untouched) so `spreadsheet/find.ts` never imports a text-product module (`spreadsheet/boundary.spec.ts` stays green)
- [ ] Cmd/Ctrl+F opens the bar and Cmd/Ctrl+H opens it with the replace row revealed, matched through `findShortcut` (`find.open`, `find.openReplace`) from `src/lib/shortcuts.ts` with `preventDefault`, so the browser's find never opens over the grid; Escape closes the bar, clears the marks and returns focus to the grid; while the bar is open on macOS, Cmd+G / Cmd+Shift+G step to the next / previous match (the platform's find-next convention — the reason Go To is Control+G there, as in Excel for Mac); Cmd+H hides the app on macOS, so the bar's `find-expand` toggle is the guaranteed way to the replace row; while the cell editor is open these keys, Ctrl+G and F5 are consumed and do nothing (Excel disables Find and Go To in edit mode)
- [ ] The `options` slot carries `sheet-find-entire-cell` ("Match entire cell", a toggle), `sheet-find-within` (select: "Sheet", "Workbook"), `sheet-find-look-in` (select: "Values", "Formulas") and `sheet-find-all` ("Find all"); the sheet passes `showWholeWords: false`, `placeholder` "Find in sheet" / "Find in workbook", and the scope strings in "Messages"; every option persists for the view session; an invalid expression shows `INVALID_REGEX_MESSAGE` and runs nothing
- [ ] A cell's search text per "Search text": Values → the SPEC-118 displayed text at unlimited width; Formulas → `"=" + formula` for a formula cell and SPEC-118's `editText` for a constant; never searched: cells that store no value, array members in Formulas mode, merge-covered cells, notes and comments, drawings and charts, names, foreign sheets, hidden and very hidden sheets; hidden rows and columns **are** searched and a current match inside one shows the hint "This match is in a hidden row or column"
- [ ] Match rules: case-insensitive unless "Match case"; "Match entire cell" anchors the compiled query to the whole search text; "Use regular expressions" per SPEC-017; non-regex queries are literal (`*` and `?` are not wildcards); a cell matches when it holds at least one non-empty match, and the count "N of M" counts cells, not occurrences
- [ ] Scope and order per "Scope and order": with more than one cell selected and Within = Sheet, only the selection; otherwise the active sheet or every visible grid sheet in tab order; row-major within a sheet; Enter / `find-next` goes to the first matching cell after the active cell and Shift+Enter / `find-prev` to the last before it, wrapping with the hint "Wrapped to top"; a match becomes the active cell (switching sheets in workbook scope) and is scrolled into view
- [ ] Visible matching cells carry `data-find-match="true"` and the current one also `data-find-current="true"` on their `sheet-cell`, driven by a `findMarks` prop on `SheetGrid.vue`; the marks clear when the bar closes
- [ ] The walk is chunked per "The walk": `createSheetSearch` visits stored cells only (binary search into the sparse rows, never per-area iteration), the view drives `step` through `scheduleEagerWork` with `FIND_SLICE_BUDGET_MS = 8`, the query debounces `FIND_DEBOUNCE_MS = 150`, the count shows the running total with the hint "Searching…" until done, and any model change (a sheet's `rows` reference or the sheet list) restarts it — a unit test with an injected clock proves no `step` call overruns its deadline by more than one 256-cell check interval
- [ ] "Find all" opens `sheet-find-results` below the bar: one `sheet-find-result` row per matching cell (`data-sheet`, `data-ref`) showing Sheet, Cell, Value and Formula, in search order; activating a row selects that cell exactly as find-next does; above `FIND_ALL_LIMIT = 5_000` rows the list shows the first 5,000 and "Showing the first 5,000 of N matches"
- [ ] "Replace" replaces every occurrence in the current cell's search text and re-enters the result per "Replace rules", then moves to the next match; a refused cell stays unchanged, the bar shows "That replacement would break the formula in B4" (the cell's address) and the current match does not move
- [ ] "Replace all" applies "Replace rules" to every matching cell in scope as **one** undo entry — one `record("Replace all", scope)` (SPEC-121; scope `"workbook"` for Within = Workbook, else `{ sheetId }`) before any edit and one `commit()` after, copy-on-write rows, `dirty` on touched sheets only, one recalculation — and reports "Replaced N occurrences in M cells" plus, when anything was skipped, " · skipped " and the parts in "Messages" (e.g. "Replaced 7 occurrences in 7 cells · skipped 1 protected")
- [ ] Replace rules: Formulas mode re-enters the new text exactly as if typed into the cell through SPEC-122's cell entry, and a new text starting with `=` must parse (`parseFormula`, SPEC-119) or the cell is refused and counted "that would break a formula"; Values mode replaces in constants only (formula cells are counted "formula results"), enters the new text through `parseInput` and never creates a formula; cells SPEC-122's protection rule forbids are skipped and counted "protected"; array formula cells are skipped and counted "in array formulas"; the replacement string is literal (`$1`, `$&` mean nothing)
- [ ] Go To (`sheet-goto-dialog`, on `BaseDialog.vue`) opens on Ctrl+G (the Control key on every platform) and F5 (`sheet.goTo` in SPEC-123's `SHEET_SHORTCUTS`; `preventDefault`, so F5 never reloads the page while a workbook view is mounted and no dialog is open), with a reference field (`sheet-goto-reference`), a list (`sheet-goto-list` of `sheet-goto-item`) of the names and tables SPEC-133's Name Box lists (visible from the active sheet; hidden and reserved names never) then the last ten Go To origins, "Special…" (`sheet-goto-special`) and OK (`sheet-goto-ok`); the reference grammar is "Go To references"; errors show in `sheet-goto-error` and leave the selection unchanged
- [ ] A successful Go To or Name Box jump selects the target's ranges with the active cell at the first range's top-left, switches the active sheet when needed (view state, never `dirty`), scrolls it into view, and first pushes the previous selection (as `Sheet!A1:B2` text) onto the recent list, so F5 then Enter on its first item returns
- [ ] Go To Special (`sheet-goto-special-dialog`) offers `sheet-goto-special-blanks`, `-constants`, `-formulas`, `-errors`, `-last-cell` and `-visible`, with type checkboxes `sheet-goto-type-numbers`, `-text`, `-logicals`, `-errors` (all checked by default) enabled for Constants and Formulas; OK (`sheet-goto-special-ok`) selects per "Go To Special" within the selection when it has more than one cell, else within the used range, as coalesced rectangles; no match → "No cells were found." in `sheet-goto-special-message` and the selection unchanged
- [ ] The Name Box keeps SPEC-133's order of interpretation (a reference navigates; a visible name or table selects its range; otherwise a new name is validated and defined; an invalid entry toasts its `NAME_PROBLEM_MESSAGES` text) and this spec replaces SPEC-133's first two steps (its own reference and name jump): both now run through `resolveGoToTarget`, so the Name Box and the Go To dialog accept exactly the same references and names
- [ ] View wiring: `src/views/spreadsheet/useSheetFind.ts` and `useSheetGoTo.ts` (each ≤ 400 lines under `composableSize.spec.ts`), `src/components/spreadsheet/SheetFindResults.vue`, `GoToDialog.vue`, `GoToSpecialDialog.vue`; framework-free logic in `src/lib/spreadsheet/find.ts`; SpreadsheetView gains only two `use*({ … })` calls, the `<FindBar>` dock and the three component hookups, and `SpreadsheetView.size.spec.ts`'s budget never rises
- [ ] Tests: `spreadsheet/find.spec.ts` (search text per cell kind and both Look-in modes, every match option, scope and order including multi-range selections and wrap, the chunked walk with an injected clock, the replace plan with every skip reason and refusal, literal replacement, Go To grammar acceptances and rejections, Go To Special per option and type filter, rectangle coalescing, the range limit), `core/findQuery.spec.ts` (if this spec performs the move), `FindBar.spec.ts` (the new props and slot; existing cases unchanged), `useSheetFind.spec.ts`, `useSheetGoTo.spec.ts`, `GoToDialog.spec.ts`, `GoToSpecialDialog.spec.ts`
- [ ] e2e (`e2e/SPEC-134.spec.ts`) over `e2e/fixtures/spec134-find.xlsx` per "Fixture": Cmd/Ctrl+F → type `north` → `find-count` "1 of 3" and `A2` is active → Enter → "2 of 3", `A4` active → "Match entire cell" → "2 of 2" (`A6`) → off → "3 of 3" → Within Workbook → "3 of 4" → Look in Formulas → "6 of 8" → Find all → eight `sheet-find-result` rows → Cmd/Ctrl+H → replacement `N.` → Replace all → `find-status` "Replaced 7 occurrences in 7 cells · skipped 1 protected" → Escape → click `C2` → `formula-bar-content` reads `=IF(A2="N.",1,0)` → Cmd/Ctrl+Z once → `A2` reads "North" and `formula-bar-content` reads `=IF(A2="North",1,0)` → Ctrl+G → `Sales!B3:C4` → Enter → `name-box` reads "B3" and `C4` is selected → F5 → choose "Totals" → OK → `name-box` reads "B10" → F5 → Special… → Formulas → OK → `C2`…`C6` and `B10` are selected and `A2` is not → type `D20` in the Name Box + Enter → `name-box` reads "D20" → type `Nowhere!` + Enter → a `toast` reads "Names must start with a letter, underscore or backslash and can't contain spaces or symbols" (SPEC-133)

## Out of scope
Find by format (Excel's "Format…" button), searching notes, comments,
shapes, charts or names (SPEC-135 may add Notes to Look in), Excel's
`*`/`?`/`~` wildcards (regex covers patterns; non-regex stays literal as in
SPEC-017), "Search: By Columns", capture groups in the replacement (SPEC-017's
stated follow-up), replacing inside array formulas, enforcing data validation
on replaced values (Excel does not either), searching hidden sheets, Go To
Special's Comments, Conditional formats, Data validation, Objects, Current
region/array, Row/Column differences, Precedents and Dependents (backlog — the
last two need a trace-precedents UI), R1C1 references in Go To, and a find
surface across workbooks or the library.

## Reference

### Dependencies restated
- **`specs/017-find-and-replace.md`** — the controlled bar
  `src/components/FindBar.vue` (its only import is `vue`): props `visible,
  expanded, query, replaceText, matchCase, wholeWords, useRegex, matchCount,
  currentIndex, error, hint`; emits `update:query, update:replaceText,
  update:matchCase, update:wholeWords, update:useRegex, toggle-expand, next,
  prev, replace, replace-all, close`; testids `find-bar, find-query,
  find-count, find-status, find-prev, find-next, find-close, find-expand,
  find-match-case, find-whole-words, find-regex, find-scope,
  find-replace-row, find-replace-input, find-replace, find-replace-all`;
  `find-status` shows `error`, else `hint`, else "No results". `compileQuery(
  query, { matchCase, wholeWords, regex }): CompiledQuery | QueryError | null`
  (empty → `null`, flags `g`/`gi`), zero-length matches never count, "Replace
  all" is one undo step, "Wrapped to top".
- **`specs/098-…`** (written, not built) — moves those query helpers
  unchanged into `src/lib/core/findQuery.ts` with `findReplace.ts`
  re-exporting them, and adds `scopeLabel` (default "Body text only") and
  `scopeTitle` (default "Search covers body text only. Headers, footers and
  footnotes are excluded.") to `FindBar.vue`'s `find-scope` span.
- **SPEC-116/117** — sparse `Sheet.rows`, `findRow`/`getCell` binary search,
  `usedRange(sheet)`, `parseRange`/`formatRange`/`formatA1`, `GridRange`
  inclusive and normalised; `Row.hidden?`, `ColumnSpan.hidden?`,
  `Sheet.merges`, `Sheet.visibility`, `Sheet.kind`; `SheetSelection { active;
  ranges; activeRangeIndex }`; `SheetGrid.vue` cells `sheet-cell` with
  `data-ref`; `name-box`; `core/platform.ts` `isMacPlatform()`.
- **SPEC-118** — `formatValue(value, code, ctx)` (no `widthPx` → no
  `########`, no width fitting) and `plainText(fv)`; `editText(cell, ctx:
  EditTextContext)` — the formula-bar text of a constant, whose callers
  ("SPEC-134's search" among them) pass the cell's effective style as
  `ctx.style`. The display locale is `Preferences.spreadsheetLocale`,
  resolved by `useSpreadsheetLocale.ts`.
- **SPEC-119/120** — `Cell.formula` canonical text without `=`;
  `tokenize`, `parseFormula` (throws `FormulaSyntaxError`), `ReferenceNode
  { sheet?; sheetEnd?; start; end? }`; `Cell.arrayRange?` on an array anchor;
  `recalculate(model, changed: CellKey[])`.
- **SPEC-121** — `useSpreadsheetUndo` `{ record(label, scope: UndoScope),
  commit(alsoRecalc?), undo, redo, reset }` with `UndoScope = "workbook" |
  { sheetId }`: `record` before mutating, `commit` after (the edit pipeline:
  identity diff, recalculation, persistence); snapshots by reference;
  `Sheet.dirty?: true`.
- **SPEC-122** — cell entry "as typed" (`parseInput(text, locale,
  dateSystem)`, a leading `=` makes a formula, canonicalised); the protection
  rule: a cell whose effective style has `protection.locked` not `false` on a
  protected sheet is not editable.
- **SPEC-133** — `names.ts`: `resolveDefinedName(model, name, qualifier,
  hostSheetId)` (host sheet's scope first, then the workbook), hidden and
  reserved (`_xlnm.*`) names never shown in the Name Box, `validateName` and
  `NAME_PROBLEM_MESSAGES`; tables (`tables.ts`) with a data body. Its Name
  Box order: reference → navigate; visible name or table → select its range
  (a table its data body; a name that is not a reference toasts "That name
  doesn't refer to a range"); else a new name, validated, defined for the
  selection as "Define name"; `name-box-dropdown` / `name-box-option` list
  visible names and tables. SPEC-133's out-of-scope line says Go To by name
  calls `resolveDefinedName`.
- **SPEC-128** — `Row.filterHidden?: true`; "Visible cells only"
  treats it as hidden.
- **SPEC-123** — `SHEET_SHORTCUTS` (`src/views/spreadsheet/sheetShortcuts.ts`)
  matched through `findShortcut(event, definitions)`, whose bindings gain
  `code` and `ctrl` (the Control key on every platform).
- `src/lib/core/idleScheduler.ts` — `scheduleEagerWork(callback(budgetMs),
  budgetMs): IdleHandle`, a macrotask per slot.

### The search model (`src/lib/spreadsheet/find.ts`)
```ts
export type FindLookIn = "values" | "formulas";
export type FindWithin = "sheet" | "workbook";
export interface SheetFindOptions { matchCase: boolean; entireCell: boolean; regex: boolean;
  lookIn: FindLookIn; within: FindWithin }
export interface FindHit { sheetId: string; row: number; col: number }
export interface CompiledCellQuery { test: RegExp; global: RegExp }
export function compileCellQuery(query: string, options: SheetFindOptions):
  CompiledCellQuery | QueryError | null;
export function searchText(model: SpreadsheetModel, sheet: Sheet, cell: Cell,
  lookIn: FindLookIn, ctx: FormatContext): string | null;   // null = not searched
export interface SheetSearch {
  readonly hits: FindHit[];          // in search order, appended as the walk proceeds
  readonly done: boolean;
  step(deadline: () => boolean): void;   // returns when done or deadline() is true
}
export function createSheetSearch(model: SpreadsheetModel, query: CompiledCellQuery,
  options: SheetFindOptions, activeSheetId: string, selection: SheetSelection,
  ctx: FormatContext): SheetSearch;
export interface ReplacePlan {
  edits: { sheetId: string; row: number; col: number; text: string; asValue: boolean }[];
  occurrences: number;
  skipped: { breaksFormula: number; protected: number; array: number; formulaResult: number };
}
export function planReplace(model: SpreadsheetModel, hits: readonly FindHit[],
  query: CompiledCellQuery, replacement: string, options: SheetFindOptions,
  ctx: FormatContext): ReplacePlan;     // pure: mutates nothing
export type GoToTarget = { sheetId: string; ranges: GridRange[] };
export function resolveGoToTarget(model: SpreadsheetModel, text: string, activeSheetId: string):
  GoToTarget | { error: string };
export type SpecialKind = "blanks" | "constants" | "formulas" | "errors" | "lastCell" | "visible";
export interface SpecialTypes { numbers: boolean; text: boolean; logicals: boolean; errors: boolean }
export function selectSpecial(model: SpreadsheetModel, sheet: Sheet, selection: SheetSelection,
  kind: SpecialKind, types: SpecialTypes): GridRange[] | { error: string };
export function coalesceCells(cells: Iterable<CellRef>): GridRange[];
```
`compileCellQuery` calls `compileQuery(query, { matchCase, wholeWords: false,
regex })` and, with `entireCell`, wraps the source as `^(?:…)$`; `test` is the
non-global copy. `ctx.locale` is the user's display locale.

### Search text
| Cell | Values | Formulas |
| --- | --- | --- |
| constant | `plainText(formatValue(value, effectiveFormat, ctx))` | `editText(cell, { ...ctx, style })` |
| formula (incl. array anchor) | its displayed value, as above | `"=" + cell.formula` |
| array member (no formula) | its displayed value | not searched |
| no stored value, merge-covered | not searched | not searched |

Rich text searches the plain `value`. Booleans display `TRUE`/`FALSE`, errors
their code. The effective format follows SPEC-117's style resolution.

### Scope and order
- **Scope.** Within = Sheet: if the selection covers more than one cell, its
  ranges (in `ranges` order); else the whole active sheet. Within = Workbook:
  every sheet with `kind: "grid"` and `visibility: "visible"`, in tab order,
  whole (the selection is ignored, as in Excel). Hidden rows and columns are
  in scope.
- **Order.** Within a range or sheet: rows top to bottom, cells left to
  right. Hits are kept in that order; "N of M" is the current hit's 1-based
  index among all hits.
- **As you type.** When the query or an option changes, the current hit
  becomes the first hit at or after the active cell in search order and is
  selected (SPEC-017's as-you-type behaviour); no hit → "No results".
- **Navigation.** Next = the first hit after the active cell in search order
  — the current hit after any navigation, or wherever the user has since
  clicked, as in Excel — wrapping to the first with "Wrapped to top";
  previous symmetrically. A request the walk cannot answer yet waits. The
  active cell moves through the view's selection setter; a hit on another
  sheet switches `activeSheetIndex` first. Navigation is never an undo entry.

### The walk
`step(deadline)` checks `deadline()` every `FIND_CHECK_INTERVAL = 256`
visited cells. For each range it binary-searches the first stored row ≥ the
range's top and, inside each row, the first cell ≥ its left, so cost is
proportional to stored cells in scope, never to its area. `useSheetFind`
runs a slot per `scheduleEagerWork(…, FIND_SLICE_BUDGET_MS)` with
`deadline = () => performance.now() - start >= FIND_SLICE_BUDGET_MS`, until
`done`. The search reads plain model arrays and mutates nothing, so the grid
does not re-render while it runs. Hits are invalidated, and the walk
restarted after the debounce, when any sheet's `rows` reference or the
`sheets` array changes (copy-on-write makes that check exact).

### Replace rules
For each hit, in order:
1. An array anchor or member → `array`. A cell SPEC-122's protection rule
   forbids → `protected`.
2. Values mode: a cell with a formula → `formulaResult`.
3. `newText` = the search text with every non-empty match of `global`
   replaced by `replacement` verbatim (a replacer function, so `$` patterns
   never apply); `occurrences` += the number replaced.
4. Formulas mode: `newText` starting with `=` must satisfy
   `parseFormula(newText.slice(1))`; a `FormulaSyntaxError` → `breaksFormula`.
   Otherwise the edit is `{ text: newText, asValue: false }`: entered as typed.
5. Values mode: `{ text: newText, asValue: true }`: `parseInput` in the
   display locale; a `formula` result is discarded and the text stored as a
   string.

The composable applies `plan.edits` through SPEC-122's cell entry between
one `record` and one `commit`, keeping each cell's style (a number format
`parseInput` proposes applies only where SPEC-122's typing rule applies it);
`commit` hands every changed key to recalculation once (SPEC-121's edit
pipeline).

### Go To references
`resolveGoToTarget(model, text, activeSheetId)`: split `text` on top-level
commas (via `tokenize`, outside quotes and parentheses); each part must be
one of: a cell, range, whole-column (`A:A`, `A:C`) or whole-row (`3:3`)
reference with an optional sheet prefix (`Sheet2!A1:C3`, `'My Sheet'!B4`),
case-insensitive and `$`-tolerant; a defined name, through
`resolveDefinedName(model, name, qualifier, activeSheetId)`, whose value is a
reference (else "That name doesn't refer to a range", SPEC-133's string); a
table name (its data body, as SPEC-133's Name Box selects it). Every part
must land on the same sheet. Rejected with `GO_TO_INVALID_MESSAGE`: 3-D and
external references, R1C1, intersections, anything outside
`A1:XFD1048576`, an unknown sheet or name. A target on a
hidden or very hidden sheet → `GO_TO_HIDDEN_SHEET_MESSAGE`; on a foreign
sheet → invalid.

### Go To Special
- **Scope.** Selection with more than one cell → its ranges, each intersected
  with `usedRange(sheet)`; otherwise `usedRange(sheet)`. Empty used range →
  "No cells were found.".
- **Blanks** — cells in scope that store no value (absent or `value: null`
  without a formula); a formula returning `""` is not blank. Computed per
  row as the complement of that row's stored columns within each range — a
  row with no stored cell is one full-width run — so the cost is rows in
  scope plus stored cells, never the area.
- **Constants** — stored non-null values without a formula, not array
  members, filtered by type: numbers (dates included), text, logicals,
  errors.
- **Formulas** — cells with a `formula` and array members, filtered by the
  type of their value; `""` is text, an uncomputed `null` counts as a
  number.
- **Errors** — any cell whose value is a `CellError`, constant or formula
  (a shortcut Excel spreads over two checkboxes).
- **Last cell** — the single cell at `usedRange`'s bottom row and right
  column (A1 for an empty sheet), ignoring the selection.
- **Visible cells only** — the scope minus rows that are `hidden` (or
  `filterHidden`) and columns that are hidden, as the cross product of the
  visible row runs and column runs within each range.
- **Coalescing.** Matching cells are swept row by row into maximal horizontal
  runs; a run extends the rectangle directly above it when both have the same
  `left`/`right` and the rows are adjacent. More than
  `GO_TO_SPECIAL_MAX_RANGES = 10_000` rectangles → the "too many" message
  and nothing changes. The active cell is the first rectangle's top-left.

### The Name Box
SPEC-133 owns the Name Box's order and its messages; this spec owns its
reference grammar. On Enter: (1) `resolveGoToTarget` succeeds → jump, as Go
To does; (2) it fails because the text named an existing name that is not a
reference → SPEC-133's "That name doesn't refer to a range" toast; (3)
otherwise SPEC-133's new-name path (`validateName`, then "Define name" or the
`NAME_PROBLEM_MESSAGES` toast). Valid names can never parse as references
(`validateName` refuses `A1` and `R1C1` shapes), so steps (1) and (3) never
compete. SPEC-133's own reference and name jump (its first two steps) is
replaced by step (1), so one resolver decides.

### Keys
SPEC-123's `SHEET_SHORTCUTS` gains `{ id: "sheet.goTo", label: "Go to",
bindings: [{ key: "g", ctrl: true }, { key: "F5" }] }` — Control+G on every
platform, so macOS keeps Cmd+G for find-next (handled by `useSheetFind`
while the bar is open: `{ key: "g", mod: true }` next, with `shift` previous,
macOS only); neither binding is taken by another entry, and the text
registry's `shortcuts.spec.ts` is unchanged. `useSheetGoTo` matches it with
`findShortcut(event, SHEET_SHORTCUTS)` and calls `preventDefault`. F5 is taken while a workbook view
is mounted and no dialog is open; the store is write-through, so this guards
against losing one's place, not data.

### Messages
| Where | String |
| --- | --- |
| `scopeLabel` | "Cells only" |
| `scopeTitle` | "Search covers cell values and formulas. Notes, comments, shapes, charts and hidden sheets are not searched." |
| hint, hidden match | "This match is in a hidden row or column" |
| hint, walking | "Searching…" |
| replace report | "Replaced N occurrence(s) in M cell(s)" + " · skipped " + parts joined by ", " |
| skip parts | "K that would break a formula", "K protected", "K in array formulas", "K formula results" |
| single refusal | "That replacement would break the formula in B4" |
| Find all cap | "Showing the first 5,000 of N matches" |
| `GO_TO_INVALID_MESSAGE` | "That isn't a cell reference or a defined name" |
| `GO_TO_HIDDEN_SHEET_MESSAGE` | "That sheet is hidden" |
| `GO_TO_SPECIAL_NONE_MESSAGE` | "No cells were found." |
| `GO_TO_SPECIAL_TOO_MANY_MESSAGE` | "Too many separate areas to select (more than 10,000)" |

Singular forms ("1 occurrence", "1 cell") are used for one.

### Fixture
`e2e/fixtures/spec134-find.xlsx`, built by `spreadsheet/fixtures/build.ts`
in Excel's shape (shared strings, `xr:uid`, `x14ac:dyDescent`, cached values,
a currency `numFmt` ≥ 164 on `B2:B6`): sheet "Sales" (active) — `A1:C1`
"Region", "Amount", "Flag"; `A2` "North", `A3` "South", `A4` "Northwest",
`A5` empty, `A6` "north"; `B2:B6` 1200, 800, 450, 300, 50; `C2`, `C3`, `C4`,
`C6` `=IF(An="North",1,0)` with cached results, `C5` `=B5/0` cached
`#DIV/0!`; `B10` `=SUM(B2:B6)` cached 2800; defined name `Totals` →
`Sales!$B$10`. Sheet "Locked" — `<sheetProtection sheet="1">`, every cell
locked, `A1` "North office", `A2` "South office". Sheet "Archive" —
`state="hidden"`, `A1` "North".

### Tests
As listed in the acceptance criteria. Unit tests drive `find.ts` over models
built in the test (no file I/O) except the e2e, which uses the fixture.

### Specs this changes
- `specs/017-…` — `FindBar.vue` gains `placeholder`, `showWholeWords` and the
  `options` slot; the text editor passes none of them and is unchanged.
- `specs/098-…` — if still unbuilt, its `core/findQuery.ts` move and scope
  props land here first, identically; SPEC-098 then finds them done.
- `specs/117-…` — `SheetGrid.vue` gains the `findMarks` prop.
- `specs/133-…` — the Name Box's reference and name steps run through
  `resolveGoToTarget`; `resolveDefinedName` is shared with Go To.
- `specs/123-…` — `SHEET_SHORTCUTS` gains `sheet.goTo` (Ctrl+G, F5).
- `specs/000-roadmap.md` — no change here (the lead owns it).
