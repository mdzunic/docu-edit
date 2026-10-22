---
id: SPEC-122
title: Cell entry and the sheet writer
status: draft
area: editing
complexity: complex
story_points: 21
depends_on: [SPEC-021, SPEC-118, SPEC-120, SPEC-121]
---
## Why
SPEC-121 lets a user create a workbook and rearrange its sheets, but not
type a single value: the grid is still read-only, and the writers cannot
serialise a sheet whose cells changed — SPEC-121 ships an interim that
re-emits the retained source with its sheet names rewritten and asks Excel
to recalculate on load. This spec makes cells editable the way Excel users
expect them to be — type to replace, F2 to edit in place, `=` and the arrow
keys to point at references, Enter and Tab to commit and move — and builds
the half of the save path that turns a changed sheet back into a part: the
worksheet serialiser for `.xlsx`, the table re-serialiser for `.ods`, the
shared-strings and style tables both grow into, and the parsing that turns
"1,200", "10%", "8-Oct" and "=A2*2" into the right value, formula and
number format for the user's locale.

The save path is where a spreadsheet editor loses people's work. Excel's
own files are dense with things this product does not model — dynamic-array
metadata on a cell, phonetic runs in a shared string, outline levels on a
row, `x14` extensions, a data-table formula — and LibreOffice's tables carry
annotations, links and frames inside cells, row groups around them and a
million-row tail of repeated empty rows. A rebuilt sheet must keep every one
of them, must leave every untouched sheet byte-identical, and must only ever
*append* to the workbook-wide string and style tables that untouched sheets
still reference by number.

**This file is the normative source for sheet-part serialisation in both
formats (the xlsx splice rules and element order, the ods table
re-serialisation and repeat compression), the append-only shared-strings and
style writers, `internStyle`, the generated ods cell and number styles, the
`fullCalcOnLoad` rule, the cell editor and formula bar editing, value input
parsing (`numberFormat/input.ts`), point mode, and protection honour.**

## Acceptance criteria
- [ ] With a cell selected, typing a printable character (or starting an IME composition) opens the in-cell editor `cell-editor` in **Enter** mode with that input replacing the cell's content: the editor element is always mounted over the active cell and holds focus while the grid does (transparent and empty in **Ready** state, `aria-label` "B3, Revenue" naming the cell and its displayed value), so the first character — composed or not — lands natively and is never lost or doubled
- [ ] F2 or a double-click opens the editor in **Edit** mode on the cell's edit text — "=" + `formula` for a formula cell (an unparsed one as stored), else SPEC-118's `editText(cell, { ...ctx, style })` with the effective style from SPEC-117's `effectiveStyleId`/`cellStyle` (a quote-prefixed string with its leading "'") — caret at the end (F2) or at the clicked character; F2 inside the editor toggles Enter ↔ Edit; the status bar shows the mode in `sheet-status-mode` ("Ready", "Enter", "Edit", "Point")
- [ ] Keys follow "Editor keys": Enter / Shift+Enter commit and move down/up, Tab / Shift+Tab right/left, Enter after a run of Tabs returns to the run's starting column one row down, and inside a multi-cell selection Enter and Tab cycle within it without collapsing it; in Enter mode the arrow keys commit and move (unless point mode takes them), in Edit mode they move the caret; Escape cancels and restores the cell; Alt/Option+Enter inserts a line break; Ctrl/Cmd+Enter writes the entry into every cell of the selection
- [ ] Point mode: when the caret follows "=", an operator, "(" or "," in a formula, arrow keys, clicks and drags insert (and then replace) a reference at the caret — Shift+arrows extend it to a range, Ctrl/Cmd+arrows jump to the data edge (SPEC-117 `edgeInDirection`), clicking another sheet's tab shows that sheet and inserts sheet-qualified references quoted as SPEC-119 prints them (`'My Sheet'!B2`), and commit or cancel returns to the edited sheet; typing any other character ends the current point; F4 cycles the reference at the caret through `A1` → `$A$1` → `A$1` → `$A1`
- [ ] Every reference in the editor text is coloured, and each one on the displayed sheet is outlined on the grid in the same colour (`sheet-ref-highlight`, `data-ref-index`), colours assigned by first appearance from the eight `--sheet-ref-1..8` tokens (light and dark values), the reference being pointed at drawn dashed
- [ ] The formula bar's `formula-bar-content` becomes editable and is the same draft as the in-cell editor — typing in either updates both, clicking it opens Edit mode at the clicked character, and Enter/Tab/Escape behave identically; while editing, registry shortcuts (`edit.undo` …) are left to the editor's native text editing, and the workbook undo applies only after commit
- [ ] Commit parses the text per "Input parsing" with the user's SPEC-066 locale and the workbook's date system: a formula is validated with SPEC-119's `parseFormula` (unbalanced trailing "(" are closed first, as Excel offers); a syntax error keeps the editor open with "There's a problem with this formula" in `cell-editor-error` and the caret at the error offset; text over 32,767 characters cannot be typed, and a formula over 8,192 characters shows "Formulas can be at most 8,192 characters"
- [ ] `numberFormat/input.ts` exports `parseInput(text, locale, dateSystem, options?)` per "Input parsing", taking every number, date and time **value** from SPEC-120's `valueFromText` (so a typed value and a coerced one never disagree) and adding only the classification that picks the number format: numbers with the locale's group and decimal separators, leading sign, "(12)", percent, the locale's currency, scientific, "1 1/2" fractions, dates in the locale's order and ISO `yyyy-mm-dd` in every locale, month-name dates, times (24-hour and AM/PM, over 24 hours), date + time, `TRUE`/`FALSE` and error literals case-insensitively, a leading apostrophe → text with `quotePrefix`, more than 15 significant digits truncated like Excel; a cell formatted as Text (`@`) stores text verbatim, a percent-formatted cell turns "10" into 0.1, and a returned number format is applied only to a cell whose format is General — table-tested for all six SPEC-118 locales
- [ ] Each commit is one undo step through SPEC-121's `record`/`commit` (label `Typing "<first 20 characters>" in B3`, scope the sheet), recalculates through SPEC-121's pipeline so dependents on every sheet update before the next paint, marks the sheet dirty and writes through; committing an unchanged entry records nothing; committing a changed text on a rich-text cell stores plain text (an unchanged one keeps its runs)
- [ ] Delete or Backspace on a selection (not editing) clears values and formulas of every selected cell — styles, notes and other preserved content stay — as one step "Clear contents"; a whole-row or whole-column selection touches only stored cells; Ctrl/Cmd+Enter is one step "Fill", writing each target the entry with its formula shifted relatively by SPEC-119's `shiftFormula` (merged ranges receive it in their anchor only), whole-row/column targets clipped to the used range
- [ ] A line break committed into a cell without wrap sets `alignment.wrap` on its style through `internStyle`, and every row without `customHeight` that a commit, Fill or Clear contents touched refits through `refitRows`/`autofitRowHeightPt` per "Row autofit", inside the same undo step (this spec owns the function; SPEC-124 reuses it for its AutoFit commands); typed text passes through `core/xmlText.ts` `sanitizeXmlText`; paste into the editor inserts plain text; Enter, Tab, arrows and Escape are ignored while an IME composition is active and the editor never re-renders mid-composition
- [ ] Protection is honoured per "Protection": a locked cell on a protected sheet (xlsx `sheetProtection@sheet="1"`, ods `table:protected="true"`) refuses the editor, Delete and Fill with the toast "This cell is on a protected sheet" while unlocked cells edit normally; a hidden formula on a protected sheet shows an empty formula bar; workbook structure protection (`workbookProtection@lockStructure="1"`, `table:structure-protected="true"`) disables SPEC-121's sheet commands with the reason "The workbook's structure is protected"; protection XML is never changed
- [ ] Editing any cell of an array formula's range other than its anchor refuses with the toast "You can't change part of an array"; editing the anchor keeps its `arrayRange`
- [ ] A dirty `.xlsx` sheet is written by `serializeWorksheet` per "The worksheet writer": `<dimension>`, `<sheetViews>`, `<sheetFormatPr>`, `<cols>`, `<sheetData>` and `<mergeCells>` are rebuilt from the model and spliced into the retained part in CT_Worksheet order; the root start tag (namespaces, `mc:Ignorable`, `xr:uid`) and every other element stay byte-verbatim; cells are written per the type table — never `t="shared"`, strings as `t="s"`, formula string results `t="str"`, array anchors `t="array"` with `ref` (a dynamic array keeping its `cm`), a data table's `dataTableXml` verbatim, an unparsed formula's wire text verbatim in its own format only, numbers as shortest round-trip text, `_xHHHH_` encoding for XML-illegal characters and literal `_x` sequences — rows carry no `spans` or `x14ac:dyDescent`, and the first `sheetView` carries `showFormulas` from SPEC-119's `Sheet.showFormulas`
- [ ] `xl/sharedStrings.xml` is append-only per "Shared strings": an unchanged string cell keeps its source index (rich runs and phonetics intact), a new plain string reuses an equal plain source item or is appended, inline strings are converted to shared ones, source `<si>` bytes are never touched, `uniqueCount` is updated and `count` removed, and a workbook without the part gains it with its relationship and override
- [ ] `internStyle(model, style)` (`spreadsheet/styles.ts`) returns the index of a structurally equal style or appends one copy-on-write; the xlsx style writer reuses `sourceXf`, appends fonts, fills, borders, number formats (`builtinFormatId` written as the `numFmtId` when set; otherwise a built-in id only for a code equal to a **locale-independent** built-in — never one of SPEC-118's `LOCALE_DEPENDENT_BUILTINS` — else custom ids from 164) and `cellXfs` entries with `count`s updated, writes `quotePrefix="1"`, and leaves every source entry and every other `xl/styles.xml` element byte-verbatim; the from-scratch base is built from `model.styles[0]`
- [ ] A dirty `.ods` table is written by `serializeTable` per "The table writer": start tag and leading children verbatim (with SPEC-121's name and style splices), columns and rows regenerated with `table:number-columns-repeated` / `table:number-rows-repeated` compression over identical neighbours (empty ones included), the source's row and column groups and header rows reproduced around the indices they covered, merges as spans plus `table:covered-table-cell`, the source's trailing repeated tail re-emitted to the same extent, and trailing table children (`table:named-expressions`, `calcext:conditional-formats`) verbatim
- [ ] ODS cells carry `office:value-type` chosen from value and number format (float, percentage, currency with `office:currency`, date, time, boolean, string), `calcext:value-type`, `<text:p>` holding the SPEC-118 formatted text (lines as paragraphs, `text:s`/`text:tab` for whitespace, `text:span` for rich runs), `table:formula` via `toOdsFormula` (an unparsed one's wire value verbatim), array anchors as `table:number-matrix-*-spanned`, and their preserved attributes, text and child content verbatim; new cell, text and number styles are generated as `de-ce<n>`, `de-T<n>` and `de-N<n>` through SPEC-121's `OdsStyleRegistry` (number styles by SPEC-118's `odsNumberStyleXml(code, name)`, borders by SPEC-117's pinned ODS border write table), and a style read from the file keeps its name
- [ ] `xl/calcChain.xml` is dropped on any edit (SPEC-121's rule, restated); `calcPr@fullCalcOnLoad="1"` is written when `model.forceFullCalcOnLoad` is set (SPEC-120 sets it whenever a pass leaves a stale cell) and kept when the source had it; SPEC-121's interim rule for dirty sheets is removed; `.ods` has no equivalent and a stale cell's last known value is written
- [ ] Round trips: for every fixture, an edit → write → read gives the edited model (`workbookForCompare`, both formats); write → read → write is `equalBytes`-stable for rebuilt parts; untouched sheet parts and tables stay `equalBytes`/splice-equal; `xl/styles.xml` and `xl/sharedStrings.xml` after an edit equal their source bytes outside the spliced appends and counts; forcing every sheet of `mini-book.ods` and `spec122-shapes.ods` dirty without edits reads back to the same model
- [ ] The readers capture the fields in "Model additions" and this spec **bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`**, whose dirty-sheet step re-reads them from the retained source onto cells whose value and formula still equal the source's
- [ ] SPEC-121's from-scratch writers now serialise every sheet's cells and styles, and "Download a copy as…" is enabled for every workbook: `mini-book` and both `spec122-shapes` fixtures, converted each way and re-opened, show the same displayed values, formulas and number formats
- [ ] Fixtures: the builder adds `e2e/fixtures/spec122-shapes.xlsx` and `spec122-shapes.ods` per "Fixtures", emulating Excel's and LibreOffice's real output
- [ ] Tests: `numberFormat/input.spec.ts`, `pointMode.spec.ts`, `cells.spec.ts`, `autofit.spec.ts` (`fixedMeasurer` cases for wrap, rotation, stacked text, the 0.75 pt step and the row cap), `styles.spec.ts`, `xlsx/sheetXml.spec.ts`, `xlsx/sharedStrings.spec.ts`, `xlsx/styles.spec.ts`, `ods/tableXml.spec.ts`, `ods/styles.spec.ts`, `roundTrip.spec.ts` and `open.spec.ts` additions, `views/spreadsheet/useCellEditing.spec.ts`, `usePointMode.spec.ts`, `components/spreadsheet/CellEditor.spec.ts`
- [ ] e2e (`e2e/SPEC-122.spec.ts`): New spreadsheet → type "Revenue" in A1, Enter → A2 is active → "1,200", Tab → A2 reads "1,200" (`#,##0` applied) → in B2 type "=A2*2", Enter → B2 reads "2400" → in B3 type "=SUM(", ArrowUp twice, Shift+ArrowDown, ")" → the editor reads "=SUM(B1:B2)", Enter → "2400" → C1 "10%" reads "10%" → C2 "2026-10-08" reads the en-US short date "10/8/2026" → C3 Alt+Enter two lines wraps → select D1:D3, type "x", Ctrl+Enter → three "x" → Delete → empty → undo → "x" back → type "=1+" in E1, Enter → "There's a problem with this formula" → Escape → download, re-open → A1–D3 as left and B2's formula bar reads "=A2*2" → download a copy as `.ods`, re-open → same displayed values; open `spec122-shapes.xlsx` → edit one cell on "Edit" → download → every other sheet part `equalBytes` with the fixture, `xl/calcChain.xml` absent → on "Locked", typing shows "This cell is on a protected sheet"

## Out of scope
Formatting commands, the Format Cells dialog and painting styles
(SPEC-123); inserting/deleting rows and columns, resizing, the AutoFit commands and
merging (SPEC-124); clipboard, fill handle and drag-move (SPEC-125 — which
reuses `fillWithEntry`'s shift); function autocomplete, argument hints and
Insert Function (SPEC-127); in-column value AutoComplete and the AutoCorrect
list; partial (rich-text) formatting inside the editor; entering legacy
array formulas with Ctrl+Shift+Enter and evaluating dynamic arrays (both
backlog — existing ones are preserved and displayed); data-table (What-If)
formulas beyond verbatim preservation; locale-specific formula syntax and
localised `TRUE`/`FALSE`; editing or removing protection; the Name Box as a
Go To field (SPEC-134); right-to-left editing. Treating a leading "+" or
"-" as the start of a formula (Lotus/Excel habit) — only "=" starts one.

## Reference

### Dependencies restated
- `specs/021-keyboard-shortcuts-and-shortcut-reference.md` — the single
  registry `src/lib/shortcuts.ts` (`findShortcut(event)`,
  `isOutsideDocumentInput(target)` — contenteditable counts as the document
  surface, `edit.undo` / `edit.redo`, `shortcutLabel`); per-component
  conventions (Enter confirms a field, Escape dismisses) are deliberately
  not registry commands. The editor keys below are such conventions of the
  cell editor, handled in `useCellEditing`, so Ctrl/Cmd+Enter never collides
  with the text product's `insert.pageBreak`.
- `specs/118-number-formats-and-dates.md` — `formatValue(value, code, ctx):
  FormattedValue`, `parseFormatCode(code)` (`FormatSection.kind`: number,
  fraction, scientific, date, text, general, literal), `BUILTIN_NUMBER_FORMATS`
  (ids 0–49, en-US spellings), `LOCALE_DEPENDENT_BUILTINS` (5–8, 14, 22,
  41–44), `formatCodeOf(style, locale)`, `CellStyle.builtinFormatId`
  (a locale-dependent built-in renders in the user's locale);
  `SpreadsheetLocale { tag; decimal; group; dateOrder; dateSeparator;
  shortDate; …; amPm; currencySymbol; currencyPosition; currencyCode }` for
  en-US, en-GB, de-DE, fr-FR, es-ES, cs-CZ (a per-user SPEC-066 preference,
  never per file); `editText(cell, ctx: EditTextContext)` with `ctx.style` =
  the cell's effective style; `odsNumberStyleXml(code, name)` in
  `numberFormat/odsNumberStyles.ts` (main style `name`, sub-styles
  `${name}P0`…). Formulas are always entered in invariant en-US syntax.
- `specs/120-calculation-engine-and-core-functions.md` —
  `valueFromText(text, locale, dateSystem): number | null` in
  `formula/values.ts` (the text → number/date/time grammar `parseInput`
  must agree with); `MAX_TEXT_LENGTH = 32_767`; `cellKey`; recalculation
  replaces a row only when a cached value changes; `model.forceFullCalcOnLoad`
  (set by any pass that leaves a stale cell, never cleared) is what the xlsx
  writer turns into `calcPr@fullCalcOnLoad="1"`.
- `specs/121-create-workbook-and-sheet-management.md` — `Sheet.dirty`,
  `structureDirty`, `copiedFrom`; `useSpreadsheetUndo` `record(label,
  scope)` before and `commit(alsoRecalc?)` after a mutation (identity diff →
  recalculate → mark dirty → `touchDocument`); `UndoScope = "workbook" | {
  sheetId }`; `SHEET_BY_REFERENCE_FIELDS`; the writers' three paths, the
  workbook-level splices, the sheet-reference pass, the calcChain drop;
  `OdsStyleRegistry`
  (`de-ta`, `de-co`, `de-ro`, collision-proof counter) in `ods/styles.ts`;
  the from-scratch empty-sheet goldens; `stripRetention`;
  `workbookForCompare`; `WriteXlsxOptions.fullCalcOnLoad`; the
  `SpreadsheetMenuBar` Edit menu.
- `specs/116-…` / `117-…` — the model and its invariants (copy-on-write
  rows, sparse, the source is the splice base; sheet objects replaced, never
  mutated); `core/xmlScanner.ts` offsets; `address.ts` (`formatA1`,
  `formatRange`, `getCell`, `usedRange`); `serial.ts`
  (`serialToOdsDateValue`, `serialToOdsTimeValue`); `CellStyle` (font, fill,
  borders, alignment, `numberFormat`, `builtinFormatId`, `quotePrefix` (read
  from `xf@quotePrefix` there), `protection { locked?: false; hidden?:
  true }`, `namedStyle`, `sourceXf`, `sourceStyleName`);
  `effectiveStyleId(sheet, row, col)` (cell → row if `customFormat` →
  column → 0) and `cellStyle(model, id)` in `styles.ts`; `ColumnSpan { min,
  max, widthPx?, hidden?, styleId?, sourceWidthChars?, sourceAttrs?,
  sourceStyleName? }`; `Row { heightPt?, customHeight?, hidden?, styleId?,
  sourceAttrs?, sourceStyleName? }`; the colour codec
  `modelColorToXlsxAttrs(color, theme)` in `xlsx/themeColor.ts` (theme slot
  + `tint`, or `rgb` when no tint expresses the reference); the px ↔ chars
  formula over the
  workbook's MDW; the pinned "ODS borders" write table; `Sheet.merges`,
  `frozen`, `showGridLines`, `zoomPercent`, `selection`; `SheetSelection`
  and `selection.ts`; `TextRun.font`.
- `specs/119-the-formula-language.md` — `Cell.formula` (canonical, or the
  **wire** text when `formulaUnparsed` — xlsx `<f>` text, ods the whole
  `table:formula` value with its prefix — re-emitted verbatim in its own
  format only), `arrayRange` (anchor only), `dynamicArray` + `cm` (xlsx
  `c@cm`), `dataTableXml` (the `<f t="dataTable">` element verbatim),
  `Sheet.showFormulas` (view state); `parseFormula` / `FormulaSyntaxError`
  (`offset`); `shiftFormula`, `findArrayAnchor`; `toXlsxFormula` /
  `toOdsFormula`; `formulaBarText` (the read-only bar this spec makes
  editable).

### Model additions
```ts
export interface Cell {
  /** xlsx: the shared-string index the value was read from; reused while
   *  value and runs are unchanged. */
  sstIndex?: number;
  /** Unmodelled start-tag attributes, verbatim with leading spaces:
   *  xlsx ` vm="2"`, ` ph="1"` (`cm` is SPEC-119's field); ods
   *  ` table:content-validation-name="v1"`, … */
  preservedAttrs?: string;
  /** Unmodelled children, verbatim: xlsx `<extLst>` (a data table's `<f>`
   *  is SPEC-119's `dataTableXml`); ods everything except the text
   *  paragraphs (annotations, frames, unknown elements), in source order. */
  preservedXml?: string;
  /** ods: the cell's `text:p` sequence verbatim when it holds inline content
   *  the model does not (links, fields, notes). */
  preservedText?: string;
}
// Row and ColumnSpan gain nothing: SPEC-117's `sourceAttrs` (xlsx unmodelled
// <row>/<col> attributes — outlineLevel, collapsed, thickTop, thickBot, ph,
// bestFit, phonetic … — verbatim) and `sourceStyleName` (ods row/column style)
// are what the writers re-emit.
export interface Sheet { protection?: { sheet: true } }
export interface SpreadsheetModel { structureProtected?: true }
```
Drop rules, applied by every command (this spec's and later ones'): setting
a cell's value or runs removes `sstIndex`, `preservedText` and ` vm="…"`;
setting or removing its formula removes `cm`, `dynamicArray` and
`dataTableXml`; changing a row's height removes its `sourceStyleName`, a column's width
likewise. A cell holding only preserved content is stored (sparseness counts
it as content). A later spec that models an attribute or child stops
capturing it here, and its reader step moves it into the model.

### Editor keys
| Key | Ready | Enter mode | Edit mode |
| --- | --- | --- | --- |
| printable, IME | open Enter, input replaces content | insert | insert |
| F2 | open Edit, caret at end | → Edit | → Enter |
| Enter / Shift+Enter | move down / up (SPEC-117) | commit, move down / up | same |
| Tab / Shift+Tab | move right / left | commit, move right / left | same |
| arrows | move (SPEC-117) | point if allowed, else commit + move | move caret |
| Escape | — | cancel | cancel |
| Alt/Option+Enter | — | line break | line break |
| Ctrl/Cmd+Enter | — | commit into the whole selection | same |
| Delete / Backspace | clear contents | edit text | edit text |
| F4 | — | cycle `$` at caret | same |
| Cmd/Ctrl+Z, +Y | workbook undo/redo | native text undo | native text undo |
Inside a multi-cell selection Enter moves down the column and wraps to the
next column's top (then the next range, then back to the first), Tab moves
along the row and wraps to the next row; Shift reverses. A Tab-run's return
column is forgotten by any navigation other than Tab/Enter commits.

**Editor DOM.** `CellEditor.vue` is a `contenteditable` div sized to the
active cell, growing right across empty neighbours and down as lines grow,
in the cell's resolved font. Input is normalise-after-input (SPEC-081's
pattern): after the browser applies an edit the DOM is read back to plain
text (`<br>`/blocks → "\n"), re-rendered as text plus one coloured `<span>`
per reference, and the caret restored by character offset — skipped while
`isComposing`. `useCellEditing.ts` owns the state machine and the draft
(`text`, `caret`), shared by `CellEditor.vue` and the editable formula bar;
`usePointMode.ts` owns point references, highlights and sheet switching.
SPEC-117's navigation moves from the grid element to the editor proxy
unchanged.

### Point mode (`src/lib/spreadsheet/pointMode.ts`)
```ts
export function canPointAt(text: string, caret: number): boolean;
export function referenceSpans(text: string): { start: number; end: number; ref: ReferenceNode; colorIndex: number }[];
export function placePointReference(text: string, span: { start: number; end: number } | null, ref: string):
  { text: string; span: { start: number; end: number } };
export function cycleAbsolute(text: string, caret: number): { text: string; caret: number };
```
`canPointAt`: the text starts with "=" and the text before the caret, less
trailing spaces, ends with "=", "(", ",", a binary operator (`+ - * / ^ & =
< > <= >= <>`) or a unary sign. All four are pure and token-based over
SPEC-119's `tokenize`. Colour index = order of first appearance of each
distinct reference text, modulo 8. In Edit mode clicks point when allowed;
arrows always move the caret.

### Input parsing (`numberFormat/input.ts`)
```ts
export interface ParsedInput {
  value: CellValue;
  formula?: string;          // canonical, without "=" — validated by the caller
  numberFormat?: string;     // applied only when the cell's format is General
  builtinFormatId?: number;  // with numberFormat, for SPEC-118's locale-dependent built-ins (14, 22)
  quotePrefix?: true;
}
export function parseInput(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904",
  options?: { existingFormat?: string }): ParsedInput;
```
Every number, date, time and date-time **value** comes from SPEC-120's
`valueFromText(text, locale, dateSystem)` — its grammar (group separators
only between three-digit groups, U+0020/U+00A0/U+202F accepted for a
space-like group, sign or parentheses, `%`, `locale.currencySymbol` at
`locale.currencyPosition`, the exponent, dates in `locale.dateOrder` with
`/`, `-` or `.`, English and locale month names, two-digit years 00–29 →
20xx and 30–99 → 19xx, a missing year from core `now()`, ISO everywhere,
times with `locale.amPm`, hours past 24) is the one grammar, so a typed
value and a text coerced by a formula never disagree. `parseInput` adds the
steps around it and classifies which shape matched, to choose the format.
Rules in order (shown for en-US):
| Input | `value` | `numberFormat` |
| --- | --- | --- |
| "" | `null` | — |
| existing format `@` | the text verbatim (even "=A1") | — |
| `=…` (more than "=") | formula = the rest | — |
| `'…` | the text after the apostrophe; `quotePrefix` | — |
| `42`, `-1.5`, `+7`, `007` | number | — |
| `1,234,567`, `1,234.5` (3-digit groups only) | number | `#,##0` / `#,##0.00` |
| `(12)` | −12 | — |
| `10%`, `12.5%` | 0.1, 0.125 | `0%` / `0.00%` |
| `$1,200`, `$1,200.50`, `-$5`, `($5)` (de-DE `1.200 €`) | number | `"$"#,##0` / `"$"#,##0.00` (de-DE `#,##0" €"`) — the locale's symbol and side |
| `1.5E3`, `1e-4` | number | `0.00E+00` |
| `1 1/2`, `0 3/16` (parsed here: an integer, a space, a fraction) | 1.5, 0.1875 | `# ?/?` / `# ??/??` (by denominator digits) |
| `10/8/2026`, `10-8-26` (locale order) | serial | `builtinFormatId: 14` (+ its en-US code) |
| `2026-10-08` (every locale) | serial | `builtinFormatId: 14` |
| `8-Oct-2026`, `Oct 8, 2026`, `8 October 2026` | serial | `[15]` `d-mmm-yy` |
| `8-Oct`, `Oct 8` | serial in the current year (core `now()`) | `[16]` `d-mmm` |
| `Oct-2026`, `Oct 2026` | the 1st of the month | `[17]` `mmm-yy` |
| `13:45` / `13:45:30` | day fraction | `[20]` / `[21]` |
| `1:45 PM`, `1:45:30 pm`, `1 PM` | day fraction | `[18]` / `[19]` |
| `36:00`, `25:30:00` | 1.5 … | `[46]` `[h]:mm:ss` |
| `10/8/2026 13:45`, `2026-10-08T13:45` | serial + fraction | `builtinFormatId: 22` |
| `TRUE`, `false` | boolean | — |
| `#N/A`, `#div/0!` (any `ErrorCode`) | `CellError` | — |
| anything else | the text as typed | — |
`[n]` is `BUILTIN_NUMBER_FORMATS[n]` — 15–21 and 46 are
locale-independent codes; 14 and 22 are SPEC-118's locale-dependent
built-ins, carried as `builtinFormatId` so they render, and are written as
`numFmtId`, in the user's locale.
- Numbers: digits after the 15th significant digit become 0 before
  `valueFromText` sees them ("1234567890123456789" → 1234567890123450000,
  Excel's truncation); a percent-formatted cell divides a plain number by
  100 (Excel's automatic percent entry).
- Dates: `valueFromText` validates them (1900-02-29 is serial 60 in the 1900
  system); a date before the system's first day or after 9999-12-31 is text.
- The current year comes from core `now()` — tests pin it with `setClock`
  (the clock-coupled-test rule).

### Cell commands (`src/lib/spreadsheet/cells.ts`)
```ts
export interface CellEntry { value: CellValue; formula?: string; numberFormat?: string; builtinFormatId?: number; quotePrefix?: true; wrap?: true }
export function setCellEntry(model: SpreadsheetModel, sheetIndex: number, at: CellRef, entry: CellEntry): void;
export function fillWithEntry(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[], origin: CellRef, entry: CellEntry): void;
export function clearContents(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): void;
export function firstLockedCell(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): CellRef | null;
export function arrayRangeTouched(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): GridRange | null;
```
Copy-on-write throughout. The new style is the cell's effective style
(`cellStyle(model, effectiveStyleId(sheet, row, col))`) with `numberFormat`
and `builtinFormatId` (only when the effective format is General; setting
`numberFormat` without a `builtinFormatId` clears it, SPEC-118's rule),
`quotePrefix` (set when the entry has it, removed otherwise) and `wrap`
changed, interned; an unchanged effective style keeps
`styleId` as it was. A cell left with no value, formula, style or preserved
content is removed, and a row left empty without properties likewise.

### Row autofit (`src/lib/spreadsheet/autofit.ts`)
```ts
export interface MeasuredFont { family: string; sizePt: number; bold: boolean; italic: boolean }
export interface CellTextMeasurer { widthPx(text: string, font: MeasuredFont): number; lineHeightPx(font: MeasuredFont): number }
export function fixedMeasurer(charPx: number, linePx: number): CellTextMeasurer;   // tests
export function autofitRowHeightPt(model: SpreadsheetModel, sheet: Sheet, row: number,
  measurer: CellTextMeasurer, ctx: FormatContext): number;
export function refitRows(model: SpreadsheetModel, sheetId: string, ranges: GridRange[],
  measurer: CellTextMeasurer, ctx: FormatContext): void;
export const AUTOFIT_MAX_ROWS_PER_COMMAND = 10_000;
```
`autofitRowHeightPt` = the tallest cell of the row: one line of its font
unwrapped, greedy word-wrap at (column width − 5 px) when wrapped, the
rotated bounding box for 1–180, characters × line height for 255 —
converted px → pt (× 0.75), rounded up to a 0.75 pt step, never below
`defaultRowHeightPt`. `refitRows` is the automatic refit: for the rows
without `customHeight` that hold a stored non-empty cell in `ranges` (the
first `AUTOFIT_MAX_ROWS_PER_COMMAND` of them) it sets `heightPt`
copy-on-write (absent when it equals the default), never `customHeight`.
This spec runs it inside the undo step of every commit, Fill and Clear
contents; SPEC-124 adds the manual AutoFit (which also clears
`customHeight`), `autofitColumnWidthPx` beside it, and further triggers.
The view supplies `createCanvasCellMeasurer()`
(`src/views/spreadsheet/cellMeasurer.ts`): `CanvasRenderingContext2D
.measureText`, line height from `fontBoundingBoxAscent +
fontBoundingBoxDescent` (1.2 × size where absent), theme fonts resolved
through `core/theme.ts` `resolveFont`.

### The worksheet writer (`xlsx/sheetXml.ts`)
`serializeWorksheet(sheet, base: string | undefined, ctx: XlsxSheetContext):
string` — `base` is the retained part text (the sheet's `partName` in
`opaque`; a copy has its own since SPEC-121), or `undefined` for a new sheet (SPEC-121's empty-sheet template, which this
function reproduces byte for byte for an empty sheet). Rebuilt elements
replace their source counterpart's exact range (located with `scanXml`); an
absent one is inserted before the first source child that comes later in
CT_Worksheet order — sheetPr, dimension, sheetViews, sheetFormatPr, cols,
sheetData, sheetCalcPr, sheetProtection, protectedRanges, scenarios,
autoFilter, sortState, dataConsolidate, customSheetViews, mergeCells,
phoneticPr, conditionalFormatting, dataValidations, hyperlinks,
printOptions, pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks,
customProperties, cellWatches, ignoredErrors, smartTags, drawing,
legacyDrawing, legacyDrawingHF, picture, oleObjects, controls,
webPublishItems, tableParts, extLst.
| Element | Rule |
| --- | --- |
| `dimension` | `ref` = `formatRange(usedRange)`, "A1" when empty |
| `sheetViews` | the first `sheetView`'s `tabSelected`, `showGridLines`, `showFormulas` (SPEC-119), `zoomScale`, `topLeftCell` attributes spliced from the model (others verbatim); its `pane` (frozen: `xSplit`/`ySplit`/`topLeftCell`/`activePane`/`state="frozen"`) and `selection` children regenerated; `pivotSelection`, `extLst` and further `sheetView`s verbatim |
| `sheetFormatPr` | `defaultRowHeight` and `customHeight` spliced; `baseColWidth`, `defaultColWidth`, `x14ac:dyDescent`, outline levels verbatim |
| `cols` | one `col` per `ColumnSpan` (1-based `min`/`max`, `width`, `customWidth="1"` when `widthPx` set, `hidden`, `style`, SPEC-117's `sourceAttrs`); `width` = `sourceWidthChars` when present, else the smallest multiple of 1/256 that SPEC-117's formula maps to `widthPx`; removed when empty |
| `sheetData` | per "Cells" |
| `mergeCells` | `count` + one `mergeCell ref` per merge; removed when none |
**Cells.** `<row r [s customFormat="1"] [ht [customHeight="1"]] [hidden="1"]
{sourceAttrs less spans and x14ac:dyDescent}>` per stored row; `<c r [s] [t]
[cm] {preservedAttrs}>` per
cell (`cm` from SPEC-119's `Cell.cm`), `s` omitted for xf 0, self-closing
when childless; children in order `f`, `v`/`is`, then `preservedXml`.
| Model | `t` | Children |
| --- | --- | --- |
| number | — | `<v>` |
| string | `s` | `<v>` shared-string index |
| boolean | `b` | `<v>1</v>` / `<v>0</v>` |
| error | `e` | `<v>#DIV/0!</v>` |
| formula + number / boolean / error | — / `b` / `e` | `<f>` + `<v>` |
| formula + string | `str` | `<f>` + `<v>` text |
| formula, no value yet | — | `<f>` only |
`<f>` text = `toXlsxFormula(formula)` (`formulaUnparsed`: the stored wire
text verbatim — it only ever reaches its own format, `stripRetention`
having dropped it from cross-format copies); an array anchor writes `<f
t="array" ref="…">`, its members their cached `<v>` only; a dynamic-array
anchor (`dynamicArray`) also keeps its `cm` (an index into the retained
`xl/metadata.xml`) while its formula is unchanged, and becomes a legacy
array formula once edited (the drop rules remove `cm`); a data-table anchor
writes its `dataTableXml` verbatim as the `<f>` child. Shared formulas are never written: each
cell carries its own text. Numbers: `Object.is(n, -0) ? "0" : String(n)`.
Strings: `encodeXlsxText` (in `xlsx/sharedStrings.ts`) maps every character
XML 1.0 cannot carry, and CR, to `_xHHHH_`, and escapes a literal `_xHHHH_`
as `_x005F_xHHHH_` — the inverse of SPEC-116's decoding, property-tested —
then XML-escapes; `<t>` gets `xml:space="preserve"` when the text has
leading/trailing whitespace or a line break.

### Shared strings (`xlsx/sharedStrings.ts`)
`SharedStringsWriter` per save: `indexFor(cell)` returns `cell.sstIndex`
when present; else for a plain string the first source `<si>` that is a lone
`<t>` with the same text (a lookup built once from the source with
`scanXml`), else an earlier append, else a new `<si><t>…</t></si>`; rich runs
(`<r><rPr>` from `TextRun.font` with the font XML below) are appended and
deduplicated among appends. The part is rewritten only when something was
appended: source bytes, appends spliced before `</sst>`, `uniqueCount` =
items, `count` removed (it is optional, and correct only with every clean
part's references counted). Items are never removed.

### Styles (`spreadsheet/styles.ts`, `xlsx/styles.ts`, `ods/styles.ts`)
`internStyle(model, style): number`: the key is the style as JSON with keys
sorted and `sourceXf`/`sourceStyleName` omitted, cached in a `WeakMap` keyed
by the `styles` array; an equal entry's index is returned (a source-backed
one first), else `model.styles = [...model.styles, style]` and the new
index. Entries are immutable once interned.

**xlsx** (`XlsxStyleWriter`): a style with `sourceXf` writes that index;
otherwise its font, fill, border and number format each reuse a source entry
SPEC-117 parses to an equal value, else an earlier append, else are
appended; `numFmtId` is `builtinFormatId` when set (no `numFmt` entry),
else the first built-in id 0–49 outside `LOCALE_DEPENDENT_BUILTINS` whose
code equals `numberFormat`, else an existing custom id with the same code,
else max(163, highest custom) + 1; the xf is `<xf numFmtId fontId fillId borderId xfId>` (the `cellStyleXfs`
index of `namedStyle`, "Normal" → 0) with `applyNumberFormat`/`applyFont`/
`applyFill`/`applyBorder`/`applyAlignment`/`applyProtection="1"` where it
differs from that parent, `quotePrefix="1"`, `<alignment>` and `<protection>`
(`locked="0"` only when false, `hidden="1"` only when true). Fonts in Excel's
child order (`b`, `i`, `strike`, `u`, `vertAlign`, `sz`, `color`, `name`,
`family`, `scheme`; a theme font writes its theme typeface + `<scheme
val="minor|major">`); solid fills `<patternFill patternType="solid"><fgColor
…/><bgColor indexed="64"/>`; colours through SPEC-117's
`modelColorToXlsxAttrs`. Appends are
spliced before each collection's close and its `count` updated; an absent
collection is inserted in CT_Stylesheet order (numFmts, fonts, fills,
borders, cellStyleXfs, cellXfs, cellStyles, dxfs, tableStyles, colors,
extLst); the part is rewritten only when something was appended. From
scratch, font 0 and xf 0 are built from `model.styles[0]`, so unstyled cells
keep the workbook's own default font in either direction.

**ods** (`OdsStyleRegistry`, families added here): a style with
`sourceStyleName` writes that name; otherwise `de-ce<n>` (`style:family=
"table-cell"`, `style:parent-style-name` = `namedStyle` or "Default",
`style:data-style-name` = `de-N<n>` for a non-General code) whose
properties are the exact inverse of SPEC-117's "Reading ODS styles" table,
so a written style reads back equal:
| CellStyle | ODS property |
| --- | --- |
| font family / size / bold / italic | `fo:font-family` (theme fonts resolved), `fo:font-size`, `fo:font-weight="bold"`, `fo:font-style="italic"` |
| underline / strike | `style:text-underline-style="solid"` + `style:text-underline-type="single|double"`; `style:text-line-through-style="solid"` |
| colour / super-, subscript | `fo:color` (resolved hex); `style:text-position="super 58%"` / `"sub 58%"` |
| fill | `fo:background-color` (pattern and gradient fills: their first colour — a documented approximation) |
| borders | `fo:border-top|bottom|left|right` per SPEC-117's pinned ODS border **write** table (thin `0.74pt solid`, medium `1.75pt solid`, thick `2.5pt solid`, hair `0.74pt fine-dashed`, double `2.6pt double`, …) + colour (`#000000` when automatic); diagonals `style:diagonal-tl-br` / `-bl-tr` |
| horizontal | `fo:text-align` start/center/end/justify + `style:text-align-source="fix"`; fill → `style:repeat-content="true"`; general → `style:text-align-source="value-type"` |
| vertical / wrap / indent | `style:vertical-align` top/middle/bottom; `fo:wrap-option="wrap"`; `fo:margin-left` = indent × `INDENT_PX_PER_LEVEL` (9) px, written in pt |
| rotation / shrink | `style:rotation-angle` r for r ≤ 90, 360 − (r − 90) for 91–180, 255 → `style:direction="ttb"`; `style:shrink-to-fit="true"` |
| protection | `style:cell-protect` none (`locked: false`) / formula-hidden (`hidden`, unlocked) / hidden-and-protected (`hidden`) / protected (default, omitted) |
Rich runs use `de-T<n>` text styles; number styles are
`odsNumberStyleXml(code, "de-N<n>")` (its `P0`… sub-styles included),
spliced with the other automatic styles; a `builtinFormatId` style is
written with the code of that built-in in the user's locale
(`formatCodeOf`), since ODS has no locale-dependent ids. Equal property sets
share one name.

### The table writer (`ods/tableXml.ts`)
`serializeTable(sheet, ctx: OdsTableContext): string`, in order:
1. The start tag — the source's with SPEC-121's `table:name` /
   `table:style-name` splices, or generated for a new sheet.
2. Leading children verbatim (`table:title`, `table:desc`,
   `table:table-source`, `office:dde-source`, `table:scenario`,
   `office:forms`, `table:shapes`).
3. Columns: one element per run of identical columns (style =
   `sourceStyleName` or a `de-co` by width, `table:visibility="collapse"`
   when hidden, `table:default-cell-style-name` from `styleId` or
   "Default"), compressed with `table:number-columns-repeated`.
4. Rows `0..rowExtent−1`: each stored row serialised; gaps as empty rows;
   consecutive rows with identical serialisation (empty ones included)
   compressed with `table:number-rows-repeated`; row style
   `sourceStyleName` or a `de-ro` by height; `table:visibility="collapse"`
   when hidden; `table:default-cell-style-name` from `Row.styleId`.
5. Within a row, columns `0..colExtent−1`: stored cells serialised, gaps as
   `<table:table-cell/>`, identical neighbours compressed with
   `table:number-columns-repeated`; a merge anchor gets
   `table:number-columns-spanned`/`-rows-spanned` and the positions it
   covers are `table:covered-table-cell` (carrying content if a cell is
   stored there).
6. The source's grouping skeleton — `table:table-header-rows`,
   `table:table-row-group` (attributes verbatim), `table:table-rows`, and
   the column equivalents — re-emitted around the indices each covered at
   read; `text:soft-page-break` (a layout cache) is dropped.
7. Trailing children verbatim (`table:named-expressions`,
   `calcext:conditional-formats`, anything else after the rows).
Extent = max(the source's total repeated extent, used range + 1) per axis,
so a LibreOffice tail (`table:number-rows-repeated="1048553"`) is re-emitted
at the same length; a new sheet's extent is its used range (at least 1 × 1).
**Cells**: `table:style-name` when the effective style differs from the
row/column default; then value attributes:
| Value | Format kind | `office:value-type` | Value attribute |
| --- | --- | --- | --- |
| number | date (with or without time) | date | `office:date-value` (`serialToOdsDateValue`) |
| number | time, duration | time | `office:time-value` (`serialToOdsTimeValue`) |
| number | percent | percentage | `office:value` |
| number | currency | currency | `office:value` + `office:currency` |
| number | other | float | `office:value` |
| boolean | — | boolean | `office:boolean-value` |
| string | — | string | — |
| error, formula cell | — | float, `office:value="0"` | — |
| error, literal | — | string | — |
plus `calcext:value-type` (the same word, or "error"), `table:formula=
"of:=…"` (`toOdsFormula`; a `formulaUnparsed` formula's stored wire value,
prefix included, verbatim), array anchors'
`table:number-matrix-columns-spanned` / `-rows-spanned`, and
`preservedAttrs`. Content: `preservedXml` first, then `preservedText`
verbatim when present, else one `<text:p>` per line of the SPEC-118
formatted text in the user's locale (informative — consumers display from
the typed value; strings as typed), with runs of spaces as `<text:s
text:c>` (a leading space too), tabs as `<text:tab/>` and rich runs as
`<text:span text:style-name="de-T…">`. A formula cell with no value has no
value type.

### Protection
Read: xlsx `sheetProtection@sheet` "1"/"true" → `Sheet.protection`; ods
`table:protected="true"` → same; `workbookProtection@lockStructure` /
`office:spreadsheet@table:structure-protected` → `structureProtected`.
Locked = the effective style's `protection.locked !== false`. The check runs
before the editor opens and before Delete or Fill (`firstLockedCell`), so
nothing partial is written; formula-bar text is empty for a cell whose
style has `protection.hidden` on a protected sheet.

### calcChain and calcPr
SPEC-121 drops `xl/calcChain.xml`, its relationship and its override on any
edit. `writeXlsx` writes `calcPr@fullCalcOnLoad="1"` (through SPEC-121's
`calcPr` splice) whenever `model.forceFullCalcOnLoad` is set — SPEC-120
sets it on any pass that leaves a stale cell and never clears it — or its
`fullCalcOnLoad` option asks; a source `fullCalcOnLoad="1"` stays. ODS has no equivalent: stale cached values are
written as they are. Dirty sheets are now always written by this spec's
serialisers, so the interim's own `fullCalcOnLoad` rule is gone.

### Reader upgrade
Both readers capture the cells' `sstIndex`, `preservedAttrs`, `preservedXml`
and `preservedText`, `Sheet.protection` and `structureProtected`
(`quotePrefix`, `cm`, `dataTableXml` and the rows' and columns'
`sourceAttrs`/`sourceStyleName` are already read by SPEC-117 and SPEC-119).
The appended `ReaderUpgrade`'s `upgradeDirtySheet` scans the sheet's retained
source and copies these onto each model cell whose value and formula equal
the source's at that address, and sets the protection flags; workbook-level
`structureProtected` is read from the retained workbook part.

### Fixtures
- `spec122-shapes.xlsx` (Excel-shaped): sheet "Edit" with a shared-formula
  block, a legacy array formula, a dynamic-array anchor with `cm="1"` and
  `xl/metadata.xml`, rich and phonetic (`rPh`) shared strings, an
  `_x000D_` string, rows with `spans`, `x14ac:dyDescent` and `outlineLevel`,
  `cols` with `customWidth` and `outlineLevel`, a merge, a frozen pane, a
  `t="dataTable"` formula, an `x14` conditional format in `extLst`,
  hyperlinks, `ignoredErrors`, `pageSetup`; sheet "Locked" with
  `sheetProtection sheet="1"` and one unlocked cell; sheet "Inline" written
  with `t="inlineStr"` like non-Excel producers; `calcChain.xml`; number
  formats 14, 22 and ≥164; theme colours with tints.
- `spec122-shapes.ods` (LibreOffice-shaped): the trailing
  `1048553`-row/`16380`-column tail, repeated non-empty runs, covered cells,
  an `office:annotation` and a `text:a` link inside cells, a
  `table:table-row-group` and `table:table-header-rows`, table-local
  `table:named-expressions` and `calcext:conditional-formats`, a protected
  table with an unprotected cell style, number styles with `style:map`, an
  existing `de-ce1` style (the collision rule), a cell with
  `table:content-validation-name`.

### Tests
- `numberFormat/input.spec.ts`: the rules table × six locales, truncation,
  format steering, `setClock` for the current year. `pointMode.spec.ts`,
  `cells.spec.ts` (identity of untouched rows and cells, sparseness, drop
  rules, protection, arrays, fill shifting), `styles.spec.ts`.
- `xlsx/sheetXml.spec.ts`: goldens per type; splice order with every
  CT_Worksheet sibling present; `encodeXlsxText` round-trip property; arrays
  and `cm`; preserved attributes. `xlsx/sharedStrings.spec.ts`,
  `xlsx/styles.spec.ts` (append-only: the source bytes are unchanged
  outside the inserted runs and `count` values).
- `ods/tableXml.spec.ts` (compression, covered cells, value types, skeleton,
  extent, preserved content), `ods/styles.spec.ts` (`de-` naming past a
  collision, number styles).
- `roundTrip.spec.ts` and `open.spec.ts` per the ACs; view and component
  specs for modes, the key table, IME (`isComposing`), the Tab-run return,
  highlights and the formula bar. e2e per the AC (Playwright's real
  `keyboard.type` — the browser pane's synthetic typing is not a keyboard).

### Specs this changes
- SPEC-116: the readers capture the fields above (reader version bumped).
- SPEC-117: focus moves from the grid element to the editor proxy
  (navigation unchanged); the formula bar (SPEC-119's `formulaBarText`)
  becomes editable; the status bar gains `sheet-status-mode`; rows without
  `customHeight` refit to their content on edit (`autofit.ts`); the rows' and
  columns' `sourceAttrs`/`sourceStyleName` become the writers' input.
- SPEC-118: `numberFormat/input.ts` joins its directory, on SPEC-120's
  `valueFromText`.
- SPEC-121: "Interim: dirty sheets before SPEC-122" is replaced; the
  from-scratch writers serialise cells; the copy's empty-workbook condition
  is lifted; sheet commands honour `structureProtected`;
  `OdsStyleRegistry` gains the `de-ce`, `de-T` and `de-N` families.
- SPEC-125 reuses `fillWithEntry`'s relative shift for Ctrl+D/R and the fill
  handle; Ctrl/Cmd+Enter is owned here.
