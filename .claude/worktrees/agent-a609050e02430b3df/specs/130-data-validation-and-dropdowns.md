---
id: SPEC-130
title: Data validation and in-cell dropdowns
status: draft
area: editing
complexity: normal
story_points: 8
depends_on: [SPEC-120, SPEC-122, SPEC-124, SPEC-125]
---
## Why
Every form somebody builds in a spreadsheet — the expense claim, the order
sheet, the survey export — leans on data validation: a dropdown of allowed
choices, a whole number between 1 and 10, a date no later than today, and a
message that says what to type before anyone types it. Excel writes these as
`<dataValidations>` in each worksheet (with cross-sheet lists in an x14
`<extLst>`), LibreOffice as `table:content-validations` referenced from each
cell by name. After SPEC-122 this product lets people type into those
cells, but it carries the rules as bytes it never reads: no dropdown arrow,
no input message, and any value is accepted — so the next person to open the
file in Excel finds entries the form was built to refuse.

SPEC-122 owns the moment a typed value becomes a cell value; SPEC-120 can
evaluate the formulas a rule is made of. What is missing is the rule model,
a check at that moment, the two pieces of grid chrome (the input-message
tooltip and the dropdown), a dialog to author rules, and both wire formats.

**This file is the normative source for the data-validation model
(`DataValidation`, its types, operators, list sources and messages), the
validation semantics and the stop/warning/information verdicts at entry, the
in-cell dropdown and input message, and the xlsx (`dataValidations`, x14
`dataValidations`) and ODS (`table:content-validations` +
`table:content-validation-name`) wire mappings.**

## Acceptance criteria
- [ ] `src/lib/spreadsheet/validation.ts` exports the model per "The validation model": `Sheet.validations: DataValidation[]` with `ranges`, `type` (`any`, `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`), `operator`, canonical `formula1`/`formula2` relative to the first range's top-left, `allowBlank`, `showDropdown` in the **model's** sense (true = show the arrow), `input { show; title; text }`, `error { show; style: "stop" | "warning" | "information"; title; text }`, and optional `x14`, `sourceXml`, `odsName`, `odsListOrder`; readers, `createBlankWorkbook` and `addSheet` (SPEC-121) initialise `[]`
- [ ] `validateValue(model, sheet, at, candidate, deps)` implements "Validation semantics" exactly — whole numbers without a fraction, decimals, dates and times as serials compared as numbers, text length of the value's text, `between`/`notBetween` with bounds in either order, `custom` evaluating `formula1` at the target with the candidate overlaid on that cell, blanks governed by `allowBlank`, and a bound or list source that cannot be evaluated accepting the entry rather than blocking it — returning `{ ok: true }` or a verdict naming the rule
- [ ] List sources per "List sources": a literal list (`formula1` is the string literal `"a,b,c"`, items split on commas and trimmed, at most `LITERAL_LIST_MAX_CHARS = 255` characters), a range on the same sheet, a range on another sheet, a defined name and any formula whose result is one row or one column; `listItems` returns `{ value, text }` items in source order (blank source cells as empty items, duplicates kept, capped at `VALIDATION_LIST_MAX = 32_767`) or `unavailable`; a typed entry matches an item case-insensitively by text, or numerically when both are numbers
- [ ] Entry is checked on commit through SPEC-122's commit path — the cell editor's Enter/Tab/click-away, the formula bar's commit, a dropdown pick, and Ctrl/Cmd+Enter (SPEC-122's fill-the-selection entry, each target checked against its own rule, the first failure deciding) — before the set-cell command runs; a formula entry is checked against its evaluated result; an invalid entry under a rule whose `error.show` is false is committed silently (Excel)
- [ ] An invalid entry opens `validation-alert` (`role="alertdialog"`, `AppIcon` `error`/`warning`/`info`) titled `error.title` or "Invalid entry", reading `error.text` or "This value doesn't match the data validation rules for this cell.": **stop** → `validation-alert-retry` ("Retry": back into the editor with the entered text selected) and `validation-alert-cancel` ("Cancel": the old value stays); **warning** → the text plus "Continue?" with `validation-alert-yes` (commits), `validation-alert-no` (back to editing), `validation-alert-cancel`; **information** → `validation-alert-ok` (commits) and `validation-alert-cancel`
- [ ] Nothing else validates — paste (SPEC-125; Excel does not validate pasted values either), fill, drag-move, Delete/Backspace clearing, undo/redo, sort, CSV import and recalculated formula results all write without a check, and `validation.ts`'s module header states it; a copy-paste does not change which rules cover the destination cells (copying rules with cells is Backlog), while a cut-paste or drag-move carries the moved cells' rules per "Moves"
- [ ] While the active cell is covered by a rule with `input.show` and a non-empty title or text (and the sheet's `validationPromptsDisabled` is unset), `validation-input-message` (`role="tooltip"`, title in bold above the text, `\n` as line breaks) sits below-right of the cell, flips to stay inside the grid viewport, follows the cell on scroll and zoom, stays up while editing, hides while the dropdown list is open, and is referenced by the active `sheet-cell`'s `aria-describedby`
- [ ] When the active cell's rule is a `list` with `showDropdown`, a `validation-dropdown` button (`aria-label="Choose from list"`, `aria-haspopup="listbox"`, `aria-expanded`) sits against the cell's right edge; click or Alt+Down opens `validation-list` (on an autofilter header cell Alt+Down stays SPEC-128's filter popover, as in Excel; the arrow button is then the path) (`role="listbox"`, at least the cell's width, eight rows visible then scrolling) of `validation-option`s with the current value's item selected (none when the value is not in the list, the first ArrowDown then selecting the first item); Up/Down, Home/End and PageUp/PageDown (eight) move, a typed character jumps to the next item starting with it, Enter or Tab picks and commits through the commit path, a click picks, and Escape closes and returns focus to the grid; an unavailable source shows one disabled `validation-list-unavailable` option "The list source isn't available"
- [ ] A picked item is entered as its value: a range source's cell value as stored (a number stays a number), a literal item through SPEC-122's `parseInput` in the user's locale (so `10` becomes the number 10)
- [ ] Data ▸ "Data validation…" (`sheet-menu-data-validation`, in SPEC-120's Data menu after a separator, below SPEC-128's filter items) opens `validation-dialog` with three tabs (`validation-tab-criteria`, `validation-tab-input`, `validation-tab-error`) and the controls and labels in "The dialog", loading the active cell's rule; `validation-mixed-notice` reads "The selection contains cells with different validation settings — OK applies these settings to all of them" when the selection's cells disagree; OK applies to the whole selection, "Clear All" (`validation-clear-all`) resets the fields so OK removes the rules, and the dialog's own errors (pinned strings) show in `validation-dialog-error` and save nothing
- [ ] Commands `setValidation(model, sheetId, ranges, rule)` and `clearValidation(model, sheetId, ranges)` subtract the target ranges from every other rule with `subtractRanges(ranges, cut): GridRange[]` (SPEC-129 adds it to `address.ts`; this spec does not depend on SPEC-129, so if that spec has not landed it adds the function with exactly this contract; a rule left with no range is removed), add the target to an existing rule whose settings are identical (ignoring `id`, `ranges`, `sourceXml`, `odsName`) or else create one, replace what they touch copy-on-write, drop touched rules' `sourceXml`, set `Sheet.dirty`, and record one undo entry labelled "Data validation"
- [ ] xlsx read and write per "xlsx wire mapping": every `dataValidation` attribute in the table with Excel's defaults, `type="none"` ↔ `any`, **`showDropDown="1"` ↔ `showDropdown: false`** (asserted both ways), `<formula1>`/`<formula2>` through `fromXlsxFormula`/`toXlsxFormula`, `dataValidations@disablePrompts` ↔ `Sheet.validationPromptsDisabled`, the block rewritten at its CT_Worksheet position (after `conditionalFormatting`, before `hyperlinks`) with `count` updated and the start tag's other attributes kept, and each rule emitted from its `sourceXml` while unchanged
- [ ] x14 per "x14": `x14:dataValidation`s in the worksheet ext `uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"` (formulas in `xm:f`, ranges in `xm:sqref`) are read with `x14: true`; on write a rule whose formulas reference another sheet goes to the x14 ext (Excel does the same), every other rule to the base block, and read rules keep their location while unchanged; only this ext is rewritten, created in Excel's ext order after SPEC-129's `{78C0D931…}`
- [ ] ODS read per "ODS wire mapping": `table:content-validations` from `odsContentEnvelope.prefix`; each `table:condition` form in the table (with `of:` or LibreOffice's older `oooc:` prefix), `table:allow-empty-cell`, `table:display-list`, `table:base-cell-address` (formulas rebased onto each sheet's origin), `table:help-message` and `table:error-message` (`table:message-type`, `table:display`, `text:p` paragraphs joined with `"\n"`); ranges are accumulated from `table:content-validation-name` on cells as row-run × column-run rectangles **without materialising repeated empty cells** — a rule over LibreOffice's trailing 1,048,553-row run costs one rectangle; a `table:error-macro` is kept and never run, and its rule enforces nothing
- [ ] ODS write: SPEC-122's table serialiser derives each emitted cell's `table:content-validation-name` from `Sheet.validations` (the first rule covering it) and includes it in the repeat-compression key, emitting empty validated cells as attribute-only runs; `table:content-validations` is regenerated in the prefix whenever any sheet is dirty or `structureDirty` — unchanged rules keep their `odsName` and `sourceXml`, others are serialised under a fresh `de-val<n>` that collides with no name in use, a rule shared by two clean sheets is written once, and validations no cell references (`SpreadsheetModel.odsOrphanValidations`) are re-emitted verbatim
- [ ] Round trip: an untouched workbook downloads byte-identical part for part (`equalBytes`); set → save → reopen reproduces `validations` under `toEqual` minus minted ids and `sourceXml`, in both formats; "Download a copy as…" maps every type both ways through the model (an x14 rule into the ODS form; an ODS `sort-ascending` list into an ordinary xlsx list, the one property that does not cross), and SPEC-121's `UNCONVERTIBLE_CONTENT` loses its "data validation" row
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`: a clean sheet is re-read; for a dirty sheet the step parses only its validations — xlsx from the retained part (base block and x14 ext), ODS from the retained prefix's `table:content-validations` plus the names on the sheet's retained `sourceXml` — the retained source has carried SPEC-124's and SPEC-125's retained-reference rewrites since it was read, so the parsed positions already match the current model — tested with a synthetic previous-version record holding one dirty and one clean sheet
- [ ] Registers `{ id: "validations", adjust }` in `REFERENCE_ADJUSTERS` per "Structural edits" and, in the same change, retires its rows from SPEC-124's "Retained references" tables (`dataValidation@sqref`, `formula1`, `formula2`; `x14:dataValidation/xm:sqref`, `x14:formula1|2/xm:f`; `table:content-validation@table:base-cell-address` and `@table:condition`): inserting two rows above a rule on `B5:B20` moves it to `B7:B22`, its custom formula `B5>$Z$1` to `B7>$Z$1` and a list source `$E$2:$E$6` below the insertion to `$E$4:$E$8`; deleting the rule's first row re-anchors its formulas with `adjustAnchoredFormula` rather than producing `#REF!` — the test joins `structure.spec.ts`'s registry contract test; sheet rename and delete run SPEC-121's `remapSheetReferences` over every rule formula
- [ ] Registers `{ id: "validations", move }` in SPEC-125's `MOVE_ADJUSTERS` per "Moves": cutting `B2` (a list rule on `B2:B10`) and pasting it at `D2` leaves the rule on `B3:B10` and creates the same rule on `D2`, while the rule that covered `D2` before loses it; a rule wholly inside the source keeps its id and moves; the list source `$E$2:$E$6` follows only if those cells were moved — the test joins `move.spec.ts`'s registry contract test
- [ ] Fixtures from `spreadsheet/fixtures/build.ts`: `e2e/fixtures/spec130-validation.xlsx` and `spec130-validation.ods` per "Fixtures", emulating Excel's and LibreOffice's real shapes (x14 cross-sheet list, `showDropDown="1"`, `xr:uid`, `imeMode`; `of:` conditions, a validation over a repeated empty-cell run, a sorted list, an orphan, an error macro)
- [ ] Tests: `spreadsheet/validation.spec.ts` (every type × operator, blanks, lists of every source, unresolvable sources, `custom` with the overlay, commands and merging of identical rules), `xlsx/dataValidations.spec.ts`, `ods/contentValidations.spec.ts`, the adjuster test in `structure.spec.ts`, the move adjuster test in `move.spec.ts`, the upgrade step in `open.spec.ts`, `roundTrip.spec.ts` additions, `views/spreadsheet/useDataValidation.spec.ts`, and component specs for `ValidationDialog`, `ValidationAlert` and `ValidationDropdown` (keyboard paths)
- [ ] e2e (`e2e/SPEC-130.spec.ts`): open `spec130-validation.xlsx` → select `B2` → `validation-input-message` reads "Priority" / "Pick one" → Alt+Down → `validation-list` lists Low, Medium, High → ArrowDown, ArrowDown, Enter → `B2` reads "Medium" → type "Urgent" and Enter → `validation-alert` (stop) → Retry → the editor holds "Urgent" selected → type "High" and Enter → `B2` reads "High" → type 42 into `B3` (whole 1–10, warning) → "Continue?" → No → Escape → `B3` unchanged → select `B8` → no `validation-dropdown` (its file says `showDropDown="1"`) → select `D2` → Data ▸ Data validation… → List, Source "Yes,No" → OK → `D2` shows `validation-dropdown` → download → reopen → `D2` still offers Yes, No; then open `spec130-validation.ods` → select `C500` → its input message shows

## Out of scope
"Circle invalid data" and "Clear validation circles" — Backlog. Copying
validation rules with cells (Excel's paste-all and Paste Special ▸
Validation — a cut-paste or drag-move does carry them, per "Moves") and
"Apply these changes to all other cells with the same settings" — Backlog. Dependent dropdowns beyond what an `INDIRECT` list
formula already gives through SPEC-120. Running a LibreOffice error macro
(never — macros are bytes this product carries). Input-method modes
(`imeMode`, preserved). Searchable or multi-select dropdowns (neither format
has them). Validating values that arrive any way other than typed entry (see
the AC). Locale-specific list separators in the Source field — it always
splits on commas, consistent with the invariant formula syntax.

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — plain-JSON model, copy-on-write
  rows, sparse storage, the retained source as the splice base (xlsx parts
  in `opaque`; ODS `sourceXml` and `odsContentEnvelope { prefix; suffix }`,
  whose prefix holds `table:content-validations`); `GridRange`, `CellRef`,
  `parseSqref`, `formatRange`, `getCell`; the ODS reader's rule never to
  materialise repeated empty rows or cells; `ReaderUpgrade`,
  `READER_UPGRADES`.
- SPEC-117 — `SheetSelection { active; ranges; activeRangeIndex }`,
  `GridMetrics` (`colLeft`, `colWidth`, `rowTop`, `rowHeight`) to place the
  tooltip and the arrow, `sheet-cell` with `data-ref`, zoom.
- SPEC-118 — `formatValue(value, code, ctx)` for item text;
  `SpreadsheetLocale`; formulas always in the invariant en-US syntax.
- SPEC-119 — canonical formula text; `parseFormula` (throws
  `FormulaSyntaxError`); `parseFormulaAt(text, host): { node; origin } |
  FormulaSyntaxError`; `shiftFormula`;
  `fromXlsxFormula`/`toXlsxFormula`/`fromOdsFormula`/`toOdsFormula`.
- SPEC-120 — `evaluate(node, ctx: EvalContext)` resolving relative parts as
  `part + (host − origin)`; `EvalContext { model; host; origin; arrayMode;
  settings; now(); random(); valueAt(sheetId, row, col); recordDependency;
  resolveName? }`; `EvalValue` (incl. `RangeValue`/`ArrayValue`),
  `Uncomputable`; comparison and text coercion; an unknown name is
  uncomputable until SPEC-133 supplies `resolveName`.
- SPEC-121 — `useSpreadsheetUndo` `{ record(label, scope) /* before */,
  commit(alsoRecalc?) /* after */, … }` with `SHEET_BY_REFERENCE_FIELDS`
  (`rows`, `sourceXml`) kept by reference and other sheet fields
  `structuredClone`d; `Sheet.dirty`, `structureDirty`;
  `writeXlsx`/`writeOds`; `stripRetention` and `UNCONVERTIBLE_CONTENT`;
  `remapSheetReferences(canonical, remap)`; `createBlankWorkbook`/`addSheet`.
- SPEC-122 — the commit path from the cell editor (`cell-editor`) and
  formula bar to the set-cell command in `cells.ts`, Ctrl/Cmd+Enter's
  fill-the-selection entry; `parseInput(text, locale, dateSystem, options?):
  { value; formula?; numberFormat?; builtinFormatId?; quotePrefix? }`; the CT_Worksheet
  splice order (… `conditionalFormatting`, **`dataValidations`**,
  `hyperlinks` … `extLst`); ODS table re-serialisation with repeat
  compression; protection honoured.
- SPEC-124 — `StructuralEdit`, `adjustRange(range, edit): GridRange | null`,
  `adjustFormula(text, hostSheetId, edit, sheetNameOf)`,
  `REFERENCE_ADJUSTERS: { id; adjust(model, edit): void }[]` (adjusters
  replace, never mutate) and `rewriteRetainedReferences`, whose
  data-validation rows this spec retires.
- SPEC-125 — `src/lib/spreadsheet/move.ts`: `RangeMove { fromSheetId; from:
  GridRange; toSheetId; toTop; toLeft }`, `MoveAdjuster { id; move(model,
  move): void }`, `MOVE_ADJUSTERS` (registered: `cells`, `formulas`,
  `merges`, `retained`), `moveRange`, `repointFormula(text, hostSheetId,
  move, sheetNameOf)` (a reference entirely inside the source follows the
  move; one entirely inside the destination becomes `#REF!`; anything else
  is unchanged) and `moveMapper`; `move.spec.ts` holds the contract test
  later move adjusters join.
- `src/lib/spreadsheet/ods/rangeAddress.ts` (SPEC-121) —
  `odsRangeAddressToCanonical(text): string | null` /
  `canonicalToOdsRangeAddress(canonical): string | null` for
  `table:base-cell-address`.
- SPEC-120's Data menu (`menu-title-data`) in `SpreadsheetMenuBar.vue`.
- `src/components/BaseDialog.vue` (dialog shell), `AppIcon` (`error`,
  `warning`, `info` exist), `MenuItem.vue` (`testid`), `core/platform.ts`.

### The validation model (`src/lib/spreadsheet/validation.ts`)
```ts
export type ValidationType = "any" | "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
export type ValidationOperator = "between" | "notBetween" | "equal" | "notEqual"
  | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
export interface DataValidation {
  id: string;                     // minted at read; never written
  ranges: GridRange[];            // the origin is ranges[0]'s top-left
  type: ValidationType;
  operator: ValidationOperator;   // used by whole/decimal/date/time/textLength
  formula1?: string;              // canonical, relative to the origin
  formula2?: string;
  allowBlank: boolean;
  showDropdown: boolean;          // true = show the arrow (xlsx showDropDown="1" means HIDE)
  input: { show: boolean; title: string; text: string };
  error: { show: boolean; style: "stop" | "warning" | "information"; title: string; text: string };
  x14?: true;
  sourceXml?: string;             // verbatim source element while unchanged
  odsName?: string;               // table:content-validation@table:name
  odsListOrder?: "sort-ascending";
}
// Sheet gains: validations: DataValidation[]; validationPromptsDisabled?: true
// SpreadsheetModel gains: odsOrphanValidations?: string[]
export const VALIDATION_LIST_MAX = 32_767;
export const LITERAL_LIST_MAX_CHARS = 255;
export type ListSource = { kind: "literal"; items: string[] } | { kind: "reference" | "name" | "formula"; formula: string };
export function listSourceOf(dv: DataValidation): ListSource | null;
export function parseListLiteral(formula1: string): string[] | null;   // "\"a,b,c\"" → ["a","b","c"]
export function formatListLiteral(items: string[]): string;            // quotes doubled
export function validationAt(sheet: Sheet, ref: CellRef): DataValidation | undefined;   // first covering rule
export interface ValidationDeps {
  base: Omit<EvalContext, "host" | "origin" | "arrayMode">;   // valueAt reads stored values
  ctx: FormatContext;
}
export function listItems(model: SpreadsheetModel, sheet: Sheet, dv: DataValidation, at: CellRef,
  deps: ValidationDeps): { items: { value: CellValue; text: string }[]; unavailable?: true };
export type ValidationVerdict = { ok: true } | { ok: false; rule: DataValidation };
export function validateValue(model: SpreadsheetModel, sheet: Sheet, at: CellRef,
  candidate: CellValue, deps: ValidationDeps): ValidationVerdict;
```
Defaults for a rule made in the dialog (Excel's): `allowBlank: true`,
`showDropdown: true`, `input.show: true`, `error.show: true`, `style:
"stop"`, `operator: "between"`. A rule's formulas are parsed once with
`parseFormulaAt(formula, origin)` and evaluated at cell `(r, c)` by
`evaluate` with `host` = that cell and `origin` = the rule's origin (the
engine's `part + (host − origin)` rule); `custom` wraps `base.valueAt` so
the target cell answers the candidate. Overlapping rules (files can hold
them; the dialog never makes them): the first in model order wins, for
enforcement and for the ODS cell attribute.

### Validation semantics
Blank first: a candidate `null` (an empty commit) or `""` is valid iff
`allowBlank`. Bounds `f1`/`f2` are evaluated at the target; blank = 0; a
bound that is an error or not a number accepts the entry. `between` uses
`[min(f1, f2), max(f1, f2)]`.

| Type | Valid when (`v` = the candidate) |
| --- | --- |
| `any` | always |
| `whole` | `v` is a number, `v === Math.trunc(v)`, and `v` satisfies the operator |
| `decimal`, `date`, `time` | `v` is a number satisfying the operator (dates and times are serials; text never passes) |
| `textLength` | the length (UTF-16 code units, as Excel's `LEN`) of `v`'s text satisfies the operator |
| `list` | some item matches: both numbers and numerically equal (15 significant digits), or texts equal case-insensitively; an unavailable source accepts |
| `custom` | `formula1` evaluated with `v` overlaid on the target cell is `TRUE` or a non-zero number |

### List sources
`formula1` of a `list` rule is either a string literal (`"Low,Medium,High"`
— `parseListLiteral` splits on `,` and trims; an empty item is kept) or any
other formula: a reference (`$E$2:$E$6`, `Lists!$A$1:$A$4`), a name
(`Colours`) or a formula (`OFFSET(…)`, `INDIRECT(…)`). The latter is
evaluated at the target: a range or array that is one row or one column
yields its cells in order (text by `formatValue` with each source cell's
format); a 2-D result, an error or an `Uncomputable` (an unknown name until
SPEC-133 supplies `resolveName`) is `unavailable`. An ODS literal item that
itself contains a comma becomes two items once the rule is edited (an
unedited rule re-emits its `sourceXml`); the xlsx form cannot hold it either.

### Entry gate
`src/views/spreadsheet/useDataValidation.ts` exposes `checkEntry(targets:
CellRef[], candidate: { value: CellValue; formula?: string }): Promise<"commit"
| "retry" | "cancel">`; this spec makes `useCellEditing`'s commit (SPEC-122,
which has no pre-commit hook) await it before the set-cell command. `"retry"` reopens the editor with the
entered text selected; `"cancel"` closes it leaving the cell unchanged. A
formula candidate's value is obtained by evaluating it at the target first.
Protection (SPEC-122) is checked before validation.

### The dialog
`src/components/spreadsheet/ValidationDialog.vue` on `BaseDialog` (title
"Data validation"), state in `src/views/spreadsheet/useValidationDialog.ts`:
- **Criteria.** `validation-allow`: "Any value", "Whole number",
  "Decimal", "List", "Date", "Time", "Text length", "Custom".
  `validation-operator` (whole/decimal/date/time/text length): "between",
  "not between", "equal to", "not equal to", "greater than", "less than",
  "greater than or equal to", "less than or equal to". `validation-value-1`
  / `validation-value-2` labelled "Minimum"/"Maximum" (between; "Start
  date"/"End date", "Start time"/"End time" for those types), "Value" /
  "Date" / "Time" / "Length" (equal, not equal), "Minimum" (greater…),
  "Maximum" (less…), "Source" (list), "Formula" (custom).
  `validation-ignore-blank` "Ignore blank"; `validation-in-cell-dropdown`
  "In-cell dropdown" (list only).
- **Input message.** `validation-show-input` "Show input message when cell
  is selected", `validation-input-title` "Title" (max 32 characters),
  `validation-input-text` "Input message" (max 255).
- **Error alert.** `validation-show-error` "Show error alert after invalid
  data is entered", `validation-error-style` "Style" (Stop, Warning,
  Information), `validation-error-title` "Title" (32), `validation-error-text`
  "Error message" (225).
- Buttons `validation-clear-all` "Clear All", `validation-cancel` "Cancel",
  `validation-ok` "OK".

A value starting with `=` is a formula (validated with `parseFormula`);
otherwise it goes through `parseInput` and a number, date or time is stored
as invariant number text; a list Source without `=` becomes a literal.
Dialog errors, verbatim: "The formula contains an error", "The Maximum must
be greater than or equal to the Minimum", "Enter a whole number of 0 or
more", "The list source can't be longer than 255 characters", "The list
source must be a delimited list, or a reference to a single row or column".
OK with "Any value" and no input message removes the rules.

### xlsx wire mapping
| `dataValidation` attribute (default) | Model |
| --- | --- |
| `type` (`none`) | `type` (`none` ↔ `any`) |
| `operator` (`between`) | `operator` |
| `allowBlank` (0) | `allowBlank` |
| `showDropDown` (0) | `showDropdown = !showDropDown` |
| `showInputMessage` (0), `promptTitle`, `prompt` | `input` |
| `showErrorMessage` (0), `errorStyle` (`stop`), `errorTitle`, `error` | `error` |
| `sqref` | `ranges` |
| `imeMode`, `xr:uid` | `sourceXml` only |
| `<formula1>`, `<formula2>` | `formula1`, `formula2` |

Attribute text is XML-escaped (`&#10;` for line breaks). The block is
`<dataValidations count="n" [disablePrompts xWindow yWindow]>`.

### x14
The ext `uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"` (declaring
`xmlns:x14` on itself) holds `<x14:dataValidations count xmlns:xm>` ›
`<x14:dataValidation type allowBlank showInputMessage showErrorMessage
errorStyle errorTitle error promptTitle prompt showDropDown operator>` ›
`<x14:formula1><xm:f>Lists!$A$1:$A$4</xm:f></x14:formula1>`,
`<x14:formula2>`, `<xm:sqref>B9</xm:sqref>`. Same attribute mapping as the
base table. Ext order: `{78C0D931-6437-407d-A8EE-F0AAD7539E65}` (SPEC-129)
first, then this, then any others as found.

### ODS wire mapping
`table:content-validation @table:name @table:condition @table:allow-empty-cell
(true) @table:display-list (unsorted) @table:base-cell-address`:

| Model | `table:condition` |
| --- | --- |
| `any` | absent |
| `whole` / `decimal` / `date` / `time` + `between` | `of:cell-content-is-whole-number() and cell-content-is-between(a;b)` (`…-is-decimal-number()`, `…-is-date()`, `…-is-time()`) |
| … + `notBetween` | `… and cell-content-is-not-between(a;b)` |
| … + comparison | `… and cell-content()>=a` (`=`, `!=`, `<`, `>`, `<=`, `>=`) |
| `textLength` | `of:cell-content-text-length()<=a`, `of:cell-content-text-length-is-between(a;b)`, `…-is-not-between(a;b)` |
| `list` literal | `of:cell-content-is-in-list("a";"b";"c")` |
| `list` other | `of:cell-content-is-in-list(<formula>)` |
| `custom` | `of:is-true-formula(<formula>)` |

Operands are OpenFormula (`fromOdsFormula`/`toOdsFormula`), rebased from
`table:base-cell-address` onto each sheet's origin with `shiftFormula` and
written with the origin as the base. `table:display-list`: `none` ↔
`showDropdown: false`, `unsorted` ↔ `true`, `sort-ascending` ↔ `true` +
`odsListOrder` (the dropdown then sorts its items ascending: numbers
numerically, before text ordered by `Intl.Collator(locale.tag)`). `table:help-message @table:title @table:display
(false)` ↔ `input`; `table:error-message @table:title @table:display (false)
@table:message-type (stop)` ↔ `error`; `table:error-macro` → `sourceXml`
only and `error.show: false`. A rule referenced from two sheets becomes one
`DataValidation` per sheet with the same `odsName`. New
`table:content-validations` goes after `table:calculation-settings`, before
`table:label-ranges` and the first `table:table`.

### Structural edits
Ranges map through `adjustRange`; `formula1`/`formula2` go through
`adjustAnchoredFormula(text, origin, firstRange, hostSheetId, edit,
sheetNameOf)` in `formula/adjust.ts` (SPEC-129 adds it; if that spec has
not landed — this one does not depend on it — this spec adds it with exactly
this contract): when the edit deletes the origin's row or column, the formula
is first rebased with `shiftFormula` onto the first surviving cell of the
first surviving range (pre-edit coordinates), then `adjustFormula` applies.
A rule with no range left is removed; changed rules lose `sourceXml`.
Duplicating a sheet copies its rules with fresh ids and no `odsName`;
renaming or deleting a sheet runs `remapSheetReferences` over every rule
formula (a deleted sheet's references become `#REF!`).

### Moves
The `validations` move adjuster (cut-paste and drag-move — rules travel
with their cells, as in Excel), for a `RangeMove` with offset `(dr, dc) =
(toTop − from.top, toLeft − from.left)`:
1. Every rule on `toSheetId` loses the destination rectangle
   (`subtractRanges`) — the overwritten cells' rules go.
2. For every rule on `fromSheetId`, the part of its ranges inside
   `move.from` is cut out. A rule wholly inside moves: ranges offset onto
   `toSheetId`, same id. Otherwise the cut-out pieces become a new rule
   on `toSheetId` (fresh ids), whose formulas are the source's first rebased
   onto the pieces' top-left with `shiftFormula` (so each piece means what
   it meant in place) — and a source left without its first range's
   top-left is re-anchored the same way onto its new origin. A moved
   rule is never merged into an identical one at the destination.
3. Every formula of every rule on every sheet then goes through
   `repointFormula`, exactly as cell formulas do — references to moved
   cells follow them, nothing else shifts. Across sheets, a moved
   rule's unqualified references are first qualified with the source
   sheet's name, so they keep meaning the same cells.
Touched rules are replaced (never mutated) and lose `sourceXml`.

### Fixtures
`spec130-validation.xlsx`, sheet "Form" (`B3` holds 5, the rest of
`B2:B10` is empty): `B2` list `"Low,Medium,High"` with `promptTitle="Priority" prompt="Pick
one"` and a stop alert; `B3` whole between 1 and 10, warning; `B4` decimal
> 0, information; `B5` date between two serials; `B6` text length ≤ 5; `B7`
custom `ISNUMBER(SEARCH("@",B7))`; `B8` list `$E$2:$E$6` with
`showDropDown="1"`; `B9` list `Lists!$A$1:$A$4` in the x14 ext; `B10` list
`Colours` (a workbook defined name); `xr:uid` and `imeMode="noControl"`
attributes; sheet "Lists".
`spec130-validation.ods` mirrors it LibreOffice-style, plus a rule over
`C2:C1000` carried by a `table:number-rows-repeated` run of empty cells,
`table:display-list="sort-ascending"` on one list, a two-paragraph help
message, an orphan `table:content-validation` and one with
`table:error-macro`.

### Tests
Per the Tests AC; entry-gate tests drive `checkEntry` with each style and
button; round-trip tests assert `equalBytes` per part for untouched
fixtures and model `toEqual` after set → save → reopen; the ODS suite
asserts the repeated-run rule is one rectangle and that no `Cell` was
created for it.

### Specs this changes
- SPEC-117 — the grid hosts `ValidationInputMessage.vue` and
  `ValidationDropdown.vue` beside the active-cell outline; the active cell
  gains `aria-describedby`.
- SPEC-121 — `SHEET_BY_REFERENCE_FIELDS` gains `validations` (commands
  replace the array); duplicate, delete and rename of sheets cover rule
  formulas; `stripRetention` drops `sourceXml`, `odsName` and
  `odsOrphanValidations`; `UNCONVERTIBLE_CONTENT` loses its "data
  validation" row.
- SPEC-122 — the commit path awaits the entry gate (Ctrl/Cmd+Enter checked
  per target); `dataValidations` joins
  the rebuilt set; the ODS table writer emits
  `table:content-validation-name` and compresses runs with it.
- SPEC-124 — the data-validation rows of "Retained references" are retired
  (the modelled path writes the same shifted `sqref` and
  `table:base-cell-address`).
- SPEC-125 — `MOVE_ADJUSTERS` gains `validations`; paste, fill and move
  write without validation.
- SPEC-120 — its Data menu gains "Data validation…" (after SPEC-128's
  filter items).
- SPEC-128 — sorting leaves rules where they are (they are range-based).
- SPEC-133 — defined-name list sources resolve through its names model.
