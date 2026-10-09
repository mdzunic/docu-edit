---
id: SPEC-120
title: Calculation — the engine, recalculation and the core functions
status: draft
area: platform
complexity: complex
story_points: 21
depends_on: [SPEC-118, SPEC-119]
---
## Why
After SPEC-119 a workbook knows its formulas but cannot compute one. Every
value on screen is whatever the producer cached when it last saved: a file
written by a library that never computes shows zeros, a sheet that says
`=TODAY()` shows the day it was saved, and from SPEC-122 on, typing a number
into a cell would leave every total that depends on it wrong. An editable
spreadsheet without a calculation engine is a liability, not a product.

There is no engine to borrow. `package.json` is frozen and the
Content-Security-Policy in `index.html` has no `'unsafe-eval'`
(`src/cspPolicy.spec.ts` pins that), so there is no formula library and no
compiling a formula to JavaScript either. The engine is an AST interpreter
over SPEC-119's `FormulaNode`, written here.

The hard part is not arithmetic; it is semantics. Excel coerces `"1"+"2"` to
3 but `"TRUE"+0` to `#VALUE!`; treats an empty cell as 0, `""` and FALSE at
once; says `-2^2` is 4 and `0.1+0.2=0.3` is TRUE; intersects a range with the
formula's own row when a single value is expected; and decides what to
recompute from a dependency graph, not by recomputing everything. A workbook
whose cached values the engine contradicts is a workbook the user stops
trusting. So this spec pins the semantics first, with a conformance table,
and the functions second.

**This file is the normative source for value semantics and coercion,
comparison and error propagation, ranges, arrays and implicit intersection,
the dependency graph (`DependencyGraph`), incremental recalculation
(`recalculate`), cycles, volatile functions, the open-time calculation
policy, stale (uncomputable) cells and the `forceFullCalcOnLoad` flag, the
function registry shape (`FunctionSpec`), and the 100 core functions.**
SPEC-127 adds the rest of the library through the same registry; SPEC-133
plugs in names and structured references; SPEC-137 may move evaluation into
a worker without changing any contract here.

## Acceptance criteria
- [ ] `formula/values.ts` exports the value model per "Values": `EvalValue = number | string | boolean | CellError | EmptyValue | RangeValue | ArrayValue`, the `EMPTY` singleton, `Area`, the coercions `toNumber`/`toText`/`toLogical`, `valueFromText`, `numberToText`, `compareScalars`, `roundTo15`, `toCellValue`, and the bounds `MAX_ARRAY_CELLS = 4_194_304` and `MAX_TEXT_LENGTH = 32_767`
- [ ] Coercion follows "Coercion" in every context — number, text and logical — and every row of the "Conformance" table holds in `formula/conformance.spec.ts`, each row labelled with its source per "Where expected values come from"
- [ ] Text→number coercion is **locale-aware**: `valueFromText(text, locale, dateSystem)` implements the grammar in "Text to number" over the user's `SpreadsheetLocale` (SPEC-118), so `="1,5"+0` is 1.5 under `de-DE` and `#VALUE!` under `en-US`, `="1,000"+0` is 1000 under `en-US`, and `="2026-10-08"+0` is 46303 everywhere (1900 system); booleans are never parsed from text and `""` is `#VALUE!`; SPEC-122's `parseInput` calls this function for every value it recognises, so a typed value and a coerced one never disagree
- [ ] Comparison follows "Comparison": type order number < text < logical, an empty operand takes the other side's type (0, `""`, FALSE), text compares case-insensitively through `Intl.Collator(locale.tag, { sensitivity: "accent" })` (case-sensitively for an `.ods` whose settings say so), numbers compare after rounding both to 15 significant digits (`=0.1+0.2=0.3` is TRUE), and the root-level `+`/`−` "final operation" rule returns 0 for an exact-at-15-digits cancellation (`=0.1+0.2-0.3` is 0)
- [ ] Errors propagate per "Errors": a binary operator returns its left operand's error first; a function returns the first error among its value arguments in argument order (row-major within a range) unless it declares `acceptsErrors`; IF, IFS, SWITCH, CHOOSE, IFERROR and IFNA evaluate only the branch they take; a non-finite numeric result is `#NUM!` and a text result longer than `MAX_TEXT_LENGTH` is `#VALUE!` — no NaN, Infinity or `-0` ever reaches `Cell.value`
- [ ] References evaluate per "Ranges and areas": a `RangeValue` holds one `Area` per sheet of a 3-D reference; iterating a range walks only the **stored** cells (binary search to the first row, then forward), so `=SUM(A:A)` over a column of 1,000 values touches 1,000 cells, not 1,048,576; an unknown sheet is `#REF!`, an empty intersection `#NULL!`, a range operator across sheets `#REF!`, and materialising more than `MAX_ARRAY_CELLS` elements `#NUM!`
- [ ] Implicit intersection follows "Implicit intersection": in a non-array formula a range in a value position yields the cell in the host's row (single column) or column (single row), `#VALUE!` otherwise; an array in a value position yields its top-left element; a top-level range or array result is reduced the same way and a top-level `EMPTY` is stored as 0
- [ ] Legacy array (CSE) formulas evaluate per "Array mode": the anchor's formula runs with array semantics (no intersection, value-kind arguments lift element-wise, single-row/single-column arrays broadcast, out-of-bounds elements are `#N/A`), and the result is distributed over `Cell.arrayRange` with the same broadcasting — every member cell's `value` is written, the anchor's included
- [ ] `formula/evaluate.ts` exports `evaluate(node, ctx: EvalContext): EvalValue` per "Evaluation" — an interpreter over SPEC-119's `FormulaNode` taking ASTs from `parseFormulaAt` and resolving each relative `RefPart` as `part + (host − origin)`; it, `values.ts`, `registry.ts`, `functions/*.ts`, `dependencies.ts` and `recalc.ts` import nothing from Vue, `window`, `document` or `core/idleScheduler.ts`, and contain no `eval`/`new Function`/`Function(` — asserted by SPEC-119's `formula/purity.spec.ts`, which this spec extends to the new files
- [ ] `formula/registry.ts` exports `FunctionSpec`, `ArgKind`, `ArgList`, `FunctionCategory`, `registerFunction(spec)` (throws on a duplicate name), `getFunction(name)` and `registeredFunctionNames()` per "The registry"; argument kinds are `"value" | "range" | "any" | "array"` and the evaluator — not each function — applies them (intersection, lifting, the reference requirement, array evaluation of `"array"` arguments)
- [ ] Exactly the 100 functions in "The core functions" are registered, from `formula/functions/math.ts`, `aggregate.ts`, `logical.ts`, `information.ts`, `text.ts`, `dateTime.ts` and `lookup.ts`, each with its category, arity, argument kinds, a one-line `summary` and named `args`; `registeredFunctionNames()` equals the list, and every name resolves through SPEC-119's `functionNames.ts` in both wire dialects
- [ ] The *IF/*IFS criteria and every wildcard match follow "Criteria": an optional operator (`=`, `<>`, `<`, `<=`, `>`, `>=`) then an operand matched numerically, logically, by error code or as text; `*`, `?` and the `~` escape apply to `=`/`<>` text criteria, `MATCH`/`VLOOKUP`/`HLOOKUP` exact mode and `SEARCH`; wildcard matching is the linear two-pointer algorithm in `formula/functions/criteria.ts` — no `RegExp` is ever built from cell or formula text
- [ ] TEXT formats through SPEC-118's `formatValue(value, code, { locale, dateSystem })` — the format code is read as the invariant (en-US) code a cell's number format is, and the output is localised (`=TEXT(1234.5,"#,##0.00")` is "1,234.50" in `en-US`, "1.234,50" in `de-DE`); an invalid code is `#VALUE!`
- [ ] Date functions work in both date systems over SPEC-116's `serial.ts`, honouring the fictitious 1900-02-29 (`=DATE(1900,2,29)` is 60, `=EOMONTH(DATE(1900,2,1),0)` is 60); TODAY and NOW read `core/clock.ts`'s `now()`, never `Date.now()`, and RAND/RANDBETWEEN read `ctx.random()`, so unit tests pin both
- [ ] `formula/dependencies.ts` exports `CellKey`, `cellKey`, `parseCellKey`, `DependencyGraph`, `graphFor(model)`, `updateGraphCells(model, keys)`, `invalidateDependencyGraph(model)` and `findCycles(graph)` per "The dependency graph": cell precedents in a map, range precedents per sheet in `BAND_WIDTH = 32`-column bands with ranges wider than `WIDE_RANGE_BANDS = 8` bands in a per-sheet wide list, 3-D references expanded to one area per sheet, a reference overlapping an array range depending on its anchor, and INDIRECT/OFFSET dependencies recorded at evaluation and replaced on each evaluation; the graph is lib state in a `WeakMap` keyed by the model, never model data
- [ ] `formula/recalc.ts` exports `changedCellsBetween(sheetId, beforeRows, afterRows)` (the identity diff: `core/arrayWindow.ts`'s `windowOf` over rows, then over each changed row's cells), `recalculate(model, changed, options?)`, `startRecalc(model, changed, options?)` and `RecalcResult` per "Recalculation": the dirty set is the transitive dependents of `changed` plus every volatile cell and its dependents, evaluated in a deterministic topological order; results are written copy-on-write (one new `Row` per touched row per step, one new `rows` array per touched sheet), an unchanged value writes nothing, and untouched rows keep their identity
- [ ] Cycles follow "Cycles": every cell of a cycle evaluates to 0 and its dependents compute from that 0; `RecalcResult.circular` lists the first cell of each cycle in sheet order then row-major; the status bar shows `sheet-status-circular` reading "Circular reference: Sheet1!A1" (sheet name quoted per SPEC-119) while any cycle exists — at open too, where cycles are found from the graph but their cached values are left alone — and there is no iterative calculation
- [ ] The open-time policy follows "Opening a workbook": cached values are trusted and painted first; a full recalculation runs only when `model.calc.fullCalcOnLoad` is set or some formula cell has `value === null`; otherwise, once idle, one pass recomputes the volatile cells and their dependents (so TODAY is today, as Excel does on open); no open-time pass sets `dirty` (SPEC-121) or changes the downloaded bytes of an untouched workbook (`equalBytes`)
- [ ] Uncomputable cells follow "Stale cells": a formula using a function not in the registry, a name, a structured or external reference, a dynamic array that spills (per "Array mode" — a single-cell one whose result fits its cell computes), an unparsed formula, or a cycle in an iterating workbook keeps its cached value, and other cells compute from that value; when such a cell is a transitive dependent of anything that changed in a pass it is **stale** — its `sheet-cell` gets `data-stale="true"` and a `title` from the "Stale cells" table, and the pass sets `model.forceFullCalcOnLoad`, which SPEC-122's xlsx writer turns into `<calcPr fullCalcOnLoad="1"/>`
- [ ] The readers bump `SPREADSHEET_READER_VERSION` (to the next integer) and append their `ReaderUpgrade` (a dirty-sheet step that throws, as SPEC-119's does) to read `model.calc` per "Calculation settings" — xlsx `calcPr@fullCalcOnLoad`, `@calcMode`, `@iterate`; ods `table:calculation-settings` (`table:case-sensitive`, `table:use-regular-expressions`, `table:use-wildcards`, `table:search-criteria-must-apply-to-whole-cell`, `table:iteration@table:status`) and the `settings.xml` AutoCalculate item
- [ ] A workbook saved in manual calculation mode is recalculated automatically all the same, and opening it raises one toast: "This workbook was set to calculate manually. It recalculates automatically here." — the file's own setting is preserved verbatim
- [ ] Large recalculations are time-sliced per "Time slicing": `startRecalc` returns a `RecalcPass` whose `step(budgetMs)` does at most that much work (`performance.now()`), the view drives it in `RECALC_SLICE_MS = 12` slices through `core/idleScheduler.ts`'s `scheduleEagerWork`, and `sheet-status-calculating` reads "Calculating…" only once a pass has run longer than `CALCULATING_STATUS_DELAY_MS = 250`; a pass started while another runs cancels it and inherits its unevaluated cells; a pass driven in 0 ms steps produces a result `toEqual` to the synchronous `recalculate`
- [ ] `SpreadsheetMenuBar.vue` gains a Data menu (`menu-title-data`, after View) whose last item, after a separator, is "Recalculate" (`sheet-menu-recalculate`, shortcut hint "F9"); it and F9 run a full recalculation of every formula cell, ignoring cached values; F9 is ignored when `isOutsideDocumentInput(event.target)` (the SPEC-021 convention), leaving it free for SPEC-122's cell editor
- [ ] The view wiring lives in `src/views/spreadsheet/useRecalc.ts` (≤ 400 lines per `composableSize.spec.ts`): open-time policy, passes, the stale set (view state only — never persisted, never in the file), the circular and calculating status, F9 and the manual-mode toast; `SpreadsheetView.vue` gains only the wiring call and template hookup and its size ratchet does not rise
- [ ] A human-generated conformance fixture, `e2e/fixtures/spec120-numbers.xlsx`, exported by Apple Numbers through `scripts/make-numbers-calc-fixture.applescript` per "The Numbers oracle", is committed **before this spec is labelled ready** (the factory cannot run Numbers), with a note recording the Numbers version; `formula/numbersOracle.spec.ts` fully recalculates it and every formula cell matches Numbers' cached value under `roundTo15` except the rows the test lists with a reason
- [ ] Tests: `formula/values.spec.ts`, `conformance.spec.ts`, `evaluate.spec.ts`, `registry.spec.ts`, `functions/*.spec.ts` (one per functions file, plus `criteria.spec.ts`), `dependencies.spec.ts`, `recalc.spec.ts`, `numbersOracle.spec.ts`, the reader and `open.spec.ts` additions, and `views/spreadsheet/useRecalc.spec.ts`
- [ ] e2e (`e2e/SPEC-120.spec.ts`), with `page.clock.setFixedTime(new Date(2026, 9, 8, 9, 30))` before navigation: drop `spec120-calc.xlsx` → open → `B2` (`=SUM(A1:A3)`, cached 999) reads "999" → `D2` (`=TODAY()`) reads "10/8/2026" → `sheet-status-circular` reads "Circular reference: Loop!A1" → press F9 → `B2` reads "6" → `E2` (`=_xll.CUSTOM(B2)`, cached 42) still reads "42" with `data-stale="true"` → open the Data menu → "Recalculate" is present → Close → drop `spec120-fullcalc.xlsx` → open → its zero-cached totals read their computed values → download → every part `equalBytes` the fixture's → drop `spec120-manual.xlsx` → open → the toast reads "This workbook was set to calculate manually. It recalculates automatically here."

## Out of scope
The other ~150 functions, function help and the function picker (SPEC-127 —
until then they are "unsupported" and their cells keep cached values). CELL
and INFO (volatile, registered by SPEC-127). Defined names and structured
references (SPEC-133 supplies `resolveName`; until then they are
uncomputable). Dynamic arrays and spilling (FILTER, SORT, UNIQUE, SEQUENCE,
LET, LAMBDA — backlog; their cached spill values show — a single-cell
dynamic-array formula whose result fits its cell is in scope, per "Array
mode"). Iterative
calculation, manual calculation mode, multi-threaded calculation, "precision
as displayed" (`calcPr@fullPrecision="0"`, ODS `table:precision-as-shown`) —
backlog. External workbook links (cached values only). What-If data tables
(never recomputed; see "Stale cells"). R1C1 text in `INDIRECT(…, FALSE)`
(`#REF!`, backlog). Excel's "word sort" (ignoring hyphens and apostrophes in
text order) — the collator is ICU's. Entering or editing formulas (SPEC-122),
which is also when edit-triggered recalculation first runs in the product.
Moving evaluation into a worker, and recalculation performance budgets
(SPEC-137). Emulating LibreOffice's application-level settings that are not
in the file (its text-to-number conversion option).

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — `SpreadsheetModel { id; format;
  sheets; dateSystem: "1900" | "1904"; … }`, `Sheet { id; name; kind:
  "grid" | "foreign"; rows }`, `Row { index; cells }` sparse and sorted,
  `Cell { col; value: CellValue; … }`, `CellError { error: ErrorCode }`; the
  copy-on-write invariant (an edit builds new `Cell`/`Row` objects and a new
  `rows` array; untouched rows keep identity) and sparseness; `getCell`,
  `findRow` (binary search), `GridRange`, `CellRef`, `MAX_ROWS`, `MAX_COLS`;
  `serial.ts` (`dateToSerial(y, m, d, system)`, `serialToParts(serial,
  system)`, 1900 serial 60 = the fictitious 1900-02-29, 1904 serial 0 =
  1904-01-01); reader versions (`SPREADSHEET_READER_VERSION`,
  `ReaderUpgrade`, `READER_UPGRADES`); `xl/workbook.xml` read with
  `parseXml`; part-level `equalBytes` for untouched downloads.
- `specs/117-sheet-grid-styles-sizes-and-selection.md` — `SheetGrid.vue`
  (`sheet-cell`, `data-ref`), `SpreadsheetStatusBar.vue` (`sheet-status-sum`
  etc.; this spec adds two regions), `SpreadsheetMenuBar.vue` on the
  product-neutral `MenuBar.vue`/`MenuItem.vue` (this spec adds its Data menu;
  SPEC-128 later puts Sort/Filter above "Recalculate"), and
  `isOutsideDocumentInput` in `src/lib/shortcuts.ts` (SPEC-021).
- `specs/118-number-formats-and-dates.md` — `formatValue(value: CellValue,
  code: string | undefined, ctx: FormatContext): FormattedValue` with
  `FormatContext { locale: SpreadsheetLocale; dateSystem; widthPx?;
  dateQuirks?: "excel" | "proleptic"; … }` (the view passes `"proleptic"`
  for `.ods`) and `FormattedValue { text; color?; alignHint?; overflow? }`;
  `SpreadsheetLocale { tag; decimal; group; dateOrder: "MDY" | "DMY" | "YMD";
  dateSeparator; amPm: [string, string]; currencySymbol; currencyPosition:
  "prefix" | "suffix"; … }` for `en-US` (default), `en-GB`, `de-DE`,
  `fr-FR`, `es-ES`, `cs-CZ`, chosen per user (SPEC-066 preferences), never
  per file; formulas are always en-US syntax.
- `specs/119-the-formula-language.md` — `FormulaNode` (number, string,
  boolean, error, array, ref, name, function with `missing` args, unary
  `+ - %`, binary incl. `:`, `" "`, `,`, paren, structured, external),
  `ReferenceNode { sheet?; sheetEnd?; start: RefPart; end? }`, `RefPart {
  row?; col?; rowAbs; colAbs }` (0-based; absent row = whole column),
  `parseFormulaAt(text, host): { node; origin } | FormulaSyntaxError` (LRU by
  `relativeFormulaKey`; nodes immutable), `tokenize`, `findArrayAnchor`,
  `Cell.formula` (canonical), `formulaUnparsed`, `arrayRange`, `dynamicArray`,
  `dataTableXml`, `functionNames.ts`, `formula/purity.spec.ts`, the
  sheet-name quoting predicate. SPEC-119 is also where "Sheet!#REF!" errors
  and external nodes come from.
- Platform: `core/clock.ts` (`now()`, `setClock` — the one clock seam),
  `core/arrayWindow.ts` (`windowOf(before, after)` → common prefix/suffix by
  reference identity), `core/idleScheduler.ts` (`scheduleEagerWork(cb,
  budgetMs)`, `runWhenIdle(cb, timeoutMs)`), `src/composables/useToasts.ts`
  (`addToast(message)`, `toast` testid), `src/composableSize.spec.ts`.

### Values (`formula/values.ts`)
```ts
export interface EmptyValue { kind: "empty" }
export const EMPTY: EmptyValue;
export interface Area extends GridRange { sheetId: string }
export interface RangeValue { kind: "range"; areas: Area[] }       // a reference, read lazily
export type ScalarValue = number | string | boolean | CellError | EmptyValue;
export interface ArrayValue { kind: "array"; rows: number; cols: number; data: ScalarValue[] } // row-major
export type EvalValue = ScalarValue | RangeValue | ArrayValue;

export function toNumber(v: ScalarValue, s: EvalSettings): number | CellError;
export function toText(v: ScalarValue): string | CellError;
export function toLogical(v: ScalarValue): boolean | CellError;
export function valueFromText(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): number | null;
export function numberToText(n: number): string;
export function compareScalars(a: ScalarValue, b: ScalarValue, s: EvalSettings): -1 | 0 | 1 | CellError;
export function roundTo15(n: number): number;               // Number(n.toPrecision(15)); 0 stays 0
export function toCellValue(v: ScalarValue): CellValue;    // EMPTY → 0, −0 → 0, non-finite → #NUM!
```
`EMPTY` is what reading an empty cell or a `missing` argument produces; it
never reaches the model.

### Coercion
| From \ To | number | text | logical |
| --- | --- | --- | --- |
| number | itself | `numberToText` | ≠ 0 |
| text | `valueFromText`, else `#VALUE!` | itself | `"TRUE"`/`"FALSE"` any case, else `#VALUE!` |
| boolean | 1 / 0 | `"TRUE"` / `"FALSE"` | itself |
| `EMPTY` | 0 | `""` | FALSE |
| error | the error | the error | the error |

Operators coerce every operand, including text **cells** (`=A1+1` with A1
the text "5" is 6). Aggregates (SUM, AVERAGE, MIN, MAX, PRODUCT, COUNT…)
read numbers only from references and arrays — text, logicals and empties
there are skipped — but coerce arguments given directly (`=SUM("3",TRUE)` is
4, `=MAX("a")` is `#VALUE!`). AND/OR/XOR read logicals and numbers from
references and arrays and coerce direct arguments logically; with nothing to
read they are `#VALUE!`.

`numberToText(n)`: round to 15 significant digits; fixed notation with
trailing zeros removed when `1e-9 ≤ |n| < 1e15`, otherwise scientific with
the shortest mantissa and `E+NN`/`E-NN` (`1/3` → `"0.333333333333333"`,
`1e15` → `"1E+15"`). Locale-independent: it is the text a number becomes
inside a formula, as in Excel.

### Text to number (`valueFromText`)
After trimming U+0020, in the given locale:
- **Number**: optional sign or enclosing parentheses (negative); digits with
  `locale.group` separators only between groups of exactly three digits (a
  space-like group separator also accepts U+0020, U+00A0 and U+202F, as
  SPEC-118 states for input); an optional `locale.decimal` fraction; an
  optional `E±digits` exponent; an optional trailing `%` (÷ 100);
  `locale.currencySymbol` at `locale.currencyPosition` (a suffix after an
  optional space).
- **Date**: day, month and year in `locale.dateOrder` separated by `/`, `-`
  or `.`; a month name or three-letter abbreviation (English, plus the
  locale's own) in place of the month number; a missing year means the
  current year (`now()`); a two-digit year is 2000–2029 for 00–29 and
  1930–1999 otherwise; ISO `YYYY-MM-DD` in every locale. The serial is in the
  workbook's date system; an impossible date (`2/30/2026`) is no match.
- **Time**: `h:mm`, `h:mm:ss`, `h:mm:ss.fff`, optional `locale.amPm` markers (any case);
  hours may exceed 24 without AM/PM. A date and a time separated by spaces
  add.
- Anything else, including `""` and `"TRUE"`, returns `null`.

### Comparison
- Different types order **number < text < logical**; FALSE < TRUE. An
  `EMPTY` operand compares as 0, `""` or FALSE to match the other side (two
  empties are equal). An error operand returns that error (left first).
- Numbers compare as `roundTo15(a)` vs `roundTo15(b)`.
- Text equality and order use a collator built once per locale:
  `Intl.Collator(locale.tag, { sensitivity: "accent" })` — `"a" = "A"`,
  `"é" ≠ "e"`; with `calc.caseSensitive` (ods only) `sensitivity: "variant"`.
  Fast path: identical strings are equal without the collator.
- **Final operation.** When the formula's root node (looking through
  `paren`) is a binary `+` or `−` whose operands are both non-zero numbers
  that cancel at 15 significant digits (`roundTo15(a) === roundTo15(−b)` for
  `+`, `roundTo15(a) === roundTo15(b)` for `−`), the result is exactly 0.
  This emulates Excel's documented final-subtraction adjustment and applies
  nowhere else (`=1*(0.1+0.2-0.3)` keeps its tiny residue).

### Errors
`0^0`, a negative base with a non-integer exponent, `SQRT` of a negative,
`LN`/`LOG` of a non-positive, and any non-finite result are `#NUM!`;
division by zero is `#DIV/0!`; a function given the wrong number of
arguments (possible only in a hand-made file) is `#VALUE!`. Value-kind
arguments are checked for errors before `impl` runs, left to right; range,
any and array arguments are the function's business (SUM propagates the
first error it reads, COUNT ignores errors).

### Ranges and areas
A `ref` node evaluates to a `RangeValue`: the sheet named (or the host's),
resolved case-insensitively against `model.sheets` (an unknown or `"foreign"`
sheet is `#REF!`); a 3-D reference yields one `Area` per grid sheet between
the two names in workbook order, either direction. Relative parts resolve
against `host − origin`; a whole column spans rows `0..MAX_ROWS−1`.
`ctx.valueAt` reads a cell; range iteration visits stored cells only, and
functions that count empties (COUNTBLANK, array materialisation) compute
them arithmetically from the area size. Materialising a range as an array
(array mode, `"array"` arguments) is bounded by `MAX_ARRAY_CELLS`; an
implementation may represent a run of empties implicitly as long as results
are identical. The reference operators: `:` takes two single-area ranges on
one sheet and returns their bounding area; ` ` returns the intersection
(empty → `#NULL!`); `,` concatenates areas.

### Implicit intersection
In a non-array formula, wherever a single value is expected (an operator
operand, a value-kind argument, the formula's result):

| Value | Becomes |
| --- | --- |
| single-cell range | that cell's value |
| one column, several rows | the cell in the host's row if within the range, else `#VALUE!` |
| one row, several columns | the cell in the host's column if within the range, else `#VALUE!` |
| 2-D range, several areas | `#VALUE!` |
| array | its top-left element |

The host's row and column apply even when the range is on another sheet.
`SINGLE(x)` applies this rule explicitly in any mode (it is the wire form of
Excel's `@`).

### Array mode
Entered for the whole formula of an `arrayRange` anchor, and for the
argument subtree of an `"array"`-kind argument (so `SUMPRODUCT((A1:A3>1)*B1:B3)`
needs no CSE, as in Excel). In array mode a range in a value position is
materialised, operators apply element-wise, and a function whose value-kind
argument receives an array is applied per element, returning an array.
Shapes broadcast: a 1×1 array or scalar fills any shape; a single row repeats
down, a single column across; otherwise the result takes the larger extent
in each dimension and positions outside an operand are `#N/A`. The anchor's
result is distributed over `arrayRange` by the same rule; `EMPTY` elements
store as 0. `ROW()`/`COLUMN()` of a multi-cell range in array mode return
the vertical/horizontal array of indices.

A dynamic-array anchor (`dynamicArray`) whose `arrayRange` is its own single
cell is an `arrayRange` anchor like any other: it evaluates in array mode and
stores its 1×1 result. Excel 365 writes an array-evaluated formula that does
not spill this way, and Numbers 14.4 writes **every** formula this way (see
"The Numbers oracle"). If the result is larger than 1×1, the formula would
spill, so `evaluateCell` returns `{ uncomputable: { kind: "dynamicArray" } }`
and the cell keeps its cached value. A dynamic-array anchor over more than
one cell, and its members, are uncomputable without being evaluated.

### Evaluation (`formula/evaluate.ts`)
```ts
export interface EvalSettings {
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
  caseSensitive: boolean;                        // ods setting; xlsx false
  criteria: "wildcards" | "regex" | "plain";     // ods setting; xlsx "wildcards"
  wholeCell: boolean;                            // ods setting; xlsx true
  dateQuirks: "excel" | "proleptic";             // SPEC-118's FormatContext value: "proleptic" for ods
}
export interface EvalContext {
  model: SpreadsheetModel;
  host: { sheetId: string; row: number; col: number };
  origin: CellRef;                               // from parseFormulaAt
  arrayMode: boolean;
  settings: EvalSettings;
  now(): Date;                                   // core/clock.ts now by default
  random(): number;                              // Math.random by default
  valueAt(sheetId: string, row: number, col: number): ScalarValue; // evaluates a dirty precedent first
  recordDependency(area: Area): void;            // INDIRECT/OFFSET results
  /** SPEC-133 adds `resolveName(name, sheetQualifier, host)` and
   *  `resolveStructured(node, host)` and owns their signatures. While a hook
   *  is absent, a name or structured node makes the cell uncomputable. */
}
export type UncomputableReason =
  | { kind: "function"; name: string } | { kind: "name"; name: string } | { kind: "structured" }
  | { kind: "external" } | { kind: "dynamicArray" } | { kind: "unparsed" } | { kind: "iterative" }
  | { kind: "regex" } | { kind: "dataTable" };
export interface Uncomputable { uncomputable: UncomputableReason }
export function evaluate(node: FormulaNode, ctx: EvalContext): EvalValue;
export function evaluateCell(model: SpreadsheetModel, sheetId: string, row: number, col: number,
  ctx: Omit<EvalContext, "host" | "origin" | "arrayMode">): CellValue | ArrayValue | Uncomputable;
export const MAX_EVAL_DEPTH = 512;
```
A recursive interpreter with no code generation. `valueAt` on a cell that
is dirty in the running pass evaluates it first (on-demand ordering for
dynamic dependencies); past `MAX_EVAL_DEPTH` nested on-demand evaluations
the outer cell is re-queued behind its precedent instead of recursing, so a
long dynamic chain finishes in several sweeps rather than overflowing the
stack. A cell found in progress is a cycle. Per-AST facts (volatile,
dynamic, uncomputable reason, array-constant values) are memoised in a
`WeakMap` keyed by the shared node.

### The registry (`formula/registry.ts`)
```ts
export type ArgKind = "value" | "range" | "any" | "array";
export type FunctionCategory = "Financial" | "Date & Time" | "Math & Trig" | "Statistical"
  | "Lookup & Reference" | "Text" | "Logical" | "Information" | "Engineering" | "Compatibility";
                                      // SPEC-127's display categories, adopted here so it narrows nothing
export interface ArgList {
  readonly length: number;
  get(index: number): EvalValue;      // lazy functions: evaluated on first get, memoised
  isMissing(index: number): boolean;  // an empty slot `F(1,,3)` or an omitted trailing argument
}
export interface FunctionSpec {
  name: string;                       // canonical, upper-case
  category: FunctionCategory;
  minArgs: number;
  maxArgs: number;                    // 255 for variadic
  argKinds: ArgKind[];                // one per entry of `args`
  /** args[from … from + count − 1] repeat as a group up to maxArgs (SPEC-127's
   *  field, defined here because argument kinds cycle the same way). */
  repeat?: { from: number; count: number };
  volatile?: true;
  lazy?: true;
  acceptsErrors?: true;
  impl(args: ArgList, ctx: EvalContext): EvalValue;
  summary: string;
  args: { name: string; description: string; optional?: true; repeating?: true }[];
}
export function registerFunction(spec: FunctionSpec): void;
export function getFunction(name: string): FunctionSpec | undefined;
export function registeredFunctionNames(): string[];
```
| Kind | The evaluator passes | Non-array mode | Array mode |
| --- | --- | --- | --- |
| `value` | a scalar | range/array → implicit intersection | array → the function lifts element-wise |
| `range` | a `RangeValue` | anything else → `#VALUE!` | same |
| `any` | range, array or scalar as evaluated | no intersection at this argument | same |
| `array` | as `any`, the subtree evaluated in array mode | — | — |

An argument at index `i ≥ argKinds.length` takes
`argKinds[from + (i − from) mod count]` from `repeat` (SUMIFS: `argKinds`
`["range", "range", "value"]`, `repeat: { from: 1, count: 2 }`). An empty
slot (`F(1,,3)`) is `EMPTY`, so it reads as 0/`""`/FALSE — which is why
`VLOOKUP(x,t,2,)` is an exact match; an omitted trailing argument takes the
function's default. This spec's `summary` and `args` follow SPEC-127's
metadata invariants as far as they are stated there (a summary ending in
".", snake_case argument names, required before optional); SPEC-127 owns
that metadata, rewrites it freely, adds `visibilitySensitive` and
`UnsupportedCaseError` (which the evaluator maps onto the uncomputable
path), and registers every further function through `registerFunction`.
`FunctionCategory`'s values are the display names themselves; SPEC-127 adds
no second category vocabulary.

### The core functions
| File | Functions |
| --- | --- |
| `math.ts` | SUM SUMIF SUMIFS SUMPRODUCT PRODUCT ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD ABS SIGN SQRT POWER EXP LN LOG LOG10 PI CEILING FLOOR MROUND RAND RANDBETWEEN |
| `aggregate.ts` | AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MIN MAX MINIFS MAXIFS |
| `logical.ts` | IF IFS IFERROR IFNA AND OR NOT XOR TRUE FALSE SWITCH CHOOSE |
| `information.ts` | ISBLANK ISNUMBER ISTEXT ISERROR ISERR ISNA ISLOGICAL NA |
| `text.ts` | CONCAT CONCATENATE TEXTJOIN LEN LEFT RIGHT MID UPPER LOWER PROPER TRIM SUBSTITUTE REPLACE FIND SEARCH TEXT VALUE REPT EXACT |
| `dateTime.ts` | DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND TODAY NOW WEEKDAY EDATE EOMONTH DAYS |
| `lookup.ts` | VLOOKUP HLOOKUP INDEX MATCH ROW COLUMN ROWS COLUMNS OFFSET INDIRECT SINGLE |

Semantics an implementer would otherwise guess:
- **Rounding.** ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR and MROUND work
  on the 15-significant-digit decimal form of their input (ROUND half away
  from zero; the quotient for CEILING/FLOOR/MROUND rounded to 15 digits
  before it is floored), so `ROUND(1.005,2)` is 1.01 and `FLOOR(0.3,0.1)` is
  0.3. CEILING/FLOOR/MROUND follow Excel's documented sign rules (`CEILING(-2.5,2)`
  = −2, `CEILING(-2.5,-2)` = −4, `FLOOR(2.5,-2)` = `#NUM!`, `MROUND(5,-2)` =
  `#NUM!`); `CEILING(x,0)` = 0, `FLOOR(x,0)` = `#DIV/0!`.
- `MOD(n,d)` = `n − d·INT(n/d)` (sign of the divisor). `INT` floors; `TRUNC`
  truncates. `RANDBETWEEN(a,b)` = `ceil(a) + floor(random()·(floor(b)−ceil(a)+1))`,
  `#NUM!` when `ceil(a) > floor(b)`. `PRODUCT` of no numbers is 0.
- **SUMIF/AVERAGEIF** resize the sum/average range to the criteria range's
  shape from its top-left cell, as Excel does; the *IFS family requires every
  range the same shape (`#VALUE!` otherwise). AVERAGE* of nothing is
  `#DIV/0!`; MIN/MAX/MINIFS/MAXIFS of nothing are 0. `SUMPRODUCT` requires
  equal dimensions and reads non-numbers as 0.
- **COUNTA** counts every non-empty cell, `""` results and errors included;
  **COUNTBLANK** counts empty cells and `""` results; **ISBLANK** is TRUE only
  for an empty cell.
- **Text** functions count UTF-16 code units, as Excel does (`LEN("😀")` is
  2). FIND is case-sensitive without wildcards; SEARCH is case-insensitive
  with wildcards; not found is `#VALUE!`. UPPER/LOWER are locale-independent
  (`toUpperCase`). PROPER capitalises every letter that follows a non-letter
  (`PROPER("it's")` is "It'S"). TRIM removes U+0020 only. CONCATENATE takes
  value arguments; CONCAT and TEXTJOIN read ranges row-major. TEXT uses
  SPEC-118's formatter (colour ignored); VALUE uses `valueFromText` (a
  number passes through, an empty cell is 0, `""` is `#VALUE!`).
- **Dates.** DATE adds 1900 to years 0–1899, normalises month and day
  overflow, and is `#NUM!` for a year outside 0–9999 or a negative serial; in
  the 1900 system February 1900 has 29 days in every date function — except
  in an `.ods` workbook (`dateQuirks: "proleptic"`, as SPEC-118 renders it),
  where it has 28 and no function produces serial 60. TIME
  truncates its arguments and returns the fraction of a day modulo 1
  (`#NUM!` if negative). HOUR/MINUTE/SECOND take the time of day rounded to
  the nearest second. WEEKDAY supports return types 1, 2, 3 and 11–17;
  in the 1900 system the weekday is Excel's arithmetic `(⌊serial⌋ + 6) mod 7`
  (so `WEEKDAY(1)` is 1, Sunday), in 1904 `(⌊serial⌋ + 5) mod 7`. EDATE and
  EOMONTH clamp the day to the target month's end. DAYS is
  `INT(end) − INT(start)`. NOW is local time to the millisecond.
- **Lookups.** Exact mode (MATCH 0, VLOOKUP/HLOOKUP FALSE or an empty slot)
  is a linear scan for the first match — wildcards for a text lookup value,
  case-insensitive unless `caseSensitive`; empty cells never match. Approximate
  mode (MATCH 1/−1, VLOOKUP/HLOOKUP TRUE or omitted) bisects the vector,
  skipping elements of a type other than the lookup value by probing to the
  nearest same-type element below, then above; on sorted data this is
  Excel's documented answer, on unsorted data Excel documents the result as
  undefined and this one is whatever the bisection finds. INDEX given a
  range returns a **reference** (row or column 0 selects the whole column or
  row; a one-row source indexes columns with a lone second argument); out of
  bounds is `#REF!`. OFFSET returns a reference (`#REF!` off-grid or for a
  height/width ≤ 0). INDIRECT accepts exactly one reference or name token,
  optionally sheet-qualified (`#REF!` otherwise; a second argument of FALSE
  is `#REF!`). OFFSET and INDIRECT record their result areas as dynamic
  dependencies. CHOOSE truncates its index (`#VALUE!` outside 1..n).
- **Volatile**: NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT (CELL and
  INFO when SPEC-127 registers them).

### Criteria (`formula/functions/criteria.ts`)
A criterion is a value. A number matches number cells equal under
`roundTo15` and text cells whose `valueFromText` is equal; a boolean matches
that boolean; an empty-cell criterion is the number 0. A text criterion is
parsed as an optional operator and an operand:

| Operand is… | `=` (or no operator), `<>` | `<`, `<=`, `>`, `>=` |
| --- | --- | --- |
| `valueFromText` succeeds | numeric equality with number cells (and numeric text) | numeric order, number cells only |
| `TRUE`/`FALSE` | booleans | booleans |
| an error literal | errors of that code | nothing |
| empty (`""` or `"="`) | `""`: empty cells and `""` results; `"="`: empty cells only; `"<>"`: non-empty cells | — |
| other text | text cells, wildcards, case per `caseSensitive` | collator order, text cells only |

Wildcards: `*` any run, `?` one UTF-16 unit, `~*` `~?` `~~` literals. The
matcher is the greedy two-pointer algorithm with one backtrack point —
O(text × pattern) worst case, never a `RegExp`. ODS settings adjust it
(see "Calculation settings").

### The dependency graph (`formula/dependencies.ts`)
```ts
export type CellKey = string;                    // built only by cellKey
export function cellKey(sheetId: string, row: number, col: number): CellKey; // `${sheetId}!${row},${col}`
export function parseCellKey(key: CellKey): { sheetId: string; row: number; col: number }; // splits at the last "!"
export const BAND_WIDTH = 32;
export const WIDE_RANGE_BANDS = 8;
export interface DependencyGraph {
  precedentsOf(key: CellKey): readonly Area[];                       // static ∪ last dynamic
  dependentsOf(sheetId: string, row: number, col: number): CellKey[];
  formulaCells(): Iterable<CellKey>;                                  // sheet order, row-major
  readonly volatile: ReadonlySet<CellKey>;
  readonly dynamic: ReadonlySet<CellKey>;
  readonly uncomputable: ReadonlyMap<CellKey, UncomputableReason>;
}
export function graphFor(model: SpreadsheetModel): DependencyGraph;   // built lazily, WeakMap-cached
export function updateGraphCells(model: SpreadsheetModel, keys: CellKey[]): void;
export function invalidateDependencyGraph(model: SpreadsheetModel): void;
export function findCycles(graph: DependencyGraph): CellKey[][];      // Tarjan, deterministic order
```
Nodes are formula cells (array anchors stand for their whole range; data
tables are not nodes). A single-cell precedent goes in a map; a range goes
in its sheet's band lists (sorted by `top`) or, if it spans more than
`WIDE_RANGE_BANDS` bands, the sheet's wide list. `dependentsOf` unions the
map hit with the band and wide-list entries containing the cell; a precedent
area overlapping an `arrayRange` also makes the formula depend on that
anchor. Sheet add/delete/rename/move (SPEC-121) and row/column edits
(SPEC-124) call `invalidateDependencyGraph`; cell edits call
`updateGraphCells`.

### Recalculation (`formula/recalc.ts`)
```ts
export type RecalcReason = "open" | "edit" | "manual";
export interface RecalcOptions {
  reason?: RecalcReason;            // default "edit"
  full?: true;                      // every formula cell is dirty
  settings?: Partial<EvalSettings>; // locale defaults to en-US
  now?: () => Date;
  random?: () => number;
}
export interface RecalcResult {
  valueChanged: CellKey[];
  changedSheetIds: string[];
  circular: CellKey[];              // first cell of each cycle, report order
  stale: CellKey[];
  evaluated: number;
}
export interface RecalcPass {
  step(budgetMs: number): boolean;  // true when done
  cancel(): CellKey[];              // the cells it had not evaluated
  readonly result: RecalcResult;
}
export function changedCellsBetween(sheetId: string, before: Row[], after: Row[]): CellKey[];
export function recalculate(model: SpreadsheetModel, changed: CellKey[], options?: RecalcOptions): RecalcResult;
export function startRecalc(model: SpreadsheetModel, changed: CellKey[], options?: RecalcOptions): RecalcPass;
export const RECALC_SLICE_MS = 12;
export const CALCULATING_STATUS_DELAY_MS = 250;
```
1. Update the graph for changed cells whose formula changed.
2. Dirty set = changed formula cells ∪ transitive dependents of every
   changed cell ∪ every volatile cell and its transitive dependents (or every
   formula cell when `full`).
3. Order the dirty subgraph by Kahn's algorithm, seeded in sheet order then
   row-major — deterministic, so the order of RAND draws and the cycle
   reported are reproducible. Leftover nodes are cycles (Tarjan names them).
4. Evaluate; uncomputable cells are skipped and recorded stale; write
   changed values copy-on-write, batched per row per step.
`changedSheetIds` is what SPEC-121 marks `dirty` after `"edit"` and
`"manual"` passes; an `"open"` pass marks nothing. Open-time and F9 results
are not persisted by this spec (nothing persists a workbook view before
SPEC-121); from SPEC-121 on they ride the write-through like any row change,
without touching the library's `lastModified`.

### Cycles
Every member of a cycle (a strongly connected component of more than one
node, or a self-reference) is 0 after an edit or F9 pass; dynamic cycles
(OFFSET onto itself) are caught by the in-progress check and treated the
same. At open, `findCycles` runs when the graph is first built and the
status bar reports them, but the cycle cells keep their cached values —
open-time never rewrites a trusted value. With `calc.iterate` set the
workbook meant its cycles: their cells are uncomputable (reason
`"iterative"`), keep their cached values, and no circular status shows.

### Opening a workbook
Order in `useRecalc`: paint with cached values → if `calc.fullCalcOnLoad`
or any formula cell has `value === null`, a full `"open"` pass (sliced) →
else, in `runWhenIdle`, build the graph (sliced), report cycles, and run one
`"open"` pass with no changed cells, which evaluates the volatile cells and
their dependents → raise the manual-mode toast if `calc.manual`. The pass
never sets `dirty` and never alters `opaque` or `sourceXml`, so an untouched
workbook still downloads `equalBytes`.

### Stale cells
`UncomputableReason` and the `title` of a stale `sheet-cell` (prefix "This
value couldn't be recalculated here — "):

| Reason | Suffix |
| --- | --- |
| `{ kind: "function"; name }` | "NAME isn't supported yet" |
| `{ kind: "name"; name }` | "named ranges aren't supported yet" |
| `{ kind: "structured" }` | "table references aren't supported yet" |
| `{ kind: "external" }` | "it links to another workbook" |
| `{ kind: "dynamicArray" }` | "dynamic array formulas aren't supported yet" |
| `{ kind: "unparsed" }` | "its formula couldn't be read" |
| `{ kind: "iterative" }` | "iterative calculation isn't supported" |
| `{ kind: "regex" }` | "regular-expression criteria aren't supported" |
| `{ kind: "dataTable" }` | "data tables aren't supported" |

A cell is stale when it is uncomputable and a transitive dependent of a
cell whose value or formula changed in the pass; after an open-time full
pass, every uncomputable cell is stale (the file itself asked for
recalculation). Every cell of a data-table range is stale after any
`"edit"` or `"manual"` pass, because its inputs cannot be traced. A pass
that produces any stale cell sets `model.forceFullCalcOnLoad = true`,
persisted with the model and never cleared by this product; SPEC-122's xlsx
writer emits `fullCalcOnLoad="1"` on `calcPr` when it is set. `.ods` has no
equivalent: a stale cell's last known value is written (SPEC-122 states it).

### Calculation settings
```ts
export interface CalcSettings {                 // SpreadsheetModel.calc?
  fullCalcOnLoad?: true;   // xlsx calcPr@fullCalcOnLoad="1"
  manual?: true;           // xlsx calcPr@calcMode="manual"; ods settings.xml AutoCalculate=false
  iterate?: true;          // xlsx calcPr@iterate="1"; ods table:iteration@table:status="enable"
  caseSensitive?: true;    // ods table:case-sensitive (ODF default "true"); never for xlsx
  criteria?: "regex" | "plain"; // ods: regex when use-regular-expressions is in effect, plain when neither it nor use-wildcards is
  partialCellMatch?: true; // ods table:search-criteria-must-apply-to-whole-cell="false"
}
// SpreadsheetModel also gains: forceFullCalcOnLoad?: true
```
ODF's defaults apply when `table:calculation-settings` or an attribute is
absent (`case-sensitive` true, `use-regular-expressions` true,
`use-wildcards` false, whole cell true) — LibreOffice writes the attributes
explicitly. In regex mode a criterion containing none of `.*+?()[]{}|^$\`
matches literally and any other is uncomputable (`"regex"`);
`partialCellMatch` makes `=` text criteria match substrings. Confirm the
`settings.xml` item name (`AutoCalculate`) against a LibreOffice-saved file
before implementing; if LibreOffice does not persist it per document, the
ods branch of `manual` is dropped. `calcMode="autoNoTable"` is automatic.

### Time slicing
`recalc.ts` stays DOM-free: a `RecalcPass` is a resumable state machine
(graph build, ordering and evaluation are all steps). `useRecalc` drives it
with `scheduleEagerWork(…, RECALC_SLICE_MS)`; a pass that finishes within
its first slice completes in the triggering task, so a small edit shows no
"Calculating…" flicker. SPEC-137 may move the pass into a worker over
`src/lib/workers/workerRpc.ts`; nothing here assumes the main thread.

### Where expected values come from
Each conformance row is labelled **Excel** — Microsoft's documented
behaviour (the "Calculation operators and precedence in Excel" article's
"Converting values in formulas" table, the function reference pages and
their examples) or Excel behaviour established well enough to verify on any
copy — or **pinned** — computed by hand from the rule this spec states,
where Excel's behaviour is undocumented. No expected value is produced by
running this engine and pasting its output.

### Conformance
| Formula (A1 etc. as stated) | Expected | Source |
| --- | --- | --- |
| `="1"+"2"` | 3 | Excel |
| `="1"+"$4.00"` (en-US) | 5 | Excel |
| `="6/1/2001"-"5/1/2001"` (en-US) | 31 | Excel |
| `=SQRT("8+1")` | `#VALUE!` | Excel |
| `="A"&TRUE` | "ATRUE" | Excel |
| `=1+""` / `="TRUE"+0` / `="a"*1` | `#VALUE!` | Excel |
| `=TRUE+TRUE` | 2 | Excel |
| `=A1+1`, A1 the text "5" | 6 | Excel |
| `=A1`, A1 empty | 0 | Excel |
| `=A1=0`, `=A1=""`, `=A1=FALSE`, A1 empty | TRUE, TRUE, TRUE | Excel |
| `=1="1"` / `=1<"1"` / `="z"<TRUE` / `=FALSE<TRUE` | FALSE / TRUE / TRUE / TRUE | Excel |
| `="A"="a"` / `="a"<"B"` | TRUE / TRUE | Excel |
| `=0.1+0.2=0.3` | TRUE | Excel |
| `=0.1+0.2-0.3` | 0 | pinned |
| `=-2^2` / `=0-2^2` / `=2^3^2` | 4 / −4 / 64 | Excel |
| `=0^0` | `#NUM!` | Excel |
| `=1/3&""` | "0.333333333333333" | Excel |
| `=#N/A+1/0` / `=1/0+#N/A` | `#N/A` / `#DIV/0!` | pinned |
| `=IF(TRUE,1,1/0)` | 1 | Excel |
| `=SUM("3",TRUE)` / `=SUM(A1:A2)`, A1 "3" text, A2 TRUE | 4 / 0 | Excel |
| `=MAX("a")` / `=MIN(A1:A3)` all empty / `=AVERAGE(A1:A3)` all empty | `#VALUE!` / 0 / `#DIV/0!` | Excel |
| `=AND(A1:A2)`, A1 "x" text, A2 TRUE / `=AND("x")` | TRUE / `#VALUE!` | pinned |
| `=ISBLANK(A1)`, A1 `=""` / `=COUNTBLANK(A1:A2)`, A2 empty / `=COUNTA(A1:A2)` | FALSE / 2 / 1 | Excel |
| `="1,5"+0` en-US / de-DE | `#VALUE!` / 1.5 | pinned |
| `="2026-10-08"+0` / `="12:00"+0` / `="(5)"+0` / `="10%"+0` | 46303 / 0.5 / −5 / 0.1 | pinned |
| `=ROUND(2.5,0)` / `=ROUND(-2.5,0)` / `=ROUND(1.005,2)` | 3 / −3 / 1.01 | Excel / Excel / pinned |
| `=MOD(-3,2)` / `=MOD(3,-2)` / `=INT(-8.9)` / `=TRUNC(-8.9)` | 1 / −1 / −9 / −8 | Excel |
| `=CEILING(-2.5,2)` / `=FLOOR(-2.5,-2)` / `=MROUND(10,3)` / `=FLOOR(0.3,0.1)` | −2 / −2 / 9 / 0.3 | Excel ×3 / pinned |
| `=DATE(108,1,2)` / `=DATE(2008,14,2)` / `=DATE(1900,2,29)` | 2008-01-02 / 2009-02-02 / 60 | Excel |
| `=WEEKDAY(1)` / `=WEEKDAY(46303)` | 1 / 5 | Excel / pinned |
| `=VLOOKUP("a",{"c",1;"a",2},2,)` | 2 | Excel |
| `=MATCH("B*",{"apple","banana"},0)` | 2 | Excel |
| `=COUNTIF(A1:A3,1)`, A1 the text "1", A2 1, A3 2 | 2 | pinned |
| `=LEN("😀")` / `=PROPER("it's")` / `=TRIM("  a   b  ")` | 2 / "It'S" / "a b" | pinned / pinned / Excel |
| `=TEXT(0.5,"0%")` | "50%" | Excel |
| `=A2:A4` in row 3, non-array / in row 9 | A3's value / `#VALUE!` | Excel |
| `{=A1:A2*{10,100}}` over B1:C3 | 2×2 products, row 3 `#N/A` | Excel |

### The Numbers oracle
`scripts/make-numbers-calc-fixture.applescript` builds one Numbers table:
input values in columns A–C and, one per row in column E, a formula from
this spec's list that Numbers also implements (arithmetic, rounding,
aggregates, *IF(S), text, dates, VLOOKUP/MATCH/INDEX), with column D holding
the formula text for the report; it exports with "Export To ▸ Excel".
`open -a Numbers --args -AppleLocale en_US` first, or `activate` fails with
-600 (the SPEC-087/116 precedent). The locale argument makes Numbers parse and
compute as en-US whatever the Mac's region. Numbers' cached
results are a third-party engine's arithmetic, which is the point: rows
where Numbers is known to differ from Excel are listed in the test with
their reason, not "fixed".

The fixture is committed (2026-10-09, Numbers 14.4; a value set through
AppleScript with a leading `=` does become a formula).
`e2e/fixtures/spec120-numbers.md` records its layout, and it is normative for
what the test reads. Three facts of the export bind `numbersOracle.spec.ts`:
- All 147 formula cells are single-cell dynamic-array formulas (`c@cm` plus
  `<f t="array" ref>` naming the cell itself). They compute through the
  single-cell rule in "Array mode", so the test recalculates the workbook
  exactly as the product does, with no special case. A dynamic-array cell
  left uncomputable fails the test.
- Numbers writes no error codes. A formula cell with no `<v>` (fourteen of
  them) is one Numbers evaluated to an error, and it matches when the
  engine's value is any `CellError`. A `t="b"` result compares as a boolean,
  `t="str"` as exact text and a number under `roundTo15`.
- The nine rows in the note's "Where Numbers differs from Excel" table are the
  test's listed exceptions, each with that table's reason. For those rows the
  test asserts the Excel / SPEC-120 value the table gives.

### Fixtures
Built by `spreadsheet/fixtures/build.ts` (Excel-shaped per SPEC-116):
- `spec120-calc.xlsx` — "Calc": `A1:A3` 1,2,3, `B2` `SUM(A1:A3)` cached 999
  (deliberately wrong, to prove the cache is trusted until F9), `D2`
  `TODAY()` with numFmt 14 and an old cached date (45000), `E2`
  `_xll.CUSTOM(B2)` cached 42 (an add-in function no registry will have), the
  conformance rows as formulas with Excel-true cached values; "Loop": `A1`
  `B1+1`, `B1` `A1+1`, cached 0; `calcPr` with no `fullCalcOnLoad`.
- `spec120-fullcalc.xlsx` — `calcPr fullCalcOnLoad="1"` and every cached
  value 0, as a library-written file has; `spec120-manual.xlsx` —
  `calcMode="manual"`; `spec120-calc.ods` — LibreOffice-shaped, explicit
  `table:calculation-settings` (case-sensitive, wildcards) and two formula
  cells without cached values.

### Tests
- `values.spec.ts` — the coercion table, `valueFromText` per locale (every
  grammar branch, group-position errors, impossible dates, two-digit years),
  `numberToText` boundaries, `compareScalars` incl. collator cases and
  `caseSensitive`.
- `conformance.spec.ts` — every row above, plus the `.ods` settings
  variants. `evaluate.spec.ts` — operators, intersection, union, `#NULL!`,
  3-D sums, array mode and broadcasting, a single-cell dynamic-array anchor
  (computes; uncomputable once its result is larger than 1×1; a multi-cell
  one is never evaluated), lazy branches, `MAX_EVAL_DEPTH`
  re-queueing, `MAX_ARRAY_CELLS`, origin offsets on a shared AST.
- `functions/*.spec.ts` — each function's documented examples plus its
  edge rows; `criteria.spec.ts` — the criteria table and a pathological
  wildcard (`*a*a*a*a*a*b` against 10,000 `a`s) finishing in linear-ish time.
- `dependencies.spec.ts` — bands, the wide list, 3-D, array anchors,
  dynamic dependencies replaced on re-evaluation, a 100,000-row shared
  formula yielding one cached AST and `dependentsOf(A50000)` returning one
  key. `recalc.spec.ts` — `changedCellsBetween` over copy-on-write edits,
  topological order, cycles and their report order, volatile seeding, stale
  marking and `forceFullCalcOnLoad`, open policy, step/sync equivalence,
  cancel-and-inherit, untouched row identity.
- `useRecalc.spec.ts` — the status texts and delay, F9 and the focus guard,
  the toast. Unit tests call `setClock` and inject a seeded `random`
  (mulberry32); none reads the wall clock.

### Specs this changes
- `specs/116-…` — `SpreadsheetModel` gains `calc` and `forceFullCalcOnLoad`;
  `SPREADSHEET_READER_VERSION` moves.
- `specs/118-…` — TEXT is a second caller of `formatValue`.
- `specs/121-…` — commands hand `changedCellsBetween` output to
  `recalculate` and mark `changedSheetIds` dirty for edit/manual passes; sheet
  commands call `invalidateDependencyGraph`.
- `specs/122-…` — `parseInput` recognises values through `valueFromText`;
  the xlsx writer honours `forceFullCalcOnLoad`; F9 inside the editor is free
  for it.
- `specs/124-…` — structural edits call `invalidateDependencyGraph`.
- `specs/127-…` — registers the rest of the library on `FunctionSpec` and
  extends the conformance table; CELL and INFO are volatile.
- `specs/133-…` — supplies `EvalContext.resolveName` and structured
  references, turning those uncomputable reasons into evaluation.
- `specs/137-…` — owns recalculation budgets and any worker move.
