---
id: SPEC-127
title: The function library and the function picker
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-120, SPEC-122, SPEC-123]
---
## Why
After SPEC-120 the engine computes the hundred functions nearly every
workbook uses. Real workbooks reach further: the loan schedule needs `PMT`,
`IPMT` and `PPMT`, the sales report `RANK.EQ` and `PERCENTILE.INC`, the HR
sheet `NETWORKDAYS` and `DATEDIF`, the filtered table `SUBTOTAL`, and anything built
since 2019 reaches for `XLOOKUP`. A function missing from the registry leaves
its cell stale (SPEC-120): the cached value shows, honestly marked, and then
never changes again when an input does. Every such function is a hole in
recalculation the user cannot see until a number is wrong.

The other half is discovery. Typing a formula here today means knowing the
function's name and its argument order by heart. Excel, Sheets and Calc all
give three helps — a list of matching names while you type, a hint showing
the signature with the current argument in bold, and an Insert Function
dialog to browse by category — and all three run on metadata (summaries,
argument names and descriptions) that SPEC-120 writes only for its own
hundred functions, to no stated standard.

This spec registers about two hundred further functions by category (some
twenty-five of them legacy names or aliases sharing an implementation),
holds every function's metadata — SPEC-120's included — to one checked
standard, and builds the three assists on SPEC-122's editors. **It is the
normative source for function metadata (category order, summaries,
argument names, descriptions, optionality and repetition), the semantics of
every function it adds, the numerical methods and iteration limits of the
iterative ones, the rule for where expected test values come from beyond
SPEC-120's table, and the autocomplete, argument hint and Insert Function
dialog.**

## Acceptance criteria
- [ ] Every `FunctionSpec` in the registry — SPEC-120's and this spec's — carries metadata per "Function metadata": its SPEC-120 `category`, a `summary`, and an `args` list with names, descriptions, `optional`, and `repeating` on a variadic function's repeated group; `formula/registry.spec.ts` asserts the invariants listed there for every name in `registeredFunctionNames()` and fails naming the function that breaks one; `FUNCTION_CATEGORY_ORDER` lists SPEC-120's `FunctionCategory` values — each already its own display name — in the dialog's order
- [ ] `formula/functionHelp.ts` exports `functionSignature(spec)` rendering Excel's signature form — `SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)`, `SUM(number1, [number2], …)`, `PMT(rate, nper, pv, [fv], [type])` — and `argumentAt(spec, index)` mapping any argument index, past the end of a repeated group, to its display name (`number7`, `criteria_range3`) or `null` beyond `maxArgs`
- [ ] Lookup & Reference per the Reference: `XLOOKUP` (match modes 0, −1, 1, 2 and search modes 1, −1, 2, −2, `if_not_found`), `XMATCH`, `LOOKUP` (vector and array forms, approximate by binary search), `ADDRESS` (abs modes 1–4, A1 and R1C1, quoted sheet text), `AREAS`, `TRANSPOSE`, `HYPERLINK` (value only), `FORMULATEXT`
- [ ] Text: `CHAR`/`CODE` over Windows-1252, `UNICHAR`/`UNICODE` over code points (surrogate pairs whole), `CLEAN`, `FIXED`, `DOLLAR` (per-locale currency codes in the Reference), `NUMBERVALUE`, `T`, `TEXTBEFORE`, `TEXTAFTER` (all six arguments), `VALUETOTEXT`, `ARRAYTOTEXT`; `TEXTSPLIT` is not registered (it spills)
- [ ] Date & Time: `DATEVALUE`, `TIMEVALUE`, `WEEKNUM` (return types 1, 2, 11–17, 21), `ISOWEEKNUM`, `WORKDAY`, `WORKDAY.INTL` and `NETWORKDAYS.INTL` (weekend codes 1–7, 11–17 and seven-character masks, `"1111111"` → `#VALUE!`), `NETWORKDAYS`, `YEARFRAC` (bases 0–4), `DATEDIF` (units `Y`, `M`, `D`, `MD`, `YM`, `YD`), `DAYS360` (US and European methods)
- [ ] Math & Trig: `SUMSQ`, `GCD`, `LCM`, `FACT`, `COMBIN`, `PERMUT`, `QUOTIENT`, `EVEN`, `ODD`, the trigonometric, inverse, reciprocal and hyperbolic functions in the Reference, `DEGREES`, `RADIANS`, `CEILING.MATH`, `FLOOR.MATH`, `CEILING.PRECISE`, `FLOOR.PRECISE`, `ISO.CEILING`, `ROMAN` (classic form), `ARABIC`, `BASE`, `DECIMAL`, `MMULT`, and `SUBTOTAL` and `AGGREGATE` per "SUBTOTAL and AGGREGATE"
- [ ] `SUBTOTAL` function numbers 1–11 skip filter-hidden rows and 101–111 skip every hidden row; both skip cells whose formula calls `SUBTOTAL` or `AGGREGATE`; hidden columns are never skipped; `AGGREGATE` supports function numbers 1–19 and options 0–7; both read `EvalContext.rowVisibility(sheetId, row)` — which this spec adds to SPEC-120's `EvalContext` (`"hidden"` for `Row.hidden`, else `"visible"`; SPEC-128 later answers `"filtered"`) — and are `visibilitySensitive`, and hiding or unhiding rows recalculates every formula cell that calls either (`visibilitySensitiveCells(model)` in `recalc.ts`)
- [ ] Statistical, with every legacy name in the Compatibility category sharing its modern implementation: `MEDIAN`, `MODE.SNGL`, `STDEV.S`/`.P`, `VAR.S`/`.P`, `LARGE`, `SMALL`, `RANK.EQ`/`.AVG`, `PERCENTILE.INC`/`.EXC`, `QUARTILE.INC`/`.EXC`, `PERCENTRANK.INC`/`.EXC`, `CORREL`, `PEARSON`, `RSQ`, `COVARIANCE.P`/`.S`, `SLOPE`, `INTERCEPT`, `STEYX`, `FORECAST.LINEAR`, `NORM.DIST`/`.INV`, `NORM.S.DIST`/`.INV`, `STANDARDIZE`, `T.DIST`/`.2T`/`.RT`, `T.INV`/`.2T`, `BINOM.DIST`, `POISSON.DIST`, `EXPON.DIST`, `CONFIDENCE.NORM`, `AVERAGEA`, `MAXA`, `MINA`, `STDEVA`, `STDEVPA`, `VARA`, `VARPA`, `GEOMEAN`, `HARMEAN`, `AVEDEV`, `DEVSQ`, `KURT`, `SKEW`, `SKEW.P`, `TRIMMEAN`, `FREQUENCY`, `GAMMA`, `GAMMALN`, `GAMMALN.PRECISE`; `COUNTUNIQUE` (Google Sheets only) is not added
- [ ] Financial: `PMT`, `IPMT`, `PPMT`, `FV`, `PV`, `NPER`, `RATE`, `NPV`, `XNPV`, `IRR`, `XIRR`, `MIRR`, `SLN`, `SYD`, `DB`, `DDB`, `CUMIPMT`, `CUMPRINC`, `EFFECT`, `NOMINAL`; `RATE`, `IRR` and `XIRR` iterate exactly per "Numerical methods" (method, starting guess, iteration cap, tolerance) and return `#NUM!` on non-convergence or a non-finite iterate, never a half-converged number
- [ ] Information: `ISEVEN`, `ISODD`, `ISREF`, `ISFORMULA`, `TYPE`, `ERROR.TYPE`, `N`, `SHEET`, `SHEETS` (hidden and very hidden sheets counted), and `CELL` and `INFO` (both `volatile`, as SPEC-120's volatile list names them) with exactly the info types the Reference lists; Engineering: the twelve base conversions, `DELTA`, `GESTEP`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`
- [ ] A valid call this library does not compute — an info type outside the supported subset, `ROMAN` forms 1–4 — throws `UnsupportedCaseError` from `impl`, which SPEC-120's evaluator turns into the new `UncomputableReason` `{ kind: "unsupportedCase"; name }` (stale title suffix "this use of NAME isn't supported yet"), so the cell keeps its cached value exactly as for a function missing from the registry; an uncomputable cell with **no** cached value (a newly typed formula — a case SPEC-120 leaves open) gets the value `#NAME?` and is stale — never a guessed value
- [ ] Dynamic-array functions (`FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `LET`, `LAMBDA` and the `LAMBDA` helpers, `TEXTSPLIT`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `EXPAND`) stay unregistered: a cell from a file keeps its cached value, stale, and its xlsx spill metadata survives verbatim (SPEC-119/120); a newly typed one shows `#NAME?` flagged stale; none is offered by the autocomplete or the dialog
- [ ] Legacy array functions (`TRANSPOSE`, `MMULT`, `FREQUENCY`) fill an array formula's `Cell.arrayRange` (SPEC-119/120); in a plain cell their multi-cell result follows SPEC-120's single-cell rule for an array result, as does an `XLOOKUP` whose `return_array` is two-dimensional
- [ ] Every registered name round-trips through SPEC-119's `formula/functionNames.ts` (`FUNCTION_NAMES`, `xlsxPrefix`, `odsName` — e.g. `_xlfn.XLOOKUP`, `_xlfn.STDEV.S`, `_xlfn.AGGREGATE` on the xlsx wire); a name this spec registers that the table lacks is added from the same two sources SPEC-119 transcribes (Microsoft's future-function list, LibreOffice's `formulabase.cxx`), never from memory; `functionNames.spec.ts` asserts `fromXlsxFormula(toXlsxFormula(t)) === t` and the ODS equivalent for one call of every registered function
- [ ] Expected values come only from the sources in "Conformance tables": one data file per category under `formula/conformance/`, labelled `Excel` / `pinned` as SPEC-120 labels its rows (plus `reference` for published numerical tables), at least three rows per registered function (SPEC-120's included), at least one an error case, an `Excel` row wherever Microsoft's page carries an example table, and no expected value produced by running this engine or another implementation; SPEC-120's `formula/conformance.spec.ts` runs every row through the real parser, evaluator and recalculation and fails naming the row
- [ ] Autocomplete per "The assists": while editing a formula in the cell editor or the formula bar (SPEC-122), an identifier being typed where a function may appear opens `function-autocomplete` (`role="listbox"`) listing registered names that start with it, alphabetically, each a `function-autocomplete-item` with `data-name`, and the highlighted item's summary in `function-autocomplete-summary`; Up/Down move the highlight, Tab or Enter inserts `NAME(` replacing the typed prefix, Escape closes the list only, and it never opens inside a string literal
- [ ] The argument hint (`function-arg-hint`, `role="tooltip"`) shows `functionSignature` of the innermost enclosing call of a registered function with the current argument in `function-arg-current` (bold), following nesting: commas inside nested calls, parentheses, array constants and quoted sheet names never advance the outer argument
- [ ] The Insert Function dialog (`insert-function-dialog`, on `BaseDialog.vue`) opens from Shift+F3 (`sheet.insertFunction` in SPEC-123's `SHEET_SHORTCUTS`, `src/views/spreadsheet/sheetShortcuts.ts`) and from an "fx" button on the formula bar (`formula-bar-fx`, `aria-label` "Insert function" — the guaranteed path, since Firefox keeps Shift+F3 for its own find-previous), with a search field (`insert-function-search`), a category select (`insert-function-category`: "All", then every category holding a registered function, in `FUNCTION_CATEGORY_ORDER`), a list (`insert-function-list` of `insert-function-item`), a detail pane (`insert-function-detail`: signature, summary, every argument's description) and "Insert" (`insert-function-insert`), which puts `=NAME(` into the editor per "The assists"; on a cell SPEC-122 will not let the user edit, the button and Shift+F3 do nothing and the button is disabled with the title "This cell is protected"
- [ ] View wiring: `src/views/spreadsheet/useFunctionAssist.ts` (≤ 400 lines under `composableSize.spec.ts`) owns all three assists; `FunctionAutocomplete.vue`, `FunctionArgHint.vue` and `InsertFunctionDialog.vue` live in `src/components/spreadsheet/`; SpreadsheetView gains only the `useFunctionAssist({ … })` call and the template hookups, and `SpreadsheetView.size.spec.ts`'s budget never rises
- [ ] Tests: `formula/registry.spec.ts`, `formula/functionHelp.spec.ts`, `formula/functions/numeric.spec.ts`, `formula/conformance.spec.ts` over all category files, `functionNames.spec.ts` additions, `views/spreadsheet/useFunctionAssist.spec.ts`, `FunctionAutocomplete.spec.ts`, `FunctionArgHint.spec.ts`, `InsertFunctionDialog.spec.ts`
- [ ] e2e (`e2e/SPEC-127.spec.ts`): open `mini-book.xlsx` → select the empty cell `J1` → type `=XL` → `function-autocomplete` lists `XLOOKUP` → Tab → the editor reads `=XLOOKUP(` and `function-arg-current` reads `lookup_value` → type `"x",` → it reads `lookup_array` → type `ROUND(` → the hint shows `ROUND` with `number` current → Escape cancels the edit → click `formula-bar-fx` → choose category Financial → search `payment` → select `PMT` → Insert → the editor reads `=PMT(` → replace the text with `=ROUND(PMT(0.05/12,60,10000),2)` → Enter → `J1` reads "-188.71" → in `J2` type `=SEQUENCE(3)` → Enter → `J2` reads "#NAME?" with `data-stale="true"` → Shift+F3 opens the dialog and no list item is `SEQUENCE`

## Out of scope
Dynamic arrays and spilling — every function above that would spill is either
legacy-array only or unregistered (whole-queue backlog). Database functions
(`DSUM`, `DGET`, …), cube and web functions, `GETPIVOTDATA` (SPEC-139), bond
and coupon functions (`PRICE`, `YIELD`, `ACCRINT`, `COUP*`), `VDB`, complex
numbers (`IM*`), `CONVERT`, bitwise functions, the chi-square, F, beta,
gamma (`GAMMA.DIST`), lognormal, Weibull, hypergeometric and
negative-binomial distributions,
`MINVERSE`/`MDETERM`/`MUNIT`, `TREND`/`GROWTH`/`LINEST`/`LOGEST`, `MODE.MULT`
— all backlog, and stale until then. Clicking `HYPERLINK` (SPEC-135). Defined
and table names in the autocomplete (SPEC-133 adds them to the same list).
Live argument evaluation inside the Insert Function dialog (Excel's argument
boxes). Localised function names and `;` separators (whole-queue backlog).
The format inference Excel applies to a cell from its function (`PMT` making
a cell currency).

## Reference

### Dependencies restated
- **`specs/120-…`** — `formula/registry.ts`: `ArgKind = "value" | "range" |
  "any" | "array"`; `FunctionCategory = "Financial" | "Date & Time" |
  "Math & Trig" | "Statistical" | "Lookup & Reference" | "Text" | "Logical" |
  "Information" | "Engineering" | "Compatibility"` (the values are the
  display names); `FunctionSpec { name; category; minArgs; maxArgs /* 255
  for variadic */; argKinds: ArgKind[]; repeat?: { from: number; count:
  number }; volatile?: true; lazy?: true; acceptsErrors?: true;
  impl(args: ArgList, ctx: EvalContext): EvalValue; summary: string; args:
  { name; description; optional?: true; repeating?: true }[] }`;
  `registerFunction(spec)` (throws on a duplicate), `getFunction(name)`,
  `registeredFunctionNames()`; "SPEC-127 registers every further function
  through `registerFunction` and may add optional metadata fields". The 100
  core functions by file — `math.ts` (SUM, SUMIF(S), SUMPRODUCT, PRODUCT, the
  ROUND family, INT, TRUNC, MOD, ABS, SIGN, SQRT, POWER, EXP, LN, LOG, LOG10,
  PI, CEILING, FLOOR, MROUND, RAND, RANDBETWEEN), `aggregate.ts`,
  `logical.ts`, `information.ts`, `text.ts`, `dateTime.ts`, `lookup.ts`.
  `EvalContext { model; host: { sheetId; row; col }; origin; arrayMode;
  settings: EvalSettings { locale; dateSystem; caseSensitive; criteria;
  wholeCell }; now(); random(); valueAt(…); recordDependency(area);
  resolveName? }`; `formula/evaluate.ts`, `values.ts`, `registry.ts`,
  `functions/*.ts` import no Vue, `window`, `document` or scheduler and
  contain no `eval`/`new Function` (`formula/purity.spec.ts`). Errors
  propagate unless a function declares `acceptsErrors`. Uncomputable cells:
  `UncomputableReason` (`function`, `name`, `structured`, `external`,
  `dynamicArray`, `unparsed`, `iterative`, `regex`, `dataTable`) — the cell
  keeps its cached value, is stale (`data-stale="true"`, `title` "This value
  couldn't be recalculated here — " + a suffix) when a dependent of a change,
  and sets `model.forceFullCalcOnLoad`. Conformance rows live in
  `formula/conformance.spec.ts`, labelled **Excel** (Microsoft's documented
  behaviour and examples) or **pinned** (hand-computed from a stated rule),
  never pasted from the engine's own output; `numbersOracle.spec.ts`
  cross-checks against an Apple Numbers export. Volatile: NOW, TODAY, RAND,
  RANDBETWEEN, OFFSET, INDIRECT, and CELL/INFO when registered here.
- **`specs/119-…`** — canonical formula text (en-US syntax, upper-case names
  without `_xlfn.`); `formula/tokens.ts` `tokenize(text): Token[]` with
  trivia; `parseFormula(text): FormulaNode`; `formula/functionNames.ts`, the
  canonical ↔ xlsx ↔ ODS name table; `fromXlsxFormula`/`toXlsxFormula`/
  `fromOdsFormula`/`toOdsFormula`; `Cell.arrayRange?: GridRange` on an array
  formula's anchor.
- **`specs/122-…`** — `CellEditor.vue` and the editable `FormulaBar.vue`,
  driven by `src/views/spreadsheet/useCellEditing.ts`: an editing session
  with its text and caret, Enter/Tab commit, Escape cancel, point mode on the
  arrow keys, and the protection rule (a cell whose effective style has
  `protection.locked` not `false` on a protected sheet cannot be edited).
- **SPEC-118** — `formatValue(value, code, ctx: FormatContext)`,
  `SpreadsheetLocale` (`en-US`, `en-GB`, `de-DE`, `fr-FR`, `es-ES`,
  `cs-CZ`); formulas are always en-US syntax; locale affects display and
  value parsing only. `parseInput(text, locale, dateSystem)` (SPEC-122).
- **SPEC-117** — `Row.hidden?: true` (user-hidden). This spec adds
  `EvalContext.rowVisibility(sheetId, row): "visible" | "filtered" |
  "hidden"`, returning `"hidden"` for `Row.hidden` and `"visible"` otherwise;
  SPEC-128 (later) adds `Row.filterHidden?: true` and widens the same function
  to answer `"filtered"`, so `SUBTOTAL` 1–11 needs no change then.
- **SPEC-123** — `SHEET_SHORTCUTS` in `src/views/spreadsheet/sheetShortcuts.ts`,
  matched through `src/lib/shortcuts.ts`'s `findShortcut(event, definitions)`,
  where `sheet.insertFunction` is registered.

### Function metadata
```ts
// src/lib/spreadsheet/formula/registry.ts (additions; SPEC-120 owns the rest)
export interface FunctionSpec {
  /* …SPEC-120 fields… */
  /** The result depends on row visibility (SUBTOTAL, AGGREGATE). */
  visibilitySensitive?: true;
}
/** SPEC-120's FunctionCategory values (already display names), in dialog order. */
export const FUNCTION_CATEGORY_ORDER: readonly FunctionCategory[] = [
  "Financial", "Date & Time", "Math & Trig", "Statistical", "Lookup & Reference",
  "Text", "Logical", "Information", "Engineering", "Compatibility",
];
/** Thrown by impl for a valid call this library does not compute. */
export class UnsupportedCaseError extends Error {}
// formula/evaluate.ts: UncomputableReason gains { kind: "unsupportedCase"; name: string }
```
Invariants `registry.spec.ts` asserts for every spec: the name matches
`/^[A-Z][A-Z0-9]*(\.[A-Z0-9]+)*$/`; `summary` is 10–100 characters, starts
with a capital and ends with "."; arg names match `/^[a-z][a-z0-9_]*$/` and
are unique; every description is 5–120 characters ending in "."; required
args precede optional ones and `minArgs` equals the number of required args;
the `repeating` args, if any, are the trailing run of `args`, all optional,
each name ending in a digit, and then `maxArgs ≤ 255`; with none,
`maxArgs === args.length`.
Legacy names (`STDEV`, `VAR`, `PERCENTILE`, `QUARTILE`, `RANK`, `MODE`,
`NORMDIST`, `NORMINV`, `NORMSDIST`, `NORMSINV`, `TDIST`, `TINV`, `BINOMDIST`,
`POISSON`, `EXPONDIST`, `CONFIDENCE`, `COVAR`, `FORECAST`, `PERCENTRANK`,
`STDEVP`, `VARP`) are category `"Compatibility"`, as Excel files them.

Summaries and descriptions are written in this repo's own words — Microsoft's
text is copyrighted and is not copied — and argument names follow Excel's
(`lookup_value`, `sum_range`), which are interface facts. With *f* the
index of the first `repeating` arg and *g* the number of them, argument
index *i* ≥ *f* displays as `args[f + (i − f) mod g]` with its trailing
integer increased by ⌊(i − f) ÷ g⌋. `functionSignature` shows the repeating
group once, bracketed, followed by `, …`. (SPEC-120's `repeat` governs
argument *kinds*; `repeating` governs names — both describe the same group.)

### Where semantics come from
In order: (1) Microsoft's reference page for the function — its examples are
test rows; (2) where that page is silent, the OASIS OpenFormula specification
(ODF 1.2 Part 2), which defines Excel-compatible behaviour precisely
(`DAYS360`, `YEARFRAC` bases, `WEEKNUM` systems, `DB`'s rounding); (3) the
explicit pins below, which win where stated. Where (1) and (2) disagree,
Microsoft's examples win and the row's `note` records the disagreement.

### Function pins
- **Lookup.** `LOOKUP` and `XLOOKUP`/`XMATCH` binary modes use SPEC-120's
  approximate-match search (largest value ≤ target, case-insensitive text);
  unsorted data gives what Excel's binary search gives, untested. `XLOOKUP`
  with mismatched array sizes → `#VALUE!`; not found without `if_not_found`
  → `#N/A`. `ADDRESS` R1C1 forms `R2C3`, `R2C[3]`, `R[2]C3`, `R[2]C[3]`;
  `sheet_text` quoted when SPEC-119 would quote it. `AREAS` counts the areas
  of a union reference. `FORMULATEXT` → `"=" + Cell.formula`, `#N/A` for a
  constant. `HYPERLINK(link, [friendly_name])` → `friendly_name` if given,
  else `link`.
- **Text.** `CHAR(1…255)` and `CODE` map through Windows-1252 (Excel for
  Windows and the web), `CODE` of an unmappable character is 63. `UNICHAR` of
  0 or above U+10FFFF → `#VALUE!`, of a surrogate → `#N/A`. `CLEAN` removes
  U+0000–U+001F. `FIXED(number, [decimals=2], [no_commas])` and
  `DOLLAR(number, [decimals=2])` round half away from zero and render through
  `formatValue(…, { locale: ctx.settings.locale, dateSystem })` as SPEC-120's
  `TEXT` does; `DOLLAR`'s codes (decimals substituted; symbols and
  positions agree with SPEC-118's `currencySymbol`/`currencyPosition`):
  en-US `$#,##0.00;($#,##0.00)`, en-GB `£#,##0.00;-£#,##0.00`,
  de-DE/fr-FR/es-ES `#,##0.00 €;-#,##0.00 €`, cs-CZ `#,##0.00 Kč;-#,##0.00
  Kč`. `NUMBERVALUE` defaults its separators to the
  locale's. `ARRAYTOTEXT` concise joins with `", "`; strict is `{…}` with `,`
  between columns, `;` between rows and strings quoted.
- **Dates.** `DATEVALUE`/`TIMEVALUE` parse through `parseInput` in the
  locale, keeping the integer / fractional part; anything that is not a
  date/time → `#VALUE!`. `WORKDAY.INTL`/`NETWORKDAYS.INTL` weekend codes:
  1 Sat–Sun, 2 Sun–Mon, … 7 Fri–Sat, 11 Sun only, … 17 Sat only, or a
  Monday-first mask of `0`/`1`. `YEARFRAC` basis 1 (actual/actual): if the
  period is at most a year, days ÷ (366 when it lies in one leap year or
  spans a 29 February, else 365); otherwise days ÷ the average year length
  over the calendar years it touches. Basis 0 is OpenFormula's US 30/360
  (February month-end rules included); `DAYS360` method FALSE follows
  Microsoft's page, TRUE is European. `DATEDIF` with start > end → `#NUM!`;
  `MD` is computed as Microsoft documents it and its rows are hand-derived.
- **Math.** The trigonometric set is `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`,
  `ATAN`, `ATAN2`, `SINH`, `COSH`, `TANH`, `ASINH`, `ACOSH`, `ATANH`, `COT`,
  `CSC`, `SEC`, `ACOT`, with the domain limits Microsoft's pages state
  (`COT`/`CSC`/`SEC` of |x| ≥ 2^27 → `#NUM!`). `ATAN2(x_num, y_num)` keeps
  Excel's argument order; `ATAN2(0,0)` → `#DIV/0!`. `FACT` above 170 and
  `GCD`/`LCM` operands ≥ 2^53 → `#NUM!`. `EVEN`/`ODD` round away from zero.
  `ROMAN(number, [form])` supports form 0/`TRUE` only.
- **Statistics.** Variance and standard deviation use the two-pass algorithm
  (mean, then squared deviations) for stability. `PERCENTRANK.*` truncate to
  `significance` digits (default 3). `MODE.SNGL` returns the first value, in
  argument order, with the highest count; none repeated → `#N/A`. `-A`
  variants count text in ranges as 0 and `TRUE` as 1.
- **Financial.** `DB` rounds its rate to three decimals (Microsoft's rule).
  `XNPV`/`XIRR` truncate dates; a date before the first → `#NUM!`.
- **Information.** `TYPE`: 1 number or blank, 2 text, 4 logical, 16 error,
  64 array. `ERROR.TYPE`: `#NULL!` 1, `#DIV/0!` 2, `#VALUE!` 3, `#REF!` 4,
  `#NAME?` 5, `#NUM!` 6, `#N/A` 7, `#GETTING_DATA` 8, `#SPILL!` 9,
  `#CONNECT!` 10, `#BLOCKED!` 11, `#UNKNOWN!` 12, `#FIELD!` 13, `#CALC!` 14
  (confirm 9–14 against Microsoft's page when writing the table), else `#N/A`.
  `CELL` supports `"address"` (references on the host sheet), `"col"`,
  `"row"`, `"contents"`, `"type"`, `"filename"` (`[<title>.<ext>]<sheet>` —
  no folder exists here, and the `MID(CELL("filename",A1),…)` sheet-name
  idiom keeps working), `"protect"` and `"prefix"`; `INFO` supports
  `"numfile"` (grid sheets in the workbook) and `"recalc"` (`"Automatic"`).
  Every other info type throws `UnsupportedCaseError` — `INFO("system")`
  included, since the engine may not read the platform (`purity.spec.ts`) and
  a value that differs per machine would differ per save. `TYPE` and
  `ERROR.TYPE` declare `acceptsErrors`; `ISREF`, `ISFORMULA`, `AREAS`,
  `FORMULATEXT`, `SHEET`, `SHEETS` and `CELL`'s reference take `range`/`any`
  arguments.
- **Engineering.** Base conversions use 10-character two's complement as
  Microsoft documents (`DEC2BIN` −512…511, `DEC2OCT` −2^29…2^29−1,
  `DEC2HEX` −2^39…2^39−1); too few `places` → `#NUM!`.

### SUBTOTAL and AGGREGATE
| `function_num` | 1/101 | 2/102 | 3/103 | 4/104 | 5/105 | 6/106 | 7/107 | 8/108 | 9/109 | 10/110 | 11/111 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| function | AVERAGE | COUNT | COUNTA | MAX | MIN | PRODUCT | STDEV | STDEVP | SUM | VAR | VARP |

Any other number → `#VALUE!`. `AGGREGATE` adds 12 `MEDIAN`, 13
`MODE.SNGL`, and the array form 14 `LARGE`, 15 `SMALL`, 16
`PERCENTILE.INC`, 17 `QUARTILE.INC`, 18 `PERCENTILE.EXC`, 19
`QUARTILE.EXC`; options 0/1/2/3 ignore nested `SUBTOTAL`/`AGGREGATE`, plus
hidden rows (1, 3, 5, 7) and errors (2, 3, 6, 7); 4 ignores nothing. Hidden
rows are skipped only inside references, never inside computed arrays. A cell
is "nested" when its formula's tokens contain a call to either function,
memoised per formula string. `visibilitySensitiveCells(model): CellKey[]`
returns every formula cell calling a `visibilitySensitive` function, through
that memo; SPEC-124's hide/unhide hands them to `recalculate`, and SPEC-128's
filter changes will do the same.

### Numerical methods
| Function(s) | Method | Start | Max iterations | Stop when |
| --- | --- | --- | --- | --- |
| `RATE` | Newton, analytic derivative | `guess` (0.1) | 20 | \|Δrate\| < 1e-7 |
| `IRR` | Newton on NPV | `guess` (0.1) | 20 | \|Δrate\| < 1e-7 |
| `XIRR` | Newton on XNPV | `guess` (0.1) | 100 | \|Δrate\| < 1e-8 |
| `T.DIST*`, `TDIST` | regularised incomplete beta, modified Lentz continued fraction | — | 300 | relative change < 1e-15 |
| `T.INV*`, `TINV` | bisection bracketing then Newton on `T.DIST` | — | 100 | \|Δx\| < 1e-12 |

`RATE`/`IRR`/`XIRR` limits are the ones Microsoft documents. Exhausting the
cap, a zero or non-finite derivative, or a rate ≤ −1 → `#NUM!`. Closed forms:
`erf`/`erfc` by Cody's rational Chebyshev approximations (netlib `specfun`
CALERF); `NORM.S.INV` by Acklam's rational approximation plus one Halley
step on `erfc`; `GAMMALN` by Lanczos (g = 7, nine coefficients); `BINOM.DIST`
and `POISSON.DIST` probabilities in log space through `GAMMALN`, cumulative
forms as sums of the point probabilities. All live in
`formula/functions/numeric.ts`, framework-free and tested alone against
published tables (Abramowitz & Stegun, NIST DLMF).

### Conformance tables
```ts
// src/lib/spreadsheet/formula/conformance/<category>.ts
export interface ConformanceRow {
  formula: string;                       // canonical, no "=": "PMT(0.08/12,10,10000)"
  cells?: Record<string, CellValue>;     // inputs on Sheet1, by A1 address
  expected: CellValue;                   // { error: "#NUM!" } for an error row; the anchor's value for an array row
  array?: string;                        // entered as a legacy array formula over this range ("D1:E2")
  expectedArray?: CellValue[][];         // every member's value, row-major, when `array` is set
  digits?: number;                       // significant digits compared (default 15, via roundTo15)
  source: "Excel" | "pinned" | "reference";
  note: string;                          // Microsoft's page title, the derivation, or the table cited
}
```
Files: `lookup.ts`, `text.ts`, `date.ts`, `math.ts`, `statistical.ts`,
`financial.ts`, `information.ts`, `engineering.ts`, `logical.ts`. `Excel` and
`pinned` mean what SPEC-120 says; `reference` is a published numerical table
(Abramowitz & Stegun, NIST DLMF) for the special functions. Microsoft's pages
print rounded results, so an `Excel` row sets `digits` to what the page
shows; a `pinned` row's `note` carries the arithmetic. No expected value is
produced by running this engine, LibreOffice, Google Sheets or a JavaScript
library — a copied bug is still a bug, and agreeing with ourselves
proves nothing. SPEC-120's Numbers oracle may gain rows for functions Numbers
implements as a cross-check; it never supplies a conformance row.

### The assists
`useFunctionAssist({ editing, text, caret, replaceRange, startEditing,
canEditActiveCell })` — the editing session's text and caret and a
`replaceRange(from, to, insert, caretAfter)` from SPEC-122's
`useCellEditing` (added there if absent). It handles keydown **before**
SPEC-122's editor: while the list is open, Up/Down, Tab, Enter and Escape are
consumed, so Enter does not commit and arrows do not enter point mode.
- **Completion context.** The text starts with `=`; tokenizing up to the
  caret, the caret ends an identifier token of at least one letter that is
  not inside a string literal, and the previous significant token is the
  leading `=`, an operator, `(` or `,`. An unterminated string or a
  tokenizer error means "inside a string": no list, no hint. Matches:
  `registeredFunctionNames()` starting with the prefix, case-insensitively, at
  most ten rows visible with scrolling. Typing refilters; no match closes.
- **Insertion.** The prefix is replaced by `NAME(` and the caret placed
  after `(`; if the next character is already `(`, only the name is
  completed and the caret steps past it.
- **The hint.** `callAtCaret(text, caret): { name: string; argIndex: number }
  | null` in `functionHelp.ts` walks the tokens before the caret with a frame
  stack: `(` after a function-name token pushes a function frame, any other
  `(` a plain frame, `{` an array frame; `,` advances `argIndex` only in a
  function frame at the top; `)`/`}` pop. The answer is the nearest function
  frame whose name is registered. `argumentAt` decides the bold argument;
  beyond `maxArgs` nothing is bold. The hint sits below the editor and is
  hidden while the list is open.
- **The dialog.** Search matches, case-insensitively, name prefix, then name
  substring, then every query word in the summary, ranked in that order and
  alphabetically within each; the empty query lists the category
  alphabetically; the search field has focus on open; Up/Down move, Enter
  inserts. Insert while editing a formula replaces the selection at the
  caret with `NAME(`; otherwise it starts editing the active cell with the
  text `=NAME(`, replacing its content, as Excel does.
- **Accessibility.** The editor gets `aria-autocomplete="list"`,
  `aria-expanded`, `aria-controls` and `aria-activedescendant` while the
  list is open, and `aria-describedby` the hint while it shows. Every
  string, separator and argument name is en-US whatever the locale.

### Tests
- `registry.spec.ts` — the metadata invariants; `FUNCTION_CATEGORY_ORDER`
  lists every `FunctionCategory` exactly once; legacy names in
  `"Compatibility"`; every name in `functionNames.ts`.
- `functionHelp.spec.ts` — signatures for fixed, optional, single- and
  double-group repeats; `argumentAt`; `callAtCaret` over nesting, plain
  parentheses, unions, array constants, `'a,b'!A1`, strings containing `,`
  and `(`, unterminated strings and empty text; completion contexts.
- `numeric.spec.ts` — `erf`, `erfc`, `GAMMALN`, the incomplete beta and the
  normal inverse against published values to 1e-13 relative or better, and
  each iteration cap tripped by a constructed non-convergent input.
- `formula/conformance.spec.ts` (SPEC-120's) — every row of every category
  file; the ≥ 3 rows / error row / `Excel` row completeness rules.
- `useFunctionAssist.spec.ts`, `FunctionAutocomplete.spec.ts`,
  `FunctionArgHint.spec.ts`, `InsertFunctionDialog.spec.ts`,
  `e2e/SPEC-127.spec.ts`.

### Specs this changes
- `specs/120-…` — `FunctionSpec` gains the optional `visibilitySensitive`;
  `registry.ts` gains `FUNCTION_CATEGORY_ORDER` and `UnsupportedCaseError`;
  `UncomputableReason` gains `unsupportedCase` and its stale-title row;
  `EvalContext` gains `rowVisibility(sheetId, row)`;
  `recalc.ts` gains `visibilitySensitiveCells`; an uncomputable cell with no
  cached value is `#NAME?`; `conformance.spec.ts` gains the category files;
  its core functions' metadata is brought up to "Function metadata" here
  wherever it falls short.
- `specs/123-…` — `SHEET_SHORTCUTS` gains `sheet.insertFunction`
  (Shift+F3).
- `specs/119-…` — `functionNames.ts` gains any name registered here that
  its transcription lacks.
- `specs/122-…` — the editors host the assists and yield keydown to them
  while the list is open.
- `specs/124-…` — hide/unhide recalculates `visibilitySensitiveCells`.
- `specs/128-…` (later) — widens this spec's `rowVisibility` to answer
  `"filtered"`, the value `SUBTOTAL` 1–11 skips.
- `specs/133-…` — defined and table names join the autocomplete list.
- `specs/135-…` — `HYPERLINK`'s value is computed here; 135 makes it a link.
