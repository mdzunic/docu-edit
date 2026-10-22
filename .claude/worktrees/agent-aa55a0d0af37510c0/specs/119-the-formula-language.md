---
id: SPEC-119
title: The formula language — parse, print and translate formulas between both formats
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-116, SPEC-117, SPEC-118]
---
## Why
After SPEC-116/117 a workbook opens, and every formula cell shows the value
its producer cached — but the formula itself is thrown away at read. Click
the total under a column and the formula bar shows `60`, not `=SUM(B2:B4)`;
there is no way to see how a sheet computes anything, and nothing later in
this queue — recalculation, editing, copy and fill, inserting a row — has a
formula to work on. A spreadsheet without its formulas is a table of numbers.

The two formats do not even agree on how a formula is spelled. `.xlsx` stores
Excel's own syntax without the `=`, with future functions disguised as
`_xlfn.CONCAT` and LAMBDA parameters as `_xlpm.x`, and with one master text
standing for a whole block of cells (`t="shared"`). `.ods` stores OpenFormula:
`of:=SUM([.A1:.A3];[$Sheet2.B1])`, semicolons, bracketed references, `~` for
union, `!` for intersection, `{1;2|3;4}` for arrays, Excel 2010+ functions
renamed `COM.MICROSOFT.*` — plus the legacy `oooc:` and Excel's own `msoxl:`
dialect for files written by older tools. A reader that gets any of this
wrong silently changes what a workbook means.

**This file is the normative source for the formula language: canonical
formula text (`Cell.formula`), the token stream and its whitespace trivia,
the AST (`FormulaNode`), the reference model (`ReferenceNode`, `RefPart`),
operator precedence, the token-level reference rewrite every later spec
builds on (`rewriteReferences`, `shiftFormula`, `renameSheetInFormula`), the
xlsx and ODS dialect translation, shared/array/data-table/dynamic-array
formula reading, and the parse cache keyed by a formula's relative form.**
SPEC-120 evaluates what this spec parses; SPEC-122 writes it back; SPEC-124
and SPEC-125 move references with the primitive pinned here; SPEC-133 gives
names and structured references their meaning.

It evaluates nothing. Values stay the producer's cached values until
SPEC-120. The visible change is small and honest: the formula bar shows
`=…` for a formula cell, and View ▸ Show formulas (Ctrl+`) shows formula text
in the grid.

## Acceptance criteria
- [ ] `src/lib/spreadsheet/model.ts` gains, per "Model additions", the optional `Cell` fields `formula` (canonical text, no leading `=`), `formulaUnparsed`, `arrayRange`, `dynamicArray`, `cm` and `dataTableXml`, and `Sheet.showFormulas`; every field is optional, the model stays plain JSON (`JSON.parse(JSON.stringify(model))` `toEqual`s it), and a formula cell's `value` is still its cached result
- [ ] `formula/tokens.ts` exports `tokenize(text): Token[]` and `joinTokens(tokens): string` per "Tokens": tokenizing is **total** — it never throws, unknown characters become `unknown` tokens — every token carries its preceding whitespace as `lead` trivia, and `joinTokens(tokenize(t)) === t` holds exactly for every string in the corpus (see "Tests") and for 10,000 seeded random strings over the formula alphabet
- [ ] The tokenizer classifies per "Token rules": a cell, range, whole-column (`A:A`, `$A:$C`), whole-row (`1:1`), sheet-qualified (`Sheet1!A1`, `'My Sheet'!A1:B2`, `'O''Brien'!A1`) and 3-D (`Sheet1:Sheet3!A1`, `'Q1 2026:Q4 2026'!C3`) reference is **one** `ref` token carrying its parsed `ReferenceNode`; an identifier immediately followed by `(` is a `function` token even when it spells a cell (`LOG10(`, `ATAN2(`, `DAYS360(`); `TRUE`/`FALSE` not followed by `(` are `boolean`; all fifteen `ErrorCode`s are `error` tokens, including the reference forms `Sheet1!#REF!` and `#REF!A1`; `[1]Sheet1!A1` and `[1]!Name` are one opaque `external` token; `Table1[[#This Row],[Amount]]` and `[@Qty]` are one `structured` token consumed by balanced brackets with `'` escapes; a space between a reference-ending and a reference-starting token is the intersection `operator`; a comma is `argSep` directly inside a function call, `union` elsewhere, and `arrayColSep` inside `{…}`
- [ ] `formula/parse.ts` exports `parseFormula(text): FormulaNode` building the `FormulaNode` union in `formula/ast.ts` per "The AST", and `FormulaSyntaxError` (`offset` = UTF-16 index of the offending token); nesting deeper than `MAX_PARSE_DEPTH = 256` is a `FormulaSyntaxError`, never a stack overflow; every node records its token span (`from`, `to`) so rewrites can map nodes to tokens
- [ ] Operator precedence is Excel's, per the "Precedence" table and asserted node-for-node: `:` > intersection > union > prefix `-`/`+` > postfix `%` > `^` > `*` `/` > `+` `-` > `&` > comparisons, all binary operators left-associative — so `-2^2` parses as `(-2)^2` (Excel's well-known result 4), `2^-2` as `2^(-2)`, `2^3^2` as `(2^3)^2`, `2^3%` as `2^(3%)`, `-A1:A3` as `-(A1:A3)`, and `1-2-3` as `(1-2)-3`
- [ ] Literals parse per "Literals": numbers `1`, `1.5`, `.5`, `1.`, `1E3`, `1e-3`, `1.5E+10` (an exponent overflowing a double is a syntax error); strings with `""` escapes; case-insensitive booleans; error literals; array constants `{1,2;3,4}` and `{"a",TRUE;#N/A,-1.5}` whose elements are signed numbers, strings, booleans or errors only, and whose rows must be equal length (`{1,2;3}` is a syntax error)
- [ ] Function calls keep missing arguments as `missing` nodes (`IF(A1,,1)` has three arguments, the middle one missing; `F()` has none; `IFERROR(1/0,)` has two), names parse as `name` nodes (with an optional sheet scope for `Sheet1!Rate`), and structured and external tokens parse as `structured`/`external` nodes that keep their source text — parsed, never interpreted here
- [ ] `formula/print.ts` exports `printFormula(node)` and `formatReference(ref)` producing canonical text (no whitespace but the intersection space, minimal parentheses, sheet names quoted per "Sheet-name quoting"); for every corpus formula, `parseFormula(printFormula(parseFormula(t)))` deep-equals `parseFormula(t)` ignoring token spans; print is a test oracle and a builder for SPEC-125/140 — **no transformation in this queue is implemented as `print(parse())`**, because that discards the user's whitespace
- [ ] `formula/references.ts` exports `RefPart`, `ReferenceNode`, `rewriteReferences(text, rewrite)` (the token-level primitive: only rewritten `ref` tokens change, every other token and all trivia are copied verbatim), and on top of it `shiftFormula(text, dRow, dCol)` (relative parts move, absolute stay; a reference any part of which leaves the grid becomes `#REF!`, keeping its sheet prefix as `Sheet2!#REF!`), `renameSheetInFormula(text, from, to)` (case-insensitive match, 3-D endpoints included, re-quoted per the quoting rule) and `relativeFormulaKey(text, host)` per "The relative key"
- [ ] `parseFormulaAt(text, host)` in `formula/parse.ts` caches parsed ASTs in an LRU of `FORMULA_AST_CACHE_SIZE = 50_000` entries keyed by `relativeFormulaKey`, returning `{ node, origin }` — so the 100,000 cells of one shared formula share **one** AST (asserted by cache size), and cached nodes are never mutated
- [ ] `formula/functionNames.ts` exports the name table per "Function names": one entry per function whose wire name differs from its canonical name in either format, transcribed from the two primary sources named there (not from memory), with lookup helpers both ways for each format; a name absent from the table is identical in all three spellings
- [ ] `formula/dialect.ts` translates xlsx per "xlsx dialect": `fromXlsxFormula(wire)` strips `_xlfn.` and `_xlfn._xlws.` from function names and `_xlpm.` from names, upper-cases function names and drops a stray leading `=`; `toXlsxFormula(canonical)` restores each prefix from the table and `_xlpm.` on LET/LAMBDA parameter names by the lexical scope rule; for every Excel-shaped corpus formula `toXlsxFormula(fromXlsxFormula(w)) === w`, and for every canonical corpus formula `fromXlsxFormula(toXlsxFormula(c)) === c`
- [ ] `formula/dialect.ts` translates ODS per "ODS dialect": `fromOdsFormula(wire, namespaces)` recognises the `of:`, `oooc:` and `msoxl:` dialects by namespace URI (unprefixed text is OpenFormula), maps every bracketed reference form in the table to canonical, `;` to `,`, `~` to a parenthesised `,` union, `!` to the intersection space, `{1;2|3;4}` to `{1,2;3,4}` and ODS function names through the table; `toOdsFormula(canonical)` is the inverse and emits `of:=`; for every LibreOffice-shaped corpus formula `toOdsFormula(fromOdsFormula(w)) === w` except the forms listed under "Known non-identities"
- [ ] A formula that does not parse after translation keeps its **wire** text verbatim in `Cell.formula` with `formulaUnparsed: true` (xlsx: the `<f>` text; ods: the whole `table:formula` value, prefix included); it is shown, never evaluated, never rewritten by `shiftFormula`/`renameSheetInFormula`/SPEC-124, and re-emitted verbatim only in its own workbook's format
- [ ] The `.xlsx` reader reads `<f>` per "Reading formulas from .xlsx": normal formulas; shared formulas expanded per follower cell by `shiftFormula` of the master's canonical text (a follower met before its master is resolved at the end of the sheet; an out-of-grid shift yields `#REF!` in the text, as Excel does); `t="array"` → anchor `formula` + `arrayRange` (members keep cached values only); `t="dataTable"` → anchor `dataTableXml` (the `<f>` element verbatim); `c@cm` → `cm`, and `t="array"` with `cm` → `dynamicArray: true`; `sheetView@showFormulas="1"` → `Sheet.showFormulas` — and every formula cell in `spec116-numbers.xlsx` reads without `formulaUnparsed`
- [ ] The `.ods` reader reads `table:formula` per "Reading formulas from .ods": the dialect prefix resolved against `content.xml`'s namespace declarations; `table:number-matrix-columns-spanned`/`-rows-spanned` → `arrayRange`; a repeated formula cell copies the formula text **verbatim** to each repetition (ODF formula text names absolute addresses — it is never shifted)
- [ ] The readers bump `SPREADSHEET_READER_VERSION` (to the next integer) and append their `ReaderUpgrade`; a stored record from the previous version re-reads its (necessarily clean) sheets from the retained source and gains its formulas, and `upgradeDirtySheet` throws `Error("SPEC-119 upgrade: a sheet cannot be dirty before SPEC-121")` per "Reader version"; `open.spec.ts` drives the path with a synthetic previous-version record
- [ ] An untouched workbook still downloads with every part `equalBytes` the source (`roundTrip.spec.ts` unchanged and green over the new fixtures) — reading formulas changes the model, never the bytes
- [ ] The formula bar (`formula-bar-content`, SPEC-117) shows `formulaBarText(sheet, row, col, style, ctx)` per "The formula bar": `=` + formula for a formula cell; `{=…}` for every cell of a legacy array range, members included; the anchor's formula, muted, with `data-spill-member="true"` for a dynamic-array member; `{=TABLE(A1,A2)}`-style text for a data-table range; an unparsed formula as stored; SPEC-118's `editText` for every other cell — the bar stays read-only (editing is SPEC-122)
- [ ] SPEC-117's View menu gains "Show formulas" (`sheet-menu-show-formulas`, a `MenuCheckItem` checked while on, after "Gridlines"); it and Ctrl+` (matched on `event.code === "Backquote"` with `ctrlKey` and no other modifier — Ctrl+Shift+` is a different binding — on **every** platform including macOS) toggle the active sheet's `showFormulas`; in that mode `spreadsheet-grid` carries `data-show-formulas="true"`, formula cells show `=` + formula left-aligned, and constants show `generalText` of their value (no number format); the toggle never sets `dirty`/`structureDirty` (SPEC-121) and is ignored when `isOutsideDocumentInput(event.target)` (the SPEC-021 convention)
- [ ] `formulaForTarget(cell, target)` in `formula/dialect.ts` is the one cross-format rule per "Crossing formats": a parsed formula crosses unchanged (it is canonical); an unparsed formula and a data table cannot, so the cell drops to its cached value and the drop is reported for SPEC-121's "Download a copy as…" toast; a dynamic array crosses without its `cm`
- [ ] Layering: every module under `src/lib/spreadsheet/formula/` imports only spreadsheet siblings and `src/lib/core`, touches no `window`/`document`, and contains no `eval`, `new Function` or `Function(` — asserted by `formula/purity.spec.ts` (a source scan over the directory), alongside the existing `spreadsheet/boundary.spec.ts`
- [ ] Fixtures per "Fixtures": `e2e/fixtures/spec119-formulas.xlsx` (Excel-shaped) and `spec119-formulas.ods` (LibreOffice-shaped), built by `spreadsheet/fixtures/build.ts` and `scripts/build-spreadsheet-fixtures.ts`
- [ ] Tests: `formula/tokens.spec.ts`, `parse.spec.ts`, `print.spec.ts`, `references.spec.ts`, `dialect.spec.ts`, `functionNames.spec.ts`, `purity.spec.ts`, additions to `xlsx/reader.spec.ts`, `ods/reader.spec.ts`, `open.spec.ts` and `roundTrip.spec.ts`, `views/spreadsheet/useFormulaDisplay.spec.ts`, and the `views/SpreadsheetView.spec.ts` wiring test
- [ ] e2e (`e2e/SPEC-119.spec.ts`): drop `spec119-formulas.xlsx` → open → select `B5` → `formula-bar-content` reads "=SUM(B2:B4)" → select `C3` (a shared-formula follower) → reads "=B3*$F$1" → select `E3` (an array member) → reads "{=A2:A4*2}" → select `G2` → reads `=IFS(A2>1,"big",TRUE,"small")` (no `_xlfn.`) → press Ctrl+` → `B5`'s `sheet-cell` reads "=SUM(B2:B4)" and the View menu shows "Show formulas" checked → Ctrl+` again → `B5` reads "60" → Close → drop `spec119-formulas.ods` → open → `B5`'s formula bar reads "=SUM(B2:B4)" → `B7` reads "='My Sheet'!A1+Sheet2!B1" → download → every part `equalBytes` the fixture's

## Out of scope
Evaluating anything (SPEC-120): values stay cached, a formula that would now
compute differently shows the producer's number. Editing a formula, entering
one, point-mode reference highlighting and colour-coding of a formula's
ranges in the grid (SPEC-122). Writing formulas back to a dirty sheet
(SPEC-122 — this spec gives it `toXlsxFormula`/`toOdsFormula`). Adjusting
references for inserted or deleted rows (SPEC-124 builds `adjustFormula` on
`rewriteReferences`). Copy, fill and move (SPEC-125, on `shiftFormula`).
What a name or a structured reference refers to, and normalising Excel's
wire `[#This Row]` to the UI's `@` (SPEC-133). Locale-specific formula syntax
(`;` separators, localised function names — whole-queue backlog: formulas are
always en-US syntax, SPEC-118). R1C1 display (`calcPr@refMode`), the spill
operator `A1#` and `@` as typed syntax (backlog — their wire forms
`_xlfn.ANCHORARRAY(A1)` and `_xlfn.SINGLE(x)` parse as ordinary calls).
Resolving `[1]` to an external workbook's file name (external links show as
stored). `c@vm` (rich values and pictures in cells — SPEC-122 decides its
preservation). LibreOffice's relative-sheet references (`[Sheet2.A1]`
without `$`) survive only on clean sheets; see "Known non-identities".

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — the workbook model
  (`SpreadsheetModel`, `Sheet { id, name, visibility, kind, rows, partName?,
  sourceXml? }`, `Row { index, cells }`, `Cell { col, value, runs? }`,
  `CellValue = number | string | boolean | CellError | null`, `ErrorCode` —
  the fifteen codes `#NULL!` `#DIV/0!` `#VALUE!` `#REF!` `#NAME?` `#NUM!`
  `#N/A` `#GETTING_DATA` `#SPILL!` `#CALC!` `#FIELD!` `#BLOCKED!` `#CONNECT!`
  `#BUSY!` `#UNKNOWN!`), its four invariants (copy-on-write rows, sparse, the
  source is the splice base, dates are serials), `MAX_ROWS = 1_048_576`,
  `MAX_COLS = 16_384`. Addressing in `spreadsheet/address.ts`: 0-based
  `CellRef { row; col }`, inclusive normalised `GridRange { top; left;
  bottom; right }`, `colToLetters`, `lettersToCol`, `formatA1`, `parseA1`
  (accepts `$`, returns `null` outside `A1..XFD1048576`), `parseRange`,
  `getCell`, `findRow`. Reader versions in `spreadsheet/open.ts`:
  `SPREADSHEET_READER_VERSION`, `ReaderUpgrade { from;
  upgradeDirtySheet(model, sheet, sourceParts) }`, `READER_UPGRADES`,
  `upgradeSpreadsheetModel(model)`. Readers scan worksheet parts and
  `content.xml` with `core/xmlScanner.ts`'s `scanXml` (exact UTF-16 source
  offsets on every event, entities decoded); `xl/workbook.xml` is small and
  read with `core/xml.ts`'s `parseXml`. Part-level `equalBytes`
  (`core/equality.ts`) is the round-trip relation. The fixture builder is
  `spreadsheet/fixtures/build.ts`, regenerated by `npx vite-node
  scripts/build-spreadsheet-fixtures.ts` (which ends in `process.exit(0)`).
- `specs/117-sheet-grid-styles-sizes-and-selection.md` — `SheetGrid.vue`
  (`spreadsheet-grid`, each rendered cell a `sheet-cell` with `data-ref`),
  the read-only formula bar `FormulaBar.vue` (`formula-bar`, `name-box`,
  `formula-bar-content`), `SpreadsheetStatusBar.vue`, `SheetSelection {
  active; ranges; activeRangeIndex }`, the sheet-view attributes read from
  `<sheetViews>` (`Sheet.selection`, `zoomPercent`, `showGridLines`, `frozen`
  — view state that never marks a sheet dirty), primary modifier from
  `core/platform.ts`.
- `specs/118-number-formats-and-dates.md` — `editText(cell, ctx:
  EditTextContext)` in `numberFormat/render.ts` (`EditTextContext extends
  FormatContext { style?: CellStyle }`, `ctx.style` = the cell's effective
  style) is the formula-bar text of a non-formula cell and never reads
  `Cell.formula`; `FormatContext { locale; dateSystem; … }`; **formulas are
  always entered and shown in the invariant en-US syntax** regardless of the
  user's locale.
- SPEC-117's `SpreadsheetMenuBar.vue` and its View menu (where "Gridlines",
  `sheet-menu-gridlines`, already lives; menu-bar item testids keep the
  `sheet-menu-` prefix) on the product-neutral
  `MenuBar.vue`/`MenuCheckItem.vue`; `isOutsideDocumentInput(target)` in
  `src/lib/shortcuts.ts` (SPEC-021: a keydown from an input, textarea,
  select or dialog is not a shortcut).

### Model additions (`spreadsheet/model.ts`)
```ts
export interface Cell {
  // ...SPEC-116/117 fields...
  /** Canonical formula text without the leading "=" (see "Canonical text").
   *  When `formulaUnparsed`, the WIRE text instead (xlsx <f> text; ods the
   *  whole table:formula value, prefix included). */
  formula?: string;
  formulaUnparsed?: true;
  /** Legacy array (CSE) formula: the range it fills. On the anchor only —
   *  the top-left cell, which holds `formula`; member cells hold values. */
  arrayRange?: GridRange;
  /** xlsx: the anchor's array formula is a dynamic-array (spill) formula —
   *  `t="array"` plus a `cm` attribute. arrayRange is its last spill extent. */
  dynamicArray?: true;
  /** xlsx `c@cm` verbatim (index into xl/metadata.xml cellMetadata), so the
   *  sheet writer (SPEC-122) re-emits it. */
  cm?: number;
  /** xlsx What-If data table: the anchor's `<f t="dataTable" …/>` element,
   *  verbatim. No `formula`; members hold cached values. Never evaluated. */
  dataTableXml?: string;
}
export interface Sheet {
  // ...
  /** View ▸ Show formulas for this sheet (xlsx sheetView@showFormulas). View
   *  state: never marks the sheet dirty. */
  showFormulas?: true;
}
```

### Canonical text
Excel's en-US UI syntax without the leading `=`: function names upper-case
and unprefixed (`CONCAT`, `STDEV.S`, `SORT`), names without `_xlpm.`, A1
references with `$` markers, sheet names quoted when the quoting rule
requires, `,` argument separator, `.` decimal point, `{1,2;3,4}` arrays.
Everything else — spacing, line breaks, the case of reference letters and of
names, number spellings like `1E3` — is the source's, verbatim. Canonical is
what the formula bar shows after `=`, what every rewrite operates on, and
what both writers translate from.

### Tokens (`formula/tokens.ts`)
```ts
export type TokenKind =
  | "number" | "string" | "boolean" | "error" | "ref" | "name" | "function"
  | "structured" | "external" | "operator"          // + - * / ^ & = <> < <= > >= % : and " " (intersection)
  | "open" | "close" | "argSep" | "union"
  | "arrayOpen" | "arrayClose" | "arrayColSep" | "arrayRowSep"
  | "unknown" | "end";                              // "end" holds trailing whitespace as its lead
export interface Token {
  kind: TokenKind;
  text: string;            // the token's own source text
  lead: string;            // whitespace (space, tab, CR, LF) immediately before it
  start: number;           // UTF-16 offset of `text` in the source
  ref?: ReferenceNode;     // kind "ref"
  unterminated?: true;     // a string or structured token that ran off the end
}
export function tokenize(text: string): Token[];
export function joinTokens(tokens: Token[]): string;  // concatenates lead + text
```

### Token rules
- **Identifiers** start with a letter, `_` or `\` and continue with letters,
  digits, `_` and `.` (letters include every Unicode letter). An identifier
  immediately followed by `(` is a `function`, whatever it spells. Otherwise
  it is a `ref` when it is a cell inside `A1..XFD1048576` (`$` allowed before
  the letters and before the digits), a `boolean` when it is `TRUE`/`FALSE`
  in any case, and a `name` otherwise (`XFE1`, `A0`, `R1C1`, `Rate`).
- **Ref tokens join greedily**: `prefix? area (":" area)?` with no
  whitespace around `:`, where `prefix` is `sheet!`, `sheet:sheet!`, `'…'!`
  (quoted, `''` escapes, may contain `:` for 3-D) and `area` is a cell, a
  column (`$`? letters, only as a column range) or a row (`$`? digits, only
  as a row range). `A:A` and `1:1` are refs only as ranges; a lone `A` is a
  name. `Sheet1!A1:Sheet1!B2` is two refs and a `:` operator.
- **Errors**: the fifteen codes, case-insensitive. `Sheet!#REF!` (quoted or
  not) and `#REF!` immediately followed by a cell (`#REF!A1`, Excel's spelling
  of a reference into a deleted sheet) are each one `error` token.
- **External**: `[digits]` followed by a sheet name and `!`, or by `!`
  directly, through the end of the reference that follows; inside quotes
  (`'[1]My Sheet'!A1`) likewise. One token, never interpreted.
- **Structured**: an identifier immediately followed by `[`, or a `[` not
  starting an external prefix, consumes balanced brackets; `'` escapes the
  next character (`[Col'#1]`).
- **Intersection**: a whitespace run between a reference-ending token (`ref`,
  `name`, `structured`, `external`, `close`) and a reference-starting one
  (`ref`, `name`, `function`, `structured`, `external`, `open`) that contains
  a U+0020 space is an `operator` token whose text is its **last** space; the
  run before it is that token's `lead`, the run after it the next token's.
  Any other whitespace is trivia.
- **Commas** are `argSep` when the innermost open bracket is a function
  call's `(`, `arrayColSep` inside `{…}`, and `union` otherwise (including
  top level, which defined names use — SPEC-133). `;` is `arrayRowSep` inside
  `{…}` and `unknown` elsewhere.
- **Numbers**: digits with an optional fraction and exponent, or `.` then
  digits. A leading sign is an operator except inside `{…}`, where it belongs
  to the element.

### The AST (`formula/ast.ts`)
```ts
interface Span { from: number; to: number }          // inclusive token indices
export type ArrayElement = number | string | boolean | CellError;
export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | "<=" | ">" | ">="
  | ":" | " " | ",";                                 // range, intersection, union
export type FormulaNode = Span & (
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "boolean"; value: boolean }
  | { kind: "error"; code: ErrorCode; text: string } // text: "Sheet2!#REF!" when spelled so
  | { kind: "array"; rows: ArrayElement[][] }
  | { kind: "ref"; ref: ReferenceNode }
  | { kind: "name"; name: string; sheet?: string }
  | { kind: "function"; name: string; args: FormulaNode[] }   // name upper-case
  | { kind: "missing" }                              // an empty argument slot
  | { kind: "unary"; op: "+" | "-" | "%"; operand: FormulaNode }
  | { kind: "binary"; op: BinaryOp; left: FormulaNode; right: FormulaNode }
  | { kind: "paren"; inner: FormulaNode }
  | { kind: "structured"; table?: string; text: string }
  | { kind: "external"; text: string });

// formula/parse.ts
export class FormulaSyntaxError extends Error { readonly offset: number }
export function parseFormula(text: string): FormulaNode;
export const MAX_PARSE_DEPTH = 256;
export const FORMULA_AST_CACHE_SIZE = 50_000;
export function parseFormulaAt(text: string, host: CellRef):
  { node: FormulaNode; origin: CellRef } | FormulaSyntaxError;     // LRU by relativeFormulaKey
```
`origin` is the host of the formula the cached AST was first parsed for; a
later caller with the same key resolves a relative `RefPart` as
`part.row + (host.row − origin.row)` (SPEC-120 does exactly this). Nodes are
never mutated after parse.

### Precedence
| Level (tightest first) | Operators | Notes |
| --- | --- | --- |
| 1 | `:` | range; operands must be references |
| 2 | ` ` | intersection (a space between references) |
| 3 | `,` | union — only where a comma is a `union` token |
| 4 | prefix `-`, `+` | binds tighter than `^`: `-2^2` = `(-2)^2` |
| 5 | postfix `%` | `50%`, `1%%` |
| 6 | `^` | left-associative: `2^3^2` = 64 |
| 7 | `*` `/` | left |
| 8 | binary `+` `-` | left |
| 9 | `&` | left |
| 10 | `=` `<>` `<` `<=` `>` `>=` | left: `A1=B1=C1` is `(A1=B1)=C1` |

The right operand of a binary operator may begin with any number of prefix
operators (`1+-+-1`, `2^-2`). This is Excel's table; LibreOffice evaluates
the same precedence for these operators.

### Literals
- Number text is kept in the token; the node's `value` is `Number(text)`.
  `1E400` is a `FormulaSyntaxError`.
- Strings: `"…"`, `""` is one quote; an unterminated string is a syntax error.
- Array constants: rows separated by `;`, columns by `,`; every row the
  same length; elements are numbers (an optional leading `-`/`+`), strings,
  booleans or error literals — references and expressions are syntax errors,
  as in Excel.

### Sheet-name quoting
A sheet name is written unquoted iff it matches
`^[\p{L}_][\p{L}\p{N}_.]*$` and is not a cell reference (`A1`, `XFD1`), not
an R1C1-shaped token (`R`, `C`, `R1`, `C1`, `RC`, `R1C1`, any case), and not
`TRUE`/`FALSE`; otherwise it is wrapped in `'` with every `'` doubled. Both
`formatReference` and `toOdsFormula` use this predicate; renames re-apply it.

### References (`formula/references.ts`)
```ts
export interface RefPart { row?: number; col?: number; rowAbs: boolean; colAbs: boolean }
// 0-based; row absent = whole column(s), col absent = whole row(s)
export interface ReferenceNode {
  sheet?: string;        // as written, unquoted and un-escaped
  sheetEnd?: string;     // 3-D: the last sheet of Sheet1:Sheet3!
  start: RefPart;
  end?: RefPart;         // absent for a single cell
}
export type RefRewrite = (ref: ReferenceNode) => ReferenceNode | null | undefined; // null = #REF!, undefined = keep verbatim
export function rewriteReferences(text: string, rewrite: RefRewrite): string;
export function shiftFormula(text: string, dRow: number, dCol: number): string;
export function renameSheetInFormula(text: string, from: string, to: string): string;
export function relativeFormulaKey(text: string, host: CellRef): string;
export function findArrayAnchor(sheet: Sheet, row: number, col: number): Cell | undefined;
```
- `rewriteReferences` is the primitive SPEC-124's `adjustFormula` and
  SPEC-125's cut-move build on. It tokenizes once, calls `rewrite` for each
  `ref` token, replaces only the returned tokens' `text` (via
  `formatReference`; `null` becomes `<sheet prefix>#REF!`), and joins.
  Trivia and every untouched token survive byte-for-byte.
- `shiftFormula`: a relative row part moves by `dRow`, a relative column by
  `dCol`; whole-column refs ignore `dRow`, whole-row refs ignore `dCol`; any
  part leaving `0..MAX_ROWS−1` / `0..MAX_COLS−1` turns the whole token into
  `#REF!`. Excel produces the same text on copy and on shared-formula
  expansion.
- `relativeFormulaKey`: the token texts joined with trivia dropped, each
  `ref` token replaced by its R1C1 form relative to `host` (`R[-1]C[2]` for
  relative parts, `R5C3` for absolute, `C[0]:C[0]` for a whole column), sheet
  prefixes kept as written. Two cells share a key exactly when one is a
  shifted copy of the other.
- `findArrayAnchor`: the cell whose `arrayRange` contains `(row, col)`,
  memoised per `sheet.rows` array reference in a `WeakMap` (lib state, never
  model data).

### Function names (`formula/functionNames.ts`)
```ts
export interface FunctionNameEntry {
  name: string;                            // canonical: "CONCAT"
  xlsxPrefix?: "_xlfn." | "_xlfn._xlws.";  // how the xlsx wire spells it
  odsName?: string;                        // OpenFormula name when it differs
}
export const FUNCTION_NAMES: readonly FunctionNameEntry[];
export function canonicalFromXlsx(wire: string): string;
export function xlsxFromCanonical(name: string): string;
export function canonicalFromOds(wire: string): string;
export function odsFromCanonical(name: string): string;
```
Transcribed, not recalled: the `_xlfn.` set from Microsoft's future-function
list ([MS-XLSX] §2.2.2 "Formulas" and the "_xlfn. prefix" support article);
the ODS names from LibreOffice's `sc/source/filter/oox/formulabase.cxx`
(Excel ↔ ODFF mapping) cross-checked with `formula/inc/core_resource.hrc`
(the ODFF English names). Test vectors this spec pins, which any transcription
must reproduce:

| Canonical | xlsx wire | ODS wire |
| --- | --- | --- |
| `CONCAT` | `_xlfn.CONCAT` | `COM.MICROSOFT.CONCAT` |
| `TEXTJOIN` | `_xlfn.TEXTJOIN` | `COM.MICROSOFT.TEXTJOIN` |
| `IFS` | `_xlfn.IFS` | `COM.MICROSOFT.IFS` |
| `SWITCH` | `_xlfn.SWITCH` | `COM.MICROSOFT.SWITCH` |
| `MINIFS` / `MAXIFS` | `_xlfn.MINIFS` / `_xlfn.MAXIFS` | `COM.MICROSOFT.MINIFS` / `…MAXIFS` |
| `XOR`, `IFNA`, `DAYS` | `_xlfn.XOR`, `_xlfn.IFNA`, `_xlfn.DAYS` | unchanged |
| `SORT`, `FILTER` | `_xlfn._xlws.SORT`, `_xlfn._xlws.FILTER` | per source |
| `SINGLE` | `_xlfn.SINGLE` (Excel's wire form of `@`) | per source |
| `SUM`, `VLOOKUP` | unchanged | unchanged |

LibreOffice-only names (`ORG.OPENOFFICE.*`, `ORG.LIBREOFFICE.*`) have no
canonical counterpart and stay as written; SPEC-120 reports them as
unsupported.

### xlsx dialect
- `fromXlsxFormula(wire)`: tokenize; for each `function` token strip a
  leading `_xlfn._xlws.` or `_xlfn.` and upper-case; for each `name` token
  strip `_xlpm.`; drop one leading `=` if a producer wrote it; join.
- `toXlsxFormula(canonical)`: tokenize; `function` tokens through
  `xlsxFromCanonical`; `name` tokens get `_xlpm.` when they are a parameter
  of an enclosing `LET` (arguments 1, 3, 5 … except the last) or `LAMBDA`
  (every argument but the last), or a use of such a parameter inside that
  call's argument span (lexical scope, innermost wins). No `=`.

### ODS dialect
`fromOdsFormula(wire, namespaces)` takes the `table:formula` value and the
in-scope prefix→URI map; the dialect comes from the prefix's URI:

| Namespace URI | Dialect | Translation |
| --- | --- | --- |
| `urn:oasis:names:tc:opendocument:xmlns:of:1.2` | OpenFormula (`of:`) | the table below |
| `http://openoffice.org/2004/calc` | legacy (`oooc:`) | same bracket grammar and `;` separators |
| `http://schemas.microsoft.com/office/excel/formula` | Excel (`msoxl:`) | strip `=`, then `fromXlsxFormula` |
| (no prefix) | OpenFormula | as `of:` |

| OpenFormula | Canonical |
| --- | --- |
| `[.A1]`, `[.$A$1]` | `A1`, `$A$1` |
| `[.A1:.B2]` | `A1:B2` |
| `[$Sheet1.A1]`, `[Sheet1.A1]` | `Sheet1!A1` |
| `[$Sheet1.A1:.B2]`, `[$Sheet1.A1:$Sheet1.B2]` | `Sheet1!A1:B2` |
| `['My Sheet'.A1]` (`''` escapes) | `'My Sheet'!A1` |
| `[$Sheet1.A1:$Sheet3.B2]` | `Sheet1:Sheet3!A1:B2` |
| `[.A:.C]`, `[.1:.3]` | `A:C`, `1:3` |
| any bracket reference containing `#REF!` | `#REF!` |
| `['file:///…'#$Sheet1.A1]` | an `external` token, text verbatim |
| `;` (argument) | `,` |
| `a ~ b` (union) | `(a,b)` — parentheses added unless the union is already a paren's whole content |
| `a ! b` (intersection) | `a b` |
| `{1;2\|3;4}` | `{1,2;3,4}` |
| `COM.MICROSOFT.CONCAT(…)` etc. | through `canonicalFromOds` |

`toOdsFormula(canonical)` inverts it: `of:=` prefix, `[.A1]` for same-sheet
refs, `[$Sheet.A1:.B2]` for another sheet (absolute sheet, LibreOffice's
default), quoting by the shared predicate, `,`→`;`, union→`~` (the
parentheses Excel needed are kept — they are harmless in OpenFormula),
intersection→`!`, arrays→`{…;…|…}`. Each token's `lead` is carried across,
so whitespace survives both directions. The translator's own OpenFormula
tokenizer emits the same `Token` shape.

**Known non-identities** (`toOdsFormula(fromOdsFormula(w)) !== w`, asserted
as such): a relative-sheet reference `[Sheet2.A1]` comes back as
`[$Sheet2.A1]`; a same-sheet second endpoint `[$S.A1:$S.B2]` comes back as
`[$S.A1:.B2]`; a `~` union gains parentheses; a `#REF!` reference comes back
as the bare error. Only dirty sheets (SPEC-122) ever go through this path —
clean sheets are spliced verbatim.

### Reading formulas from .xlsx
| Source | Model |
| --- | --- |
| `<c><f>text</f><v>…</v></c>` | `formula = fromXlsxFormula(text)`; `value` = cached `<v>` as before |
| master `<f t="shared" ref="C2:C9" si="0">B2*$F$1</f>` | `formula` as above |
| follower `<f t="shared" si="0"/>` | `formula = shiftFormula(masterCanonical, row − masterRow, col − masterCol)` |
| `<f t="array" ref="E2:E4">A2:A4*2</f>` | anchor `formula` + `arrayRange` (also for a one-cell `ref`) |
| `<c cm="1"><f t="array" ref="H2:H6">_xlfn._xlws.SORT(A2:A6)</f>` | as array, plus `cm: 1`, `dynamicArray: true` |
| `<f t="dataTable" ref="B2:D5" dt2D="1" r1="A1" r2="A2"/>` | anchor `dataTableXml` = the element's source text (sliced by scanner offsets) |
| `<sheetView showFormulas="1">` (first sheetView) | `Sheet.showFormulas: true` |

A formula with no `<v>` keeps `value: null` — SPEC-120 treats that as "no
cached value". Followers whose `si` is never defined keep their values and no
formula. `ca`, `aca`, `bx` and `xl/calcChain.xml` are not read (the source
keeps them). The parse check uses `parseFormulaAt`, so a shared block costs
one parse.

### Reading formulas from .ods
`table:formula` on `table:table-cell` → `fromOdsFormula(value, ns)` where
`ns` are the `xmlns:*` declarations on `office:document-content` (recorded
from the scanner's first open event). `table:number-matrix-columns-spanned`
and `-rows-spanned` on the anchor → `arrayRange`. A cell repeated by
`table:number-columns-repeated` (or in a repeated row) copies `formula`
unchanged to every repetition. A formula cell whose `office:value-type` is
absent keeps `value: null`.

### Reader version
Bumps `SPREADSHEET_READER_VERSION` to the next integer and appends
`{ from: <previous>, upgradeDirtySheet() { throw … } }`. The re-read path
covers every real record: a sheet can only become dirty under a build that
has SPEC-121, and such a build has already upgraded the record past this
step on load. The throwing step keeps `READER_UPGRADES` contiguous (one entry
per version, asserted by `open.spec.ts`) rather than leaving a hole.

### The formula bar
`formulaBarText(sheet, row, col, style, ctx): string` in `formula/print.ts`
(`style` is the cell's resolved `CellStyle`, passed through to `editText`):

| Cell | Text |
| --- | --- |
| formula, parsed | `"=" + formula` |
| anchor or member of a legacy `arrayRange` | `"{=" + anchor.formula + "}"` |
| dynamic-array anchor | `"=" + formula` |
| dynamic-array member | the anchor's `"=" + formula`, rendered muted, `data-spill-member="true"` |
| data-table range | `{=TABLE(r1,r2)}`; one-dimensional: `{=TABLE(r1,)}` when `dtr="1"`, else `{=TABLE(,r1)}` |
| unparsed, xlsx | `"=" + formula` |
| unparsed, ods | `formula` from its first `=` |
| anything else | SPEC-118 `editText(cell, { ...ctx, style })` |

### Show formulas
`src/views/spreadsheet/useFormulaDisplay.ts` owns the toggle, the Ctrl+`
listener and the formula-bar text; `SheetGrid.vue` gains a `showFormulas`
prop; `SpreadsheetView.vue` gains only the wiring call and template hookup
(its size ratchet does not rise). Ctrl+` and not Cmd+` on macOS: Cmd+` is the
system's window-cycling shortcut and never reaches the page, and Excel for
Mac uses Control+` too. Matching on `event.code` keeps it working on layouts
where `` ` `` is a dead key. Columns are not widened (Excel doubles them;
that is a display nicety this spec does not copy).

### Crossing formats
`formulaForTarget(cell, target: SpreadsheetFormat): { cell: Cell; dropped?:
"unparsed" | "dataTable" | "dynamicArray" }`: a parsed formula crosses
unchanged. An unparsed formula and a data table drop to the cached values
(`formula`/`dataTableXml` removed). A dynamic array crossing to `.ods` keeps
`formula` + `arrayRange` (written as a matrix formula) and loses `cm` and
`dynamicArray`. SPEC-121's `stripRetention` calls it for every cell, and any
`"unparsed"` or `"dataTable"` drop puts SPEC-121's "untranslatable formulas"
label into its "Not included in the … copy" toast.

### Fixtures
- `spec119-formulas.xlsx` (Excel-shaped, the SPEC-116 shapes: `mc:Ignorable`,
  `xr:uid`, `x14ac:dyDescent`, `spans`, `calcChain.xml`): sheet "Calc" with
  `B5` `SUM(B2:B4)` (cached 60); a shared block `C2:C9` `B2*$F$1` in
  Excel's order (master first); a second shared block shaped like a
  non-Excel producer's — one follower written before its master, and a
  relative reference whose shift runs past row 1,048,576 on the last
  follower (→ `#REF!`); a CSE array `E2:E4` `A2:A4*2`; `G2`
  `_xlfn.IFS(A2>1,"big",TRUE,"small")`; `H2` dynamic `_xlfn._xlws.SORT(A2:A6)`
  with `cm="1"` and an `xl/metadata.xml`; `_xlfn.LET(_xlpm.x,5,_xlpm.x*2)`;
  `_xlfn.SINGLE(A2:A4)`; a data table `t="dataTable"`; a 3-D sum over
  `Jan:Mar`; `'O''Brien'!A1`; `A:A` and `1:1`; `[1]Sheet1!$A$1` with an
  `xl/externalLinks/externalLink1.xml`; `Sheet2!#REF!`; a multi-line formula
  with indentation; and one unparseable formula Excel really writes —
  `_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)` (an immediately-invoked lambda, which
  this grammar does not parse). A second sheet has `showFormulas="1"`.
- `spec119-formulas.ods` (LibreOffice-shaped: the SPEC-116 shapes plus
  `xmlns:of`/`xmlns:oooc`/`xmlns:msoxl` on the root): `B5`
  `of:=SUM([.B2:.B4])`; `B7` `of:=[$'My Sheet'.A1]+[$Sheet2.B1]`; `~` and
  `!`; `{1;2|3;4}`; `COM.MICROSOFT.CONCAT`; a matrix formula with both span
  attributes; a repeated formula cell; one `oooc:=` and one `msoxl:=`
  formula; a formula cell without a cached value.

### Tests
- `tokens.spec.ts`: the corpus round trip — at least these, plus every
  fixture formula: `IF(A1="","",VLOOKUP(A1,'Price List'!$A$2:$D$500,4,FALSE))`,
  `SUMPRODUCT((Data!$B$2:$B$1000="West")*Data!$C$2:$C$1000)`,
  `"He said ""hi"""&A1`, `INDEX(A:A,MATCH(MAX(B:B),B:B,0))`, `SUM(Jan:Dec!B5)`,
  `SUM((A1:A3,C1:C3))`, `SUM(A1:C3 B2:D4)`, `  SUM( A1 , B1 ) `, `-2^2`,
  `1+-+-1`, `50%%`, `.5+1E3`, `LOG10(100)+ATAN2(1,1)`, `XFD1048576+XFE1`,
  `Sales[@Qty]*Sales[@Price]`, `Table1[[#Headers],[Col'#1]]`, `[1]!GlobalName`,
  `#REF!A1+Sheet1!#REF!`, `A1:INDEX(B:B,5)`, `'Übersicht'!A1`, `true+TRUE()`,
  a CRLF-indented `IF`; kind assertions per rule; the seeded fuzz.
- `parse.spec.ts`: the precedence table as fully-parenthesised prints; every
  literal form; arrays incl. the ragged error; missing args; offsets of
  errors; the depth bound; `parseFormulaAt` sharing (100,000 shifted texts →
  cache size 1) and origin offsets.
- `references.spec.ts`: shift tables (relative/absolute mixes, whole
  columns/rows, 3-D, out-of-grid at all four edges, trivia untouched),
  renames (quoting added and removed, 3-D endpoints, case-insensitive),
  `relativeFormulaKey` equal for shifted copies and different otherwise.
- `dialect.spec.ts`: every row of both dialect tables in both directions,
  the `_xlpm.` scope rule (nested LETs, a name shadowing a parameter), the
  three ODS dialects, the non-identities, `formulaForTarget`.
- `functionNames.spec.ts`: the pinned vectors; no duplicate canonical or
  wire names; every entry round-trips.
- Reader additions: each "Reading formulas" row; the Numbers fixture; ods
  repetition copies text unshifted. `open.spec.ts`: the upgrade. All
  equality in these suites is `toEqual` on the model; byte round trips are
  `equalBytes`.

### Specs this changes
- `specs/116-…` — `Cell` and `Sheet` gain the fields above; `<f>` and
  `table:formula` are now read; `SPREADSHEET_READER_VERSION` moves.
- `specs/117-…` — the formula bar's content comes from `formulaBarText`;
  the View menu gains `sheet-menu-show-formulas`.
- `specs/118-…` — `editText` becomes the formula-bar text of non-formula
  cells only, reached through `formulaBarText`.
- `specs/121-…` — `stripRetention` applies `formulaForTarget` and counts drops.
- `specs/122-…` — the writers emit `toXlsxFormula`/`toOdsFormula` output,
  unparsed text verbatim, `t="array" ref`, `cm`, `dataTableXml` and ODS
  matrix spans; `Sheet.showFormulas` joins the rebuilt `<sheetView>`.
- `specs/124-…` — `adjustFormula` is built on `rewriteReferences`; a data
  table's `ref`/`r1`/`r2` are reference-bearing retained XML.
- `specs/125-…` — copy and fill use `shiftFormula`.
- `specs/133-…` — names and structured nodes get their meaning there.
