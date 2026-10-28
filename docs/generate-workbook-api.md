# Workbook generation API (SPEC-140)

The third product of the generation API (see [generate-api.md](generate-api.md)
for documents and [generate-deck-api.md](generate-deck-api.md) for decks): a
fluent `WorkbookBuilder` and a 1:1 plain-JSON workbook description for
generating `.xlsx`/`.ods` workbooks from data — the monthly sales export, the
reconciliation a finance system mails out, the CSV somebody wants as a
formatted workbook with totals that actually compute — plus `readWorkbook`
to read any workbook's bytes back.

It composes the product's own spreadsheet engine and adds no serialisation
code of its own: SPEC-121's `createBlankWorkbook()` and sheet-name rules,
SPEC-123's `applyStylePatch` (a builder style and an editor style mean the
same thing), SPEC-124's `mergeCells` and band commands, SPEC-126's CSV
inference, and SPEC-120's recalculation. Bytes come only from SPEC-121's
writers (`writeXlsx`/`writeOds`, on their from-scratch path); a construct the
writers cannot emit is fixed in the writer.

**Every formula's cached value is computed before the file is written.**
Excel can be told to recalculate on open, but LibreOffice by default does
not, Numbers shows what it is given, and every preview pane and search
indexer reads the cached value. So the builder evaluates every formula with
SPEC-120's engine and refuses one it cannot evaluate — `toWorkbook()` holds
the values Excel would show.

Everything lives in `src/lib/generate/`:

```ts
import { WorkbookBuilder, workbookFromDescription, readWorkbook } from "./src/lib/generate"; // all three builders
import { WorkbookBuilder } from "./src/lib/generate/workbook";                               // the workbook side alone
```

The workbook modules (`workbook.ts`, `workbookPlan.ts`, `workbookStyles.ts`,
`workbookDescription.ts`, `readWorkbook.ts`) import only `src/lib/spreadsheet/`,
`src/lib/core/` and generate's product-neutral `errors.ts`, `validate.ts` and
`nodeDom.ts` — never the text or deck builders, `documentModel.ts` or
`openDocument.ts` (`workbookBoundary.spec.ts`); and nothing under
`src/lib/spreadsheet/` imports `generate/`.

Runnable examples: `node examples/sales-report.js`,
`node examples/csv-to-xlsx.js`, `node examples/budget.js` — see
[Examples](#examples).

## Quick start

The Reference's "Q3 sales" workbook:

```ts
import { WorkbookBuilder } from "./src/lib/generate";

const wb = new WorkbookBuilder({ title: "Q3 sales" });
wb.defineStyle("header", { font: { bold: true }, fill: "#dce6f1", borders: { bottom: "thin" } })
  .addSheet("Sales")
  .setRow("A1", ["Region", "July", "August", "September", "Total"])
  .setRows("A2", [["East", 1200, 1350, 1410], ["West", 980, 1010, 1130]])
  .setFormula("E2", "SUM(B2:D2)").setFormula("E3", "SUM(B3:D3)")
  .style("A1:E1", "header")
  .numberFormat("B2:E3", "$#,##0.00")
  .columnWidth("A", { chars: 14 })
  .freeze({ rows: 1 });

const bytes = await wb.toBytes();           // .xlsx (the builder's own format)
const ods = await wb.toBuffer("ods");       // Node
wb.toWorkbook().sheets[0].rows[1].cells[4]; // { col: 4, value: 3960, formula: "SUM(B2:D2)", styleId: … }
```

Every method validates its arguments on entry and throws a
`DocumentGenerationError` naming the offending path —
`sheets[0].cells.B5.formula: unknown function SUMM` — before it records
anything, so an error always precedes any byte and a failed call leaves the
builder as it was.

## `WorkbookBuilder`

### `new WorkbookBuilder(options?)`

| Option | Type | Default |
| --- | --- | --- |
| `title` | string | "Untitled spreadsheet" (as a new workbook's) |
| `format` | `"xlsx"` \| `"ods"` | `"xlsx"` — what `toBytes()`/`toBlob()`/`toBuffer()`/`fileName()` use when called without a format |
| `dateSystem` | `"1900"` \| `"1904"` | `"1900"` |
| `locale` | `"en-US"` \| `"en-GB"` \| `"de-DE"` \| `"fr-FR"` \| `"es-ES"` \| `"cs-CZ"` | `"en-US"` — the locale `parseText` reads fields in and formulas evaluate in. Formula *text* is always en-US syntax. |

The workbook starts as SPEC-121's `createBlankWorkbook()` — the default
style (Calibri 11 through the Office theme's minor font), the Office theme —
with its sheets replaced by the ones the builder adds. A builder that adds
none produces one empty "Sheet1".

### Sheets

#### `addSheet(name, options?)`

Appends a sheet and makes it current. The name follows SPEC-121's rules —
1–31 characters, none of `: \ / ? * [ ]`, no leading or trailing `'`, not
"History", unique ignoring case — and a broken rule's own message becomes
the error at `sheets[i].name` (`sheets[1].name: There's already a sheet called "sales"`).

| Option | Type | Default |
| --- | --- | --- |
| `tabColor` | `"#rrggbb"` | none |
| `hidden` | boolean | `false` — at least one sheet must stay visible; the first visible sheet is active |
| `showGridLines` | boolean | `true` |

#### `sheet(name)`

Makes an existing sheet current (its name in any case). Every cell, range,
size, merge and freeze method acts on the current sheet; with none it throws
`no sheet — call addSheet first`.

### Cells

#### `setCell(ref, input)`

Writes one cell, `ref` in A1 notation (`"B3"`, `"$B$3"`).

#### `setRow(ref, values)`

Writes `values` left to right from `ref`.

#### `setRows(ref, rows, options?)`

Writes a 2-D block with its top-left at `ref`, row by row. `options.parseText:
true` routes every string through SPEC-126's `inferCsvField(field, locale,
dateSystem)` — the lossless inference the editor's CSV import uses — so in
the builder's locale `"1,234.50"`, `"45%"` and a date become typed values with
their number formats, while anything that would change when written back
stays text: `"007"` keeps its zeros, and in en-US `"10/08/2026"` stays text
(the US short date writes it `10/8/2026`) where en-GB types it. **parseText
never makes a formula**: SPEC-122's `parseInput` is deliberately not used,
because it turns a leading "=" into a formula, and CSV formula injection
(`=HYPERLINK(…)`, DDE payloads) is exactly the attack a data-import option
must not open — `"=SUM(A1:A2)"` stays the text it is. It applies to plain
string inputs; a `{ value: "…" }` object is always stored as given.

#### `setFormula(ref, text)`

The same as `setCell(ref, { formula: text })`.

#### Cell input

| Input | Stored `value` | Number format when none given |
| --- | --- | --- |
| `"text"` (≤ 32,767 characters; `"=…"` is text) | the string (`""` stores nothing, as `null` does) | — |
| a finite number (`-0` → `0`; `NaN`/`Infinity` are errors) | the number | — |
| `true` / `false` | the boolean | — |
| `null` | no cell (a style may still store one) | — |
| `{ error: "#N/A" }` — any of SPEC-116's `ErrorCode`s | the error | — |
| a JS `Date`, read in **UTC** | the serial in the workbook's date system | `"yyyy-mm-dd"`, or `"yyyy-mm-dd hh:mm"` with a time |
| `{ date: "YYYY-MM-DD[THH:MM[:SS[.sss]]]" }`, read by its calendar fields | the serial | the same |
| `{ formula: "SUM(B2:B4)" }` | the computed result | — |
| a string with `parseText` | `inferCsvField`'s value | `inferCsvField`'s format |

The object form is `{ value?, date?, error?, formula?, numberFormat?, style? }`
with at most one of `value`, `date`, `error` and `formula`; `numberFormat` and
`style` (a [CellStyleSpec](#cellstylespec) or a defined style's name) are the
cell's own and count as a style call made when the cell was set. An object
with none of the four content keys writes no value but keeps its style.

Dates are converted with SPEC-116's `serial.ts`: 1900-01-01 is serial 1,
1900-03-01 is 61, and `{ date: "1900-02-29" }` is the 1900 system's
fictitious leap day, serial 60, exactly as the reader counts (the 1904
system has no such day, and refuses it). A JS `Date` is read with its UTC
fields, so the same instant gives the same cell on every machine; an ISO
string with a UTC offset (`Z`, `+02:00`) is refused — give the calendar
date and time to store. An invalid `Date` is refused too.

#### Formulas

Formula text is SPEC-119's canonical, en-US syntax (`SUM(B2:D2)`,
`'Q3 Notes'!A1&"!"`); a leading `=` is accepted and stripped, and the
formula is stored as SPEC-119 prints it (function names upper-case, no
whitespace but the intersection operator). On entry:

- the syntax is checked by `parseFormula` — `sheets[0].cells.A9.formula: unexpected ")" at 7`;
- every function must be in SPEC-120/127's registry — `…formula: unknown function SUMM`;
- defined names, table (structured) references and external references are
  refused: this API builds none of them.

At output:

- a reference to a sheet the workbook does not have is refused —
  `…formula: refers to a sheet the workbook does not have ("Costs")`;
- any circular reference is refused at its first cell —
  `sheets[0].cells.B2.formula: circular reference (Sales!B2)`;
- every formula's cached value is computed by SPEC-120's `recalculate`, a
  cell the engine cannot compute being refused. Volatile functions (`NOW`,
  `TODAY`, `RAND`) are evaluated at generation time through
  `src/lib/core/clock.ts`'s `now()` (tests pin it with `setClock`).

### Styles

#### `defineStyle(name, spec)`

Registers a named style that `style`, a cell's `style` and later calls can
use by name. Its errors are at `styles.<name>…`.

#### `style(range, spec | name)`

Applies a [CellStyleSpec](#cellstylespec) to a range through SPEC-123's
`applyStylePatch`. Ranges: `"A1:E1"`, a single cell `"B3"`, whole columns
`"B:D"`/`"B"`, whole rows `"3:5"`/`"3"` — whole rows and columns are styled
by SPEC-123's range classes (a column's own style, never a million cells).

#### `numberFormat(range, code)`

Applies a number format code — SPEC-118's en-US code syntax
(`"$#,##0.00"`, `"0.0%"`, `"yyyy-mm-dd"`, `'0.0 "kg"'`), validated by
`parseFormatCode`; `"General"` removes one.

Every style call — `style`, `numberFormat` and a cell's own `style`/
`numberFormat` — is applied in call order, so a later call wins where they
overlap, exactly as the same edits in the editor would. A refusal from
`applyStylePatch` (say, a range too large to format cell by cell) becomes the
error at that call's path.

#### `CellStyleSpec`

A validated subset of SPEC-117's `CellStyle`, mapped onto it field for field;
colours are `"#rrggbb"` only (no theme colours).

```ts
interface CellStyleSpec {
  font?: { family?: string; sizePt?: number;            // 1–409
           bold?: boolean; italic?: boolean; underline?: "single" | "double";
           strike?: boolean; color?: string };
  fill?: string;                                         // a solid fill
  borders?: Partial<Record<"top" | "right" | "bottom" | "left" | "all" | "outline",
            SheetBorderStyle | { style: SheetBorderStyle; color?: string }>>;
  alignment?: { horizontal?: "left" | "center" | "right" | "justify";
                vertical?: "top" | "center" | "bottom"; wrap?: boolean;
                indent?: number;                         // 0–250
                rotation?: number };                     // textRotation: 0–180, or 255
  numberFormat?: string;
}
```

- `fill` → `{ kind: "solid", color }`; `rotation` → `textRotation` (1–90
  counter-clockwise, 91–180 the same clockwise, 255 stacked vertically).
- `false` removes: `bold: false` undoes an earlier `bold: true` (SPEC-123's
  `null` in the patch); an `indent` or `rotation` of 0 removes it too.
- Border styles are SPEC-117's: `thin`, `medium`, `thick`, `dashed`,
  `dotted`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`,
  `dashDotDot`, `mediumDashDotDot`, `slantDashDot`; no colour is automatic.
- A named edge (`top`, `right`, `bottom`, `left`) is `CellStyle`'s — every
  cell's own edge. `all` sets every edge of every cell in the range,
  `outline` the range's outer edges only. They apply in that order — `all`,
  then `outline`, then the named edges — so a named edge wins where they
  meet.

### Sizes, merges and panes

#### `columnWidth(columns, { chars } | { px })`

`columns` is a column or columns: `"A"`, `"B:D"`. `{ chars: w }` (0–255)
stores `widthPx` by SPEC-117's maximum-digit-width formula —
`trunc(((256·w + trunc(128/MDW))/256)·MDW)`, MDW 7 for the default Calibri 11,
so 14 characters are 98 px — and keeps `sourceWidthChars: w`, so the `.xlsx`
carries the exact width; `{ px }` stores `widthPx` alone. Adjacent columns
with equal settings share one span. As in the editor (SPEC-124's
`resizeColumns`), a width under 4 px hides the column.

#### `rowHeight(rows, { pt } | { px })`

`rows` is 1-based as the user sees it: `3`, `"3"` or `"3:5"`. Sets `heightPt`
and `customHeight` (0–409 pt; `{ px }` converts at 0.75 pt per px) through
SPEC-124's `resizeRows`, which rounds to whole pixels (0.75 pt) as the
editor does; a height under 3 pt hides the row.

#### `merge(range)`

Merges a range of two or more cells through SPEC-124's
`mergeCells(model, sheetId, [range], "cells")`. Stricter than the editor,
which asks or extends: a merge that would discard a value outside its
top-left cell is refused at output
(`sheets[0].merges[2]: A1:D1 would discard the value in C1`) — a generator
has no one to ask — and one that overlaps an earlier merge is refused on
entry. Merges are applied after values and styles.

#### `freeze({ rows, cols })`

Freezes the first `rows` rows and `cols` columns (both default to 0):
`freeze({ rows: 1 })` freezes the header row, `{ rows: 0, cols: 0 }` clears.
The last call wins.

### Output

| Method | Returns |
| --- | --- |
| `toWorkbook()` | the `SpreadsheetModel`, every formula's value computed |
| `toBytes(format?, options?)` | `Promise<Uint8Array>` — the `.xlsx`/`.ods` package |
| `toBlob(format?, options?)` | `Promise<Blob>` with the format's MIME type (browser, or Node 18+) |
| `toBuffer(format?, options?)` | `Promise<Buffer>` — a `DocumentGenerationError` where `Buffer` is not global (a browser) |
| `fileName(format?)` | `spreadsheetDownload`'s naming rule: the title plus `.xlsx`/`.ods`, "Untitled spreadsheet" for a blank title |

`format` defaults to the builder's own. `options` accepts `modified` only
(any other key is an unknown property): it is passed straight to the
writers' own `modified` option (the package's stamped modified time;
defaults to now), so the same input with the same `modified` yields
byte-identical packages, part by part. The bytes are exactly
`writeXlsx(model, { modified })` / `writeOds(model, { modified })` of the
materialised model — nothing else writes them.

#### Materialisation

The builder accumulates cells in private state and builds the model once
per output — never one copy-on-write command per cell, which would be
quadratic on a 100,000-cell sheet. Each `to*` call builds a fresh model, so
one builder serves any number of outputs (and more calls after them):

1. `createBlankWorkbook()` with the title, format, date system and the
   builder's sheets;
2. per sheet, the cells sorted into rows — the last write to a cell wins;
3. the style calls, in call order, through `applyStylePatch` (consecutive
   cell-own calls batched into one call per patch);
4. merges;
5. column widths, row heights, frozen panes and the sheet options;
6. whole-workbook validation: sheet references, a visible sheet, the cell
   ceiling — a workbook past `MAX_WORKBOOK_CELLS` (2,000,000 stored cells) is
   refused with the product's own message, `This workbook has more than
   2,000,000 filled cells — too large to open here`, rather than written;
7. `recalculate` over every formula cell, a cycle refused at its first cell;
8. a structural self-check against SPEC-116's invariants — sorted, sparse,
   plain JSON (`JSON.parse(JSON.stringify(model))` equals it), an empty
   `opaque` and no `partName`/`sourceXml` anywhere.

### `WorkbookBuilder.fromDescription(description)`

The static factory: a builder holding the [description](#json-workbook-description)'s
workbook, current on its last sheet, ready for more calls or for output.

## JSON workbook description

A plain-JSON object mapping 1:1 onto the builder, validated by
`workbookFromDescription(description): SpreadsheetModel` and published as
[`schema/workbook-description.json`](../schema/workbook-description.json)
(JSON Schema draft-07, `additionalProperties: false` throughout).

```json
{ "title": "Q3 sales", "format": "xlsx", "dateSystem": "1900", "locale": "en-US",
  "styles": { "header": { "font": { "bold": true }, "fill": "#dce6f1" } },
  "sheets": [ {
    "name": "Sales", "tabColor": "#4472c4", "hidden": false, "showGridLines": true,
    "cells": { "A1": "Region", "B5": { "formula": "SUM(B2:B4)", "numberFormat": "$#,##0.00" },
               "F2": { "date": "2026-10-08" } },
    "data": [ { "at": "A2", "rows": [["East", 1200], ["West", 980]], "parseText": false } ],
    "styles": [ { "range": "A1:E1", "style": "header" },
                { "range": "B2:E5", "numberFormat": "$#,##0.00" } ],
    "columns": [ { "columns": "A", "width": { "chars": 14 } } ],
    "rows": [ { "rows": "1", "height": { "pt": 24 } } ],
    "merges": ["A10:E10"], "freeze": { "rows": 1, "cols": 0 } } ] }
```

| Key | Builder call |
| --- | --- |
| `title`, `format`, `dateSystem`, `locale` | constructor options |
| `styles` (top level, name → CellStyleSpec) | `defineStyle` per entry |
| `sheets[i]` + `name`, `tabColor`, `hidden`, `showGridLines` | `addSheet` |
| `cells` (A1 → cell input) | `setCell`, in key order |
| `data[]` (`at`, `rows`, `parseText?`) | `setRows` |
| `styles[]` (`range` + `style` and/or `numberFormat`; style first) | `style` / `numberFormat` |
| `columns[]` (`columns`, `width`), `rows[]` (`rows`, `height`) | `columnWidth`, `rowHeight` |
| `merges`, `freeze` | `merge`, `freeze` |

Within a sheet the keys apply in the table's order — cells, data, styles,
columns, rows, merges, freeze — whatever order the object lists them in.
Every field is optional but a sheet's `name`; `{}` describes a one-sheet
empty workbook. A key this spec does not build — `"names"`, `"charts"`,
`"validations"`, `"conditionalFormats"`, `"hyperlinks"`… — is an unknown
property, never silently ignored.

## Validation and paths

Every error is a `DocumentGenerationError` whose message starts with the
offending field's path in the JSON description's shape, and the fluent API
reports the same path for the same mistake — counting sheets and calls in
order — so a description and the program that would build it fail
identically:

| Mistake | Message |
| --- | --- |
| an unknown function | `sheets[0].cells.B5.formula: unknown function SUMM` |
| a typo'd key | `sheets[1].cells.C4.fomula: unknown property` |
| an unbuilt construct | `names: unknown property` |
| a bad value in a block | `sheets[0].data[1].rows[3][2]: expected a finite number` |
| a bad style field | `styles.header.font.sizePt: expected a size between 1 and 409 pt` |
| a colour | `sheets[0].styles[0].style.fill: expected a "#rrggbb" colour, got "red"` |
| a sheet name | `sheets[2].name: A sheet name can't begin or end with an apostrophe` |
| a merge | `sheets[0].merges[2]: A1:D1 would discard the value in C1` |

The fluent calls count as the description's entries do: `setRow` and
`setRows` are `data[j]`, `style` and `numberFormat` are `styles[k]`,
`columnWidth` is `columns[k]`, `rowHeight` is `rows[k]`, `merge` is
`merges[k]`, `setCell`/`setFormula` are `cells.<ref>`, and the constructor's
options are the description's top-level keys. Errors that need the whole
workbook — a missing sheet, a cycle, the cell ceiling, a merge's discarded
value — are raised by `toWorkbook()`/`toBytes()`, still before any byte.

## `readWorkbook(bytes, title?)`

Reads `.xlsx` or `.ods` bytes — a `Uint8Array`, `ArrayBuffer` or `Buffer` —
into a `SpreadsheetModel`, in the browser or in Node, with no `File`, no disk
and no browser: core's `unzip`, `detectPackage` and SPEC-116's
`spreadsheetFormatOf` decide the format from the package's own declarations
(never a file name), and `readSpreadsheetBytes` — the app's own open path —
reads it. `title` defaults the way a new workbook's does. Anything that is
not an `.xlsx`/`.ods` workbook (not a ZIP, a document or deck package, an
`.xlsb`) is `bytes: not an .xlsx or .ods workbook`.

### The round trip

Build → write → `readWorkbook` reproduces every construct this API emits, in
both formats (`workbookRoundTrip.spec.ts`): values and formulas exactly; each
cell's effective style (style indices and source names differ by design;
colours and fonts compared as they render); column widths within 1 px; row
heights; merges; frozen panes; sheet names, order, visibility and tab
colours. Number formats round-trip by code in `.xlsx` and by rendered text
(SPEC-118 `formatValue`, en-US) in `.ods`, whose number styles spell some
codes differently. ODF's own limits: `slantDashDot` borders are written as
`mediumDashDot` (ODF has no slanted dash-dot), and the 1900 system's
fictitious 1900-02-29 has no ODF date, so it re-reads as 1900-02-28 from an
`.ods`.

## Opening a generated workbook in the editor

The generated `Blob` goes through the same funnel the library's file picker
uses — no download round trip (`openWorkbookInEditor.spec.ts`):

```ts
import { openAnyDocumentFile } from "./src/lib/openDocument";
import { storeOpenedFile } from "./src/lib/documentStore";
import { editorPathFor } from "./src/lib/documentKinds";

const blob = await wb.toBlob("xlsx");
const opened = await openAnyDocumentFile(new File([blob], wb.fileName("xlsx"), { type: blob.type }));
storeOpenedFile(opened);
// router.push(editorPathFor("spreadsheet", opened.model.id)) — /sheets/:id
```

The model-level alternative stores the materialised model itself:

```ts
import { storeSpreadsheet } from "./src/lib/documentStore";

const model = wb.toWorkbook();
storeSpreadsheet(model);
// router.push(`/sheets/${model.id}`)
```

## Node notes

- Generation is DOM-free: building, materialising and both writers run in
  plain Node with no polyfill (`workbookPerformance.spec.ts` runs them with
  no `document`, `window` or `DOMParser` in scope). Only *reading* —
  `readWorkbook` — borrows `DOMParser`/`XMLSerializer`/`Node` from the
  project's `jsdom` dependency there (`nodeDom.ts`'s `ensureNodeXmlGlobals`).
- `toBuffer()` needs a global `Buffer` (Node); in a browser use `toBlob()` or
  `toBytes()`.
- A 10,000-row × 10-column workbook with a formula column, a styled header
  row and two number-formatted columns builds and writes in under 2 seconds
  per format in Node.
- Fix `modified` (and pin `NOW()`/`TODAY()` with `setClock`) for
  byte-identical output.

## Examples

| File | Demonstrates |
| --- | --- |
| `examples/sales-report.js` | monthly sales by region: SUM/AVERAGE totals, a growth-percentage formula column, currency and percent formats, a bold filled header, a frozen header row, column widths → `.xlsx` |
| `examples/csv-to-xlsx.js` | the committed `examples/data/orders.csv` through SPEC-126's `parseCsvText`, `setRows` with `parseText` (a field starting with `=` stays text), widths from content, a totals row → `.xlsx` |
| `examples/budget.js` | Income, Expenses and Summary sheets with cross-sheet formulas and a merged title, from one JSON description → `.xlsx` and `.ods` |

```sh
node examples/sales-report.js
node examples/csv-to-xlsx.js
node examples/budget.js
```

Output lands in `examples/output/`; `src/lib/generate/examples.spec.ts` runs
each and re-reads its output with `readWorkbook`. Like the other examples,
each `examples/<name>.js` is a small bootstrap running
`examples/_impl/<name>.ts` through `examples/lib/runTs.js`.
