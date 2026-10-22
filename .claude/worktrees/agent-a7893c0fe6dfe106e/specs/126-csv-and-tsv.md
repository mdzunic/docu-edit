---
id: SPEC-126
title: CSV and TSV — open, save and export
status: draft
area: io
complexity: normal
story_points: 8
depends_on: [SPEC-009, SPEC-116, SPEC-118, SPEC-121, SPEC-122]
---
## Why
More spreadsheets arrive as CSV than as anything else: the bank statement, the
CRM export, the analytics download, the file one program writes so another can
read it. Today this product refuses all of them. A dropped `.csv` fails the
drop name check with "We can only open …"; picked through the file chooser it
is not a ZIP, not Markdown, and falls through to the `.docx` path, which
reports a Word error about a file that was never a Word document. SPEC-116 put
CSV out of scope for the honest reason that it has no package to detect.

The way other tools open CSV is also the way they damage it. Excel strips the
leading zeros from postcodes and account numbers, rounds a 16-digit card or
order number to 15 significant digits, turns `1-2` and `SEPT2` into dates —
gene nomenclature was renamed in 2020 because of it — and a save writes the
damage back into the file. This product's first promise is round-trip
fidelity (SPEC-008), so it opens CSV under one rule instead: **a field becomes
a number, date or boolean only when that is lossless** — when the typed value,
written back the way this spec writes it, is exactly the text that came in.
Everything else stays text and survives untouched.

This file adds `"csv"` and `"tsv"` to the spreadsheet kind, admits them by
extension as the second deliberate exception to content sniffing (Markdown
being the first), and gives them a byte-preserving download, an "as shown"
save in the file's own dialect, an import bar to re-read the file with other
options, and "Download a copy as CSV/TSV" from any workbook. **It is the
normative source for the CSV dialect model (`CsvDialect`), the RFC 4180
parser and writer (`spreadsheet/csv/`), delimiter, encoding, line-ending and
locale sniffing, CSV field type inference, and the CSV byte-preservation
rule.**

## Acceptance criteria
- [ ] `documentKinds.ts` per the Reference: `SPREADSHEET_FORMATS` becomes `["xlsx", "ods", "csv", "tsv"]`, `kindOfFormat("csv")`/`("tsv")` is `"spreadsheet"`, `OPENABLE_EXTENSIONS` gains `csv` and `tsv` after SPEC-116's spreadsheet extensions, and `titleFromFileName` strips both; `SpreadsheetFormat` in `spreadsheet/model.ts` widens to `"xlsx" | "ods" | "csv" | "tsv"` and every switch over it is exhaustive (the typecheck proves it)
- [ ] Behavioural string change, every existing assertion updated in the same change: `UNSUPPORTED_FILE_TYPE_MESSAGE` becomes "We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .csv, .tsv, .pdf and image files" (`src/lib/openDocument.spec.ts`, `src/lib/openDocument.keynote.spec.ts`, `src/components/AppShell.spec.ts`, `e2e/SPEC-009.spec.ts`, `e2e/SPEC-050.spec.ts`, and SPEC-116's assertions of its own version of the string)
- [ ] Detection per "Detection": a file that is not media and not a ZIP is CSV/TSV when its name ends `.csv`/`.tsv` — stated in `openDocument.ts` as the exception exactly as SPEC-050 states Markdown's; a ZIP named `.csv` is still a ZIP and opens by its own declarations (an `.xlsx` renamed `.csv` opens as `.xlsx`); an OLE compound file named `.csv`/`.tsv` is refused with SPEC-116's `LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE`; `openDocumentFile` (text-only) rejects a CSV with `SPREADSHEET_NOT_TEXT_MESSAGE`
- [ ] Encoding per "Decoding": a UTF-8 BOM, UTF-16 LE BOM or UTF-16 BE BOM decides the encoding and is stripped; without a BOM the whole buffer is decoded with `new TextDecoder("utf-8", { fatal: true })` and, if that throws, with `new TextDecoder("windows-1252")`; decoded text containing U+0000 is refused with `CSV_NOT_TEXT_MESSAGE`
- [ ] `spreadsheet/csv/parse.ts` parses RFC 4180 per "Parsing": quoted fields with `""` escapes, delimiters and line breaks inside quotes, CRLF/LF/CR record ends outside quotes, a quote mid-field kept literally, text after a closing quote appended (Excel's leniency), an unterminated quote running to end of file, blank lines as empty rows, and no record for the line ending that ends the file — one pass over the string with `slice`-based field extraction, never per-character concatenation
- [ ] Delimiter sniffing per "Sniffing": a first line exactly `sep=X` (X one of `,` `;` tab `|`) decides the delimiter, is not a data row and is re-emitted on save; otherwise each candidate parses the first `SNIFF_RECORDS = 50` records, the highest consistency wins, ties go tab → pipe → semicolon → comma, and a file where no candidate yields two or more fields uses the extension's own delimiter (`,` for `.csv`, tab for `.tsv`) — asserted over a table including the "decimal commas inside a `;` file" case where splitting on `,` is also perfectly consistent
- [ ] `CsvDialect` per "The CSV model" is recorded on `SpreadsheetModel.csv` (present exactly when `format` is `"csv"` or `"tsv"`): delimiter, encoding, BOM, the first record terminator, whether the file ended with one, `quoteAll` (every non-empty field in the sample was quoted), `width` (present when every non-blank record had the same field count), `sepLine`, `localeTag` and `sheetId`; the source bytes are kept as the single `opaque` part `CSV_SOURCE_PART`; the one sheet is named per "The sheet"
- [ ] Locale inference per "Locale inference": decimal-comma evidence outvoting decimal-point evidence in the sample selects a comma-decimal locale (the user's spreadsheet locale if it is one, else `de-DE`), the reverse a point-decimal one (the user's if it is one, else `en-US`), a tie the user's; then day-first or month-first date evidence swaps `en-US` ↔ `en-GB`; the chosen tag is `CsvDialect.localeTag`, used to read the file and to write it back — the user's display locale (SPEC-118) is untouched
- [ ] Type inference per "Type inference" through SPEC-122's `parseInput` in the dialect's locale, with the deliberate divergence from Excel: a field that is empty stores nothing; a field starting with `=` or `'` stays text (never a formula); a pure-digit field with a leading zero (other than `0` and `0.x`/`0,x`) stays text; a numeric field with more than 15 digits stays text; and any other field becomes a number, date, boolean or error **only** if `csvFieldText` of the typed value and its number format equals the trimmed field exactly — so `007`, `1234567890123456`, `1-2`, `SEPT2`, `true`, `.5` and `1E5` stay text and `1.50`, `12%`, `2026-10-08`, `TRUE` and `123456789012` become typed with a format that reproduces them
- [ ] Bounds: more than `MAX_WORKBOOK_CELLS` non-empty fields → SPEC-116's `WORKBOOK_TOO_LARGE_MESSAGE`, raised as the count passes the ceiling (the parse stops; the remainder is never materialised); more than `MAX_ROWS` records → `CSV_TOO_MANY_ROWS_MESSAGE`; a record with more than `MAX_COLS` fields → `CSV_TOO_MANY_COLUMNS_MESSAGE` naming its 1-based line; `MAX_FILE_SIZE_BYTES` (50 MB) applies unchanged — nothing is silently truncated
- [ ] Byte preservation: while `sheets[0]` is the sheet read from the file (`csv.sheetId`) and is not `dirty` (SPEC-121), `spreadsheetDownload` returns the exact bytes of `CSV_SOURCE_PART` — `equalBytes` with the source for every fixture, including after a sheet was added, renamed or deleted elsewhere in the workbook and after the import options were changed
- [ ] Saving an edited CSV writes `sheets[0]` through `writeCsv` per "Writing": every cell's text is `csvFieldText` in `csv.localeTag` (SPEC-118 "as shown" at unlimited width, with General as the 15-significant-digit form), in the source's delimiter, encoding and BOM, line ending, final line ending, `sep=` line, `quoteAll` or minimal quoting, and `width` padding — so editing one field of `spec126-excel-utf8.csv` and downloading yields the source bytes with exactly that field's text replaced (asserted byte-for-byte)
- [ ] A Windows-1252 source is written back as Windows-1252 when every character is representable, and otherwise as UTF-8 with a BOM plus the toast `CSV_ENCODING_FALLBACK_NOTICE`; UTF-16 sources are written back in their own byte order with their BOM; the encoder is `encodeCsvText` per "Writing", never a `TextEncoder` label it does not support
- [ ] On download of a csv/tsv workbook whose model holds anything CSV cannot carry (`csvLosses` non-empty: another sheet with content, a formula, a merge, or a font/fill/border/alignment style on the written sheet), the download still happens and one toast reads "Formatting, formulas and other sheets aren't saved in CSV — download a copy as .xlsx to keep them" (`CSV_LOSS_NOTICE`); a download with nothing lost raises no toast
- [ ] Adding or duplicating a sheet in a csv/tsv workbook is allowed and raises, once per view session, the toast "CSV files keep only the first sheet — download a copy as .xlsx to keep the others" (`CSV_SECOND_SHEET_NOTICE`); only `sheets[0]` is ever written to a CSV
- [ ] The import bar (`csv-import-bar`) shows on every csv/tsv workbook view until dismissed (`csv-import-dismiss`) and reads per "The import bar", e.g. "Opened as semicolon-separated, Windows-1252, numbers and dates read as German · Change…"; Change… (`csv-import-change`) opens `csv-import-dialog` with Separator, Character set and "Numbers and dates" selects and a live preview of the first 20 records (`csv-import-preview`)
- [ ] The dialog re-reads the **retained source bytes** with the chosen options (`reimportCsv`): while the workbook is unedited (no sheet `dirty`, no `structureDirty`) its primary action "Apply" (`csv-import-apply`) replaces `sheets[0].rows` in place — same id, same sheet id, still unedited, undo stack reset, download still the source bytes; once edited the primary action is "Open as new workbook" (`csv-import-open-new`), which stores a new workbook titled "<title> (re-imported)" from the same bytes and navigates to it, leaving the edited one untouched
- [ ] SPEC-121's `spreadsheet-download-menu` gains "Download a copy as CSV (current sheet)" (`spreadsheet-download-copy-csv`) and "Download a copy as TSV (current sheet)" (`spreadsheet-download-copy-tsv`) in every workbook of every format: the active sheet's displayed values per "Copies", UTF-8 with a BOM, CRLF, the user's locale's `listSeparator` or a tab, the losses toast when applicable, and the open model never mutated; for a csv/tsv workbook `spreadsheet-download-copy` reads "Download a copy as Excel workbook (.xlsx)…" and a sibling `spreadsheet-download-copy-ods` reads "…as OpenDocument Spreadsheet (.ods)…", both through SPEC-121's `stripRetention` from-scratch path, reopening with the same values and inferred number formats
- [ ] The library treats both formats as first-class: `formatBadgeLabel` → `CSV`/`TSV`, `FORMAT_CHIPS` gains a `csv` and a `tsv` chip after `ods`, `FILTERABLE_FORMATS` gains both, `spreadsheetFormatLabel` gives "Comma-separated values (.csv)" / "Tab-separated values (.tsv)", the card shows the SPEC-116 `SheetMiniature` and "1 sheet", and search, sort, star, trash, restore and bulk download need no new code path
- [ ] `upgradeSpreadsheetModel` never runs a `ReaderUpgrade` on a csv/tsv model: an entirely unedited one behind the current `SPREADSHEET_READER_VERSION` is re-read through `readCsvBytes` with its recorded dialect (keeping `id`, `title`, starred state and `sheetId`), any other is only re-stamped; this spec adds no element to the xlsx/ods readers and does not bump the version
- [ ] Fixtures per "Fixtures", written as literal bytes by `spreadsheet/fixtures/build.ts` (never by `writeCsv`, which would only prove our writer agrees with itself): `e2e/fixtures/spec126-excel-utf8.csv`, `spec126-semicolon-1252.csv`, `spec126-quoted.csv`, `spec126-sep.csv`, `spec126-unicode.tsv`, `spec126-mac-pipe.csv`
- [ ] Tests: `spreadsheet/csv/parse.spec.ts`, `sniff.spec.ts`, `write.spec.ts`, `infer.spec.ts` (the inference table in "Type inference", both directions of every row), `csv/roundTrip.spec.ts` (untouched `equalBytes` for every fixture; one-field edits byte-exact; re-open of every written file reproduces the cells under `toEqual`), `openDocument.csv.spec.ts` (detection order, the ZIP and OLE cases, text-only rejection), `views/spreadsheet/useCsvImport.spec.ts`, `HomeView.spreadsheet.spec.ts` additions
- [ ] e2e (`e2e/SPEC-126.spec.ts`): drop `spec126-semicolon-1252.csv` → a row with a `CSV` badge → open → `csv-import-bar` reads "Opened as semicolon-separated, Windows-1252, numbers and dates read as German · Change…" → `A2` reads "Müller", `B2` reads "01234" left-aligned, `C2` reads "1234.5" right-aligned → Change… → Comma → Apply → `A1` reads "Name;Zip;Amount" → Change… → Semicolon → Apply → download → bytes `equalBytes` the fixture → type `99` into `C2` → download → bytes equal the fixture with `1234,5` replaced by `99` → type `=1+1` into `D2` → download → the `CSV_LOSS_NOTICE` toast → open `mini-book.xlsx` → `spreadsheet-download-menu` → `spreadsheet-download-copy-csv` → bytes start `EF BB BF`, records end CRLF, and one record contains `Revenue`

## Out of scope
Fixed-width text, `.txt`/`.prn`/`.dif`/`.slk` import, and any other extension
(`.tab`, `.psv`) — the exception is exactly two extensions. Excel's Text
Import Wizard: per-column data types, skipped rows, a chosen header row,
"treat consecutive delimiters as one", custom or multi-character delimiters.
Encodings other than the four here (Shift-JIS, ISO-8859-2, UTF-32, …): they
decode as Windows-1252 and the user sees it. Formulas in CSV — never
evaluated; a field starting `=` is text. Sanitising formula-like text on
export ("CSV injection" prefixes): the writer writes the data, as Excel's own
CSV export does. Opening a CSV URL whose derived file name lacks the
extension (the Content-Type is not consulted, as for Markdown). Parsing in a
worker and streaming large files (SPEC-137). Converting a CSV workbook into an
`.xlsx` in place — the path is "Download a copy as .xlsx" and open the copy.
Autofit of column widths at open (SPEC-124 owns autofit). Locale-specific
formula syntax (whole-queue backlog).

## Reference

### Dependencies restated
- **`specs/116-workbooks-in-the-library.md`** — `SpreadsheetModel { kind:
  "spreadsheet"; id; title; format; packageVariant?; readerVersion; sheets;
  activeSheetIndex; dateSystem; opaque: OpaquePart[]; preview?; … }`, `Sheet
  { id; name; visibility; kind; rows; … }`, `Row { index; cells }`, `Cell
  { col; value; … }`, `CellValue = number | string | boolean | CellError |
  null`, the copy-on-write, sparse and plain-JSON invariants;
  `MAX_ROWS = 1_048_576`, `MAX_COLS = 16_384`, `MAX_WORKBOOK_CELLS =
  2_000_000`; `OpaquePart { name; data /* base64 */ }` (`core/packageParts.ts`,
  with `bytesToBase64`/`decodedPartBytes`); `readSpreadsheetBytes`,
  `upgradeSpreadsheetModel`, `READER_UPGRADES`, `SPREADSHEET_READER_VERSION`
  in `open.ts`; `spreadsheetDownload` in `download.ts` (as SPEC-121 extends
  it: `spreadsheetDownload(model, targetFormat = model.format): Promise<{
  bytes; fileName; mimeType; notices: string[] }>`, `notices` being whole
  one-line toast messages); `buildPreview(model)`; `spreadsheetFormatLabel`;
  `storeSpreadsheet`; `WORKBOOK_TOO_LARGE_MESSAGE`,
  `LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE`, `SPREADSHEET_NOT_TEXT_MESSAGE`.
- **`specs/009-…`** — `openAnyDocumentFile` in `src/lib/openDocument.ts` is
  the one funnel (picker, drop, `/open?url=`); today's order inside
  `openFile`: size check → `detectMedia` → one `unzip` → if not a ZIP:
  flat-ODF check, `.md`/`.markdown` by extension, `.key` message, `.docx`
  path. `hasSupportedExtension` gates drops by `OPENABLE_EXTENSIONS`.
  `storeOpenedFile(opened, notify)` in `documentStore.ts` raises a toast
  through `useToasts().addToast` (`src/composables/useToasts.ts`).
- **`specs/050-markdown-open-edit-and-convert.md`** — the precedent this spec
  follows to the letter: "A file that is not a ZIP package is markdown if its
  name ends `.md` or `.markdown`, and unsupported otherwise. ZIP packages
  continue to be identified by content exactly as they are today."
- **SPEC-117** (through 118) — `SpreadsheetModel.styles: CellStyle[]` (0 =
  Normal), resolution cell → row (if `customFormat`) → column → 0,
  `CellStyle.numberFormat?` (absent = General), `font`/`fill`/`borders`/
  `alignment`; `Sheet.merges: GridRange[]`.
- **`specs/118-…`** — `formatValue(value, code, ctx: FormatContext):
  FormattedValue`, `FormatContext { locale: SpreadsheetLocale; dateSystem;
  widthPx?; … }` — with no `widthPx` there is no `########` and no width
  fitting; `plainText(fv)` (the text with one fill character, "for CSV,
  clipboard, preview"); `editText(cell, ctx)`, whose plain notation for a
  number is up to 15 significant digits, no grouping, `locale.decimal`,
  scientific only when |n| ≥ 1e15 or 0 < |n| < 1e−9.
  `numberFormat/locale.ts`: `SpreadsheetLocaleTag` (`en-US`, `en-GB`,
  `de-DE`, `fr-FR`, `es-ES`, `cs-CZ`), `SPREADSHEET_LOCALES`,
  `SpreadsheetLocale { tag; decimal; group; listSeparator /* "CSV (SPEC-126)" */;
  dateOrder: "MDY" | "DMY" | "YMD"; … }`. The locale is
  `Preferences.spreadsheetLocale`, resolved only by the app layer
  (`useSpreadsheetLocale.ts` in the view) and passed down — the product layer
  never reads preferences. Format codes are stored in en-US invariant
  syntax. `buildPreview(model, locale)`.
- **`specs/121-…`** — `Sheet.dirty?: true`, `SpreadsheetModel.structureDirty?:
  true`; `useSpreadsheetUndo` `{ record(label, scope), commit, undo, redo,
  reset }`; `stripRetention(model)`, `writeXlsx`, `writeOds`; the app bar's
  split "Download ▾" (`spreadsheet-download`, caret `spreadsheet-download-menu`,
  item `spreadsheet-download-copy` reading "Download a copy as Excel workbook
  (.xlsx)…" or "…OpenDocument Spreadsheet (.ods)…"); `addSheet(model,
  afterIndex)` and duplicate.
- **`specs/122-…`** — `parseInput(text, locale, dateSystem, options?): {
  value: CellValue; formula?: string; numberFormat?: string;
  builtinFormatId?: number; quotePrefix?: true }` in `numberFormat/input.ts`
  (a locale-dependent date or date-time carries `builtinFormatId` 14/22 with
  its en-US code in `numberFormat`); `internStyle(model, style): number` in
  `styles.ts`.

### The CSV model
```ts
// src/lib/spreadsheet/model.ts (additions)
export type SpreadsheetFormat = "xlsx" | "ods" | "csv" | "tsv";
export type CsvDelimiter = "," | ";" | "\t" | "|";
export type CsvEncoding = "utf-8" | "utf-16le" | "utf-16be" | "windows-1252";
export interface CsvDialect {
  delimiter: CsvDelimiter;
  encoding: CsvEncoding;
  bom: boolean;                         // the source began with its encoding's BOM
  lineEnding: "\r\n" | "\n" | "\r";     // the first record terminator outside quotes ("\r\n" if none)
  finalLineEnding: boolean;             // the source's last record was terminated
  quoteAll: boolean;                    // every non-empty field in the sniff sample was quoted
  width?: number;                       // fields per record, when every non-blank record agreed
  sepLine?: true;                       // the source began with Excel's "sep=X" line
  localeTag: SpreadsheetLocaleTag;      // the locale the fields were read in
  sheetId: string;                      // the sheet read from the file
}
export interface SpreadsheetModel { /* …SPEC-116/117/121 fields… */ csv?: CsvDialect }

// src/lib/spreadsheet/csv/parse.ts
export const CSV_SOURCE_PART = "csv-source";
export interface CsvReadOptions {
  userLocaleTag: SpreadsheetLocaleTag;  // Preferences.spreadsheetLocale, passed by the app layer
  delimiter?: CsvDelimiter;             // given → that facet is not sniffed
  encoding?: CsvEncoding;
  localeTag?: SpreadsheetLocaleTag;
}
export class CsvOpenError extends Error {}
export function readCsvBytes(bytes: Uint8Array, format: "csv" | "tsv", title: string,
  options: CsvReadOptions): SpreadsheetModel;
export function reimportCsv(model: SpreadsheetModel, options: CsvReadOptions): SpreadsheetModel;
export function parseCsvText(text: string, delimiter: CsvDelimiter,
  onRecord: (fields: string[], quoted: boolean[], line: number) => void | "stop"): CsvShape;
export interface CsvShape { lineEnding: CsvDialect["lineEnding"]; finalLineEnding: boolean;
  width?: number; records: number }
export function inferCsvField(field: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"):
  { value: CellValue; numberFormat?: string };
```
`opaque` is exactly `[{ name: CSV_SOURCE_PART, data: bytesToBase64(source) }]`
— the bytes as they came in, BOM included — and is what byte preservation,
re-import and the reader-version re-read all read. `dateSystem` is `"1900"`.
`readCsvBytes` is synchronous; SPEC-137 owns moving it off the main thread.
`reimportCsv` returns a new model object (it mutates nothing it was given).

### The sheet
One `kind: "grid"`, visible sheet whose `name` is the workbook title with
each of `[ ] : * ? / \` replaced by `_`, truncated to 31 characters, or
"Sheet1" if that leaves nothing — as Excel names an opened CSV's sheet after
its file.
Record *n* (0-based, after any `sep=` line) is row *n*, field *m* column *m*;
blank lines advance the row index and store nothing; empty fields store
nothing. A typed field's `numberFormat` becomes the cell's `styleId` through
`internStyle(model, { numberFormat })`, so `styles` holds Normal plus one
style per distinct inferred format. `preview` is `buildPreview(model,
SPREADSHEET_LOCALES[userLocaleTag])`.

### Detection
`openFile`'s non-ZIP branch gains one test, placed after the flat-ODF check
and the Markdown test and before the `.key` message. Stated in
`openDocument.ts` as SPEC-050 states Markdown's:

> A file that is not media and not a ZIP package is CSV if its name ends
> `.csv` and TSV if it ends `.tsv`. ZIP packages, media and flat ODF continue
> to be identified by content exactly as they are today.

So media sniffing (SPEC-107) and the unzip attempt still run first: a ZIP
named `.csv` is a ZIP. `openFile` passes `loadPreferences().spreadsheetLocale`
as `userLocaleTag`. Before decoding, a `.csv`/`.tsv` beginning with the
CFB signature `D0 CF 11 E0 A1 B1 1A E1` is refused with
`LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE` (an `.xls` renamed to look like a
CSV). SPEC-107's sentence ".md keeps its extension exception; nothing else
gains one" now reads ".md, .csv and .tsv keep their extension exception".

### Decoding
| Leading bytes | `encoding` | `bom` | Decoder |
| --- | --- | --- | --- |
| `EF BB BF` | utf-8 | true | `TextDecoder("utf-8")` after the BOM |
| `FF FE` | utf-16le | true | `TextDecoder("utf-16le")` after the BOM |
| `FE FF` | utf-16be | true | `TextDecoder("utf-16be")` after the BOM |
| anything, valid UTF-8 | utf-8 | false | `TextDecoder("utf-8", { fatal: true })` succeeded |
| anything else | windows-1252 | false | `TextDecoder("windows-1252")` |

A pure-ASCII file is `utf-8`. A U+0000 anywhere in the decoded text is
`CSV_NOT_TEXT_MESSAGE` — binary data with a CSV name, not a spreadsheet.

### Parsing
A single-pass state machine over the decoded string (after the BOM):
- A field is **quoted** only if its first character is `"`. Inside it, `""`
  is one `"`; the delimiter, CR and LF are data; CRLF and lone CR inside
  quotes become LF in the cell value (the grid's in-cell line break). The
  closing `"` ends quoting; any characters between it and the next delimiter
  or record end are appended to the value as they are (`"ab"cd` → `abcd`).
  End of file inside quotes ends the field there — no error.
- In an unquoted field every character but the delimiter, CR and LF is data,
  `"` included (`ab"cd` stays `ab"cd`). A space before an opening quote makes
  the field unquoted (` "b"` stays ` "b"`), per RFC 4180.
- Outside quotes CRLF, LF and lone CR each end a record. A line ending at the
  very end of the file ends the last record and starts none. A record of
  exactly one empty field is a blank line.
- `quoted[]` tells the sniffer which fields were quoted (for `quoteAll`).

### Sniffing
1. **`sep=` line.** If the first line (up to the first CR/LF) is exactly
   `sep=` followed by one candidate character, that is the delimiter,
   `sepLine` is set, and parsing starts on the next line.
2. **Delimiter.** Otherwise for each candidate `d` in `\t`, `|`, `;`, `,`:
   parse the first `SNIFF_RECORDS = 50` non-blank records with `d`; the
   candidate is eligible when its modal field count is at least 2;
   `consistency(d)` = records with the modal count ÷ non-blank records,
   compared exactly as fractions. The eligible candidate with the highest
   consistency wins; ties go to the earlier candidate in that order. The
   order is by how rarely each character appears inside ordinary field
   content: a `;` file with decimal commas also splits perfectly on `,`, but
   a `,` file almost never has the same number of semicolons on every line.
   No eligible candidate → `,` for `.csv`, tab for `.tsv`.
3. **Line ending** — the first terminator found outside quotes.
4. **`quoteAll`** — true when the sample has at least one non-empty field
   and every non-empty field in it was quoted.
5. **`width`** — computed over the whole file during the full parse: set to
   the shared field count when every non-blank record has the same count.

### Locale inference
Over the fields of the sniff sample, trimmed:
- decimal-comma evidence: `/^[+-]?\d+,\d+$/` and `/^[+-]?\d{1,3}(\.\d{3})+,\d+$/`,
  excluding the ambiguous `/^[+-]?\d{1,3},\d{3}$/`;
- decimal-point evidence: `/^[+-]?\d+\.\d+$/` and `/^[+-]?\d{1,3}(,\d{3})+\.\d+$/`,
  excluding the ambiguous `/^[+-]?\d{1,3}\.\d{3}$/`.
More comma evidence → a comma-decimal locale: the user's if its `decimal` is
`,`, else `de-DE`. More point evidence → the user's if its `decimal` is `.`,
else `en-US`. Equal (including none) → the user's. Then, over fields matching
`/^\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}$/`: if the chosen locale is `en-US`, some
first component exceeds 12 and no second component does, use `en-GB`; if it
is `en-GB`, some second component exceeds 12 and no first does, use `en-US`.
(The comma-decimal locales are all day-first; nothing swaps them.) A `;`
delimiter with decimal commas therefore reads as German for an English user,
and the import bar says so.

### Type inference
`inferCsvField(field, locale, dateSystem)`, with `t = field.trim()`:
1. `field === ""` → nothing stored.
2. Text, verbatim (`value: field`, untrimmed), when `t` starts with `=` (CSV
   carries data, not formulas — and `=HYPERLINK(…)`/DDE payloads are the
   classic CSV attack) or `'` (CSV has no quote-prefix convention; the
   apostrophe is data); when `t` matches `/^[+-]?0\d+([.,]\d*)?$/` (a
   leading zero: `007`, `00`, `-0123`; `0`, `0.5`, `0,5` are not caught);
   or when `t` is made only of digits, an optional sign, `.`, `,` and spaces
   and holds more than 15 digits.
3. `r = parseInput(t, locale, dateSystem)`. `r.formula` present or
   `r.value` a string → text verbatim.
4. For each candidate format, in order — `r.numberFormat` (absent =
   General), then the fidelity formats below — if `csvFieldText(r.value,
   candidate, locale, dateSystem) === t`, store `r.value` with that format.
5. Otherwise text verbatim.

Fidelity formats, derived from the shape of `t` (en-US code syntax): for a
number with *n* decimals in `t`, `0` (n = 0) or `0.` + *n* zeros; the same
with a `#,##0` integer part when `t` contains the locale's group separator;
each with `%` appended when `t` ends in `%`. For a date or time:
`yyyy-mm-dd`, `yyyy-mm-dd hh:mm`, `yyyy-mm-dd hh:mm:ss`, `hh:mm`,
`hh:mm:ss`. Nothing else — a spelling no format reproduces stays text.

| Field (en-US) | Stored | Why |
| --- | --- | --- |
| `007`, `-0123` | text | leading zero |
| `1234567890123456` | text | 16 digits; a double keeps 15 |
| `1-2`, `SEPT2`, `1E5`, `.5`, `+3`, `true` | text | no format writes them back unchanged |
| `1.50` | 1.5, `0.00` | fidelity format |
| `12%` | 0.12, `0%` | |
| `2026-10-08` | serial, `yyyy-mm-dd` | |
| `10/8/2026` | serial, `parseInput`'s date format | if it renders back identically |
| `123456789012` | number, General | General here is the 15-digit form |
| `TRUE`, `#N/A` | boolean, error | |

The divergence from Excel is the point: every rule above keeps a value as
text that Excel would have changed, and each one is a known, documented
data-loss pattern (postcodes, phone and account numbers, card numbers,
identifiers that look like dates). Text still sorts, filters and matches; a
user who wants a number converts it deliberately.

### Writing
`src/lib/spreadsheet/csv/write.ts`:
```ts
export function csvFieldText(value: CellValue, numberFormat: string | undefined,
  locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): string;
export type CsvLoss = "sheets" | "formulas" | "merges" | "formatting";
export function csvLosses(model: SpreadsheetModel, sheetIndex: number): CsvLoss[];
export function sheetToCsv(model: SpreadsheetModel, sheetIndex: number,
  dialect: Omit<CsvDialect, "sheetId">): { bytes: Uint8Array; encodingFellBack: boolean };
export function writeCsv(model: SpreadsheetModel): { bytes: Uint8Array; notices: string[] };
export function encodeCsvText(text: string, encoding: CsvEncoding, bom: boolean): Uint8Array | null;
```
- **Field text.** `null` → `""`; a string → itself; booleans `TRUE`/`FALSE`;
  an error → its code; a number under General → SPEC-118's `editText` plain
  notation (15 significant digits, `locale.decimal`, no grouping, scientific
  only beyond 1e15 / below 1e−9) rather than the 11-character General — that
  one is a fit-to-a-column rule and a file has no column; any other format →
  `plainText(formatValue(value, code, { locale, dateSystem }))` with no
  `widthPx`. A formula cell contributes its value (cached, or computed by
  SPEC-120).
- **Records.** Rows 0 … `usedRange(sheet).bottom`, from column A. A row with
  no stored cell is an empty line. With `width` set, every non-empty record
  has `max(width, usedRange.right + 1)` fields; without it, each record ends
  at its last stored cell. Merge-covered cells are empty; an empty sheet
  writes no records.
- **Quoting.** A field is quoted when it contains the delimiter, `"`, CR or
  LF, or when `quoteAll` holds and the field is non-empty; `"` is doubled.
  In-cell LF is written as LF inside the quotes.
- **Line endings.** `lineEnding` between records, and after the last when
  `finalLineEnding`. `sepLine` → `sep=` + delimiter + `lineEnding` first.
- **Encoding.** UTF-8 via `TextEncoder`, BOM `EF BB BF` when `bom`; UTF-16
  by code unit in the source's byte order, BOM when `bom`; Windows-1252
  through a reverse table built once by decoding bytes `00`–`FF` with
  `TextDecoder("windows-1252")` — a character outside it makes
  `encodeCsvText` return `null`, and `writeCsv` then writes UTF-8 with a BOM
  and adds `CSV_ENCODING_FALLBACK_NOTICE`. A download never mutates
  `model.csv`; the next one decides again.
- **Losses** (`csvLosses`): `"sheets"` when any other sheet stores a cell;
  `"formulas"` when the written sheet has a `formula`; `"merges"` when it has
  merges; `"formatting"` when any cell, row or column on it resolves to a
  style with `font`, `fill`, `borders` or `alignment` set. Number formats
  are not a loss — their effect is written. Column widths, row heights,
  hidden rows and frozen panes are view settings and are not reported.
- `spreadsheetDownload` (SPEC-121's signature, unchanged) for csv/tsv returns
  the source bytes or `writeCsv`'s output, appending `CSV_LOSS_NOTICE` when
  `csvLosses(model, 0)` is non-empty and `writeCsv`'s
  `CSV_ENCODING_FALLBACK_NOTICE` when the encoding fell back, to the same
  `notices` list SPEC-121's cross-format copy fills. File name
  `<title>.csv` / `<title>.tsv`; MIME `text/csv` / `text/tab-separated-values`.

### Copies
"CSV (current sheet)" / "TSV (current sheet)" call `sheetToCsv` on the
active sheet with `{ delimiter: the user's locale's `listSeparator` (`,` for
en-US/en-GB, `;` where the decimal is a comma), or "\t"; encoding: "utf-8";
bom: true; lineEnding: "\r\n";
finalLineEnding: true; quoteAll: false; width: usedRange.right + 1;
localeTag: the user's }` — Excel's "CSV UTF-8" shape, because a BOM-less
UTF-8 file is the one CSV Excel misreads. File name `<title>.csv` for a
one-sheet workbook, else `<title> - <sheet name>.csv`. `CSV_LOSS_NOTICE`
applies with `csvLosses(model, activeSheetIndex)`, minus `"sheets"`.

### The import bar
`src/components/spreadsheet/CsvImportBar.vue`, `CsvImportDialog.vue`, wired
by `src/views/spreadsheet/useCsvImport.ts`; SpreadsheetView gains only the
`useCsvImport({ … })` call and the two template hookups.
Text: "Opened as " + delimiter word ("comma-separated", "semicolon-separated",
"tab-separated", "pipe-separated") + ", " + encoding word ("UTF-8",
"UTF-16", "Windows-1252") + (when `localeTag` ≠ the user's locale) ", numbers
and dates read as " + `languageLabel(localeTag)` (`core/language.ts`) + " ·
Change…". Dialog title "Text import options"; selects Separator (Comma,
Semicolon, Tab, Pipe), Character set ("Unicode (UTF-8)", "Unicode (UTF-16
LE)", "Unicode (UTF-16 BE)", "Western (Windows-1252)") and "Numbers and
dates" (the six locales by `languageLabel`), preselected from `model.csv`;
the preview re-parses the retained bytes with `"stop"` after 20 records on
every change. An explicitly chosen encoding strips its BOM when present.

### Messages
| Constant | String |
| --- | --- |
| `CSV_NOT_TEXT_MESSAGE` | "This file doesn't contain text we can read as CSV" |
| `CSV_TOO_MANY_ROWS_MESSAGE` | "This file has more than 1,048,576 rows — too many to open here" |
| `CSV_TOO_MANY_COLUMNS_MESSAGE` | "Line N has more than 16,384 fields — too many to open here" |
| `CSV_LOSS_NOTICE` | "Formatting, formulas and other sheets aren't saved in CSV — download a copy as .xlsx to keep them" |
| `CSV_SECOND_SHEET_NOTICE` | "CSV files keep only the first sheet — download a copy as .xlsx to keep the others" |
| `CSV_ENCODING_FALLBACK_NOTICE` | "Some characters can't be saved as Windows-1252, so this file was saved as UTF-8" |

### Fixtures
Literal bytes from `spreadsheet/fixtures/build.ts`, each emulating a producer:
- `spec126-excel-utf8.csv` — Excel "CSV UTF-8": BOM, CRLF after every
  record, minimal quoting, a field with a comma, one with an in-cell LF, one
  with `""`; `10/8/2026`, `12%`, `1.50`, `TRUE`, `007`, a 16-digit number,
  `1-2`, `SEPT2`, empty fields; rectangular.
- `spec126-semicolon-1252.csv` — a German Excel export: Windows-1252, `;`,
  CRLF, header `Name;Zip;Amount`, row 2 `Müller;01234;1234,5`, a `€` and
  `1.234,56` elsewhere, a `31.12.2025` date.
- `spec126-quoted.csv` — a database dump: UTF-8, no BOM, LF, every non-empty
  field quoted, ragged records, no line ending after the last record.
- `spec126-sep.csv` — `sep=;` first line, then three `;` records.
- `spec126-unicode.tsv` — Excel "Unicode Text": UTF-16 LE BOM, tab, CRLF, a
  quoted field holding a tab and a `""`, Czech, Greek and an astral-plane
  emoji.
- `spec126-mac-pipe.csv` — classic Mac: lone CR records, `|` delimiter,
  Windows-1252.

### Tests
- `csv/parse.spec.ts` — every rule in "Parsing" (including LF normalisation
  inside quotes, the unterminated quote, the trailing-newline record, blank
  lines, the cell ceiling stopping the parse) and `readCsvBytes` limits.
- `csv/sniff.spec.ts` — the decoding table, NUL refusal, `sep=`, the
  delimiter table and tie order, line endings, `quoteAll`, `width`, the
  locale table (all six user locales × comma/point/tie × date evidence).
- `csv/infer.spec.ts` — the inference table, both locales of every numeric
  row, `1900`/`1904`.
- `csv/write.spec.ts` — `csvFieldText` per type, quoting, padding, every
  encoding including the Windows-1252 fallback, `csvLosses`.
- `csv/roundTrip.spec.ts` — untouched downloads `equalBytes`
  (`src/lib/core/equality.ts`) for all six fixtures; a one-field edit is
  byte-exact for `spec126-excel-utf8.csv`, `spec126-semicolon-1252.csv` and
  `spec126-unicode.tsv`; write → `readCsvBytes` → cells `toEqual` for all.
- `openDocument.csv.spec.ts`, `useCsvImport.spec.ts`,
  `HomeView.spreadsheet.spec.ts` additions, `e2e/SPEC-126.spec.ts`.

### Specs this changes
- `specs/009-…`/`specs/050-…`/`specs/116-…` — the accepted extensions and the
  rejection string grow; their assertions are updated with them. SPEC-050's
  exception gains a sibling, stated the same way.
- `specs/107-…` — the "nothing else gains one" sentence, as above.
- `specs/116-…` — `SpreadsheetFormat` widens; `spreadsheetDownload` handles
  csv/tsv (adding its notices to SPEC-121's `notices` list);
  `upgradeSpreadsheetModel` skips `READER_UPGRADES` for csv/tsv.
- `specs/121-…` — the Download menu gains two entries (three for csv/tsv
  workbooks); `spreadsheetDownload`'s `notices` gains the CSV notices; the
  add/duplicate
  sheet commands raise `CSV_SECOND_SHEET_NOTICE` on a csv/tsv workbook.
- `specs/122-…` — `parseInput` gains a headless caller; the sheet writer is
  never used for csv/tsv.
- `specs/000-roadmap.md` — no change here (the lead owns it).
