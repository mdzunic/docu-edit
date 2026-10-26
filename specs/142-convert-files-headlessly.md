---
id: SPEC-142
title: Convert files headlessly — within each product, in Node
status: draft
area: platform
complexity: normal
story_points: 8
depends_on: [SPEC-141, SPEC-008, SPEC-036, SPEC-050, SPEC-087, SPEC-106, SPEC-107, SPEC-116, SPEC-121, SPEC-126]
---
## Why
`docu-edit convert example.docx --to md` is the line most people will type
first, and the product already knows how to do every part of it: the open
funnel identifies and reads every extension in `OPENABLE_EXTENSIONS`, the writers emit
`.docx`, `.odt`, `.md`, `.pptx`, `.odp`, `.xlsx`, `.ods`, `.csv` and `.tsv`,
and each product's "download a copy as…" path already reports what a
cross-format save cannot carry. What none of it can do is run outside a
browser tab.

It nearly can. Run under plain Node behind SPEC-045's `ensureNodeXmlGlobals()`,
every committed fixture opens and every same-product conversion writes — the
one thing that breaks is the funnel's very first step, `core/fileReader.ts`,
which turns the picked `File` into bytes with a `FileReader`, and Node has
none. (`core/packageParts.ts` already falls back when `FileReader` is
missing; `merge.ts` uses one but is not on the open path.) A measured cold start (open a
`.docx`, write it as `.odt` and `.md`, open an `.xlsx`, write `.ods` and
`.csv`, open a `.pptx` and a `.key`, write `.odp`, `.pptx` and `.html`) took
1.5 seconds on the development Mac.

So this spec is mostly plumbing, and deliberately so: it removes the
`FileReader` dependency, puts one headless open function in front of the
funnel, and registers the conversions that stay inside one product — text ↔
text, deck ↔ deck (plus a deck's self-contained HTML), workbook ↔ workbook —
under SPEC-141's grammar, report and exit codes. It also adds `info`, the
cheap question a script asks before it converts. Conversions that need a
rendering engine (PDF, images) are SPEC-143/144; conversions between products
are SPEC-145. **This file is the normative source for `openBytes` and the
Node-environment rule, the same-product conversion rows, input expansion and
output naming for `convert`, the sheet-selection options, and the `info`
record.**

## Acceptance criteria
- [ ] The open funnel runs in Node: `core/fileReader.ts`'s `readFileAsArrayBuffer` keeps its `FileReader` path when `FileReader` is defined and otherwise resolves `file.arrayBuffer()`, with identical bytes; nothing else in the funnel changes (`packageParts.ts`'s `bytesToBase64Async` already falls back, and `loadPreferences()` already returns its defaults when storage is unavailable — a test pins both under Node), and every existing jsdom test passes unchanged
- [ ] `src/lib/convert/open.ts` exports `openBytes(bytes: Uint8Array, fileName: string, options?: { spreadsheetLocale?: string }): Promise<OpenedFile>` — `await prepareHeadless()` then `openAnyDocumentFile(new File([bytes], fileName), …)`, or `openTemplateFile` for a `.dotx`/`.ott` name (returned as `{ kind: "text", model }`) — and `OpenFileOptions` gains `spreadsheetLocale?: string`, which the CSV/TSV branch uses in place of `loadPreferences().spreadsheetLocale` when given (the app passes none, so its behaviour is unchanged); `src/lib/convert/open.spec.ts` (`@vitest-environment node`, no `FileReader`, no `DOMParser` until `prepareHeadless()`) opens every **committed** file in `e2e/fixtures/` (the list comes from `git ls-files e2e/fixtures`, so on-demand fixtures that `.gitignore` names — `spec-047-fixture.docx`, `spec137-*.xlsx` — are never swept) whose extension is an accepted input, and asserts the kind its extension implies — except `spec107-encrypted.pdf`, which must fail with SPEC-107's password-protected message
- [ ] `docu-edit convert <input>… --to <ext>` (alias `--ext`, `-t`; so `docu-edit -convert example.docx -ext md` works as typed) converts each input through `openBytes` → `writeAs` and adds these `CONVERSIONS` rows, all tier `"node"`: text (`docx` `odt` `md`, plus `dotx`/`ott` templates) → `docx`, `odt`, `md`; presentation (`pptx` `odp` `key`) → `pptx`, `odp`, `html`; workbook (`xlsx` `xlsm` `xltx` `xltm` `ods` `ots` `csv` `tsv`) → `xlsx`, `ods`, `csv`, `tsv`; a target the input's product has no row for is refused (exit 3) with `can't convert <a presentation> to .<ext> — try one of: pptx, odp, html` listing that product's rows in registry order
- [ ] Bytes equal the app's: for every row, the file `convert` writes is `equalBytes`-identical to `writeAs(await openBytes(bytes, name), target, { modified })` with the same pinned `--modified`, so the command adds nothing to what the app's own "download a copy as…" writes; same-format conversion is allowed and is the app's same-format save (byte-preserving where the writers retain the source package)
- [ ] Accepted inputs are `OPENABLE_EXTENSIONS` plus `dotx` and `ott`; a template input opens through SPEC-046's `openTemplateFile` (which strips the template declaration, as "New from template" does), so `convert letter.dotx --to docx` writes an ordinary document; a workbook **package variant** converted to its own base format — `xlsm`, `xltx`, `xltm` → `xlsx`, `ots` → `ods` — is refused (exit 3) with `<file> is a <macro-enabled workbook | workbook template>, and docu-edit can't change a workbook's package variant yet — convert it to another format, or open it in the app`, because `spreadsheetDownload`'s same-format branch would write the variant's bytes under a plain extension
- [ ] A presentation to `html` is SPEC-106's `exportDeckHtml(model)` written as one self-contained `<stem>.html` (UTF-8); a Keynote input's `importNotice` (SPEC-087) becomes a report notice on every conversion of it
- [ ] Sheet selection per "Sheet selection" — one rule every workbook conversion of the phase follows: the active sheet by default, `--sheet <name | 1-based index>` for another, `--all-sheets` for all; an unknown name or index is exit 2 with `<file> has no sheet "<name>" — its sheets are: <names>` for every target; both options are usage errors only with `xlsx`/`ods` targets (which keep every sheet) and with non-workbook inputs
- [ ] Workbook to `csv`/`tsv` writes the selected sheet through SPEC-126's `spreadsheetSheetCopy(model, target, locale)`, or with `--all-sheets` one file per sheet — hidden sheets included, since CSV is data — named `<stem> - <sheet name>.<ext>` with `/ \ : * ? " < > |` and control characters in the sheet name replaced by `_` (names derived from content, so collisions are checked at write time per SPEC-141); SPEC-126's "Formatting, formulas and other sheets aren't saved in CSV…" notice reaches the report as it reaches the app's toast
- [ ] `--locale <tag>` (one of `SPREADSHEET_LOCALE_TAGS`; default `en-US`) is passed as the tag `spreadsheetLocale` when opening `.csv`/`.tsv` and, resolved to its `SpreadsheetLocale` object through `SPREADSHEET_LOCALES[tag]`, as the `locale` of `spreadsheetDownload`/`spreadsheetSheetCopy` when writing — exactly where the app passes the user's number-locale preference; an unknown tag is a usage error listing the tags
- [ ] Inputs per "Inputs": files; directories only with `--recursive` (`-r`), walked in code-point order of their relative paths, skipping names that begin with `.`, never following symbolic links (`io.lstat`), and taking only files whose extension is an accepted input; `-` reads standard input and requires `--stdin-name <file name>` (the name the funnel detects by); a directory without `--recursive` and an input that does not exist are each reported for that input (exit 1 and exit 2 respectively) without stopping the others
- [ ] Outputs per "Output naming": `<stem>.<ext>` beside its input by default; `--out-dir <dir>` (created, with parents, when missing) instead, mirroring each file's path relative to the directory argument it was found under; `-o <file>` (or `-o -` for standard output) only with exactly one input that yields one file; an output equal to its own input path is refused (exit 4) unless `--force`; a collision between outputs whose paths follow from the arguments and the walk alone is a usage error found before any input is read, and one between content-derived names is exit 4 at write time; SPEC-141's overwrite and atomic-write rules apply to every file
- [ ] `--strict` refuses (exit 3, nothing written for that input) any conversion whose `ConversionReport` has a dropped item or a notice, printing the report as it would have been; without it the report is printed and the file written
- [ ] Inputs are converted one at a time in argument order (directory contents in walk order), each model released before the next is opened, and one failure never stops the batch; SPEC-141's first-failure exit code and `--json` records (one per output file set, `files` listing every path written) apply
- [ ] `docu-edit info <input>…` prints, per input, the record under "The info record" — kind, detected format, title, size, document properties and the product's counts — as aligned `key: value` lines, or as JSON records with `--json`; it opens through `openBytes` and writes nothing; an input the funnel refuses is exit 2 with the funnel's message
- [ ] The Node path never touches browser storage: no `indexedDB`, `localStorage` or `sessionStorage` access happens during `convert` or `info` (a test defines throwing getters for all three on `globalThis` and runs every row)
- [ ] Performance: `src/cli/convertPerformance.spec.ts` spawns the launcher on `e2e/fixtures/lists.docx --to odt` and finishes in under 5 seconds, and converts `e2e/fixtures/mini-book.xlsx`, `mini-deck.pptx`, `lists.docx`, `images.docx` and `mini-deck.key` each to every target its product offers, in one process, in under 20 seconds
- [ ] `docs/cli.md` gains `convert` (the conversion table, inputs, output naming, sheet and locale options, `--strict`) and `info`; `docu-edit formats` lists the new rows with no further code
- [ ] e2e (`e2e/SPEC-142.spec.ts`): the Playwright file spawns `docu-edit -convert lists.docx -ext md` (the user's own spelling, against a temp copy), `convert mini-deck.pptx --to odp`, `convert mini-deck.key --to pptx`, `convert mini-book.xlsx --to ods` and `convert mini-book.xlsx --to csv --all-sheets --out-dir <tmp>/csv`, then opens each output through `open-document-input`: the Markdown document shows the fixture's list items, both decks' filmstrips show the source's slide count, the `.ods` shows the fixture's known cell value, and every sheet of the workbook produced one `.csv` that opens

## Out of scope
PDF, images and any conversion that needs layout or a canvas (SPEC-143,
SPEC-144); conversion between products, including `.md` → deck (SPEC-145).
Media inputs: `convert` offers no Node-tier conversion from an image or a PDF,
so they are refused with the "can't convert" message until SPEC-144 adds rows;
`info` reads them. Page counts for text documents in `info` (they need layout).
Raising the funnel's 50 MB limit. Parallel conversion within one process.
Watching a directory. Converting password-protected files. Any change to what
a reader or writer does with a document.

## Reference

### Dependencies restated
- `specs/141-the-command-line-and-creating-files.md` — the grammar and its
  dash rules, `CliIo`, exit codes (1 usage, 2 input, 3 refused, 4 output, 6
  failed), the atomic write, `--force`, `--modified`/`SOURCE_DATE_EPOCH`,
  `CONVERT_FORMATS`/`CONVERSIONS`/`ConversionRow { from; to; tier; spec;
  note? }`, `writeAs(subject, format, { modified?, locale? }): Promise<
  WrittenFile { bytes; fileName; mimeType; report: ConversionReport {
  dropped; notices } }>`, the human report lines and the `--json` record
  (`command`, `input`, `output`, `files`, `from`, `to`, `kind`, `ok`,
  `exitCode`, `bytes`, `dropped`, `notices`, `ms`, `error`),
  `prepareHeadless()` in `src/lib/convert/headless.ts`, and the boundary
  (`src/lib/convert/` imports no `node:` module).
- `src/lib/openDocument.ts` — `openAnyDocumentFile(file: File, options?:
  OpenFileOptions): Promise<OpenedFile>`; `OpenFileOptions extends
  ParseMdOptions { onSpreadsheetProgress? }`; `OpenedFile` =
  `{ kind: "text"; model }` | `{ kind: "presentation"; model; importNotice? }`
  | `{ kind: "spreadsheet"; model }` | `{ kind: "media"; model }`;
  `MAX_FILE_SIZE_BYTES` (50 MB) and `FILE_TOO_LARGE_MESSAGE`; the error
  classes `DocxOpenError`, `OdtOpenError`, `LegacyDocError`, `FlatOdfError`,
  `KeynoteOpenError`, `SpreadsheetOpenError`, `UnsupportedFileError`, whose
  messages are user-facing and printed unchanged. Its CSV/TSV branch reads
  with `readCsvBytes(bytes, format, title, { userLocaleTag })`, today taking
  `userLocaleTag` from `loadPreferences().spreadsheetLocale`.
- `src/lib/documentKinds.ts` — `OPENABLE_EXTENSIONS` (`docx odt md markdown
  pptx odp key xlsx xlsm xltx xltm ods ots csv tsv pdf png jpg jpeg gif svg
  webp`; no `dotx`/`ott`). `src/lib/openDocument.ts` —
  `openTemplateFile(file)` (SPEC-046) reads a `.dotx`/`.ott` and strips its
  template declaration; `openAnyDocumentFile` does not.
- `src/lib/core/fileReader.ts` — `readFileAsArrayBuffer(file)`, today always
  through `new FileReader()`; the funnel's first call. `src/lib/core/
  packageParts.ts` — `bytesToBase64Async` already uses the synchronous path
  when `FileReader`/`Blob` are missing. `src/lib/preferences.ts` —
  `loadPreferences()` already catches storage errors.
- `src/lib/spreadsheet/download.ts` — `spreadsheetDownload(model, target,
  locale?)` takes its same-format branch whenever `target === model.format`,
  and a `.xlsm`/`.xltx`/`.xltm` model has `format: "xlsx"` with a
  `packageVariant` (an `.ots`, `format: "ods"`): that branch writes the
  variant's package unchanged.
- `specs/087-import-keynote.md` — a `.key` opens as a presentation whose
  `importNotice` is a whole sentence. `specs/106-…` —
  `exportDeckHtml(model, options?)` in `src/lib/generate/deckHtml.ts` returns
  one self-contained HTML string.
- `specs/118-…` / `spreadsheet/numberFormat/locale.ts` —
  `SpreadsheetLocaleTag` (`en-US en-GB de-DE fr-FR es-ES cs-CZ`),
  `SPREADSHEET_LOCALE_TAGS`, `SPREADSHEET_LOCALES` (tag → `SpreadsheetLocale`
  object), `DEFAULT_SPREADSHEET_LOCALE`. `specs/126-csv-and-tsv.md` —
  `spreadsheetSheetCopy(model, target: "csv" | "tsv", locale?:
  SpreadsheetLocale)` writes the active sheet (`model.activeSheetIndex`)
  with a locale-dependent delimiter; selecting another sheet is a copy of the
  model with `activeSheetIndex` set, never a mutation of the opened model;
  the not-included notice text.
- `specs/036-word-count-…` — `documentStats(model, { includeNotes })` in
  `src/lib/wordCount.ts`, the Word count dialog's figures. `specs/107-images-and-pdfs-in-the-library.md` —
  `MediaModel` (image parts with pixel size and content type; a PDF's
  `PdfSummary` with its page count and page sizes).
- `specs/008-…` — a same-format text save re-emits retained parts verbatim;
  cross-format drops are `DropRecord`s.
- Equality: `equalBytes` (`src/lib/core/equality.ts`); packages compared part
  by part after `unzip`, text formats whole.

### Inputs
```
docu-edit convert <input>… --to <ext>
                  [-o <file|-> | --out-dir <dir>] [-r] [--stdin-name <name>]
                  [--sheet <name|n> | --all-sheets] [--locale <tag>]
                  [--strict] [--force] [--modified <instant>]
docu-edit info <input>… [-r] [--stdin-name <name>]
```
Expansion happens once, before any conversion: each argument becomes a list
of files (a file → itself; a directory with `-r` → its walk; `-` → one
standard-input entry, allowed once). The expanded list is then checked for
output collisions. A file whose extension is not an accepted input, named directly, is
reported as exit 2 with `UNSUPPORTED_FILE_TYPE_MESSAGE`; found by a walk, it
is skipped silently.

### Output naming
| Situation | Output |
| --- | --- |
| `convert a/report.docx --to md` | `a/report.md` |
| `… --out-dir out` (file argument) | `out/report.md` |
| `convert docs -r --to odt --out-dir out` | `out/<path under docs>/<stem>.odt` |
| `… -o final.md` | `final.md` (one input, one file) |
| `… -o -` | standard output |
| `convert book.xlsx --to csv --all-sheets` | `book - Sales.csv`, `book - Notes.csv` |

`<stem>` is the input's file name without its last extension
(`report.v2.docx` → `report.v2`). A standard-input entry's stem comes from
`--stdin-name`. Outputs never take the writer's suggested `fileName` (which
follows the document title): the command names files after files.

### Sheet selection
| Target | Default | `--sheet <s>` | `--all-sheets` |
| --- | --- | --- | --- |
| `csv`, `tsv` (this spec) | active sheet, one file | that sheet, one file | one file per sheet, hidden included |
| `pdf` (SPEC-143), images (SPEC-144) | active sheet's pages | that sheet's pages | every visible sheet, one job |
| `docx`, `odt`, `md` (SPEC-145) | active sheet as a table | that sheet | every visible sheet |
| `xlsx`, `ods` | the whole workbook | usage error | usage error |

`<s>` matches a sheet name exactly, else case-insensitively, else as a
1-based index when it is all digits.

### The info record
```json
{ "command": "info", "input": "report.docx", "ok": true, "exitCode": 0,
  "kind": "text", "format": "docx", "title": "Quarterly report", "bytes": 18234,
  "properties": { "author": "…", "created": "…", "modified": "…", "subject": "…", "keywords": "…" },
  "text": { "words": 1234, "characters": 7012, "paragraphs": 56, "tables": 2, "images": 3, "comments": 1 },
  "notices": [] }
```
Per kind, exactly one of these product blocks, keyed by the `kind`:
- `text` — `words`, `characters`, `paragraphs` from `documentStats(model, {
  includeNotes: false })`, plus counts of tables, images and comments in the
  model.
- `presentation` — `{ "slides": 12, "hiddenSlides": 1, "slideSize": { "widthCm":
  33.867, "heightCm": 19.05 }, "slidesWithNotes": 4 }`.
- `spreadsheet` — `{ "sheets": [ { "name": "Sales", "hidden": false, "usedRange":
  "A1:E3", "cells": 15 } ], "activeSheet": "Sales" }`; `usedRange` is `null`
  for an empty sheet.
- `media` — an image: `{ "contentType": "image/png", "widthPx": 640,
  "heightPx": 480 }`; a PDF: `{ "pages": 3, "pageSizes": ["21 × 29.7 cm"] }`
  (each distinct size once, in page order, as SPEC-107's card label words it).

`properties` carries only the fields the file has. `format` is the detected
format (an `.xlsx` renamed `.csv` reports `xlsx`, as the funnel opens it).
Human output prints the same fields as `key: value` lines, product block
flattened (`slides: 12`), one blank line between inputs.

### Conversion rows added
| From | To | Note |
| --- | --- | --- |
| text | `docx`, `odt`, `md` | |
| presentation | `pptx`, `odp`, `html` | `html`: self-contained, SPEC-106 |
| workbook | `xlsx`, `ods` | |
| workbook | `csv`, `tsv` | "Sheet selection"; package variants to their base format refused |

### Tests
- `open.spec.ts` — the fixture sweep above; `spreadsheetLocale` overriding the
  preference for a CSV whose numbers read differently in `de-DE`.
- `fileReaderless.spec.ts` — `readFileAsArrayBuffer` gives identical bytes
  with and without `FileReader` (jsdom provides one; the test deletes it for
  the second run); `bytesToBase64Async` and `loadPreferences` under Node.
- `convert.spec.ts` (in-memory `CliIo`) — every row's bytes against `writeAs`
  with a pinned instant (`equalBytes`), the "can't convert" refusal text,
  same-format conversion, `-o`/`--out-dir`/mirroring/standard output, every
  collision and overwrite case, `--strict` with a dropping conversion
  (`lists.docx` → `md` drops a record today), `--sheet` by name and index,
  `--all-sheets` naming with a sheet called `Q1/Q2`, the unknown-sheet
  message, `.dotx` → `.docx` as a document, the package-variant refusal
  (`spec116-macros.xlsm`, `spec116-template.xltx`), `--locale`, directory
  walk order, hidden-name and symlink skipping, stdin with and without
  `--stdin-name`, first-failure exit code across a mixed batch.
- `info.spec.ts` — one record per kind against fixtures with known counts, the
  renamed-`.xlsx` case, human formatting.
- `storageFree.spec.ts`, `convertPerformance.spec.ts`; every existing open,
  download and generate test passes unchanged; `e2e/SPEC-142.spec.ts`.

### Specs this changes
- `specs/141-…` — its registry gains the rows above and `docs/cli.md` two
  sections; nothing else.
- `src/lib/openDocument.ts` — `OpenFileOptions.spreadsheetLocale`, additive.
- `src/lib/core/fileReader.ts` — a second branch, behaviour unchanged where
  `FileReader` exists.
- `specs/000-roadmap.md` — Phase 19 row.
