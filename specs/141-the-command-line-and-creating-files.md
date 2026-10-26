---
id: SPEC-141
title: The docu-edit command, and creating files from the command line
status: draft
area: platform
complexity: normal
story_points: 8
depends_on: [SPEC-045, SPEC-046, SPEC-050, SPEC-080, SPEC-106, SPEC-121, SPEC-126, SPEC-140]
---
## Why
Everything this product can make, it makes inside a browser tab. A person who
wants twenty blank workbooks for a course, a report built from a JSON file a
build pipeline wrote, or a deck from the Markdown notes they keep in a
repository has to open the app, click, download, rename — once per file. The
generation API (SPEC-045, SPEC-106, SPEC-140) already does the work headlessly,
but only for someone willing to write a TypeScript program against it.

This spec gives the product a command: `docu-edit`. It is the first of five
(SPEC-141–145) that together let a shell create and convert every format the
product handles without a window — `docu-edit create report.docx`,
`docu-edit convert example.docx --to md`, `docu-edit convert deck.pptx --to
pdf`. This one builds the command itself and its first verb, **create**; the
others add `convert` and `info` in Node (SPEC-142), PDF through a headless
browser (SPEC-143), images (SPEC-144) and conversion between products
(SPEC-145).

The command adds **no format code of its own**. It is a thin shell over the
writers the app already uses: a file the command creates is byte-for-byte the
file the app would have downloaded from the same model. Everything a later
spec needs is fixed here once — the grammar (including the single-dash
spellings people actually type, `docu-edit -convert example.docx -ext md`),
the exit codes, the output-path and overwrite rules, the report every
conversion prints, and the machine-readable record `--json` emits. **This file
is the normative source for the command grammar and its option conventions,
the exit codes, `CliIo`, the output-path and atomic-write rules, the format
registry (`src/lib/convert/formats.ts`), `writeAs` and `ConversionReport`, the
human report lines and the `--json` record.**

## Acceptance criteria
- [ ] `bin/docu-edit.mjs` is a plain-JavaScript launcher with a `#!/usr/bin/env node` shebang and the executable bit set: it refuses a Node outside the range Vite 8 supports (`^20.19.0 || >=22.12.0`; exit 1, "docu-edit needs Node.js 20.19+ or 22.12+ — this is <version>"), loads `src/cli/main.ts` **once** through Vite's `runnerImport` exactly as `examples/lib/runTs.js` loads an example (no flags, no new dependency), takes both `run` and `nodeCliIo` from that one module graph (a second `runnerImport` would split module state such as core `clock.ts`), awaits `run(process.argv.slice(2), nodeCliIo(repoRoot))`, waits for standard output and error to drain, and then calls `process.exit(code)` — never merely `process.exitCode`, because `core/idleScheduler.ts` holds a `MessageChannel` port open for the life of the process and Node would never exit (the reason `runTs.js` exits explicitly too); it resolves the repository root from its own real path, so it works when symlinked onto `PATH`
- [ ] `package.json` gains exactly one field, `"bin": { "docu-edit": "bin/docu-edit.mjs" }` — no dependency, script or version change — so `npm link` puts `docu-edit` on `PATH`; `README.md` gains a "Command line" section with the `npm link` step and three example commands, linking `docs/cli.md`
- [ ] `src/cli/main.ts` exports `run(argv: string[], io: CliIo): Promise<number>` per the Reference: it never calls `process.exit`, never touches `process`, `fs`, signals or the console directly (everything goes through `io`, including interrupt cleanup through `io.onInterrupt`), and imports each command's module lazily, so `docu-edit --version` and `docu-edit help` load no reader or writer
- [ ] The grammar per "Command grammar": a command word may be bare or carry one or two leading dashes (`create`, `-create`, `--create`); every option with a multi-letter name may be written with one or two dashes and with its value separated by a space or `=` (`-ext md`, `--ext md`, `--to=md`); single-letter options take one dash and are never bundled (`-qf` is an unknown option); `--` ends options; so `docu-edit -create example.docx` is `docu-edit create example.docx` and `docu-edit -convert example.docx -ext md` is `docu-edit convert example.docx --to md` (SPEC-142 registers `convert`, whose `--ext` alias this spec's option table reserves)
- [ ] Mistakes are usage errors (exit 1) with one line naming the problem and, where one exists, a suggestion: an unknown command, option or target extension within an optimal-string-alignment distance of 2 of a known one prints `did you mean "<name>"?` (`docu-edit -covert x.docx` → `docu-edit: unknown command "covert" — did you mean "convert"?`), a missing option value, a repeated option, a value outside its option's set, and two outputs of one invocation whose paths are known from the arguments alone resolving to the same path are each refused before anything is read or written; a usage error that only the input can reveal (an option that does not apply to the product the input turns out to be, a page range past its page count) is reported for that input once it is opened, per "Exit codes"
- [ ] `docu-edit create <output>…` writes a new file per output, its product and format chosen from the output's extension through the format registry (`docx` `odt` `md` → text; `pptx` `odp` → presentation; `xlsx` `ods` `csv` `tsv` → workbook), titled `--title` if given, else the description's `title` (with `--from x.json`), else `titleFromFileName(<output file name>)`; a blank file is `createDocument()` / `createBlankPresentation()` / `createBlankWorkbook()` with that title and format, written through `writeAs`, so `docu-edit create a.docx` and the app's "New document" downloaded as `.docx` with the same title and `modified` are `equalBytes`-identical
- [ ] `--type <ext>` supplies the format for an output with no extension (`create notes --type md` writes `notes.md`) and is a usage error beside an output that has one; an output whose extension the registry cannot create (`create a.pdf`) is a usage error listing the creatable extensions
- [ ] `--template <id>` (text outputs only) builds the output from SPEC-046's `BUILTIN_TEMPLATES` entry with that `id` (`letter`, `memo`, `report`, `invoice`, `meeting-minutes`, `cv`), retitled by the title rule above; an unknown id is a usage error listing the ids, and `--template` beside a presentation or workbook output, or beside `--from`, is a usage error
- [ ] `--from <file>` builds the output from a source per "Create from a source": a `.json` file through the target product's description reader — `docFromDescription` (SPEC-045), `deckFromDescription` (SPEC-106), `workbookFromDescription` (SPEC-140) — whose `DocumentGenerationError` exits 2 printing `<file>: <the error's message>` (path included, unchanged); a `.md` file through SPEC-050's `parseMd` for a text output and SPEC-106's `parseSlideMarkdownWithReport` for a presentation output, whose drop records enter the report; `--from -` reads standard input (JSON when its first non-whitespace character after an optional BOM is `{`, Markdown otherwise); any other source, or a `.md` source for a workbook, is a usage error that names `docu-edit convert` as the way to turn an existing file into another format
- [ ] Output paths per "Output paths and writing": an output that exists is refused (exit 4, `<path> already exists — pass --force to replace it`) unless `--force`; an output whose parent directory does not exist is refused (exit 4); every file is written to a temporary sibling (`.<name>.<random>.docu-edit-tmp`) and renamed over the target only once complete, so an interrupted or failed run never leaves a partial file or damages the file it would have replaced; `-` as the output writes the bytes to standard output and requires `--type`
- [ ] `--modified <ISO 8601 instant>`, or else the `SOURCE_DATE_EPOCH` environment variable (seconds since the epoch, as reproducible builds set it), pins the instant every writer stamps and core `clock.ts`'s `now()` for the whole run (`setClock`), so two runs with the same arguments and the same instant produce `equalBytes`-identical files; neither given, the real time is used; an unparsable value is a usage error
- [ ] `src/lib/convert/formats.ts` exports the format registry per "The format registry": one `ConvertFormat` row per extension (canonical name, aliases such as `jpg`→`jpeg` and `markdown`→`md`, MIME type, the product it opens as, whether `create` can make it) and the `CONVERSIONS` table of offered conversions, each row naming its tier (`"node"` or `"browser"`) and the spec that added it; this spec adds the format rows and no conversion rows, every later spec of the phase adds its own, and a unit test asserts every `CONVERSIONS` row has an implementation reachable from `convert` so the registry can never advertise a conversion the command cannot do
- [ ] `src/lib/convert/writeAs.ts` exports `writeAs(subject: ConvertSubject, format, options?): Promise<WrittenFile>` per "Writing a model": text through `documentDownload` (new in `src/lib/downloadDocument.ts`), presentations through `presentationDownload` (new in `src/lib/presentation/download.ts`), workbooks through SPEC-121/126's existing `spreadsheetDownload`; `downloadDocument`, `downloadPresentation`, `downloadSpreadsheet` and the document editor's `handleDownload` (`src/views/document/useFileActions.ts`) become thin callers of those byte-producing halves, so the app and the command cannot drift — every existing download, export and `src/lib/generate/` test passes unchanged
- [ ] Every created or converted file produces one report per "The report": by default one line `<source> → <output>` (`create` uses `new` as the source) on standard error, followed by `  not kept: 2 × comment, 1 × tracked change` grouping `ConversionReport.dropped` by `kind` in first-seen order and one `  note: <sentence>` line per notice; `--verbose` adds one `  - <kind> (<location>): <reason>` line per dropped item; `--quiet` prints nothing for a success; errors always print `docu-edit: <message>` to standard error
- [ ] `--json` replaces the human lines with one JSON record per input/output on standard output (newline-delimited, the shape in "The --json record"), in argument order, failures included with `ok: false`, `exitCode` and `error`; `--json` with output `-` is a usage error
- [ ] Exit codes per "Exit codes", and with several outputs the command continues past a failed one and exits with the code of the first failure in argument order (0 when none failed)
- [ ] `docu-edit formats` prints the registry — one row per format with what it opens as, whether `create` makes it, and every target `convert` offers from it (marking browser-tier targets) — and, below it, the template ids with their titles; `--json` prints `{ formats: ConvertFormat[], conversions: ConversionRow[], templates: { id; title }[] }`
- [ ] `docu-edit help` (also `-h`, `--help`, and no arguments at all) prints the usage summary; `docu-edit help <command>` and `<command> --help` print that command's options; `docu-edit version` / `--version` prints `docu-edit <package.json version> (Node <version>)`; a unit test asserts every option in "Command grammar" appears in the help text and in `docs/cli.md`
- [ ] Boundaries per "Where the code lives", stated as prohibitions: nothing outside `src/cli/` imports `src/cli/`; `src/lib/convert/` imports no `node:` module, no `vue`, no `.vue` file and nothing under `src/cli/` (it must stay loadable by the app); `src/cli/` itself may import `node:` built-ins, `src/lib/**`, and — lazily, from SPEC-143 on — `vite` and `@playwright/test`; asserted by a new `src/cli/boundary.spec.ts` in the style of `src/lib/core/boundary.spec.ts`, and `src/lib/core/boundary.spec.ts`, `src/lib/presentation/boundary.spec.ts` and `src/lib/spreadsheet/boundary.spec.ts` pass unchanged
- [ ] `docs/cli.md` (new) documents installation (`npm install`, `npm link`), the grammar, every command and option this spec adds, the report, `--json`, exit codes and reproducible output; later specs of the phase extend it
- [ ] Unit tests per "Tests" drive `run()` with an in-memory `CliIo` (no real file system), plus one `src/cli/launcher.spec.ts` that spawns `node bin/docu-edit.mjs --version` and `node bin/docu-edit.mjs create <tmp>/a.xlsx` and checks the exit code, the output and that the created file opens with `readWorkbook` (SPEC-140); each process exits by itself — `--version` within 3 seconds, the `create` within 10 — so a launcher that sets only `process.exitCode` fails the test by timing out
- [ ] e2e (`e2e/SPEC-141.spec.ts`): the Playwright file spawns the launcher in Node to create `blank.docx`, `letter.odt --template letter`, `talk.pptx --from <tmp>/talk.md` (three `---`-separated slides) and `budget.xlsx --from <tmp>/budget.json` (the description under "e2e fixtures"), then opens each through `open-document-input`: the text editor shows the letter's "[Your Name]" line, the deck's filmstrip has three slides and the workbook's `sheet-cell[data-ref="B4"]` reads "$1,650.00"

## Out of scope
`convert` and `info` (SPEC-142); PDF (SPEC-143) and images (SPEC-144);
conversion between products (SPEC-145). A standalone binary or a published npm
package — the command runs from a checkout with its dev dependencies
installed, because the Node readers borrow `DOMParser` from the `jsdom` dev
dependency (SPEC-045's `nodeDom.ts`) and SPEC-143's browser tier uses the
Playwright dev dependency; shipping it alone would move both into
`dependencies`, which needs its own approval. Windows (macOS and Linux are the
supported hosts; nothing deliberately excludes Windows, but nothing tests it).
Shell completion scripts. Creating from a user's own template files (`.dotx`,
`.ott`) or the template store SPEC-046 keeps in the browser. Mail merge from
the command line. Any change to what a writer emits.

## Reference

### Dependencies restated
- `specs/045-document-generation-api.md` — `src/lib/generate/`:
  `docFromDescription(description: unknown): DocumentModel` (validates, throws
  `DocumentGenerationError` whose message names the JSON path),
  `ensureNodeXmlGlobals()` in `nodeDom.ts` (supplies `DOMParser`,
  `XMLSerializer` and `Node` from `jsdom` when absent), and the
  `examples/lib/runTs.js` launcher pattern: `runnerImport(path, { configFile:
  false, logLevel: "error" })` from the `vite` package.
- `specs/046-template-gallery-and-mail-merge.md` — `BUILTIN_TEMPLATES:
  BuiltinTemplate[]` in `src/lib/builtinTemplates.ts`, each `{ id; title;
  preview; build(): DocumentModel }`.
- `specs/050-markdown-…` — `parseMd(data: ArrayBuffer | Uint8Array, options?):
  Promise<ParsedMdDocument>` in `src/lib/md/reader.ts`; `writeMd(model):
  Promise<WriteResult>`.
- Text writers (SPEC-006/007/008): `writeDocx(model, { modified })`,
  `writeOdt(model, { modified })`, `writeMd(model)` each resolve to
  `WriteResult { bytes; dropped: DropRecord[] }`; `DropRecord { kind; location;
  reason }` (`src/lib/opaque.ts`); `docxFileName`, `odtFileName`,
  `mdFileName`; `DOCX_MIME_TYPE`, `ODT_MIMETYPE`, `MD_MIME_TYPE`. The
  editor's `handleDownload` stamps `properties.created` (only when absent) and
  `properties.modified` before writing; `documentDownload` takes that stamping
  over and returns the stamped properties for the editor to mirror.
- `specs/080-create-presentation-and-slide-management.md` —
  `createBlankPresentation(): PresentationModel` (`presentation/slides.ts`),
  `stripRetention(model)`, `unconvertibleFrameCount(model)`,
  `writePptx(model, { modified? })`, `writeOdp(model)`; `downloadPresentation
  (model, targetFormat?)` → `{ droppedObjects }`, which strips retention for a
  cross-format copy. `presentationDownload(model, targetFormat?, options?)` is
  its byte-producing half: `{ bytes; fileName; mimeType; droppedObjects }`.
- `specs/106-presentation-generation-api-and-markdown.md` —
  `deckFromDescription(description: unknown): PresentationModel`;
  `parseSlideMarkdownWithReport(text, options?: { title?; format?; size? }):
  { deck; dropped: DeckDropRecord[] }`; `DeckDropRecord { kind; location;
  reason }`.
- `specs/121-create-workbook-and-sheet-management.md` —
  `createBlankWorkbook(): SpreadsheetModel` (`spreadsheet/sheets.ts`).
  `specs/126-csv-and-tsv.md` / `src/lib/spreadsheet/download.ts` —
  `spreadsheetDownload(model, targetFormat?, locale?): Promise<{ bytes;
  fileName; mimeType; notices: string[] }>`; notices are whole sentences.
- `specs/140-workbook-generation-api.md` — `workbookFromDescription(description):
  SpreadsheetModel`, `readWorkbook(bytes, title?)`, the "Q3 sales"
  description.
- `src/lib/openDocument.ts` — `titleFromFileName(fileName)`; `OpenedFile`, the
  `{ kind: "text" | "presentation" | "spreadsheet" | "media"; model; … }` union
  the open funnel returns. `src/lib/documentKinds.ts` — `DocumentKind`,
  `kindOfFormat`. `src/lib/core/clock.ts` — `now()`, `setClock(instant |
  null)`.
- Equality: `equalBytes` (`src/lib/core/equality.ts`), part by part after
  `unzip` for packages and whole-file for `.md`/`.csv`/`.tsv`.

### Command grammar
```
docu-edit <command> [arguments…] [options]

docu-edit create <output>… [--type <ext>] [--title <text>]
                 [--template <id> | --from <file|->]
                 [--force] [--modified <instant>]
docu-edit formats
docu-edit help [<command>]
docu-edit version
```
Global options, accepted by every command: `--json`, `--quiet` (`-q`),
`--verbose`, `--help` (`-h`), `--version`. Reserved here for SPEC-142 so no
later spec can give them another meaning: `--to` (`-t`; alias `--ext`),
`--output` (`-o`), `--out-dir`, `--force` (`-f`), `--strict`.

| Spelling | Meaning |
| --- | --- |
| `create`, `-create`, `--create` | the command `create` (any command word) |
| `--title Report`, `-title Report`, `--title=Report` | option with a value |
| `-o out.md` | single-letter option; never `-oout.md`, never bundled |
| `--` | everything after it is an argument |
| `-` (as an argument) | standard input (`--from -`) or output (`create - --type md`) |

Precedence is positional-free: options may appear before or after arguments.
A command word must be the first argument that is not an option. With no
arguments at all, `help` runs and the exit code is 1 (it is still a usage
error to give no command).

Suggestions use the optimal string alignment distance (Damerau–Levenshtein
without repeated transpositions), case-insensitively, over the command names,
the option names of that command, and the registry's extensions and aliases;
the single nearest name within distance 2 is suggested, and none is suggested
on a tie.

### `CliIo`
```ts
export type PathKind = "file" | "directory" | "symlink" | "other";
export interface CliIo {
  cwd: string;
  repoRoot: string;                                   // the checkout the launcher lives in
  env: Readonly<Record<string, string | undefined>>;
  readFile(path: string): Promise<Uint8Array>;        // rejects with { code: "ENOENT" | "EISDIR" | "EACCES" | … }
  stat(path: string): Promise<{ kind: PathKind } | null>;   // follows symbolic links
  lstat(path: string): Promise<{ kind: PathKind } | null>;  // does not
  readdir(path: string): Promise<string[]>;           // entry names, unsorted
  mkdirp(path: string): Promise<void>;
  remove(path: string): Promise<void>;                // a file; no error when absent
  writeFileAtomic(path: string, bytes: Uint8Array): Promise<void>;  // temp sibling + rename
  readStdin(): Promise<Uint8Array>;
  writeStdout(bytes: Uint8Array | string): void;
  writeStderr(text: string): void;
  isStdoutTty: boolean;
  onInterrupt(handler: () => Promise<void> | void): () => void;   // returns an unregister function
}
```
`nodeCliIo(repoRoot)` (in `src/cli/nodeIo.ts`, re-exported by `main.ts` for
the launcher) implements it over `node:fs/promises` and `process` signals;
`memoryCliIo(files)` (in `src/cli/memoryIo.ts`, test support) over a `Map`,
with `interrupt()` to fire the handlers. This spec uses `readFile`, `stat`,
`writeFileAtomic`, `remove` and the streams; the rest exist so SPEC-142/143
add no members.
Relative paths resolve against `cwd`. The report never prints an absolute path
the user did not type: each path is printed as given on the command line (or,
for a derived output, joined from that).

### Exit codes
| Code | Meaning |
| --- | --- |
| 0 | Every output written. |
| 1 | Usage — unknown command/option/extension, missing or repeated option, bad value, outputs that collide by their argument-derived paths. Found from the arguments alone: nothing read or written. Found only once an input is open (an option that does not apply to its product, a range past its page count): that input fails and the batch continues. |
| 2 | Input — a file or `--from` source that is missing, unreadable, unrecognised, password-protected, too large or invalid (a description's validation error). |
| 3 | Refused — the target format cannot represent this input, or `--strict` and the report is not empty. Nothing written for that input. |
| 4 | Output — the target exists without `--force`, its directory is missing, a name derived from the content (a sheet name, a page number) collides with a file this run already wrote, or the write failed. |
| 5 | The headless browser is unavailable (SPEC-143). |
| 6 | The conversion failed — an exception inside a reader, writer or renderer, or a timeout. The message is the error's own. |

### Output paths and writing
`create`'s argument is the output path. SPEC-142 derives outputs from inputs.
In every case: the parent directory must exist (only SPEC-142's `--out-dir`
creates directories); an existing target needs `--force`; an output whose
path follows from the arguments alone and equals another such output of the
same invocation is a usage error found before any work starts; a name that
can only be known from the content (SPEC-142's per-sheet names, SPEC-144's
numbered pages) is checked when it is written, and colliding with a file the
same run already wrote is exit 4 for that input — `--force` never lets one
output of a run replace another. A temporary sibling is created in the target's own
directory (so the final rename never crosses file systems) and removed on any
failure; on SIGINT/SIGTERM, `nodeCliIo` runs the handlers registered through
`io.onInterrupt` (which remove every temporary file the run created and, from
SPEC-143 on, close the headless browser) and exits 130.

### Create from a source
| Output product | `--from x.json` | `--from x.md` |
| --- | --- | --- |
| text (`docx` `odt` `md`) | `docFromDescription` | `parseMd` (the model `convert x.md` would open) |
| presentation (`pptx` `odp`) | `deckFromDescription` | `parseSlideMarkdownWithReport(text, { title })` |
| workbook (`xlsx` `ods` `csv` `tsv`) | `workbookFromDescription` | usage error |

The model is then retitled by the title rule and written with `writeAs` in the
output's format — a description's own `format` field never overrides the
output's extension (a `"format": "odp"` description written to `talk.pptx` is
a `.pptx`). A JSON source that does not parse is exit 2 with
`<file>: not valid JSON — <the parser's message>`.

### The format registry
```ts
export type ConvertTier = "node" | "browser";
export interface ConvertFormat {
  ext: string;                  // canonical, lower-case, no dot: "docx", "jpeg"
  aliases: readonly string[];   // "jpg", "markdown", "htm"
  mimeType: string;
  opensAs: DocumentKind | null; // what openAnyDocumentFile makes of it; null = output-only (html)
  create: boolean;              // docu-edit create can make it
}
export interface ConversionRow {
  from: DocumentKind;           // the opened product
  to: string;                   // canonical target ext
  tier: ConvertTier;
  spec: string;                 // "SPEC-142"
  note?: string;                // shown by `formats`, e.g. "active sheet; --sheet/--all-sheets"
}
export const CONVERT_FORMATS: readonly ConvertFormat[];
export const CONVERSIONS: readonly ConversionRow[];   // empty after this spec
export function formatByName(name: string): ConvertFormat | undefined; // ext or alias, case-insensitive, leading dot ignored
```
Rows this spec adds: `docx`, `dotx`, `odt`, `ott`, `md` (alias `markdown`) —
text; `pptx`, `odp`, `key` — presentation; `xlsx`, `xlsm`, `xltx`, `xltm`,
`ods`, `ots`, `csv`, `tsv` — workbook (`OPENABLE_EXTENSIONS`' list); `pdf`, `png`, `jpeg` (alias `jpg`), `gif`, `webp`, `svg` —
media; `html` (alias `htm`, `opensAs: null`). `create: true` exactly for
`docx odt md pptx odp xlsx ods csv tsv`. MIME types are the constants the
writers already export.

### Writing a model
```ts
export type ConvertSubject =
  | { kind: "text"; model: DocumentModel }
  | { kind: "presentation"; model: PresentationModel }
  | { kind: "spreadsheet"; model: SpreadsheetModel };
export interface WriteAsOptions { modified?: Date; locale?: SpreadsheetLocale }
export interface ConversionReport { dropped: { kind: string; location: string; reason: string }[]; notices: string[] }
export interface WrittenFile { bytes: Uint8Array; fileName: string; mimeType: string; report: ConversionReport }
```
`writeAs` dispatches on `subject.kind` and the target's product; a target of
another product is SPEC-145's and throws here. Text: `documentDownload(model,
format, { modified })` → `WriteResult` dropped records become
`report.dropped`. Presentation: `presentationDownload` → a non-zero
`droppedObjects` becomes one notice, `droppedObjectsNotice(count)` — a new
export of `src/lib/presentation/download.ts` returning the sentence
`useDeckActions.ts`'s `downloadCopy` builds inline today (`"<n> object(s)
couldn't be converted"`), which `downloadCopy` then calls instead. Workbook: `spreadsheetDownload` → `notices`
verbatim. `WrittenFile.fileName` is the writer's own naming (the app's
download name); the command decides the actual path itself and uses
`fileName` only for standard-output headers in `--json`.

### The report
```
new → report.docx
talk.md → talk.pptx
  not kept: 2 × unsupported block
  note: …
```
`--verbose` also prints, per run, the elapsed time per output
(`  took 412 ms`). Counts group by `kind` exactly as the record carries it;
the line lists at most eight kinds and then `, and N more` .

### The --json record
```json
{ "command": "create", "input": null, "output": "report.docx", "files": ["report.docx"],
  "from": null, "to": "docx", "kind": "text", "ok": true, "exitCode": 0,
  "bytes": 8123, "dropped": [], "notices": [], "ms": 87 }
```
`input` is the source path (`--from` for `create`, `null` for a blank),
`from` its canonical extension, `files` every path written (later specs write
several per input), `bytes` their total. A failure carries `ok: false`,
`exitCode`, `error` (the message printed without `--json`) and `files: []`.

### Where the code lives
```
bin/docu-edit.mjs            launcher (plain JS)
src/cli/main.ts              run(argv, io) — parse, dispatch, report, exit code
src/cli/args.ts              the grammar; parseArgs(argv, commandTable)
src/cli/suggest.ts           OSA distance + did-you-mean
src/cli/report.ts            human lines + JSON records
src/cli/nodeIo.ts            nodeCliIo()
src/cli/memoryIo.ts          memoryCliIo() (tests)
src/cli/commands/create.ts   create
src/cli/commands/formats.ts  formats
src/cli/help.ts              help text
src/lib/convert/formats.ts   the registry
src/lib/convert/writeAs.ts   writeAs
src/lib/convert/headless.ts  prepareHeadless(): Promise<void> — awaits ensureNodeXmlGlobals()
```
`src/lib/convert/` is a **composition point**, like `src/lib/generate/`: it
imports the text product, `lib/presentation`, `lib/spreadsheet` and
`lib/generate`, which is why `core/boundary.spec.ts`'s product blocks do not
list it. Unit tests under `src/cli/` declare `// @vitest-environment node`.

### e2e fixtures
Written by the test into a temporary directory, not committed:
```json
{ "title": "Budget", "sheets": [ { "name": "Budget", "cells": {
  "A1": "Item", "B1": "Cost", "A2": "Rent", "B2": 1200, "A3": "Food", "B3": 450,
  "A4": "Total", "B4": { "formula": "SUM(B2:B3)", "numberFormat": "$#,##0.00" } } } ] }
```
`talk.md` is three slides: `# One`, `---`, `# Two` with two `-` bullets,
`---`, `# Three`.

### Tests
- `args.spec.ts` — every spelling in the grammar table, `--` handling, every
  usage error with its exact message, suggestions (including `covert` →
  `convert`, ties suggesting nothing).
- `create.spec.ts` — each creatable format blank, the title rule's three
  sources, `--type`, every template id, `--from` JSON (all three products, a
  validation error's path text, invalid JSON) and Markdown (text and deck,
  slide drops in the report), stdin detection, `--force`, missing directory,
  `-` output, `--modified` and `SOURCE_DATE_EPOCH` byte determinism
  (`equalBytes` over two runs), the first-failure exit code with several
  outputs.
- `writeAs.spec.ts` — `writeAs` equals the bytes of `documentDownload`,
  `presentationDownload` and `spreadsheetDownload` for the same model and
  `modified`, report mapping for each product.
- `report.spec.ts` — human lines (grouping, eight-kind cap, `--verbose`,
  `--quiet`) and JSON records for success and each failure code.
- `formats.spec.ts` — registry invariants (unique extensions and aliases, MIME
  types match the writers' constants, every `CONVERSIONS` row resolvable).
- `boundary.spec.ts`, `launcher.spec.ts`, `help.spec.ts`; every existing
  download, export, `src/lib/generate/` and `useFileActions` test passes
  unchanged; `e2e/SPEC-141.spec.ts`.

### Specs this changes
- `specs/045-…`, `specs/106-…`, `specs/140-…` — their description readers gain
  a caller; nothing about them changes.
- `specs/080-…` / `src/lib/presentation/download.ts`, `src/lib/downloadDocument.ts`
  and `src/views/document/useFileActions.ts` — split into byte-producing
  halves and delivery, behaviour unchanged.
- `package.json` gains the `bin` field. It is a protected path in the
  factory, so this spec's pull request stops for a human merge by design.
- `specs/000-roadmap.md` gains Phase 19.
