---
id: SPEC-143
title: The headless render host, and PDF from documents, decks and workbooks
status: draft
area: platform
complexity: complex
story_points: 13
depends_on: [SPEC-142, SPEC-034, SPEC-104, SPEC-107, SPEC-136]
---
## Why
`docu-edit convert report.docx --to pdf` is the conversion people want most
and the one Node cannot do. A PDF of a text document is a picture of its
*layout* — where every line breaks, where every page ends — and this product
computes layout the only place it can: in a browser, measuring text in the
editor's own paginated DOM. SPEC-034's export extracts each realised `.page`
off the live DOM; SPEC-104's deck export mounts each sheet page, extracts it
and unmounts it; SPEC-136's sheet export does the same. Node has no layout
engine and no canvas, and `package.json` will not take one.

But the repository already ships a headless browser. Playwright is pinned at
exactly 1.63.0 for the e2e suite; its Chromium is installed on the
development Mac and baked into the factory's lane image. So the render tier
does the honest thing: the command starts the app itself, headless, and asks
it for the PDF through the **same code path** the File menu uses. A PDF the
command writes is the PDF the app would have downloaded on the same machine,
byte for byte — not a reimplementation that drifts from it.

That needs three pieces this spec builds. A **render host** in the command:
Vite's dev server started in-process on a private port plus one Chromium,
started only when a browser-tier conversion is asked for. An **automation
entry**: a second HTML page that boots the identical app and exposes a small,
versioned bridge. And **page producers**: each product's export composable
split so that "produce the pages" is callable without its dialog, registered
while its view is mounted. SPEC-144 builds images on exactly these pieces.
**This file is the normative source for the render host, the automation
entry and bridge, the transfer routes, `PageJob` and the page-producer
registry, browser-tier isolation and timeouts, and the PDF options of
`convert`.**

## Acceptance criteria
- [ ] Tier selection: a `CONVERSIONS` row with tier `"browser"` is converted through the render host; the host starts lazily at the first such input of a run and is reused by every later one; a run with only Node-tier inputs never imports Vite's server API or Playwright (asserted by replacing both dynamic imports with throwing stubs and running SPEC-142's conversions)
- [ ] `src/cli/renderHost.ts` exports `startRenderHost(options): Promise<RenderHost>` per "The render host": Vite's `createServer` in-process from the repository's own `vite.config.ts` on `127.0.0.1` with port 0, no HMR, no file watching and its own `cacheDir` (`node_modules/.vite-docu-edit-cli`), so it never collides with a dev server on 5173 or its dependency cache; Chromium launched headless through `@playwright/test`'s `chromium`; one server and one browser per process, closed when the run ends, on any uncaught error and on SIGINT/SIGTERM (exit 130), leaving no process behind
- [ ] An unavailable browser is exit 5 with one of three exact messages per "Host failures" — Chromium not installed (naming `npx playwright install chromium`), Playwright not installed (naming `npm install`), or any other launch or server failure (with the first line of its error) — and every input of the run that needed the host reports the same code without retrying the launch
- [ ] `automation.html` (beside `index.html`) loads `src/automation/main.ts`, which boots the app through a new `src/appSetup.ts` `startApp()` — the body of today's `src/main.ts` (router, `initLibrary()`, `initTemplates()`, mount), which `src/main.ts` now calls — and then installs the bridge; `vite build`'s inputs are unchanged, so the production bundle carries no bridge; a boundary test asserts nothing outside `src/automation/` imports `src/automation/bridge.ts` or `src/automation/main.ts`
- [ ] `window.__docuEditAutomation` implements `AutomationBridge` (version 1) per "The automation bridge": `ready`, `setClock(iso | null)` (core `clock.ts`), `open(input)` — fetch the input route, `openAnyDocumentFile` + `storeOpenedFile` into this context's own library (the id is `opened.model.id`; `storeOpenedFile` returns nothing), navigate to `editorPathFor(kind, id)`, wait for the view's page producer — and `render(request)`, which asks the producer for a `PageJob`, builds the output and posts it to the output route; every failure rejects with an `Error` whose message the command prints
- [ ] Bytes cross the boundary only through two routes per "Transfer": `GET /__docu-edit-cli/input/<token>` fulfilled by the host from the file it read, and `POST /__docu-edit-cli/output/<token>/<n>` captured by the host and answered 204 — never as base64 through `page.evaluate`, so a 50 MB input costs one copy; tokens are fresh `randomUUID()`s valid for one input, and any other request under `/__docu-edit-cli/` is answered 404
- [ ] Isolation per "Contexts": every input runs in a new browser context, closed after it, so its IndexedDB and storage never outlive it and the command never reads or writes the library a person keeps in their own browser; the context aborts every request whose origin is not the host's own, blocks service workers, and pins viewport 1280 × 720, device scale 1, locale `en-US`, light colour scheme and the host machine's time zone
- [ ] `src/automation/registry.ts` exports `PageJob`, `PageJobRequest`, `PageProducer`, `registerPageProducer(producer): () => void` and `waitForPageProducer(documentId, timeoutMs)` per "Page jobs"; it is framework-free and inert when no bridge reads it
- [ ] Each export composable gains `preparePageJob(request)` and its own menu export becomes `buildPdf(await preparePageJob(<the dialog's choices>))` + download: `usePdfExport` (text, SPEC-034), `useDeckExport` (SPEC-104 — with the request's layout, paper, orientation, range, note lines, frame slides and hidden-slide choices in place of the dialog's, never reading or writing its module-level `session` choices), `useSheetPdfExport` (SPEC-136 — with the request's sheet selection in place of `planFor`'s preview/active-sheet choice); each composable registers its producer on mount and unregisters on unmount itself, so `DocumentView.vue` gains no lines (its size budget is unchanged) and every existing export test passes unchanged
- [ ] The deck's export sheet can render without its dialog: `DeckPrintSheet`, today mounted only inside `DeckPrintDialog` — which `PresentationDialogs.vue` renders under `v-if="deckExport?.isOpen.value"`, so without the dialog `runSheet` throws "the export sheet did not render" — moves to an always-mounted place in the presentation view that renders nothing while no pages are planned; the dialog's own export and print keep working through it unchanged (the text and sheet export hosts are already always mounted)
- [ ] `convert` gains three browser-tier rows — text → `pdf`, presentation → `pdf`, workbook → `pdf` — with the options under "PDF options", each validated against its product (`--layout` for a text input is exit 1 for that input once it is opened — SPEC-141's per-input usage error — naming the products the option applies to)
- [ ] Text `--pages <range>` takes SPEC-034's range syntax through `parseNumberRange(input, total, "page")` (its error text is the message, exit 1): the PDF holds only those pages, in document order; an internal link whose destination page is excluded is removed and the others are renumbered; an outline heading on an excluded page is removed
- [ ] Deck PDFs default to `DEFAULT_DECK_PRINT_OPTIONS` with hidden slides excluded (the dialog's defaults); `--slides <range>` is SPEC-104's custom range, its errors exit 1 for that input; like the dialog's PDF export, there is no page limit (`PRINT_PAGE_LIMIT` guards printing only)
- [ ] Workbook PDFs follow SPEC-142's "Sheet selection": the active sheet by default — the job the workbook's direct PDF export builds (`jobFor("sheet")`) — another with `--sheet`, every visible sheet with `--all-sheets` (`scope: "workbook"`); each sheet's own page setup, print area and headers/footers apply as in the app; `PDF_LIMIT_MESSAGE` and `NOTHING_TO_PRINT_MESSAGE` refuse (exit 3) with those exact texts
- [ ] Same pipeline, same bytes: `e2e/SPEC-143.spec.ts` exports one file per product from the app with the page clock pinned — the text document through File → Export → PDF, the deck through the export dialog's PDF tab left at its defaults, the workbook through its direct PDF export (active sheet) — and converts the same file with the command at the same `--modified` instant and default options; the two PDFs are `equalBytes`; two command runs with the same `--modified` are `equalBytes`
- [ ] `--timeout <seconds>` (default 120, at least 1) bounds each input from open to the last output byte; past it the input fails with exit 6 `timed out after <n> s while <phase>` (`opening`, `laying out`, `rendering page <i> of <n>`, `writing`), its context is closed and the batch continues; an exception inside the page is exit 6 with its message, and with `--verbose` the page's console errors for that input follow it
- [ ] `--verbose` adds per-input timings (`host 2.9 s · open 0.6 s · render 3.1 s · write 0.1 s`), the host's start time appearing only on the input that started it
- [ ] `docs/cli.md` gains "PDF output and the headless browser": the one-time `npx playwright install chromium`, the PDF options per product, isolation (no network, no library), and the honest limit — pagination follows the fonts installed on the machine that runs the command, exactly as the editor on that machine paginates, so the same file can break pages differently on another machine
- [ ] e2e (`e2e/SPEC-143.spec.ts`, each test given 180 s): the equal-bytes checks above for `multipage.docx` (built by the test in Node with SPEC-045's `DocumentBuilder` — six Heading 1 sections, each a few paragraphs, separated by `addPageBreak()`; the committed fixtures are short, and `spec-047-fixture.docx` is generated on demand and never committed), `mini-deck.pptx` and `mini-book.xlsx`; `convert mini-deck.pptx --to pdf --layout handout-3` yields a PDF whose page count is `ceil(slides / 3)` by SPEC-107's structural reader; `convert multipage.docx --to pdf --pages 2-3` yields a two-page PDF by the same reader, and `--pages 0` is exit 1 with `parseNumberRange`'s own message; a run with `PLAYWRIGHT_BROWSERS_PATH` pointed at an empty directory exits 5 with the install message

## Out of scope
Images (SPEC-144) and PDF from images or between products (SPEC-144,
SPEC-145). Serving the host to anything but the command's own browser, a
long-running conversion server, or a host shared between processes. Using a
browser other than Playwright's Chromium, or an already-running browser.
Making pagination independent of the machine's fonts. Printing to a printer.
Reading or writing a person's real browser library. A `vite build` of the
automation entry or running the render tier from `dist/`. PDF/A, tagged PDF,
encryption or any change to what `buildPdf` writes. Page counts in `info`.

## Reference

### Dependencies restated
- `specs/141-…`/`specs/142-…` — the grammar, `CliIo`, exit codes (1 usage, 3
  refused, 5 browser unavailable, 6 failed), the report, the `--json` record,
  `--modified`/`SOURCE_DATE_EPOCH` and `setClock`, the registry
  (`ConversionRow.tier`), input expansion and output naming (`<stem>.pdf`),
  `openBytes`.
- `specs/034-print-and-export-to-pdf.md` — `buildPdf(input: PdfBuildInput):
  Promise<Uint8Array>` in `src/lib/core/pdfBuild.ts` with `PdfBuildInput {
  title; info?; pages: PdfPageInput[]; fonts: Map<FaceKey, SfntFont>;
  images?: Map<string, PdfImageData> }`; `PdfPageInput` carries positioned
  text items, links (`destPage` 0-based for internal links), headings, images
  and rects; `extractPageContent(element, widthPt, heightPt, options?)`
  (`core/pdfExtract.ts`); `loadPdfFonts(pages)` / `loadPdfImages(pages)`
  (`src/lib/pdfAssets.ts`); `usePdfExport` realises every page
  (`realizeAllPages`) and extracts each `.page`. `parseNumberRange(input,
  total, noun)` in `core/pageRange.ts` → `NumberRangeResult`.
- `specs/104-export-deck-to-pdf-and-images.md` — `useDeckExport` runs a
  rolling mount → settle → extract → unmount over a plan of `SheetPage`s;
  `DeckPrintOptions` (`layout: DeckPrintLayout` = `"slides" | "notes" |
  "handout-1" | "handout-2" | "handout-3" | "handout-4" | "handout-6" |
  "handout-9"`, `paper: DeckPrintPaper` = `PaperSizeName` minus "Custom",
  `orientation`, `range: "all" | "current" | "custom"`, `rangeText`,
  `includeHidden`, `noteLines`), `DEFAULT_DECK_PRINT_OPTIONS`,
  `defaultFrameSlides(layout)`, `pageHeadings`, `sheetPageRects`; the
  dialog's choices live in a module-level `session` object; the mounted
  pages come from `DeckPrintSheet` (`components/presentation/`), which today
  exists only inside `DeckPrintDialog.vue`. `PRINT_PAGE_LIMIT` and
  `printRefusal` guard Print only — PDF export has no page limit.
- `specs/136-sheet-page-setup-print-and-pdf.md` — `PrintScope = "sheet" |
  "selection" | "workbook"`, `paginateWorkbook(model, scope, options)`,
  `PDF_LIMIT_MESSAGE`, `NOTHING_TO_PRINT_MESSAGE`; `useSheetPdfExport`
  extracts each mounted print page, its `planFor` taking the print preview's
  chosen pages when the preview is open and otherwise
  `deps.print.jobFor("sheet")` — the active sheet — which is what the
  workbook's direct PDF export writes.
- `src/lib/openDocument.ts` — `openAnyDocumentFile`; `src/lib/documentStore.ts`
  — `storeOpenedFile` (the funnel's store step, returning the stored id);
  `src/lib/documentKinds.ts` —
  `editorPathFor(kind, id)` (`/doc/:id`, `/slides/:id`, `/sheets/:id`,
  `/media/:id`). `src/main.ts` today: `createRouter({ history:
  createWebHistory(), routes })`, `Promise.all([initLibrary(),
  initTemplates()])`, then `createApp(App).use(router).mount("#app")`.
- Playwright 1.63.0 (`@playwright/test`, a dev dependency pinned exactly; see
  CLAUDE.md for why its version moves only with the lane image's).
  `chromium.launch({ headless: true })`, `browser.newContext(options)`,
  `context.route(pattern, handler)`, `route.fulfill`, `route.abort`,
  `request.postDataBuffer()`. Vite: `createServer(inlineConfig)`,
  `server.listen()`, `server.resolvedUrls`, `server.close()`.
- Equality: `equalBytes` (`src/lib/core/equality.ts`) on whole PDF files.
  Page counts in e2e come from SPEC-107's structural PDF reader
  (`core/pdfDocument.ts`).

### The render host
```ts
export interface RenderHostOptions { timeoutMs: number; verbose: boolean; io: CliIo }   // repo root: io.repoRoot
export interface RenderHost {
  origin: string;                                   // "http://127.0.0.1:<port>"
  convert(input: { bytes: Uint8Array; fileName: string }, request: RenderRequest,
          modified: Date | null): Promise<RenderedFile[]>;
  close(): Promise<void>;
}
export interface RenderedFile { bytes: Uint8Array; mimeType: string; suffix: string; report: ConversionReport }
```
Start order: (1) `await import("vite")` → `createServer({ configFile:
"<repo>/vite.config.ts", root: io.repoRoot, cacheDir:
"node_modules/.vite-docu-edit-cli", logLevel: "error", clearScreen: false,
server: { host: "127.0.0.1", port: 0, strictPort: false, hmr: false, watch:
null } })`, `listen()`; (2) `await import("@playwright/test")` →
`chromium.launch({ headless: true })`. The host registers `io.onInterrupt(close)` when it starts. `convert` opens a
context, routes,
`goto(origin + "/automation.html")`, awaits `ready`, calls `setClock`,
`open`, `render`, collects the posted outputs and closes the context.
`suffix` is `""` for a single-file output (SPEC-144 numbers pages).

### Host failures
| Cause | Message (after `docu-edit: `) |
| --- | --- |
| Playwright's browser executable missing | `PDF and image output need Playwright's Chromium — run "npx playwright install chromium" once` |
| `@playwright/test` cannot be imported | `PDF and image output need the project's dev dependencies — run "npm install" in <repo>` |
| any other launch or server error | `couldn't start the headless browser — <first line of the error>` |

### The automation bridge
```ts
declare global { interface Window { __docuEditAutomation?: AutomationBridge } }
export interface AutomationBridge {
  readonly version: 1;
  readonly ready: Promise<void>;                     // app mounted, library hydrated
  setClock(iso: string | null): void;
  open(input: { url: string; fileName: string }): Promise<{ kind: DocumentKind; id: string; title: string }>;
  render(request: RenderRequest & { outputUrl: string }): Promise<{ files: number; notices: string[] }>;
}
export type RenderRequest =
  | { target: "pdf"; kind: "text"; pages?: string }
  | { target: "pdf"; kind: "presentation"; layout?: DeckPrintLayout; paper?: DeckPrintPaper;
      orientation?: Orientation; slides?: string; includeHidden?: boolean; noteLines?: boolean; frameSlides?: boolean }
  | { target: "pdf"; kind: "spreadsheet"; sheet?: string; allSheets?: boolean };
```
`open` waits for the producer with the remaining timeout. `render` builds
`buildPdf(job)` from the producer's job and POSTs it to `outputUrl + "/0"`.
SPEC-144 widens `RenderRequest` with image targets; the version stays 1 while
changes are additive.

### Transfer
| Request | Who answers | Body |
| --- | --- | --- |
| `GET /__docu-edit-cli/input/<token>` | host route | the input bytes, `application/octet-stream` |
| `POST /__docu-edit-cli/output/<token>/<n>` | host route | one output file; header `x-docu-edit-mime` its type; answered 204 |
| anything else under `/__docu-edit-cli/` | host route | 404 |

### Contexts
`browser.newContext({ viewport: { width: 1280, height: 720 },
deviceScaleFactor: 1, locale: "en-US", colorScheme: "light", timezoneId:
<Intl.DateTimeFormat().resolvedOptions().timeZone>, serviceWorkers: "block",
acceptDownloads: false })`, then `context.route("**/*")` continuing only
requests whose origin equals the host's and aborting the rest. The viewport
equals Playwright's default so the e2e comparison exports under identical
conditions.

### Page jobs
```ts
export interface PageJob extends PdfBuildInput { notices: string[] }
export type PageJobRequest = Omit<RenderRequest, "target">;
export interface PageProducer {
  kind: DocumentKind;
  documentId: string;
  prepare(request: PageJobRequest, onProgress?: (done: number, total: number) => void): Promise<PageJob>;
}
export function registerPageProducer(producer: PageProducer): () => void;
export function waitForPageProducer(documentId: string, timeoutMs: number): Promise<PageProducer>;
```
A producer's `prepare` is its composable's `preparePageJob`: the same plan,
mount/settle/extract loop, fonts and images its menu export uses, with the
request's choices in place of the dialog's. It never downloads and never
touches the dialog's session choices. `onProgress` feeds the timeout's phase
text.

### PDF options
| Option | Products | Meaning |
| --- | --- | --- |
| `--pages <range>` | text | pages to keep, SPEC-034 range syntax |
| `--layout <name>` | presentation | a `DeckPrintLayout` |
| `--paper <name>` | presentation | a `DeckPrintPaper` (`A4`, `Letter`, …; case-insensitive) |
| `--orientation portrait\|landscape` | presentation | sheet orientation |
| `--slides <range>` | presentation | SPEC-104's custom range |
| `--include-hidden` | presentation | print hidden slides too |
| `--note-lines`, `--frame-slides` | presentation | SPEC-104's layout extras where the layout allows them |
| `--sheet <name\|n>`, `--all-sheets` | workbook | SPEC-142's "Sheet selection" (default: the active sheet) |
| `--timeout <seconds>` | all browser-tier | per-input limit, default 120 |

### Tests
- `src/automation/registry.spec.ts` — register/unregister, waiting, timeout.
- `src/automation/bridge.spec.ts` (jsdom, stub producer and stub `fetch`) —
  `open` → store → navigation → producer; `render` posting; error messages.
- `src/cli/renderHost.spec.ts` (stubbed `vite` and Playwright modules) — the
  exact `createServer`/`launch`/`newContext` options, route behaviour
  (including the 404 and the non-host abort), the three failure messages,
  close on error and signal, lazy start and reuse across inputs.
- `src/cli/pdfOptions.spec.ts` — every option's validation per product.
- `src/components/presentation/` — the deck export dialog's existing tests
  pass with the sheet host moved; a new test renders a producer job with the
  dialog closed.
- Each composable's existing spec, plus one per composable asserting its menu
  export and its `preparePageJob` give the same `PdfBuildInput` for the same
  choices.
- `src/automation/boundary.spec.ts`; `DocumentView.size.spec.ts` unchanged;
  `e2e/SPEC-143.spec.ts`.

### Specs this changes
- `specs/034-…`, `specs/104-…`, `specs/136-…` — each export composable is
  split into `preparePageJob` + delivery and registers a producer, and
  SPEC-104's `DeckPrintSheet` moves out of the dialog to an always-mounted
  host; nothing a user sees changes.
- `src/main.ts` becomes a caller of `src/appSetup.ts`'s `startApp()`.
- `specs/141-…` — exit code 5 gets its implementation; the registry gains
  three rows; `docs/cli.md` a section.
- `specs/000-roadmap.md` — Phase 19 row.
