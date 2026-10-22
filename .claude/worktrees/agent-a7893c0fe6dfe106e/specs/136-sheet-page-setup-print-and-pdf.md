---
id: SPEC-136
title: Page setup, printing and PDF export for sheets
status: draft
area: layout
complexity: complex
story_points: 13
depends_on: [SPEC-034, SPEC-104, SPEC-118, SPEC-123, SPEC-124, SPEC-131, SPEC-132, SPEC-133, SPEC-135]
---
## Why
A grid has no pages. The budget still has to go to the board as four sheets of
A4, landscape, the header row repeated on each, "Page 2 of 4" in the footer —
and every workbook that arrives already says so: `<pageSetup>`,
`<pageMargins>`, `<headerFooter>`, `<rowBreaks>` and the `_xlnm.Print_Area`
name in an `.xlsx`; a master page, a page layout, `table:print-ranges` and
`table:table-header-rows` in an `.ods`. Today those bytes ride verbatim in the
retained worksheet XML and `styles.xml`, nothing reads them, and the only way to
get paper out of a workbook is the browser printing the scrolling grid — headers,
tabs and all, cut wherever the screen ends.

The machinery is half here. SPEC-034 writes PDF 1.7 by hand from realised DOM
pages (real text over embedded font subsets, `PdfPathItem` vectors for table
cell fills and per-edge borders, canvases as images) and prints through
`window.print()` under an injected `@page` rule; SPEC-104 moves that layer into
`src/lib/core` so a second product can reach it; `core/pageSetup.ts` holds the
paper catalogue and margin validation; SPEC-118 formats a value for a given
width and SPEC-123 resolves a cell's full style. What is missing is a page model
for a sheet, the algorithm that cuts a grid into pages, and the surfaces in front
of them.

**This file is the normative source for the sheet page-setup model and both of
its wire mappings, the header/footer code language, the pagination algorithm,
the mapping of the reserved print-area and print-titles names onto ODS, and the
sheet print/PDF pipeline.**

## Acceptance criteria
- [ ] `Sheet` gains `pageSetup?: SheetPageSetup` (replaced, never mutated, by every command, so it joins SPEC-121's `SHEET_BY_REFERENCE_FIELDS`) and `pageSetupDirty?: true` per "The page-setup model": orientation, paper (portrait dimensions in cm plus the xlsx `paperSize` code), scale or fit-to W × H pages, six margins (incl. header and footer), centring, print gridlines, print headings, page order, first page number, header/footer with odd/even/first variants, and manual `rowBreaks`/`colBreaks`; plain JSON, absent fields meaning the defaults the Reference pins, and a sheet whose file says nothing has no `pageSetup` at all
- [ ] xlsx read and write per "xlsx wire mapping": `printOptions`, `pageMargins` (inches), `pageSetup`, `sheetPr/pageSetUpPr@fitToPage`, `headerFooter`, `rowBreaks`/`colBreaks` (`brk@man="1"` only); a `pageSetupDirty` sheet that is not `dirty` keeps `sheetData` and every other element byte-verbatim, its page elements spliced in CT_Worksheet order and their **attributes** spliced in place — unmodelled ones (`r:id` to `printerSettings*.bin`, `horizontalDpi`, `copies`, `blackAndWhite`, `draft`, `cellComments`, `errors`) keep their bytes and position, the SPEC-104 `@show` precedent
- [ ] Paper sizes map both ways through `XLSX_PAPER_SIZES` in `spreadsheet/pageSetup.ts` (codes 1 Letter, 3 Tabloid, 5 Legal, 8 A3, 9 A4, 11 A5 ↔ `core/pageSetup.PAPER_SIZES`; 7 Executive, 12 B4, 13 B5, 14 Folio, 66 A2, 70 A6 by dimension; `paperWidth`/`paperHeight` measures as custom); an unknown code is kept and paginates as Letter; a sheet with no `<pageSetup>` prints on the user's `defaultPageSizeId` preference (SPEC-066), an explicit `<pageSetup>` without `paperSize` on Letter (the schema default)
- [ ] ods read and write per "ods wire mapping": the table's style → `style:master-page-name` → master page → `style:page-layout` properties (`fo:page-width`/`-height`, `fo:margin-*`, `style:print-orientation`, `style:scale-to`, `style:scale-to-X`/`-Y` in the `style` or `loext` namespace, `style:scale-to-pages`, `style:print`, `style:print-page-order`, `style:first-page-number`, `style:table-centering`), header/footer heights folded into the six margins by the Reference rule, and `fo:break-before="page"` on row and column styles → breaks; a change to one sheet **never** edits a master page another table shares — it mints `de-pm{n}` / `de-mp{n}` in `styles.xml` and a `de-ta{n}` table style pointing at them, spliced, with `styles.xml`'s other bytes verbatim
- [ ] Print area and print titles are SPEC-133's sheet-scoped reserved names `_xlnm.Print_Area` and `_xlnm.Print_Titles` (moved by its names adjuster): this spec maps them onto ODS — `table:print-ranges` on `table:table` ↔ `_xlnm.Print_Area`, rows inside `table:table-header-rows` / columns inside `table:table-header-columns` ↔ `_xlnm.Print_Titles` addresses converted by `odsRangeAddressToCanonical`/`canonicalToOdsRangeAddress` (`ods/rangeAddress.ts`, SPEC-121) — and SPEC-133's ODS writer emits these two names only in that form, never as `table:named-range`; `printAreaOf(model, sheet)` / `printTitlesOf(model, sheet)` / `setPrintArea` / `setPrintTitles` in `spreadsheet/pageSetup.ts`, built on SPEC-133's `getReservedName(model, sheetId, short)` / `setReservedName(model, sheetId, short, formula | null)`, are the one accessor every surface uses
- [ ] Headers and footers are stored as Excel's code text per section (`left`/`center`/`right`, `&L`/`&C`/`&R` split by `parseHeaderFooter`, `formatHeaderFooter` the inverse, `formatHeaderFooter(parseHeaderFooter(t)) === t` for every fixture string) and rendered by `renderHeaderFooterSection(section, ctx)` per "Header and footer codes": `&P` (and `&P+n`/`&P-n`), `&N`, `&D` and `&T` (the SPEC-118 locale's short date and time), `&F`, `&A`, `&&`, `&"font,style"`, `&nn`, `&Krrggbb`, `&B`, `&I`, `&U`, `&E`, `&S` are rendered; `&Z` renders empty, `&G` renders nothing (its `legacyDrawingHF` part stays verbatim), `&X`/`&Y`/`&O`/`&H` and theme `&K` forms render as plain text in the default colour; anything unknown is literal and preserved; ODS `text:page-number`/`page-count`/`date`/`time`/`sheet-name`/`file-name` map to `&P`/`&N`/`&D`/`&T`/`&A`/`&F`, `text:span` bold/italic to `&B`/`&I`
- [ ] `src/lib/spreadsheet/print.ts` (framework-free) exports `paginateSheet(model, sheet, options): SheetPrintJob` and `paginateWorkbook(model, scope, options)` implementing "The pagination algorithm" exactly — print area or printable used range (SPEC-131 drawing and SPEC-132 chart extents included), hidden and filter-hidden rows/columns skipped, percent scale or fit-to computing the largest integer percent in 10–100 that fits, column bands then row bands honouring manual breaks (ignored under fit-to, as Excel does), merged ranges never split unless larger than a page, print titles repeated, page order, centring, page numbering — unit-tested step by step against hand-computed fixtures
- [ ] File ▸ "Page setup…" opens `SheetPageSetupDialog.vue` (`sheet-page-setup-dialog`, on `BaseDialog.vue`) with tabs Page / Margins / Header/Footer / Sheet and every control named in "The dialog"; margins use the SPEC-064 `measurementUnit` preference and are validated with `textAreaTooNarrow` / `MARGINS_TOO_LARGE_MESSAGE` from `core/pageSetup.ts`; scale outside 10–400 shows "Enter a scale between 10% and 400%"; OK records one undo entry ("Page setup") via SPEC-121's `record`, Cancel changes nothing
- [ ] Ctrl/Cmd+P and File ▸ "Print…" open the print preview: the grid area shows `SheetPrintPages.vue` (one `sheet-print-page` per page of the job, scaled to fit the area, `aria-label="Page N of M"`) beside `PrintPreviewDialog.vue` — reused, gaining only an optional `options` slot after its Pages fieldset, an optional `actions` slot before Print, an `emptyMessage` prop (default "This document has no content to print.") and a `showBackgroundOption` prop (default `true`), so the text view's preview and every SPEC-034/074 assertion are unchanged — whose `options` slot holds the scope select `sheet-print-scope` ("Active sheet" | "Selection" | "Entire workbook"), and whose `actions` slot holds "Page setup…" (`sheet-print-page-setup`) and "Download PDF" (`sheet-print-download-pdf`); the page range input parses through `core/pageRange.ts`'s `parseNumberRange(input, total, "page")` (SPEC-104; `printRange.ts`'s `parsePageRange` before it lands); an empty job shows "This sheet has nothing to print." and disables Print
- [ ] Printed pages are content, not screenshots of the grid: each `sheet-print-cell` carries the cell's resolved style as CSS through SPEC-117's `cellCss(style, theme)` (`spreadsheet/cellRender.ts`, the mapping `SheetGrid.vue` uses), its value through `formatValue` (SPEC-118) at the **printed** width (`"########"` when it does not fit), text overflow into empty neighbours by SPEC-117's `overflowExtent` (bounded by the page band) and clipping otherwise; drawings print through SPEC-131's `SheetDrawingView.vue` and charts through SPEC-132's chart branch of it (`paintChartIntoCanvas`); gridlines (when `printGridLines`) are 1 CSS px `#000000` edges, headings (when `printHeadings`) are `sheet-print-heading` cells with the column letters / row numbers centred in the default font; the scale is applied **by layout** (widths, heights, font sizes, borders and padding multiplied), never by a CSS transform, because the extractor reads computed font sizes
- [ ] Print goes through the browser: the panel's Print (`print-confirm`) marks `body.is-sheet-printing`, injects `@page` rules built by `sheetPrintPageRuleCss(job)` (one unnamed rule from the first page's paper and one named rule per distinct paper/orientation, each page carrying `page: <name>` — the `printPageRuleCss` technique, restated here because that function is text-product), `margin: 0`; `App.vue`'s global print CSS shows only the pages at their natural size (the preview's fit-to-area transform is dropped), one per sheet of paper (`break-after: page`); screen-only affordances carry `.print-exclude`; a job over `PRINT_PAGE_LIMIT = 250` pages is refused with "This print job is over 250 pages — download a PDF instead"
- [ ] "Download PDF" (and the spreadsheet download menu's "PDF (.pdf)…", `spreadsheet-export-pdf`) builds a real PDF through the core writer: per page `extractPageContent(pageEl, widthPt, heightPt, { cellSelector: ".sheet-print-cell", clipSelector: ".sheet-print-cell", imageSelector, canvasSelector })`, faces loaded and subset exactly as SPEC-034 does, one level-1 `PdfHeading` per sheet (its name) on its first page, `/Title` from `model.title`, saved as `pdfFileName(model.title, DEFAULT_SPREADSHEET_TITLE)` via `triggerBlobDownload`; cell fills and borders are `PdfPathItem` vectors, images and SPEC-132 charts (painted with `core/chartCanvas.ts`'s `paintChartIntoCanvas` into `canvas.sheet-print-chart`) are images, text is selectable and searchable; an `http`/`https`/`mailto` SPEC-135 cell link (`Sheet.hyperlinks`, or a top-level `HYPERLINK()` call) becomes a `/Link` annotation
- [ ] The extractor generalises without changing text-side behaviour (`core/pdfExtract.ts` if SPEC-104 has landed, `src/lib/pdf/domExtract.ts` otherwise): `ExtractOptions` gains `cellSelector` (default `"td.doc-table-cell"`, the selector the cell-chrome loop hard-codes today) and `clipSelector` (default absent: text inside a matching element is clipped to its box — characters whose box extends past it are dropped, no partial glyphs, the one sanctioned difference from the screen); with neither option every existing SPEC-034/076 extraction test is untouched
- [ ] Large jobs stay usable: PDF export mounts, extracts and unmounts pages off-screen at 1:1 (`position: fixed; left: -20000px`, `aria-hidden`, `inert` — never the preview's fit-scaled pages, because `extractPageContent` multiplies computed font sizes by the page's on-screen scale) in a rolling window of `EXPORT_WINDOW_PAGES = 4` with `yieldToMainThread()` between pages, `sheet-print-progress` (`role="status"`) reads "Page {done} of {total}", `sheet-print-abort` stops at the next page boundary with no download and no model change, and export over `MAX_SHEET_PDF_PAGES = 2000` pages is refused with "This export is over 2,000 pages — set a print area first"; preview pages outside the visible scroll window render as sized empty boxes
- [ ] Page breaks show in the grid: View ▸ "Page breaks" (`sheet-menu-page-breaks`, a checkable menu item, session-only, never written) draws `sheet-page-break-line` elements (`data-axis="row|col"`, `data-index` = the first row/column of the next page, `data-kind="auto|manual"`; automatic dashed, manual solid, both from design tokens) from `paginateSheet` of the active sheet, recomputed on an idle callback after a model or setup change; closing a print preview turns it on for that sheet, as Excel does
- [ ] Manual breaks and print areas are editable: Insert ▸ "Page break" (`sheet-menu-insert-page-break`) adds a row break above the active row, a column break left of the active column, or both (Excel's rule by active cell; disabled at `A1`), "Remove page break" (`sheet-menu-remove-page-break`) and "Reset all page breaks" (`sheet-menu-reset-page-breaks`); File ▸ "Print area" ▸ "Set print area" (`sheet-menu-set-print-area`, from the selection's ranges) and "Clear print area" (`sheet-menu-clear-print-area`) — each one undo step ("Insert page break", "Remove page break", "Reset page breaks", "Set print area", "Clear print area") through SPEC-121's `record(label, { sheetId })` before mutating and `commit()` after
- [ ] Structure follows the cells: `pageSetup.ts` registers a `"pageBreaks"` entry in SPEC-124's `REFERENCE_ADJUSTERS` — a break at or after an inserted band shifts by `count`, a break inside a deleted band moves to `at` (deduplicated), a column break likewise — with a test that a row insert above a break moves it, and **no** `MoveAdjuster` (a break belongs to a row or column boundary, not to cells, so SPEC-125's cut-paste leaves it where it is, as Excel does); there is **no** print-area adjuster — print areas and titles are SPEC-133's names, moved by its `names` adjusters in both registries (a second adjuster would shift them twice), asserted once here end to end; the same change deletes this spec's rows from SPEC-124's "Retained references" table (`rowBreaks/brk@id`, `colBreaks/brk@id`, `table:table@table:print-ranges`)
- [ ] Setup edits mark only what they touch, through `markPageSetupEdited(model, sheet, what)`: xlsx — `pageSetupDirty`; ods — a layout, margin, header/footer, centring or scale edit splices `styles.xml` and the table's start tag (`table:style-name`, `table:print-ranges`) without re-serialising rows, while a break or print-titles edit sets `Sheet.dirty` (both live in row/column structure); `stripRetention` drops `pageSetupDirty` and keeps `pageSetup`, so "Download a copy as…" converts setups field for field and the "page setup" row leaves SPEC-121's `UNCONVERTIBLE_CONTENT` table
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`: clean sheets are re-read; a dirty xlsx sheet reads its page elements from its retained worksheet XML (SPEC-122 kept them verbatim and SPEC-124's `retained` adjuster kept `brk@id` current); a dirty ods sheet reads its master page from `styles.xml` in `opaque`, its print ranges from its retained `sourceXml` start tag (kept current the same way) and its header rows and row/column breaks from the retained table body — whose positions may predate a structural edit made before this spec landed, since no adjuster rewrote them
- [ ] Round trip: a no-edit download is byte-identical part for part (`equalBytes`, `src/lib/core/equality.ts`, after `unzip`); edit every setup field → save → reopen reproduces the model in both formats (field-level `toEqual`); writing the reopened model twice is byte-stable; sanctioned deltas, each pinned by a test: margins round to 1/1000 inch in xlsx, an ods header region holding an unmapped field (`text:title`, `text:author-name`) is regenerated with that field's current text once its setup is edited, and `style:scale-to-pages` converts to xlsx as fit 1 × N
- [ ] Printing and exporting mutate nothing: the model is deep-frozen in the unit tests, no dirty flag or store write follows a preview, print or PDF, and `buildPdf` twice on the same pages is `equalBytes`-identical
- [ ] Fixtures: `spreadsheet/fixtures/build.ts` gains `spec136-print.xlsx` (Excel-shaped: A4 landscape, fit 1 wide × 0 tall with `pageSetUpPr fitToPage="1"`, custom margins, gridlines and headings on, centred horizontally, `&L&"Arial,Bold"&12Budget&C&P of &N&R&D` header, different first page footer, a manual row break, `_xlnm.Print_Area` and `_xlnm.Print_Titles` in `workbook.xml`, `r:id` to a dummy `xl/printerSettings/printerSettings1.bin`, a merged range straddling a natural break, a hidden column) and `spec136-print.ods` (LibreOffice-shaped: two tables sharing master page `Default`, `style:scale-to-X="1"`, header regions with `text:page-number`/`text:page-count`/`text:sheet-name`, `table:print-ranges`, `table:table-header-rows`, a row style with `fo:break-before="page"`)
- [ ] e2e (`e2e/SPEC-136.spec.ts`): open `spec136-print.xlsx` → Page setup shows A4, landscape, fit 1 × automatic and the header text → Print preview reads "3 pages" with the title row repeated at the top of page 2 and the merged range whole on one page → close; page-break lines are shown → insert a page break at row 30 → preview reads "4 pages" → Download PDF → the captured file starts `%PDF-1.7`, its page tree reads `/Count 4`, its first `/MediaBox` is `[0 0 841.89 595.28]` and its Info `/Title` is the workbook title → scope "Entire workbook" sums both sheets' pages → set scale 50% → fewer pages → download the workbook, reopen → the setup, the break and the print area come back; the read half repeats against `spec136-print.ods`, and editing its first sheet's header leaves the second table's master page untouched

## Out of scope
Page Layout view and Page Break Preview (dragging breaks in a page view) — the
grid shows lines only. Header/footer pictures (`&G`, carried verbatim, never
rendered or authored). Printing notes "at end of sheet" or "as displayed"
(`pageSetup@cellComments` carried verbatim). Print error display options
(`@errors`), black-and-white and draft printing, copies, printer selection and
duplex — the browser's dialog owns the printer. Custom views. Print preview
editing of margins by dragging. Tagged-PDF structure, PDF/A, password
protection. Printing a chart sheet beyond SPEC-132's own rendering. PDF export
of a range of pages other than through the preview's range. Server-side or
headless printing. Fitting text by shrinking fonts beyond the sheet's own
scale.

## Reference

### Dependencies restated
- `specs/034-print-and-export-pdf.md` — `buildPdf(input: PdfBuildInput):
  Promise<Uint8Array>` over `PdfPageInput { widthPt; heightPt; items; links;
  headings; images; paths? }` with `PdfPathItem { subpaths; stroke?; fill?; onTop? }`,
  `PDF_MIME_TYPE`, `pdfFileName(title, fallback)`; `extractPageContent(pageEl,
  widthPt, heightPt, options)` reads text via `Range.getClientRects` (fonts mapped by
  `faceKeyFor`), images, canvases, `td.doc-table-cell` fills and per-edge borders
  (`filledRectItem`, `cellBorderItems`, `edgePaintItems`), honouring `.print-exclude`;
  faces are fetched and parsed once per session (`usePdfExport.ts`'s private
  `loadFace`/`loadImages` — lifted, behaviour-identical, into an exported app-layer
  `src/views/pdfAssets.ts` unless SPEC-104 already shared them, in which case its
  module is reused). Printing is `window.print()` under an injected `@page` rule with
  `margin: 0`. **If SPEC-104 has landed (it is a dependency), these are
  `src/lib/core/pdfBuild.ts`, `core/pdfExtract.ts`, `core/pdfImages.ts`,
  `core/pdfFonts.ts`, `core/sfnt.ts` and `core/pageRange.ts`; before it they are
  `src/lib/pdf/buildPdf.ts`, `domExtract.ts`, `imageEncode.ts`, `fontMap.ts`,
  `sfnt.ts` and `src/lib/printRange.ts`.** Either way only the view layer imports
  them — `lib/spreadsheet/print.ts` imports nothing PDF.
- `specs/104-…` — `ExtractOptions.imageSelector`/`canvasSelector`, the rolling
  `EXPORT_WINDOW_PAGES = 4` export window and `PRINT_PAGE_LIMIT = 250`, both
  restated here as this spec's own constants.
- `specs/027-page-setup.md` — `core/pageSetup.ts`: `PageSize { widthCm; heightCm }`,
  `Margins`, `PAPER_SIZES` (A4, Letter, Legal, A3, A5, Tabloid, portrait),
  `matchPaperSize`, `Orientation`, `swapOrientation`, `textAreaTooNarrow`,
  `MARGINS_TOO_LARGE_MESSAGE`; units in `core/units.ts` (`PX_PER_CM`, `cmToPx`).
- `specs/116-…` model and invariants, reader versions; `specs/117-…` `GridMetrics`
  (`colWidth`, `rowHeight` — px and pt at 100%), merges, hidden rows/cols, the overflow
  rule; `specs/118-…` `formatValue(value, code, { locale, dateSystem, widthPx })` →
  `{ text; color?; alignHint?; overflow?: "hash" }`; `specs/121-…` undo `record(label,
  scope)` / `commit()`, `Sheet.dirty`, `stripRetention`, `UNCONVERTIBLE_CONTENT`,
  `DEFAULT_SPREADSHEET_TITLE` ("Untitled spreadsheet"); `specs/122-…`
  CT_Worksheet order (`printOptions`, `pageMargins`, `pageSetup`, `headerFooter`,
  `rowBreaks`, `colBreaks` after `hyperlinks`), the generated `de-` ODS style names;
  `specs/123-…` style resolution; `specs/124-…` `StructuralEdit { kind; sheetId; at;
  count }`, `REFERENCE_ADJUSTERS`, the `retained` adjuster; `specs/133-…`
  `DefinedName` with `localSheetId` scope, `getReservedName`/`setReservedName`, the
  `names` adjuster. SPEC-117's `cellCss`/`overflowExtent`/`buildGridMetrics`
  (`cellRender.ts`, `geometry.ts`); SPEC-131's `SheetDrawingView.vue` and
  `drawingRectPx`, SPEC-132's `paintChartIntoCanvas(canvas, chart, w, h, theme)`;
  SPEC-135's `Sheet.hyperlinks` and `hyperlinkFormulaCall(cell)`.

### The page-setup model (`src/lib/spreadsheet/model.ts`)
```ts
export interface SheetMargins {               // cm; header/footer = page edge → header top / footer bottom
  topCm: number; rightCm: number; bottomCm: number; leftCm: number; headerCm: number; footerCm: number;
}
export interface HeaderFooterSections { left?: string; center?: string; right?: string }   // code text
export interface HeaderFooterPair { header?: HeaderFooterSections; footer?: HeaderFooterSections }
export interface SheetHeaderFooter {
  odd: HeaderFooterPair;
  even?: HeaderFooterPair;                     // used iff differentOddEven
  first?: HeaderFooterPair;                    // used iff differentFirst
  differentOddEven?: true; differentFirst?: true;
  scaleWithDoc?: false; alignWithMargins?: false;   // absent = true (Excel's defaults)
}
export interface SheetPageSetup {
  orientation?: "landscape";                   // absent = portrait
  paper?: { widthCm: number; heightCm: number; xlsxPaperSize?: number };   // portrait dimensions
  scale?: number;                              // 10..400; absent = 100
  fitTo?: { width: number; height: number };  // pages, 0 = unconstrained; present ⇒ fit mode
  odsScaleToPages?: number;                    // ods style:scale-to-pages: fit into N pages in total
  margins?: SheetMargins;                      // absent = DEFAULT_SHEET_MARGINS
  centerHorizontally?: true; centerVertically?: true;
  printGridLines?: true; printHeadings?: true;
  pageOrder?: "overThenDown";                  // absent = down, then over
  firstPageNumber?: number;                    // absent = automatic
  headerFooter?: SheetHeaderFooter;
  rowBreaks?: number[];                        // 0-based rows that START a page; ascending, unique
  colBreaks?: number[];
}
```
`DEFAULT_SHEET_MARGINS` is Excel's "Normal": top/bottom 0.75 in, left/right
0.7 in, header/footer 0.3 in (1.905, 1.778, 0.762 cm). Commands in
`spreadsheet/pageSetup.ts`: `setPageSetup(model, sheet, patch)`,
`insertPageBreak(model, sheet, at: CellRef)`, `removePageBreak`,
`resetPageBreaks`, `setPrintArea(model, sheet, ranges | null)`,
`setPrintTitles(model, sheet, { rows?: [number, number]; cols?: [number, number] } | null)`,
`printAreaOf`, `printTitlesOf`, `effectivePaper(sheet, preferences)`.

### xlsx wire mapping

| Model | xlsx |
| --- | --- |
| orientation | `pageSetup@orientation="landscape"` (`portrait`/`default` → absent) |
| paper | `pageSetup@paperSize` code, or `@paperWidth`/`@paperHeight` (`"210mm"`, `"8.5in"`) |
| scale / fitTo | `pageSetup@scale`; `sheetPr/pageSetUpPr@fitToPage="1"` + `pageSetup@fitToWidth`/`@fitToHeight` (default 1, 0 = automatic) |
| margins | `pageMargins@left right top bottom header footer`, inches (written to 3 decimals) |
| centring, gridlines, headings | `printOptions@horizontalCentered verticalCentered gridLines headings` (`gridLinesSet` retained) |
| pageOrder | `pageSetup@pageOrder="overThenDown"` |
| firstPageNumber | `pageSetup@firstPageNumber` + `@useFirstPageNumber="1"` |
| headerFooter | `headerFooter@differentOddEven differentFirst scaleWithDoc alignWithMargins` > `oddHeader`, `oddFooter`, `evenHeader`, `evenFooter`, `firstHeader`, `firstFooter` |
| rowBreaks / colBreaks | `<rowBreaks count manualBreakCount><brk id max man="1"/>` — `id` is the 0-based row the break is above (column: left of); `max` 16383 / 1048575 |
| print area / titles | `definedName name="_xlnm.Print_Area" localSheetId` / `_xlnm.Print_Titles` (`Sheet1!$1:$2,Sheet1!$A:$A`) — SPEC-133 |

`pageSetUpPr` is spliced inside `sheetPr` the way SPEC-121's tab colour splices
`sheetPr/tabColor`. Header/footer text is XML-escaped code text; `&` codes are
literal characters in the XML (`&amp;P`).

### ods wire mapping
The table's automatic style (`ta1`) names `style:master-page-name`; the master
page names its `style:page-layout-name`.

| Model | ods |
| --- | --- |
| paper, orientation | `style:page-layout-properties@fo:page-width fo:page-height style:print-orientation` (dimensions swapped for landscape) |
| scale / fitTo | `style:scale-to="75%"`; `style:scale-to-X`/`-Y` (or `loext:`), 0 = unconstrained; `style:scale-to-pages` → `odsScaleToPages` |
| margins | page `fo:margin-left/right` → left/right; header shown: `headerCm = fo:margin-top`, `topCm = fo:margin-top + header fo:min-height + header fo:margin-bottom`; header hidden: `topCm = fo:margin-top`, `headerCm` default; footer symmetric. Written back as `fo:margin-top = headerCm`, header `fo:margin-bottom = 0.25cm`, `fo:min-height = max(0, topCm − headerCm − 0.25cm)` |
| centring | `style:table-centering="horizontal|vertical|both|none"` |
| gridlines, headings | `style:print` list containing `grid`, `headers` (other tokens retained) |
| pageOrder | `style:print-page-order="ltr"` = over then down, `ttb` = down then over |
| firstPageNumber | `style:first-page-number="N"` (`continue` = automatic) |
| headerFooter | master page `style:header`/`style:footer` with `style:region-left/center/right` (or a single `text:p` = centre); `style:header-left` = even; `style:header-first` (ODF 1.3, or `loext:`) = first; `style:display="false"` = none |
| breaks | `fo:break-before="page"` in the row's `style:table-row-properties` / column's `style:table-column-properties` |
| print area / titles | `table:table@table:print-ranges="Sheet1.A1:Sheet1.D20"`; `table:table-header-rows` / `table:table-header-columns` wrappers |

SPEC-122's ODS table writer gains the break flag in its generated row/column
style key and emits a row element for a broken row even when it stores nothing.

### Header and footer codes
`parseHeaderFooter(text)` splits on `&L`/`&C`/`&R` (text before the first
section code is centre); `tokenizeHeaderFooter(section): HfToken[]` reads `&&`,
`&P`, `&P+n`, `&P-n`, `&N`, `&D`, `&T`, `&F`, `&Z`, `&A`, `&G`, `&"family,style"`,
a run of digits (font size, pt), `&K` + six hex digits or a theme form
(`&K01+000`), `&B`, `&I`, `&U`, `&E`, `&S`, `&X`, `&Y`, `&O`, `&H`; any other
`&` + character is literal. `renderHeaderFooterSection(tokens, ctx: { page;
pages; date: Date; sheetName; fileName; locale; dateSystem })` → styled runs;
the default font is the workbook's default style font at 11 pt. Header text sits
`headerCm` from the top edge, footer text `footerCm` from the bottom; with
`alignWithMargins` the left/right sections align to the page margins, and with
`scaleWithDoc` the sections scale with the sheet scale.

### The pagination algorithm (`paginateSheet`, normative)
All lengths in pt; column widths px × 0.75, row heights pt; `s` = scale ÷ 100.
1. **Area.** The print area's ranges if set (each range starts on a new page, in
   order), else the selection's ranges for scope "Selection", else the printable
   used range: the bounding box of cells with a value, a fill or a border, merged
   ranges touching it, and drawing/chart anchors. Empty → no pages.
2. **Visible tracks.** Hidden and filter-hidden rows and columns are dropped.
3. **Printable box.** Paper (orientation applied) minus left/right and top/bottom
   margins; minus the print-title rows' total height and title columns' total width
   (× s); minus, with `printHeadings`, a heading column of `max(24, 6 × digits + 6)`
   pt and a heading row of 15 pt (× s).
4. **Scale.** Percent mode: `s` as set. Fit mode: scanning `p` from 100 down to 10,
   the first integer percent at which steps 5–6 yield ≤ `fitTo.width` column bands
   and ≤ `fitTo.height` row bands (0 = no limit); 10 if none does. `odsScaleToPages`:
   the first `p` at which bands-wide × bands-tall ≤ N. Manual breaks are ignored in
   both fit modes. Scanning, not bisection: the merge rule makes counts
   non-monotone in rare cases.
5. **Column bands.** Left to right, accumulate scaled widths; start a new band
   before a column with a manual break (percent mode) or whose addition would exceed
   the box width. A boundary that would fall strictly inside a merged range moves
   back to the merge's first column if the band keeps at least one column; a merge
   wider than the box splits at the box edge. A single column wider than the box is
   its own band, clipped.
6. **Row bands.** The same, top to bottom, with heights and row breaks.
7. **Order.** `downThenOver`: every row band of column band 1, then column band 2…;
   `overThenDown`: every column band of row band 1, then row band 2….
8. **Titles.** Each page draws the title rows above its band and the title columns
   left of it, unless the band already contains them.
9. **Placement.** The content block (headings + titles + band) sits at the top-left
   of the printable box, centred on each axis that is set to centre.
10. **Numbering.** Pages number from `firstPageNumber` (else 1) and continue across
    sheets in a workbook job; a sheet with its own `firstPageNumber` restarts; `&N`
    is the job's total.
Output: `SheetPrintJob { pages: SheetPrintPage[] }`, each page `{ sheetId;
number; paperPt: { width; height }; marginsPt; scale; rows: number[]; cols:
number[]; titleRows: number[]; titleCols: number[]; originPt: { x; y };
header?: HeaderFooterSections; footer?: HeaderFooterSections }`, plus
`autoBreaks(job, sheetId)` for the grid's lines.

### The dialog
Page: `sheet-page-portrait`/`-landscape`, `sheet-page-paper` (the
`XLSX_PAPER_SIZES` names + "Custom"), `sheet-page-scale-mode-percent` +
`sheet-page-scale`, `sheet-page-scale-mode-fit` + `sheet-page-fit-wide` /
`sheet-page-fit-tall` (blank = automatic), `sheet-page-first-number` (blank =
auto). Margins: `sheet-margin-top`/`-right`/`-bottom`/`-left`/`-header`/`-footer`,
`sheet-center-horizontally`, `sheet-center-vertically`. Header/Footer:
`sheet-header-preset`/`sheet-footer-preset` (None; Page 1; Page 1 of ?; sheet
name; file name; date), `sheet-header-left`/`-center`/`-right` and
`sheet-footer-left`/`-center`/`-right` code-text inputs with insert buttons
`sheet-hf-insert-page`/`-pages`/`-date`/`-time`/`-file`/`-sheet`,
`sheet-hf-different-first`, `sheet-hf-different-odd-even`. Sheet:
`sheet-print-area` and `sheet-print-title-rows`/`-cols` (validated with
`parseRange`), `sheet-print-gridlines`, `sheet-print-headings`,
`sheet-page-order-down`/`-over`. Actions `sheet-page-setup-ok`/`-cancel`.

### Composables and components
`src/views/spreadsheet/usePageSetup.ts` (dialog state, commands, page-break
lines), `useSheetPrint.ts` (preview, scope, range, print), `useSheetPdfExport.ts`
(window, progress, abort, faces) — each ≤ 400 lines. Components
`SheetPageSetupDialog.vue`, `SheetPrintPages.vue`, `SheetPrintPage.vue`,
`SheetPageBreakLines.vue` under `src/components/spreadsheet/`. The print CSS
(`body.is-sheet-print-preview`, `body.is-sheet-printing`) lives in `App.vue`, the
only home for global styles. `SpreadsheetView.vue` gains three `use*` calls and
template hookup; `SpreadsheetView.size.spec.ts`'s budget never rises.

### Tests
- `spreadsheet/print.spec.ts` — every algorithm step with hand-computed cases:
  print areas, hidden tracks, fit-to scanning, ignored breaks under fit, merge
  boundaries (moved, and split when wider than a page), titles, both orders,
  centring, numbering across sheets.
- `spreadsheet/pageSetup.spec.ts` — commands, paper mapping, defaults, the
  `"pageBreaks"` adjuster, `printAreaOf`/`printTitlesOf`.
- `spreadsheet/headerFooter.spec.ts` — parse/format identity, every code's
  rendering, unknown codes literal, ODS field mapping.
- `xlsx/pageSetup.spec.ts`, `ods/pageSetup.spec.ts` — read, attribute-level splice,
  shared master page minting, margin folding, `loext:` tolerance; the upgrade path in
  `open.spec.ts`; `pageSetupRoundTrip.spec.ts` per the round-trip AC.
- `pdfExtract.spec.ts` (or `domExtract.spec.ts`) — `cellSelector` and `clipSelector`
  defaults and behaviour; `PrintPreviewDialog.spec.ts` — the new slots and props with
  the text-side render unchanged; `usePageSetup.spec.ts`, `useSheetPrint.spec.ts`,
  `useSheetPdfExport.spec.ts` (abort, limits, no mutation); `e2e/SPEC-136.spec.ts`.

### Specs this changes
- `specs/034-…` / `specs/104-…` — the extractor gains `cellSelector` and
  `clipSelector`; `PrintPreviewDialog.vue` gains two slots and two props; the face
  loader is shared. Their ACs and tests are unchanged.
- `specs/122-…` — row/column style keys carry the break flag; page elements are
  modelled, not verbatim. `specs/133-…` — the two reserved names map onto ODS
  print ranges and header rows/columns, never `table:named-range`.
- `specs/124-…` — `pageBreaks` joins the registry (print areas need none:
  SPEC-133's `names` adjuster moves them), and this spec's three rows leave
  "Retained references". `specs/121-…` — the "page setup" row leaves
  `UNCONVERTIBLE_CONTENT`; `SHEET_BY_REFERENCE_FIELDS` gains `pageSetup`.
- `specs/135-…` — `pageSetup@cellComments` stays verbatim; cell links become PDF
  links.
