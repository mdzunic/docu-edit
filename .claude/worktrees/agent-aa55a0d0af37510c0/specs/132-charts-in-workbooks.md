---
id: SPEC-132
title: Charts in workbooks
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-052, SPEC-092, SPEC-120, SPEC-125, SPEC-131]
---
## Why
SPEC-131 places a workbook's chart frames and paints each one as a dashed box
labelled "Chart". In a spreadsheet that is the wrong half to stop at: the
chart is usually the reason the workbook was sent. And a sheet chart is a
different animal from the document and slide charts SPEC-052 and SPEC-092
render. Those carry their data inside the chart part, in caches nobody edits
anywhere else. A sheet chart's data **is the cells**: every series names its
source with a formula (`c:f` in `.xlsx`, `chart:values-cell-range-address` in
`.ods`), and the cached numbers beside it are only a snapshot of what the
cells said when the file was saved. Type a new month's revenue into the
column a chart plots and the chart is expected to grow a bar.

The engine to paint, and the codecs to read and regenerate, are already in
core: `charts.ts` (`ChartVisual`, the catalogue), `chartCanvas.ts`
(`paintChart`), `chartPartOoxml.ts` and `chartPartOdf.ts` (SPEC-092's hoisted
part codecs), `chartFeatures.ts` (what a rewrite costs), and the gallery,
toolbar and dialogs in `src/components/` import nothing but core — with one
exception this spec removes (`ChartOptionsDialog.vue` reaches into
`documentModel.ts` for a one-line theme pass-through). SPEC-120 gives cells
their values. What is missing is the link: references read from the part,
resolved against the grid, re-resolved whenever the cells change, adjusted
when rows move, and written back so that Excel and LibreOffice see caches
that agree with the cells.

**This file is the normative source for the sheet chart — the `chart`
content kind of a `SheetDrawing` holding a core `ChartVisual` plus its live
source references — the refresh rule, the three write paths (verbatim, cache
patch, regenerate), insert-from-selection heuristics, chart-reference
adjustment, and chart sheets.** SPEC-008's law still outranks the feature: a
chart nobody touched is written back byte-identical, a chart whose data
changed has its caches patched inside the producer's own XML, and only a
chart the user restyles through the dialogs is regenerated — with the
SPEC-092 notice saying what that cost.

## Acceptance criteria
- [ ] SPEC-131's `{ kind: "chart"; sourcePart? }` content widens to `SheetChartContent` per "The chart model" — `visual: ChartVisual`, optional `refs: SheetChartRefs` (canonical range formulas per series: `name`, `categories`, `values`), `sourcePart`, and the write flags `edited`, `cachesDirty`, `refsDirty`, plus `readOnlyReason` — and `Sheet.kind` widens to `"grid" | "chart" | "foreign"`; the model stays plain JSON
- [ ] xlsx read: a chart drawing's part is parsed with core `parseChartXml` into `visual` (out of catalogue → `unsupportedChartXml` = the part text, painted as `paintChart`'s placeholder) and with `readChartPartRefs` (`spreadsheet/xlsx/chartXml.ts`) into `refs` — each `c:ser`'s `c:tx/c:strRef/c:f`, `c:cat` or `c:xVal` (`c:strRef`/`c:numRef`)`/c:f` and `c:val` or `c:yVal` `c:numRef/c:f`, through `fromXlsxFormula`; a literal `c:tx/c:v` name leaves `name` absent; a part with `c:pivotSource` or a `c:multiLvlStrRef` gets `readOnlyReason` per "Read-only charts"
- [ ] Charts this spec does not model keep their references right without being parsed: an out-of-catalogue chart (`unsupportedChartXml`, painted as `paintChart`'s placeholder) and an Excel chartex chart (SPEC-131's `foreign` "Chart" drawing) pass their parts through untouched by this spec, still adjusted by SPEC-124's `retained` rewrite (whose chart row this spec narrows to exactly those parts and widens to chartex `xl/charts/chartEx{n}.xml` `cx:f`) and renamed by SPEC-121's write-time `sourceSheetRemap`
- [ ] ods read: a chart drawing's `Object N/content.xml` is parsed with core `parseOdfChartXml` into `visual` and with `readOdfChartRefs` (`spreadsheet/ods/chartXml.ts`) into `refs` — `chart:series@chart:values-cell-range-address`, `@chart:label-cell-address`, `chart:domain@table:cell-range-address` (scatter x), the x axis's `chart:categories@table:cell-range-address` and `chart:plot-area@table:cell-range-address` — converted by `odsRangeAddressToCanonical`; addresses into the chart's own `local-table` mean the chart has no `refs`
- [ ] Refresh per "Refresh": after every committed edit, undo, redo and recalculation pass, `useSheetCharts` re-resolves every chart whose references intersect a changed cell (or whose sheet's row/column visibility changed while it plots visible cells only) through `refreshSheetChart(model, sheetId, drawingId)`; a refresh that changes the visual sets `cachesDirty` and is part of the triggering edit's undo entry, never its own; a refresh that changes nothing leaves the drawing object identical
- [ ] Open trusts caches exactly as SPEC-120 trusts cached cell values: no chart is refreshed at open, except one whose part carries a reference but no cache for it, which is refreshed for display without setting `cachesDirty` — so opening and downloading an untouched workbook writes every chart part byte-identical
- [ ] Resolution per "Resolving references": `resolveChartRange(model, canonical, resolveName?)` accepts a cell or range reference (sheet-qualified), a parenthesised union of them and, through the optional `resolveName` hook (which SPEC-133 later supplies), a defined name that resolves to one — without the hook a name resolves to `null`; anything else (a function, an external `[1]` reference, `#REF!`) returns `null` and that series keeps its cached values; categories show each cell's SPEC-118 `formatValue` text, values take numbers, blank and `#N/A` cells are gaps (`null`), every other value plots as 0, and with `plotVisibleOnly` cells in hidden or filtered rows/columns are skipped
- [ ] Write paths per "Writing": an untouched chart's part (and its `.rels`, `chartStyle`/`colors` parts, embeddings) is re-emitted verbatim; `cachesDirty`/`refsDirty` without `edited` patches the producer's part in place — each `c:f` text and each `c:strCache`/`c:numCache`'s `c:ptCount` + `c:pt` list (`c:formatCode` and `c:extLst` kept), or the ODS series/plot-area/categories addresses plus the `local-table` data cache — leaving every other byte as it was; `edited` regenerates the part through core `chartPartXml(visual, refs)` / `odfChartContentXml(visual, w, h, theme, refs)` and shows the first-edit `chart-rewrite-notice` built from `lostChartFeatures`
- [ ] Core's chart codecs change only additively — two optional trailing parameters, their two wire types and one newly exported builder: `chartPartXml(chart, refs?: ChartPartRefs)` writes the given wire `c:f` texts (and a literal `c:tx/c:v` for a series with no name reference), `odfChartContentXml(chart, widthCm, heightCm, theme?, refs?: OdfChartRefs)` writes sheet addresses, `chart:categories` and the plot-area range, and `odfChartTableXml(chart)` exports the existing `local-table` builder; with the new arguments absent every existing output is unchanged, asserted by `chartPartOoxml.spec.ts`, `chartPartOdf.spec.ts` and both products' chart round-trip suites passing untouched
- [ ] Reference adjustment per "Reference adjustment": a `REFERENCE_ADJUSTERS` entry `{ id: "charts" }` runs SPEC-124's `adjustFormula` over every modelled chart's `refs` (host = the chart's sheet; charts on every sheet, since a chart may plot another sheet's data), sets `refsDirty` and refreshes (read-only charts only adjust); every ODS chart frame's `draw:notify-on-update-of-ranges` in its `sourceXml` is rewritten whenever it no longer matches (marking the frame's sheet `dirty`, since the frame lives in its table); the same change retires SPEC-124's `Object N/content.xml` and `draw:object@draw:notify-on-update-of-ranges` rows for modelled charts; a SPEC-125 `MoveAdjuster` `{ id: "charts" }` re-points every modelled chart's `refs` with `repointFormula` when cells are cut-pasted or drag-moved (Excel's charts follow moved data; the chart drawing itself stays put, per SPEC-131) and refreshes; `charts.spec.ts` asserts that a row insert above a chart's data moves its references and caches, that a row insert inside the data range extends them, and that `moveRange` of the whole data block re-points them
- [ ] Sheet rename and delete reach charts on the model side: SPEC-121's rename and delete commands apply their `SheetRemap` through `remapSheetReferences` to every chart's `refs` too (this spec adds that step to `sheets.ts`), so a reference into a deleted sheet becomes `#REF!` and its series keeps its caches; a patched or regenerated chart part is model text and is excluded from SPEC-121's write-time `sourceSheetRemap`, which keeps covering the parts that pass through
- [ ] Insert ▸ Chart (`sheet-menu-insert-chart`) opens `sheet-chart-gallery-popover` around `ChartGallery.vue` unchanged (`chart-gallery`, `chart-gallery-<chartType>`); picking a type builds the chart from the selection with `buildChartFromRange` per "Insert from selection" — source range, header detection, series orientation and the scatter x column pinned there — anchors it `twoCell` (`editAs: "twoCell"`) at `NEW_CHART_WIDTH_CM` × `NEW_CHART_HEIGHT_CM` with its top-left two columns right of the source range on the range's top row, and selects it; an empty selection with an empty current region shows the toast "Select the data you want to chart first" and inserts nothing
- [ ] A selected chart shows `sheet-chart-toolbar` hosting `ChartToolbar.vue` (testids reused: `chart-edit-data-button`, `chart-options-button`, `chart-duplicate-button`, `chart-delete-button`, `chart-unsupported-reason`); "Chart options" opens `ChartOptionsDialog.vue` unchanged apart from its import (testids reused, theme = `model.theme ?? DEFAULT_THEME`); "Edit data" on a chart with `refs` opens `SheetChartDataDialog.vue` per "Select Data", and on a chart without `refs` opens `ChartDataSheetDialog.vue` unchanged; an unsupported or read-only chart disables both and shows its reason through ChartToolbar's new optional `reason` prop
- [ ] `SheetChartDataDialog.vue` (`sheet-chart-data`) edits the source per "Select Data": `sheet-chart-data-range` (SPEC-122 point mode), `sheet-chart-data-series-rows`/`-columns`, `sheet-chart-data-first-row`, `sheet-chart-data-first-col`, `sheet-chart-data-apply`, `sheet-chart-data-cancel`; references that do not tile one rectangle show "The data range is too complex to be displayed" (`sheet-chart-data-complex`) until a new range is chosen; applying keeps the patch path when the series count and every reference element kind are unchanged and regenerates (`edited`) otherwise
- [ ] Commands in `src/lib/spreadsheet/charts.ts` (framework-free, core + siblings only): `insertSheetChart`, `setSheetChartData`, `setSheetChartOptions`, `setSheetChartSeriesColor`, `setSheetChartSource`, `duplicateSheetChart` (a new chart regenerated from the source's `visual` + `refs`, offset one row and one column, its own new part), `refreshSheetChart`, `chartsAffectedBy`; each is copy-on-write over `Sheet.drawings`, one undo entry with labels "Insert chart", "Edit chart data", "Chart options", "Chart series colour", "Chart data range", "Duplicate chart", and SPEC-131's "Delete chart"/"Move chart"/"Resize chart" unchanged
- [ ] New charts serialise per "New chart parts", golden-first (`spreadsheet/xlsx/chartGolden.ts`, `ods/chartGolden.ts` hand-authored against ECMA-376 §20.5.2.16/§21.2 and ODF 1.2 §10.4.6/§16 before the serialisers): xlsx `xl/charts/chart{n}.xml` + `CHART_CONTENT_TYPE` Override + a `CHART_RELATIONSHIP` in the drawing `.rels` + the canonical `xdr:graphicFrame`; ods `Object {n}/content.xml` (+ the `styles.xml`/`meta.xml` the golden settles on) + manifest entries + a `draw:frame` > `draw:object` with `draw:notify-on-update-of-ranges`; a deleted chart's part leaves the package by SPEC-131's reachability rule
- [ ] Chart sheets per "Chart sheets": an xlsx `chartsheet` whose drawing resolves to one chart becomes `kind: "chart"` with that drawing in `drawings`; its tab opens `SheetChartSheet.vue` (`sheet-chart-sheet`) painting the chart fitted to the view at its anchor's aspect ratio, refreshed and editable through the same toolbar and dialogs, never moved, resized or deleted as a drawing; SPEC-121's rename, move, hide, tab colour and delete apply, Duplicate is disabled with the title "Chart sheets can't be duplicated yet"; an unresolvable chartsheet stays `foreign` with SPEC-116's placeholder
- [ ] Cross-format "Download a copy as…" carries every catalogue chart across through the target's regenerate path with references converted (`toXlsxFormula` / `canonicalToOdsRangeAddress`), and reports each unsupported, chartex or read-only chart it drops; a chart sheet becomes, in `.ods`, a grid sheet holding that chart at the sheet's top-left
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`: clean sheets and chart sheets are re-read; a dirty sheet's step upgrades each SPEC-131 chart placeholder in place by parsing its `sourcePart` from `opaque` — charts live in their own parts, whose `c:f` texts SPEC-124's and SPEC-125's `retained` rewrites have kept current through every structural edit and move made before this spec landed, so the step reads already-adjusted references (sheet renames are still source-time text: the step applies SPEC-121's `sourceSheetRemap` to the references it reads)
- [ ] Performance: refreshing and repainting a 12-series × 100-category chart after one cell edit takes under 100 ms (SPEC-052's budget), and a commit that touches no charted cell runs no refresh — `chartsAffectedBy` is an interval test over each chart's resolved ranges, cached per chart and invalidated by `refsDirty`
- [ ] Fixtures from `spreadsheet/fixtures/build.ts`: `e2e/fixtures/spec132-charts.xlsx` (Excel-shaped: a clustered column over `Data!$A$1:$D$5` with caches, `c:extLst` (`c15`), `chartStyle`/`colors` parts and their rels; a line chart on another sheet plotting `'Q1 Data'!$B$2:$B$13`; a pie; a scatter; a `c:bar3DChart`; a waterfall in `mc:AlternateContent Requires="cx1"` with Excel's Fallback; a chart sheet) and `spec132-charts.ods` (LibreOffice-shaped: a cell-anchored bar chart with `draw:notify-on-update-of-ranges`, its `Object 1/content.xml` with sheet addresses and `local-table`, `styles.xml`, `meta.xml`, `ObjectReplacements/Object 1`; a `chart:surface` chart)
- [ ] Tests: `spreadsheet/charts.spec.ts` (heuristics table, resolution, refresh rules incl. gaps/hidden/#N/A, commands, the adjuster incl. the row-insert-above test, rename/delete), `xlsx/chartXml.spec.ts` and `ods/chartXml.spec.ts` (reference reading, cache and formula patching leave every other byte identical, golden equality), `ods/rangeAddress.spec.ts` additions (chart address forms), `chartsRoundTrip.spec.ts` (`equalBytes` untouched; patched parts reopen with refreshed caches; created charts `equalStructural` after save → reopen in both formats; second write byte-stable), core's two codec specs gaining cases for the new parameters, `useSheetCharts.spec.ts`
- [ ] e2e (`e2e/SPEC-132.spec.ts`): open `spec132-charts.xlsx` → the column chart paints bars, the 3-D chart shows the placeholder with disabled edit buttons and the waterfall its "Chart" placeholder → type a larger number into `Data!D3` → its bar grows → insert a row above the data → the chart still plots the same values → select `A1:C4` on a sheet → Insert ▸ Chart ▸ Line → a line chart lands right of the data → Chart options → legend "bottom" → Apply → open the chart sheet tab → it paints → download, reopen → the edited value's cache, the moved references, the new chart and the options persist, and the 3-D chart's part differs from the source only in its adjusted `c:f` texts; the open-edit-download leg repeated on `spec132-charts.ods`

## Out of scope
Changing an existing chart's type, combo charts, secondary axes, trendlines,
error bars, per-point formatting, data-label formats, chart styles and
colour sets, chart templates — and editing any out-of-catalogue chart (3-D,
radar, stock, bubble, surface, chartex) beyond keeping its references right.
Pivot charts' data (SPEC-139 owns pivots; they render from caches here).
Highlighting a chart's source ranges in the grid, dragging those
highlights, and the "Recommended charts" gallery. Sparklines (queue backlog).
Charts in the library miniature. Printing charts (SPEC-136 paints through the
same `paintChart`). Editing the embedded workbook of a chart pasted from
another product.

## Reference

### Dependencies restated
- `specs/052-charts.md` / `specs/092-charts-on-slides.md` — `core/charts.ts`:
  `ChartType` (`columnClustered | columnStacked | barClustered | line |
  lineMarkers | area | pie | doughnut | scatter`), `ChartSeries { name;
  values: (number | null)[]; color: ModelColor }` (`null` = gap),
  `ChartOptions`, `ChartVisual { chartType; categories; series; options;
  unsupportedChartXml? }`, `CHART_CATALOGUE`, `SINGLE_SERIES_CHART_TYPES`,
  `chartAccentRef(i)`, `NEW_CHART_WIDTH_CM = 12`, `NEW_CHART_HEIGHT_CM = 7.5`;
  `core/chartCanvas.ts` `paintChartIntoCanvas(canvas, chart, w, h, theme)`;
  `core/chartPartOoxml.ts` `parseChartXml(xml): ParsedChartPart | null`,
  `chartPartXml(chart)`, `CHART_RELATIONSHIP`, `CHART_CONTENT_TYPE`;
  `core/chartPartOdf.ts` `parseOdfChartXml`, `odfChartContentXml(chart, w, h,
  theme?)`, `CHART_OBJECT_CONTENT_TYPE`; `core/chartFeatures.ts`
  `lostChartFeatures(xml, "ooxml" | "odf")`. SPEC-052's rule that only the
  data sheet and the options panel set `edited`, and SPEC-092's embedded-data
  answer (an edited part drops `c:externalData`; the bytes it pointed at stay).
- `specs/131-images-and-shapes-on-sheets.md` — `SheetDrawing { id; anchor;
  name; altText?; hidden?; rotationDeg?; sourceXml?; content }`, `Sheet.drawings`,
  `Sheet.drawingPart`, `drawingsDirty`, the anchor behaviours and adjuster,
  `drawingRectPx`, the layer (`SheetDrawingView.vue` gains a chart branch
  painting `paintChartIntoCanvas`), the MCE rules, the reachability rule for
  dropping parts, `useSheetDrawings`.
- `specs/120-…` — `evaluate`, `recalculate(model, changed: CellKey[]):
  RecalcResult`, the open policy (cached values trusted, open never marks
  dirty). `specs/119-…` — `parseFormula`, `ReferenceNode { sheet?; sheetEnd?;
  start; end? }`, `fromXlsxFormula`/`toXlsxFormula`, `renameSheetInFormula`.
  `specs/118-…` — `formatValue(value, code, ctx)`, date formats.
  `specs/124-…` — `StructuralEdit`, `adjustFormula(text, hostSheetId, edit,
  sheetNameOf)`, `REFERENCE_ADJUSTERS`, the "Retained references" rows for
  chart parts. `specs/121-…` — undo (`record(label, scope)`, `commit()`),
  `SheetRemap` + `remapSheetReferences`, `sourceSheetRemap`, the part-closure
  copy on duplicate, `stripRetention`, the conversion report.
  `specs/117-…` — `currentRegion`, hidden rows/columns, `model.theme`.
  `specs/125-…` — `MOVE_ADJUSTERS`, `MoveAdjuster { id; move(model, move) }`,
  `RangeMove`, `moveRange`, `repointFormula(text, hostSheetId, move,
  sheetNameOf)` (`move.ts`) and its `retained` move adjuster.

### The chart model
```ts
// src/lib/spreadsheet/model.ts (SPEC-132 widens SPEC-131's chart content)
export interface SheetChartSeriesRefs {
  name?: string;        // canonical single-cell reference; absent = the literal visual name
  categories?: string;  // canonical range (c:cat / c:xVal; ODS chart:categories or chart:domain)
  values?: string;      // canonical range (c:val / c:yVal; ODS chart:values-cell-range-address)
}
export interface SheetChartRefs {
  series: SheetChartSeriesRefs[];   // index-aligned with the part's series (and visual.series)
  plotArea?: string;                // ODS chart:plot-area@table:cell-range-address
  plotVisibleOnly: boolean;         // xlsx c:plotVisOnly (absent = true); ODS !chart:include-hidden-cells
}
export interface SheetChartContent {
  kind: "chart";
  visual: ChartVisual;
  refs?: SheetChartRefs;            // absent: a literal chart (data sheet editing, SPEC-052 style)
  sourcePart?: string;              // "xl/charts/chart1.xml" | "Object 1/content.xml" — a NAME into opaque
  edited?: true;                    // data sheet / options / colour / source change of shape → regenerate
  cachesDirty?: true;               // visual refreshed from cells → patch caches
  refsDirty?: true;                 // references changed → patch c:f / addresses
  readOnlyReason?: string;          // rendered from caches; never refreshed or edited
}
```
`ChartPartRefs` (core, wire text — the codec never interprets it):
`{ series: { name?: string; categories?: string; values?: string }[] }`;
`OdfChartRefs`: `{ plotArea?: string; categories?: string; series: { values?:
string; label?: string; domain?: string }[] }`.

### Read-only charts

| Condition | `readOnlyReason` |
| --- | --- |
| `c:pivotSource` present | "Pivot charts can't be edited yet" |
| a `c:multiLvlStrRef` category reference | "Charts with multi-level categories can't be edited yet" |
| out of catalogue | none — `unsupportedChartXml` and ChartToolbar's default text (a chartex chart is a SPEC-131 `foreign` drawing and has no chart toolbar) |

A read-only chart paints from its caches and is never refreshed, so it never
gains `cachesDirty`; its `refs` are still adjusted like any chart's, and
`refsDirty` then patches only its `c:f` texts (write path 2).

### Resolving references
`resolveChartRange(model, canonical, resolveName?: (name: string, sheetId:
string) => string | null): { sheetId: string; range: GridRange }[] | null`
parses with `parseFormula` and walks: a `ReferenceNode` (no `sheetEnd`), a
paren node around a union (`,`) of them, and a `name` node through
`resolveName` (returning the name's canonical reference text, resolved
again) — absent here, so a name yields `null` until SPEC-133 supplies a
hook built on its `resolveDefinedName` (the found name's `formula`). Unqualified references are relative to the
chart's own sheet. A union's areas are read in order, each row-major for a
row-shaped area and column-major for a column-shaped one; a multi-row,
multi-column area is read row-major.

### Refresh
`refreshSheetChart` rebuilds `visual.categories`, each series' `values` and
each referenced `name` from the cells (pie/doughnut keep all series, as the
model always does), compares with the current visual (deep) and returns the
same drawing when equal. Categories come from the first series that has a
`categories` reference; with none, they are `"1"`, `"2"`, … (Excel's
numbering). Series colours, options and `unsupportedChartXml` are never
touched by a refresh. The trigger set comes from SPEC-120's `RecalcResult`
plus the commit's directly edited cells.

### Writing
For each chart drawing, in this order of precedence:
1. **`edited`** → regenerate: `chartPartXml(visual, wire(refs))` (xlsx) or
   `odfChartContentXml(visual, wCm, hCm, theme, odfWire(refs))` (ods) at
   `sourcePart`, or a newly minted part for a chart created here; the part's
   `.rels` and its siblings stay as they were (SPEC-092's answer).
2. **`refsDirty` or `cachesDirty`** on a supported chart → patch the source
   part from `opaque` with `scanXml` offsets: series *i* is the *i*-th `c:ser`
   of the part's one type element; its `c:f` texts become `toXlsxFormula(ref)`;
   its caches are rewritten from the **cells** (not the visual): every cell of
   the reference in order, `c:ptCount` = the cell count, no `c:pt` for a blank
   cell, a `c:numCache` holding raw numbers, a `c:strCache` formatted text;
   ODS rewrites the five address attributes through
   `canonicalToOdsRangeAddress` and replaces the `local-table` with
   `odfChartTableXml(visual)`.
3. Otherwise the part's bytes pass through as the package holds them —
   which, for an unsupported or chartex chart, already carries SPEC-124's and
   SPEC-121's retained adjustments. An ODS chart frame's
   `draw:notify-on-update-of-ranges` is written from the chart's current
   references whenever the two differ.

### Insert from selection
`buildChartFromRange(model, sheetId, selection, chartType)`:
- **Source.** A multi-cell selection is the source; a single cell expands to
  SPEC-117's `currentRegion`. Several selected areas are accepted when they
  span the same rows (side-by-side columns, read left to right) or the same
  columns; otherwise toast "Select one block of data, or blocks that share the
  same rows".
- **Labels.** A cell *counts as a label* when it holds a string, or is empty
  in the top-left corner, or holds a number whose number format is a date or
  time format (SPEC-118). The first row is series names when every cell of it
  but the top-left counts as a label and some cell below it is a non-label
  number; the first column is categories when every cell of it below the
  name row counts as a label.
- **Orientation.** After removing the label row/column, series run down the
  columns when `rows ≥ columns`, across the rows otherwise (Excel's rule;
  ties → columns).
- **Scatter.** The first data column (or row) is the x values (`xVal`), each
  following one a `yVal` series; labels are names only.
- Series colours are `chartAccentRef(i)`; options `defaultChartOptions()`
  with `title` set from the single series' name when there is exactly one.

`chartRefsFromRange` is the same builder with explicit flags, used by the
Select Data dialog; `chartRangeOf(refs)` is its inverse, `null` when the
references do not tile one rectangle on one sheet.

### Select Data
Applying a range computes new `refs` with the dialog's flags. When the
series count, chart type and the presence of every `c:f` (name, categories,
values) per series are unchanged, the chart takes the patch path
(`refsDirty` + refresh); otherwise it is regenerated (`edited`) and the
rewrite notice shows.

### Reference adjustment
The `charts` adjuster covers every modelled chart drawing on every sheet:
`adjustFormula(ref, chartSheetId, edit, sheetNameOf)` per reference, then a
refresh (read-only charts only adjust). A reference adjusted to `#REF!` stays
`#REF!` and its series keeps its caches. Parts this spec does not model
(out-of-catalogue, chartex) are left to SPEC-124's `retained` rewrite, so
nothing is adjusted twice: its `xl/charts/chartN.xml` `c:f` row and its
`Object N/content.xml` row now apply only to those parts. The `charts`
`MoveAdjuster` applies `repointFormula(ref, chartSheetId, move, sheetNameOf)`
the same way (a reference entirely inside the moved block follows it; a
partial overlap stays; one inside the overwritten destination becomes
`#REF!`), and SPEC-125's `retained` move adjuster keeps covering the
unmodelled parts.

### ODS range addresses (`src/lib/spreadsheet/ods/rangeAddress.ts`, SPEC-121's, restated)
```ts
/** "Sheet1.$A$1:.$B$3", "$'Q1 Data'.A1:'Q1 Data'.B3", "Sheet1.A1:Sheet1.B3 Sheet1.D1:Sheet1.D3"
 *  → "Sheet1!$A$1:$B$3", "'Q1 Data'!A1:B3", "(Sheet1!A1:B3,Sheet1!D1:D3)"; null if unparseable. */
export function odsRangeAddressToCanonical(text: string): string | null;
/** The inverse; writes "$Sheet1.$A$1:.$B$3" (sheet on the start only, `$` as canonical has it),
 *  a union as a space-separated list. */
export function canonicalToOdsRangeAddress(canonical: string): string | null;
```
SPEC-121 created the module (its `remapOdsAddresses` is built on it); this
spec only calls it, and adds chart-shaped cases to its spec file. SPEC-133
and SPEC-135 use the same pair.

### New chart parts
xlsx, inside the anchor (golden-first — the `xdr:nvGraphicFramePr`
scaffold and Excel's zero `xdr:xfrm` are settled in the golden):
`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{n}"
name="Chart {k}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off
x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData
uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart
r:id="{rel}"/></a:graphicData></a:graphic></xdr:graphicFrame>`. Part names
`xl/charts/chart{n}.xml`, `n` = 1 + the highest across the package; no
`chartStyle`/`colors` siblings are generated. ods: `draw:frame` with SPEC-131's
anchor attributes holding `<draw:object draw:notify-on-update-of-ranges="…"
xlink:href="./Object {n}" xlink:type="simple" xlink:show="embed"
xlink:actuate="onLoad"/>`; manifest gains `Object {n}/` (media type
`application/vnd.oasis.opendocument.chart`) and its `content.xml` entry.

### Chart sheets
`xl/chartsheets/sheetN.xml` → `<drawing r:id>` → a drawing part holding one
`xdr:absoluteAnchor` graphicFrame → its chart part. The sheet keeps
`partName` (the chartsheet part, verbatim) and `drawingPart`; only the chart
part is ever rewritten, by the paths above. The view letterboxes the chart at
the anchor's `extEmu` aspect, centred, at most the visible area.

### View wiring
`src/views/spreadsheet/useSheetCharts.ts` (gallery arming, dialogs, refresh
trigger, rewrite notice; ≤ 400 lines), `src/components/spreadsheet/
SheetChartLayer.vue` (toolbar + dialogs, modelled on
`presentation/SlideChartLayer.vue`), `SheetChartDataDialog.vue`,
`SheetChartSheet.vue`; `SpreadsheetView.vue` gains one `useSheetCharts({ … })`
call and its hookup. The theme pass-through `asDocumentTheme` moves from
`documentModel.ts` into `core/theme.ts` (two-phase: `documentModel.ts`
re-exports it, so every text import is unchanged) and `ChartOptionsDialog.vue`
imports it from core — after which none of the four reused chart
components (`ChartGallery`, `ChartToolbar`, `ChartDataSheetDialog`,
`ChartOptionsDialog`) imports a product layer.

### Tests
- `charts.spec.ts`: each heuristic (label rule, date categories,
  orientation ties, scatter, multi-area), resolution (union, cross-sheet,
  unresolvable), refresh (gaps, `#N/A`, text as 0, hidden with and without
  `plotVisibleOnly`, identity when unchanged), every command, the adjuster's
  row-insert-above and insert-inside cases, rename and delete.
- `xlsx/chartXml.spec.ts`, `ods/chartXml.spec.ts`: reference reading on
  Excel- and LibreOffice-shaped parts, patches touching only `c:f`/caches/
  addresses (the rest `equalBytes`), goldens. `ods/rangeAddress.spec.ts`
  additions.
- `chartsRoundTrip.spec.ts` (`equalBytes`, `equalStructural`, byte-stable
  second write), core codec specs for the new parameters, `useSheetCharts.spec.ts`,
  `e2e/SPEC-132.spec.ts`.

### Specs this changes
- SPEC-116: `Sheet.kind` gains `"chart"`; a resolvable chartsheet is no
  longer a foreign placeholder.
- SPEC-125: `MOVE_ADJUSTERS` gains `charts`.
- SPEC-121: rename and delete remap chart references on the model side;
  chart sheets cannot be duplicated (as SPEC-121 already states);
  `UNCONVERTIBLE_CONTENT` loses its "charts" and "chart sheets" rows (catalogue
  charts cross through the model; dropped ones are reported item by item);
  chart content lives in SPEC-131's `Sheet.drawings`, already in
  `SHEET_BY_REFERENCE_FIELDS`, and every command here replaces the drawing
  objects copy-on-write, so the undo snapshot needs no new field.
- SPEC-124: `REFERENCE_ADJUSTERS` gains `charts`; the chart rows of
  "Retained references" narrow to unmodelled chart parts (gaining chartex
  `cx:f`) and the `notify-on-update-of-ranges` row is retired.
- SPEC-131: the `chart` content kind widens; `SheetDrawingView.vue` paints
  charts; cross-format copies keep catalogue charts.
- SPEC-052/SPEC-092: `chartPartXml`/`odfChartContentXml` gain an optional
  trailing parameter and `odfChartTableXml` is exported — no output changes;
  `ChartToolbar.vue` gains an optional `reason` prop (default text unchanged);
  `asDocumentTheme` moves to core behind a re-export.
- `specs/000-roadmap.md` gains this spec's row.
