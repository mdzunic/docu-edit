---
id: SPEC-092
title: Charts on slides — render, insert and edit data
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-052, SPEC-079, SPEC-080, SPEC-081, SPEC-082]
---
## Why
A deck is where a chart is most often the point of the slide, yet SPEC-079
paints it as a grey box labelled "Chart" (pptx) or "Object" (odp): open a
quarterly deck and the revenue is a rectangle.

This is a generalisation exercise, not a new chart engine. SPEC-052 shipped one
and its product-neutral half already sits in `src/lib/core` (`charts.ts`,
`chartCanvas.ts`, `chartIcons.ts`); what is still stranded in the text product
is the *chart part codec* — and a `.pptx` chart part is the identical
`c:chartSpace` schema a `.docx` one is, while an `.odp` embedded chart is the
identical ODF chart document an `.odt` one is. So: hoist those two codecs into
core in the repo's two-phase move, add a `ChartFrame`, and mount SPEC-052's
gallery, data sheet and options dialogs — all four already import nothing but
`lib/core` and UI primitives — on the slide stage. The rule that outranks the
feature stays SPEC-008's: a chart nobody edited is written back byte-identical,
one outside the catalogue is never rewritten. **This file is the normative
source for the `ChartFrame` slide frame, the slide-side chart-part resolution
and serialisation rules in both formats, and the embedded-workbook policy for a
regenerated chart part.**

## Acceptance criteria
- [ ] Two-phase core move, behaviour-preserving: phase one lifts `src/lib/docx/chart.ts`'s `parseChartXml`/`chartPartXml`/`ParsedChartPart`/`C_NAMESPACE`/`CHART_RELATIONSHIP`/`CHART_CONTENT_TYPE` into `src/lib/core/chartPartOoxml.ts` and `src/lib/odf/chart.ts`'s `parseOdfChartXml`/`odfChartContentXml`/`ParsedOdfChartPart`/`CHART_NAMESPACE`/`CHART_OBJECT_CONTENT_TYPE` into `src/lib/core/chartPartOdf.ts` — `cmToOdfLength`, the one `src/lib/odf/units.ts` helper `odfChartContentXml` reaches, moving in the same phase into `src/lib/core/odfXml.ts` beside `odfLengthToCm` and re-exported from `odf/units.ts` as the read direction already is, so no core module ends up importing `../odf/*` — all re-exported from their old homes so `docx/chart.spec.ts`, `odf/chart.spec.ts`, `docx/chartRoundTrip.spec.ts` and `odf/chartRoundTrip.spec.ts` pin identical behaviour (`odf/chart.spec.ts`'s calls gaining AC-2's arity, every value they assert unchanged); phase two has the presentation readers/writers import core directly, with `core/boundary.spec.ts` and `presentation/boundary.spec.ts` both green
- [ ] The only signature delta in that move: the builders take `ChartVisual` (`core/charts.ts`) instead of `Chart`, and `odfChartContentXml(chart, widthCm, heightCm)` takes geometry explicitly (it read `Chart.widthCm`/`heightCm`, which `ChartVisual` does not carry); the container codecs — `docx/chart.ts`'s `parseChartDrawing`/`chartDrawingXml`, `odf/chart.ts`'s `parseChartFrame`/`chartFrameXml` and both `chartWritePlan.ts` modules — stay put, being `w:drawing`/`text:anchor-type` shaped, and no `src/lib/core` module gains a non-core import
- [ ] `SlideFrame` gains `ChartFrame` per the Reference: `kind: "chart"`, the `ChartVisual` fields (`chartType`, `categories`, `series`, `options`, optional `unsupportedChartXml`), `sourcePart?`, `edited?: true` and `xml?` (the verbatim source container element, dropped by any geometry or data change); `slideFrameKinds` and `stripRetention` handle it, and `JSON.parse(JSON.stringify(model))` still `toEqual`s the model
- [ ] pptx read: a `p:graphicFrame` whose `a:graphic/a:graphicData@uri` is `http://schemas.openxmlformats.org/drawingml/2006/chart` holding `<c:chart r:id="rIdN"/>` resolves rIdN through the slide's rels (`resolvePartPath`, `core/xml.ts`) to `ppt/charts/chartN.xml` and parses via `parseChartXml` → `ChartFrame` with `sourcePart` set; `SlideContext` gains `parts: Map<string, Uint8Array>` for the lookup; an out-of-catalogue part (3-D, combo, radar/stock/bubble/surface, or the `chartex` uri) sets `unsupportedChartXml` to the part text; an unresolvable relationship keeps today's `ForeignFrame` labelled "Chart" unchanged
- [ ] odp read: a `draw:frame` holding `draw:object xlink:href="./Object 1"` resolves `Object 1/content.xml` out of `OdpContext.parts` and parses `office:document-content/office:body/office:chart/chart:chart` (with its `chart:plot-area`, `chart:series` and `table:table` data cache) via `parseOdfChartXml` → `ChartFrame`, `sourcePart` = that path; this replaces the reader's current `foreignAt(frame, base, "Object")` fall-through; any `ObjectReplacements/…` preview image is ignored for rendering and preserved as an untouched part
- [ ] Rendering: `SlideFrameView.vue` gains a `chart-frame` branch painting `paintChart` (`core/chartCanvas.ts`) into a devicePixelRatio-crisp `<canvas>` exactly as the shape branch uses `paintShapeIntoCanvas`; `slideRender.ts` gains `chartVisualOf(frame): ChartVisual | null` beside `shapeVisualOf`; `frameAccessibleName` → "Chart", `frameNoun` → "chart"; the stage, filmstrip thumbnails and `PresentationPlayView` all paint through the same component; an `unsupportedChartXml` chart paints `paintChart`'s dashed "Chart not supported" placeholder at true bounds
- [ ] Insert: `SlideInsertBar.vue` gains a "Chart" button (`insert-chart`) opening `chart-gallery-popover` around the existing `src/components/ChartGallery.vue` unchanged (testids `chart-gallery`, `chart-gallery-<chartType>`); picking a type inserts a `ChartFrame` centred on the slide at `NEW_CHART_WIDTH_CM` × `NEW_CHART_HEIGHT_CM` (12 × 7.5cm) with `sampleChartCategories()`/`sampleChartSeries()`/`defaultChartOptions()`, selects it, and fits the 720px e2e viewport via SPEC-081's overflow rule
- [ ] Edit: double-clicking a chart frame, or `chart-edit-data-button` / `chart-options-button` on a `slide-chart-toolbar` hosting `src/components/ChartToolbar.vue` unchanged, opens `ChartDataSheetDialog.vue` / `ChartOptionsDialog.vue` unchanged — every SPEC-052 testid reused verbatim (`chart-data-grid`, `chart-data-cell-<r>-<c>`, `chart-data-add-row`, `chart-data-add-col`, `chart-data-switch-rows-columns`, `chart-data-apply`, `chart-options-title`, `chart-options-legend`, `chart-options-data-labels`, `chart-options-gridlines`, `chart-options-value-min`, `chart-options-value-max`, `chart-options-series-color-<i>`, `chart-options-apply`); an `unsupportedChartXml` chart shows `chart-unsupported-reason` and both buttons are disabled
- [ ] Commands live in `src/lib/presentation/charts.ts` (framework-free, core-only imports): `insertChartFrame(model, slideIndex, chartType, rect): string | null`, `setSlideChartData`, `setSlideChartOptions`, `setSlideChartSeriesColor` — each sets `edited = true`, deletes `xml` and ends in `markSlideEdited(slide, format)`; SPEC-082's `moveFrame`/`resizeFrame`/`reorderFrame` delete `xml` but must **never** set `edited` (they do not touch the chart part), so a chart dragged across a slide still round-trips its part byte-identical
- [ ] Undo/persist per SPEC-080: one `record(label)` → mutate → `commit()` per gesture with labels in the existing voice — "Insert chart", "Edit chart data", "Chart options", "Chart series colour" — write-through persisted, and `useSlideCharts` exposes fresh-array computeds per the `useSlideRender` reactivity rule
- [ ] pptx write: an untouched chart re-emits `xml` verbatim; a touched one emits the canonical `p:graphicFrame` of the Reference, reusing `sourcePart` when present or minting `ppt/charts/chart{n}.xml` past the highest existing, adding the `…/relationships/chart` relationship to the rebuilt slide `.rels` and a `CHART_CONTENT_TYPE` Override to `[Content_Types].xml` for a genuinely new part; `edited` substitutes the regenerated part text at that name, an untouched part's bytes pass straight through; a deleted chart's `ppt/charts/*` part and its `.rels` are excluded from the pass-through (`planChartWrite`'s `orphanedPartNames` rule, restated for the deck writer)
- [ ] odp write: an untouched chart re-emits `xml` verbatim; a touched one emits the slide-flavoured `draw:frame` + `draw:object` of the Reference (`svg:x/y/width/height`, **not** `text:anchor-type` — that is the text product's container), reusing `sourcePart`'s directory or minting `Object {n}`, with `META-INF/manifest.xml` gaining the `Object N/` entry (media type `application/vnd.oasis.opendocument.chart`) and its `Object N/content.xml` entry (`text/xml`)
- [ ] Embedded workbook, decided and reported: an untouched pptx chart keeps `c:externalData`, `ppt/charts/_rels/chartN.xml.rels` and `ppt/embeddings/*.xlsx` byte-identical; an **edited** chart's regenerated part omits `c:externalData` — the same answer SPEC-052 already ships for `.docx`, where `chartPartXml` emits none — while the workbook bytes and the chart part's `.rels` stay in the package untouched; the first edit of an imported chart shows a one-line `chart-rewrite-notice` naming what the rewrite cost, built from `core/chartFeatures.ts` plus the workbook line per the Reference
- [ ] Round trip: a no-edit download stays byte-identical — `equalBytes` (`src/lib/core/equality.ts`) per part after `unzip(...)` for every part including `ppt/charts/*`, `ppt/embeddings/*`, `Object N/*` and `ObjectReplacements/*`, extending `roundTrip.spec.ts`'s existing `expectUntouched` sweep; editing one chart leaves every other part `equalBytes`, other charts' parts and workbooks included; an inserted or edited chart save → reopen reproduces its `ChartFrame` under `equalStructural` (`src/lib/equality.ts`) in both formats, geometry within ±0.01cm, series colours as the same lowercase `#rrggbb`; writing twice, and writing the reopened model, produce identical slide and chart part bytes
- [ ] Fixture coverage: `src/lib/presentation/fixtures/build.ts` gains `buildPptxChartsFixtureBytes()`/`buildOdpChartsFixtureBytes()` and `EXPECTED_FIXTURE_CHARTS` appended under a banner comment (the four mini-deck slides untouched, the builders DOM-free — the hoisted part writers are pure string builders), each deck carrying a clustered column, a line, a pie **and** one deliberately out-of-catalogue chart (`c:bar3DChart` / `chart:class="chart:surface"`) to pin the preserve-don't-parse path; `scripts/build-presentation-fixtures.ts` gains writes for `e2e/fixtures/spec092-charts.pptx` and `.odp`
- [ ] e2e (`e2e/SPEC-092.spec.ts`): open `spec092-charts.pptx` → slide 1 shows a painted chart, not a grey box; the 3-D chart shows the dashed placeholder with its edit buttons disabled → select the column chart → Edit data → change a cell → Done → the canvas repaints → Chart options → legend "bottom", data labels on → Apply → Insert → Chart → Pie → a new chart lands selected → download, reopen → both charts present with the edited values, the options applied and the placeholder chart untouched

## Out of scope
3-D, combo, radar, surface, bubble, stock and treemap charts — preserved per
SPEC-008, never rewritten, never editable. Trendlines, error bars, secondary
axes, per-point formatting and custom number formats (reported, not modelled).
Rewriting or generating the embedded `.xlsx` workbook, or any live link to it.
Chart-from-table on a slide — SPEC-091 models the grid; nothing here converts
one into a chart. Animating a chart by series or category (SPEC-086 animates the
frame as a whole). Theme-aware palettes, chart templates, fill/stroke styling
(SPEC-082's popover offers neither), rotating a chart.

## Reference

### Dependencies restated
- `specs/052-charts.md` — the chart model and its preservation rule. Its
  product-neutral half is already `src/lib/core/charts.ts`: `ChartType =
  "columnClustered"|"columnStacked"|"barClustered"|"line"|"lineMarkers"|
  "area"|"pie"|"doughnut"|"scatter"`; `ChartSeries { name: string; values:
  (number|null)[]; color: string }` (`null` is a gap, distinct from `0`);
  `ChartOptions { title?; categoryAxisTitle?; valueAxisTitle?: string; legend:
  "none"|"top"|"right"|"bottom"|"left"; dataLabels: boolean; gridlines:
  boolean; valueAxisMin?; valueAxisMax?: number }`; `ChartVisual { chartType;
  categories: string[]; series: ChartSeries[]; options: ChartOptions;
  unsupportedChartXml?: string }`; plus `CHART_CATALOGUE`,
  `SINGLE_SERIES_CHART_TYPES` (pie/doughnut persist series 1 only),
  `DEFAULT_CHART_COLORS`, `NEW_CHART_WIDTH_CM = 12`, `NEW_CHART_HEIGHT_CM =
  7.5`, `MIN_CHART_SIZE_CM = 1`, `sampleChartCategories()`/
  `sampleChartSeries()`/`defaultChartOptions()`, `core/chartCanvas.ts`'s
  `paintChart(ctx, chart: ChartVisual, widthPx, heightPx, chrome?)` and
  `core/chartIcons.ts`'s `chartIconMarks`. SPEC-052's `edited` switch and its
  "only the data sheet and the options panel set it" rule carry over.
- `specs/079-slide-canvas-rendering.md` — the slide model (`SlideFrame` union,
  `SlideFrameBase { id; xCm; yCm; widthCm; heightCm; rotationDeg? }`, `id` minted
  at read and never serialised, cm lengths, pt sizes, lowercase `#rrggbb`
  colours) and the `ForeignFrame { xml; label }` behaviour this supersedes.
- `specs/080-create-presentation-and-slide-management.md` — the writers' rebuild
  rules (`writePptx`: package verbatim when nothing is dirty, else
  `buildTextEditSubstitution`/`buildSubstitution`; `writeOdp`: verbatim unless
  `structureDirty` or any `slide.dirty`) and the undo stack (`record(label)`
  **before** the mutation, then `commit()`).
- `specs/081-slide-text-editing.md` — `serializeSlideXml(slide, rels,
  slidePartPath)` / `serializePage(slide, registry, usedNames, master,
  styleSource)`, the `modelled ?? preserved ?? ""` precedence, the wire-id
  minting (`p:cNvPr id` 2, 3, … in z-order; odp `de-f{n}` on `xml:id` +
  `draw:id`), and `markSlideEdited(slide, format)` (`dirty = true`; odp also
  deletes `slide.sourceXml`).
- `specs/082-slide-objects-images-shapes-text-boxes.md` — frame selection/move/
  resize/z-order/duplicate/delete (`src/lib/presentation/frames.ts`,
  `FrameRect`, `MIN_FRAME_SIZE_CM`), `frameAccessibleName`/`frameNoun`
  (`src/views/presentation/useFrameManipulation.ts`), the `SlideInsertBar`
  popover pattern.

### The model (extends the union SPEC-079 defines)
```ts
// src/lib/presentation/model.ts
export interface ChartFrame extends SlideFrameBase, ChartVisual {
  kind: "chart";
  /** "ppt/charts/chart1.xml" | "Object 1/content.xml" — a NAME into
   *  model.opaque, never bytes; absent for a chart created here. Not
   *  called `id`, which equalStructural's IGNORED_KEYS would drop. */
  sourcePart?: string;
  edited?: true;          // set only by the data sheet / options panel
  /** Verbatim source p:graphicFrame / draw:frame, re-emitted while the frame
   *  is untouched; deleted by any geometry or data change, the serialiser then
   *  writing the canonical container below. In XML_STRING_KEYS, so it
   *  compares canonically under equalStructural. */
  xml?: string;
}
export type SlideFrame = TextFrame | ImageFrame | ShapeFrame
                       | ChartFrame | GroupFrame | ForeignFrame;
```
Regenerating the container costs the source `p:cNvPr` name and any `p:extLst` on
it — the price of moving a chart, and why `xml` survives every other operation.
After `duplicateSlide`'s verbatim-bytes path (SPEC-080) two slides may share one
`sourcePart` — legal in both formats; the write planner makes the first claimant
in deck order that part's owner and regenerates a fresh part for each later one.

### Canonical containers (write; read tolerantly)
pptx, inside `p:spTree` in z-order, consuming one minted `p:cNvPr id`:
```xml
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{n}" name="Chart {k}"/>
  <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
  <p:nvPr/></p:nvGraphicFramePr>
 <p:xfrm><a:off x="{emu}" y="{emu}"/><a:ext cx="{emu}" cy="{emu}"/></p:xfrm>
 <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
   <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
     r:id="{relId}"/></a:graphicData></a:graphic></p:graphicFrame>
```
odp, inside `draw:page`, one minted `de-f{n}`: a `draw:frame` with
`draw:id`/`xml:id` `de-f{n}`, `svg:x`/`svg:y`/`svg:width`/`svg:height` in cm
and `presentation:class="chart"`, holding `<draw:object xlink:href="./Object
{n}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` alone.

Part bodies are the hoisted codecs' output unchanged, over SPEC-052's
catalogue mapping: `c:chartSpace/c:chart/c:plotArea` holding exactly one type
element (`c:barChart` + `c:barDir`/`c:grouping`, `c:lineChart` + `c:marker`,
`c:areaChart`, `c:pieChart`, `c:doughnutChart`, `c:scatterChart`) whose
`c:ser` carry `c:tx`, `c:cat`/`c:xVal` and `c:val`/`c:yVal` caches; and
`chart:chart` with `chart:class`, `chart:plot-area`, `chart:series` and a
`table:table` data cache. Reading tolerates any attribute order and unknown
children; a second type element means a combo chart → unsupported.

**Golden-first (normative process).** Two details are **not certain and must
not be invented**: the exact `p:nvGraphicFramePr` scaffold PowerPoint requires
and whether it accepts a chart part with no sibling `colors1.xml`/`style1.xml`
and no `c:externalData`; and whether Impress needs `Object N/styles.xml` and
`meta.xml`, an `ObjectReplacements/Object N` preview, or
`presentation:class="chart"`. So follow SPEC-086's process exactly:
hand-author the canonical bytes into `pptx/chartGolden.ts`
(`CHART_GRAPHIC_FRAME_GOLDEN`, `CHART_PART_GOLDEN`, `GOLDEN_CHART_FRAME`) and
`odp/chartGolden.ts` (`CHART_DRAW_FRAME_GOLDEN`) against ECMA-376
§19.3.1.21/§21.2 and ODF 1.2 §10.4.6/§16; assert serialiser output equals them
and that the readers accept them plus PowerPoint- and Impress-flavoured
variants with reordered attributes. Cross-app checks are manual QA in the PR.

### The embedded workbook, and what the rewrite costs
A real PowerPoint chart part carries `<c:externalData r:id="rIdX"><c:autoUpdate
val="0"/></c:externalData>` pointing through `ppt/charts/_rels/chartN.xml.rels`
at an `ppt/embeddings/*.xlsx` workbook part. This repo has no spreadsheet writer and
`package.json` is protected, so it cannot be regenerated. The decision,
following SPEC-052:
- **Untouched chart** — everything survives verbatim, `equalBytes`.
- **Edited chart** — its regenerated part **drops the `c:externalData`
  reference**, while the workbook bytes and the chart part's `.rels` stay untouched
  in the package (deleting bytes we did not author is what CLAUDE.md's first law
  forbids; an orphaned workbook costs a few KB and no correctness). PowerPoint still
  renders from the `c:cat`/`c:val` caches; the cost is that "Edit Data" no longer
  opens the original workbook, offering a sheet seeded from the caches.
- **Rejected**: keeping `c:externalData` pointed at the stale workbook — Edit
  Data would reload from stale numbers and silently discard the user's edit.

`chart-rewrite-notice`'s text comes from `presentation/charts.ts`'s
`lostChartFeatures(sourceXml, "ooxml"|"odf"): string[]` — phase one of a second
two-phase move, lifting `src/lib/opaque.ts`'s `DOCX_CHART_FEATURE_MARKERS`/
`ODT_CHART_FEATURE_MARKERS` (trendline, error bars, per-point formatting,
custom number format, secondary axis, gradient/pattern/picture fill) unchanged
into `core/chartFeatures.ts` with `opaque.ts` re-importing them, so
`collectEditedChartDropRecords` and every `chartFeature` expectation is
unchanged — coverage preserved, never weakened. The notice adds its own
workbook line when the source part had `c:externalData`.

### View wiring
`src/views/presentation/useSlideCharts.ts` owns gallery arming, dialog state,
the rewrite notice and the record → mutate → `commit()` sequencing, over
`{ model, currentSlide, currentIndex, record, commit }` plus `{ selectedFrameId,
selectFrame, refresh, addToast }`; `components/presentation/SlideChartToolbar.vue`
hosts the reused `ChartToolbar.vue`, and gallery and dialogs mount unchanged
from `src/components/`. `PresentationView.vue` gains exactly one
`useSlideCharts({ … })` call plus its hookup, paid for by extracting the deck header
into `presentation/DeckHeader.vue` — SPEC-090 owns that extraction, so an equivalent
one pays where it has landed — then `LINE_BUDGET` ratchets down with a ledger row.

### Tests
- `src/lib/core/chartPartOoxml.spec.ts`, `chartPartOdf.spec.ts` and
  `chartFeatures.spec.ts` cover the moved codecs on the `ChartVisual` surface,
  while `docx/chart.spec.ts`, `odf/chart.spec.ts`, both `chartRoundTrip.spec.ts`
  and `opaque.spec.ts` keep every asserted value through the re-exports, only
  `odf/chart.spec.ts`'s eleven `odfChartContentXml` calls changing, for AC-2's
  arity — coverage preserved, never weakened; both `boundary.spec.ts` stay green.
- `src/lib/presentation/charts.spec.ts` — the commands, `edited` set only by
  data+options, `xml` dropped on move, `markSlideEdited` called, the
  duplicate-slide part rule, and `lostChartFeatures` text.
- `src/lib/presentation/pptx/serialize.spec.ts` and `odp/serialize.spec.ts` —
  golden equality, minted ids, rels/Content-Types/manifest splices, orphans.
- `src/lib/presentation/chartsRoundTrip.spec.ts` (`equalBytes` per part,
  `equalStructural` per frame, byte stability on a second write) and
  `chartsFixtures.spec.ts`; `roundTrip.spec.ts` keeps its assertions and gains
  the chart parts in its `expectUntouched` sweep. Then
  `useSlideCharts.spec.ts`, `PresentationView.size.spec.ts`, `e2e/SPEC-092.spec.ts`.

### Specs this changes
- SPEC-079's `ForeignFrame` "Chart" label now applies only to a chart whose
  part cannot be resolved, and the odp reader's `draw:object` → "Object"
  fall-through is superseded.
- SPEC-052's four chart-part codecs and its chart-feature marker tables move into
  `src/lib/core` behind re-exports; behaviour, testids and every asserted
  text-side value are unchanged, only `odf/chart.spec.ts`'s calls taking AC-2's
  arity.
- SPEC-082's frame-kind lists gain `"chart"` and its style popover offers no
  fill/stroke for one; `unconvertibleFrameCount`'s rule stays SPEC-080's word
  for word — every non-text frame counts — so a chart frame counts as
  unconvertible and a cross-format copy drops it, `unsupportedChartXml` or not.
- SPEC-086's animatable frame kinds gain `"chart"`; SPEC-088 exports chart
  frames through `core/chartCanvas.ts`, not a placeholder box; and
  `specs/000-roadmap.md` gains this spec's row in the Phase 13 table.
