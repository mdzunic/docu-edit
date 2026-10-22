---
id: SPEC-052
title: Charts — render, insert and edit data
status: draft
area: structure
depends_on: [SPEC-008, SPEC-022, SPEC-051]
---
## Why
SPEC-008 made a promise and deliberately underdelivered on it: a chart in an
opened document survives every save, and shows as a placeholder box. That was
the right first move — never destroy what you cannot draw — but it leaves the
most visible hole in the product. Open a quarterly report and there is a grey
rectangle where the revenue is.

This spec draws them, and lets people make new ones. The hard part is not the
drawing; it is not breaking the SPEC-008 guarantee while doing it. A chart
carries far more in its part than this product will ever model — themes, styles,
number formats, per-point overrides — so the rule here is that an untouched
chart is still written back byte-identical, and only a chart the user actually
edits is regenerated from the model, with an honest account of what that
rewrite cost.

## Acceptance criteria
- [ ] Charts in opened `.docx` and `.odt` documents render with their real type, series, categories, colours, legend and titles, replacing SPEC-008's placeholder for every type in "Chart catalogue" below
- [ ] "Insert → Chart" offers the catalogue types, and inserting places a `Chart` block (SPEC-051's `DrawingPlacement`) at the caret sized 12cm × 7.5cm with sample data
- [ ] Double-clicking a chart, or "Edit data" on its toolbar, opens a data sheet: an editable grid of categories and series where cell edits, row/column insert and delete, and pasted TSV/CSV all re-render the chart on commit
- [ ] The data sheet offers "First row as series names", "First column as categories" and "Switch rows/columns", each changing the chart without changing the sheet
- [ ] A chart options panel sets the title, axis titles, legend position (none/top/right/bottom/left), data labels on/off, gridlines, per-series colour, and axis minimum/maximum with an auto default
- [ ] With a table selected, "Insert → Chart from table" builds a chart from that table's cells and reports any cell it could not read as a number
- [ ] Charts are painted on the canvas layer with the DOM proxy, alt text and selection behaviour SPEC-051 defines — they are drawing objects, and get anchoring, wrap, z-order and grouping for free
- [ ] An opened chart that the user has not edited is written back **byte-identical**: `equalBytes` on its part, exactly as SPEC-008 requires today, and a test asserts this after editing an unrelated paragraph
- [ ] Editing a chart regenerates its part from the model, and everything the model does not carry is reported per "Rewriting a chart" below — never dropped silently
- [ ] A chart whose type is outside the catalogue keeps SPEC-008's behaviour exactly: placeholder, preserved, never rewritten, with editing disabled and the reason shown ("We can't edit 3-D charts yet — this one is kept exactly as it was")
- [ ] Deleting a chart removes its part, its relationship and its data on save; SPEC-008's part-preservation rule does not resurrect it
- [ ] Charts print and export to PDF (SPEC-034) rendered at the export resolution, not as an upscaled screen bitmap
- [ ] Charts round-trip through `.docx` (`a:graphic` → `c:chart` referencing a `word/charts/chartN.xml` part with `c:barChart`/`c:lineChart`/`c:pieChart`, and its cached `c:numCache`/`c:strCache`) and `.odt` (an embedded `application/vnd.oasis.opendocument.chart` object with `chart:chart`, `chart:plot-area` and its `table:table` data)
- [ ] A document with 20 charts opens and paints within SPEC-047's budget, and a data-sheet edit re-renders in under 100ms for a 12-series, 100-category chart
- [ ] Opening the fixtures below renders one chart of each catalogue type, one with negative values, one with a gap (empty cell), and one unsupported type as the placeholder

## Out of scope
3-D charts, combo charts, radar, surface, bubble, stock and treemap — all
preserved per SPEC-008 and never rewritten. Trendlines, error bars, secondary
axes and per-point formatting. Sparklines. Live links to an external
spreadsheet, or to a table in this document that updates when the table
changes — "Chart from table" is a one-time copy. Pivot charts. Inheriting the
document theme's colour scheme; the palette is this spec's own. Chart
templates.

## Reference

### Document model — the implemented one
`specs/002-create-blank-document.md` defines the model normatively and now
describes the shipped `src/lib/documentModel.ts` directly. **Read it before
starting.** The facts this spec rests on: `DocumentModel` holds
`sections: Section[]` whose blocks are
`Block = Paragraph | Table | Image | BreakBlock | TableOfContents`,
discriminated by a `type` string literal; lengths are centimetres and type
sizes points; and equality drops every key named `id`, so a reference that must
survive a round-trip test is never called `id`.

### Anchored drawing object
Defined normatively in `specs/051-drawing-shapes-and-text-boxes.md`, which
lifts `Image`'s placement fields into a `DrawingPlacement` interface. A chart is
a third block extending it:

```ts
// src/lib/documentModel.ts
export interface DrawingPlacement {          // SPEC-051
  widthCm: number; heightCm: number; wrap: ImageWrapMode; altText?: string
  offsetXCm?: number; offsetYCm?: number
  relativeToH?: "page" | "margin" | "column"
  relativeToV?: "page" | "margin" | "paragraph"
  zOrder?: number; rotationDeg?: number; groupOrdinal?: number
}

export interface Chart extends DrawingPlacement {
  type: "chart"
  chartType: ChartType                       // the closed set below
  categories: string[]
  series: { name: string; values: (number | null)[]; color: string }[]
  options: {
    title?: string
    categoryAxisTitle?: string
    valueAxisTitle?: string
    legend: "none" | "top" | "right" | "bottom" | "left"
    dataLabels: boolean
    gridlines: boolean
    /** Absent = auto. */
    valueAxisMin?: number
    valueAxisMax?: number
  }
  /** Part name of the preserved original in `DocumentModel.opaque`; absent for
   *  a chart created here. Deliberately not named `id`. */
  sourcePart?: string
  /** True once the data sheet or the options panel changed it. */
  edited?: boolean
  changes?: BlockTrackedChanges
}

export function isChart(block: Block | undefined): block is Chart {
  return block?.type === "chart";
}
```
Anchoring, wrap, z-order, grouping, hit testing, the canvas layer and the
focusable ARIA proxy all come from SPEC-051 unchanged. This spec adds no
positioning behaviour of its own.

**`sourcePart` is a name, not bytes.** The original chart part stays in
`DocumentModel.opaque` where SPEC-008 already puts it (`OpaquePart`,
`src/lib/opaque.ts`), and `Chart` references it by part name. Holding bytes in
the block would give the same part two owners and would break the model's
plain-JSON invariant.

`null` in a series' `values` is a gap — an empty cell — and is distinct from
`0`. Line charts break the line at a gap; bar and column charts draw nothing.

**`edited` is the switch the writer reads**, and only the data sheet and the
options panel set it. Moving, resizing, rotating, re-anchoring or re-ordering a
chart does not set it: none of those touch the chart part, so a chart dragged
across the page still round-trips byte-identical.

### Chart catalogue
The closed set this spec renders and can write. Everything else is preserved.

| Editor type | `ChartType` | `.docx` | `.odt` `chart:class` |
| --- | --- | --- | --- |
| Clustered column | `columnClustered` | `c:barChart` + `c:barDir val="col"`, `grouping="clustered"` | `chart:bar` (vertical) |
| Stacked column | `columnStacked` | same, `grouping="stacked"` | `chart:bar`, stacked |
| Clustered bar | `barClustered` | `c:barChart` + `barDir val="bar"` | `chart:bar` (horizontal) |
| Line | `line` | `c:lineChart`, `c:marker val="0"` | `chart:line` |
| Line with markers | `lineMarkers` | `c:lineChart`, `c:marker val="1"` | `chart:line` with symbols |
| Area | `area` | `c:areaChart` | `chart:area` |
| Pie | `pie` | `c:pieChart` | `chart:circle` |
| Doughnut | `doughnut` | `c:doughnutChart` | `chart:ring` |
| Scatter (XY) | `scatter` | `c:scatterChart` | `chart:scatter` |

Pie and doughnut render the first series only; the data sheet says so rather
than silently ignoring the rest.

### Preservation, and the table this reads from
Preservation is defined normatively in
`specs/008-round-trip-preservation.md` and implemented in `src/lib/opaque.ts`:
original package parts are retained in `DocumentModel.opaque` as `OpaquePart`,
written back byte-identical on a same-format save, and both writers return

```ts
// src/lib/opaque.ts
export type DropKind = "chart" | "contentControl" | "customXml" | "embeddedObject"
  | "macro" | "opaqueBlock" | "opaqueInline" | "opaquePart"
  | "runShading" | "svgImage" | "trackedPropertyChange";
export interface DropRecord { kind: DropKind; location: string; reason: string }
export interface WriteResult { bytes: Uint8Array; dropped: DropRecord[] }
```
`DropKind` already carries `"chart"`, which today means "a chart we could not
convert". That part rule is what makes an untouched chart survive, and this
spec must not weaken it.

"Chart from table" reads the table node defined in
`specs/022-tables-insert-and-edit.md` — rows of cells whose content is
`Block[]`. It reads the cells' text and parses numbers from it; it holds no
reference to the table afterwards.

### Rewriting a chart
The chart part carries far more than `Chart` models. When `edited` is `true`
the writer regenerates the part, and that regeneration loses whatever was in
the source that the model has no field for.

SPEC-008 says `dropped` is empty for a same-format save, because opaque content
round-trips untouched — a rule `src/lib/opaque.ts` records in a comment on
`WriteResult`. **This spec makes the one exception**, and it is narrower than
it sounds: a same-format save may now return `DropRecord`s, but only for charts
whose `edited` flag is `true`. An untouched document still returns an empty
`dropped`. `DropKind` widens by one member, as SPEC-013, SPEC-024 and SPEC-039
have already widened it:

```ts
| "chartFeature"    // a property of an edited chart the model cannot carry
```

with `location` naming the chart ("Chart 2, page 4") and `reason` naming the
property ("custom number format on the value axis"). The conversion report
SPEC-008 already builds renders these with no change; its heading text gains a
same-format variant, since "Conversion report" is wrong for a `.docx` → `.docx`
save. The user sees this once, when they edit a chart that came from Word,
which is exactly when they should.

The alternative — regenerating on every save — would silently strip styling
from charts nobody touched, and is what makes people stop trusting an editor.

### Existing I/O surface
```ts
// src/lib/docx/reader.ts   parseDocx(data): Promise<ParsedDocxDocument>   — SPEC-004
// src/lib/odf/reader.ts    parseOdt(data):  Promise<ParsedOdtDocument>    — SPEC-005
// src/lib/docx/writer.ts   writeDocx(…): Promise<WriteResult>            — SPEC-006/008
// src/lib/odf/writer.ts    the .odt writer, same WriteResult             — SPEC-007/008
// src/lib/openDocument.ts  openDocumentFile(file): Promise<DocumentModel>
// src/lib/opaque.ts        WriteResult = { bytes: Uint8Array; dropped: DropRecord[] }
```
Per `CLAUDE.md`, chart parsing and serialisation are plain TypeScript with no
Vue imports and are unit-testable without mounting a component; the renderer is
the only part that touches the canvas. Reading a chart means reading a second
part (`.docx`) or an embedded object directory (`.odt`) and resolving the
relationship from the `c:chart`/`draw:object` reference — the first time either
reader follows a reference out of the main document part, which is why this
spec depends on SPEC-008 rather than only on SPEC-051.

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalBytes`,
`canonicalizeXml` and `firstStructuralDifference`. SPEC-002 also defines
`equalSemantic`, **which is not implemented**; this spec does not need it.
- **`equalBytes`** — on the chart part, before and after a save, for a chart
  that was not edited. This is the SPEC-008 guarantee and this spec's most
  important test. It takes `Uint8Array`, which is what `WriteResult.bytes`
  already is.
- **`equalStructural`** — for the same-format round trip of a chart *created*
  in the editor, where there is no source part to preserve and the model is
  the only truth.
- For cross-format saves, assert `equalStructural` between the source model
  masked by the writer's `dropped` records and the re-opened model. Apply the
  mask as a transform over `dropped`, never as a hand-rolled subset comparison.

Series colours are compared as the same lowercase `#rrggbb` strings the rest of
the model uses, so a reader that resolves a theme colour to RGB on the way in
must resolve it identically on the way back or the round trip fails — which is
the intent.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). SPEC-008
already commits `fixtures/docx/word-charts-and-controls.docx`, containing a
chart; this spec reuses it for the byte-identical preservation test rather than
adding a second one. It adds `fixtures/docx/word-chart-types.docx` and
`fixtures/odt/lo-chart-types.odt` — one chart of each catalogue type, one with
negative values, one with an empty cell, one 3-D chart as the unsupported case —
with manifest entries naming each chart's type, series names, category count
and first-series values.
