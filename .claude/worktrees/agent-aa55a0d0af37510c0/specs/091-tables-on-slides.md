---
id: SPEC-091
title: Tables on slides — grid model, cell editing and both wire formats
status: draft
area: structure
depends_on: [SPEC-023, SPEC-079, SPEC-081, SPEC-082, SPEC-086]
---
## Why
Every real deck has a table on it, and today every one of them is a grey
box: `parseGraphicFrame` in `src/lib/presentation/pptx/reader.ts` sees
`a:tbl` and returns `foreignFrame(gf, base, "Table")`, and
`parseDrawFrame` in `odp/reader.ts` does the same for `table:table`. The
bytes survive, but nobody can fix a typo in a cell.

`src/lib/core/tables.ts` already anticipates this — "a Slides product's
tables carry the same border and range semantics; only cell *content*
(paragraph blocks) is product-specific" — so this spec consumes those
types, moves the rest of the neutral arithmetic down beside them, and
adds a `TableFrame` whose cells hold `SlideParagraph[]`. **It is the
normative source for the slide table model, its `a:tbl`/`table:table`
wire mappings, and the resolve-never-author style rule.**

## Acceptance criteria
- [ ] `SlideFrame` gains `TableFrame` per the Reference: `kind: "table"`, `rows: SlideTableRow[]`, `columnWidthsCm: number[]`, optional `firstRow`/`bandRow` look flags and `styleId`; every row holds exactly `columnWidthsCm.length` cells, each `SlideTableCell` holding `paragraphs: SlideParagraph[]` plus optional `colSpan`/`rowSpan`/`covered`/`fill`/`borders`/`paddingCm` (default 0.13)/`verticalAlign` (default `"top"`) — covered slots carry `covered: true` and `paragraphs: []` and are never an editing or navigation target
- [ ] pptx read: a `p:graphicFrame` whose `a:graphicData@uri` is the DrawingML table uri and whose `a:tbl` is in the Reference's subset becomes a `TableFrame` (`a:tblGrid/a:gridCol@w` EMU → `columnWidthsCm`, `a:tr@h` → `heightCm`, `a:tc` + `a:gridSpan`/`a:rowSpan`/`a:hMerge`/`a:vMerge` → the regular grid, `a:tcPr` → fill/borders/padding/`anchor`, `a:txBody` through SPEC-081's run reader); anything out of subset stays a `ForeignFrame` `label: "Table"` with its `xml` verbatim — tables never half-parse
- [ ] odp read: a `draw:frame` containing an in-subset `table:table` becomes a `TableFrame` (`table:table-column` + `table:number-columns-repeated` → columns via the automatic style's `style:column-width`, `table:table-row` → `style:row-height`, `table:table-cell` with `table:number-columns-spanned`/`table:number-rows-spanned`, `table:covered-table-cell` → covered slots, `table:table-header-rows` → `isHeader`); out of subset stays a `ForeignFrame` `"Table"`
- [ ] `SlideStage` renders a `TableFrame` as a real `<table>` (`data-testid="slide-table"`, cells `slide-table-cell`, `<colgroup>` from `columnWidthsCm`, covered slots emitting nothing while their owner carries `colspan`/`rowspan`), with the effective look resolved per "Table styles": explicit cell values win, then the style's header/banding look, then the neutral built-in fallback
- [ ] Insert → Table (`insert-table` on `SlideInsertBar`) opens a `PICKER_MAX_ROWS`×`PICKER_MAX_COLS` (10×10) grid picker (`table-grid-picker`) plus rows/columns inputs bounded by `MAX_TABLE_ROWS` (200) / `MAX_TABLE_COLS` (63) — all four from `src/lib/core/tables.ts`; the table lands centred at 80% of slide width with equal columns, 0.9cm rows and `firstRow: true`, selected with the caret in cell (0,0)
- [ ] A cell is edited by SPEC-081's machinery unchanged: double-click or Enter starts a session inside the cell, `applyRunFormat`/`applyParagraphFormat` and the whole `slides-*` toolbar act on that cell's `SlideParagraph[]`, Tab/Shift+Tab move to the next/previous non-covered cell (Tab in the last cell appends a row), and arrow keys leave the table at its edges
- [ ] Row and column commands in `src/lib/presentation/tables.ts` — testids `table-insert-row-above`, `table-insert-row-below`, `table-insert-column-left`, `table-insert-column-right`, `table-delete-row`, `table-delete-column`, `table-delete-table`; an inserted column takes its neighbour's width and the frame's `widthCm` grows to the new `columnWidthsCm` sum; deleting the last row or column deletes the frame
- [ ] Sizing: dragging a column boundary resizes it live (floor `MIN_COLUMN_WIDTH_CM` 0.5cm) and dragging a row boundary sets `heightCm` (floor `MIN_ROW_HEIGHT_CM` 0.4cm — a minimum, not a fixed height); resizing the *frame* through SPEC-082's handles scales `columnWidthsCm` and every `heightCm` proportionally so the grid keeps its relative proportions
- [ ] Merge and split: a rectangular range (`CellRange`, `cellRangeBounds`, `isCellInRange` from `src/lib/core/tables.ts`) merges into its top-left cell — spans set, content concatenated, every other slot becoming `covered: true` — and `table-split-cell` repartitions a spanning cell; both resolve through `computeOwnerGrid`, so a click anywhere in a merged rectangle addresses one cell and the grid stays regular R×C after every command
- [ ] The Table panel (`slides-table-toggle` → `slides-table-panel`, open state persisted as `presentationTableOpen` in `src/lib/preferences.ts` beside the other three keys, default `false`) edits the selected cells: fill (`table-cell-fill`, `table-cell-no-fill`), borders (`table-border-preset`, `table-border-style`, `table-border-width`, `table-border-color`, reusing `BorderPreset`/`BorderLineStyle`/`BorderSide`), padding (`table-cell-padding`), vertical alignment (`table-cell-valign`), plus the table-level `table-header-row` and `table-band-rows` toggles
- [ ] A `TableFrame` is a frame *and* an animation target: SPEC-082's move, nudge, frame-resize, delete, duplicate, z-order and Tab cycling apply unchanged, `frameNoun`/`frameAccessibleName` return "table"/"Table", `FrameStylePopover` is suppressed (fill and stroke are per-cell); `isAnimatableKind` in `src/views/presentation/useAnimations.ts` and both SPEC-086 wire-id passes (`spidFor` in `pptx/serialize.ts`, `wireIdFor` in `odp/serialize.ts`) gain `kind === "table"`, so a table consumes a minted `p:cNvPr id` / `de-f{n}` and is reachable by `p:spTgt spid` / `smil:targetElement`
- [ ] Preservation: unmodelled markup survives via `pptxTblPr`/`pptxTcPr`/`pptxGraphicFramePr` and `odpTableStyleName`/`odpColumnStyleNames`/`odpRowStyleName`/`odpCellStyleName`, each cleared by the command that edits the modelled property it shadows — exactly `Table.docxTblPr`'s documented rule — and an unedited table re-emits from those strings rather than being regenerated
- [ ] Byte identity: a no-edit download is byte-identical part-for-part under `equalBytes` from `src/lib/core/equality.ts`; an edit on slide 2 leaves every other part byte-identical; and a second write of the reopened model yields identical slide-part bytes (SPEC-081's byte-stability rule extended to tables)
- [ ] Round-trip fidelity: edit → save → reopen reproduces every modelled field in both formats under field-level `toEqual` plus `equalStructural` from `src/lib/equality.ts`, whose `IGNORED_KEYS` drops frame `id`s and whose `XML_STRING_KEYS` does not list the new preserved fields, so those compare raw rather than canonically
- [ ] The neutral half moves into `src/lib/core/tables.ts` by the Reference's two-phase move; `src/lib/core/boundary.spec.ts` and `src/lib/presentation/boundary.spec.ts` stay green, and `src/lib/tables.ts`, `tables.spec.ts`, `tables.formatting.spec.ts` and the `.docx`/`.odt` round-trip suites keep passing with no expectation changed
- [ ] Fixtures: `src/lib/presentation/fixtures/build.ts` gains `buildPptxTablesFixtureBytes`/`buildOdpTablesFixtureBytes` + `EXPECTED_FIXTURE_TABLE` under a new banner (the four mini-deck slides are never mutated), `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec091-tables.pptx` and `.odp`, and each deck carries one deliberately out-of-subset table pinning the preserve-don't-parse path
- [ ] `PresentationView.vue` gains exactly one `useTables({ … })` wiring call and its template hookup, paid for by extracting the four panel `v-if` blocks into `src/components/presentation/SlidePanelStack.vue`, or by an equivalent extraction where a sibling spec has already landed that one; `PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 5 from whatever it stands at when this lands (550 → 545 today) with a ledger row appended — the budget never rises
- [ ] e2e (`e2e/SPEC-091.spec.ts`): open `spec091-tables.pptx` → the table renders as `slide-table` with its cells → click a cell, type, see the text → insert a row below → merge two header cells → tick "Banded rows" → download and reopen → the same grid, spans, text and banding come back, and the deck's out-of-subset table is still a foreign frame

## Out of scope
Charts on slides (still `ForeignFrame` `"Chart"`). Nested tables, diagonal
cell borders, cell text rotation, autofit, gradient/picture/pattern cell
fills (SPEC-079's solid-or-none rule applies), per-cell or per-row
animation builds (SPEC-086 owns the animation model; this spec only widens
its target kinds), sorting, formulas, text↔table conversion (SPEC-076's),
authoring a table *style* part, cross-format conversion.

## Reference

### Dependencies restated
- `specs/079-slide-canvas-rendering.md` — `SlideFrame` is the union
  `TextFrame|ImageFrame|ShapeFrame|GroupFrame|ForeignFrame`, each
  extending a non-exported `SlideFrameBase { id: string; xCm; yCm;
  widthCm; heightCm; rotationDeg? }` (`id` minted at read, never
  serialised); `SlideFill` = `{type:"solid";color}|{type:"none"}`.
- `specs/081-slide-text-editing.md` — `SlideParagraph { runs: SlideRun[];
  align?; bullet?; level? }`, `SlideRun { text; bold?; italic?;
  underline?; fontFamily?; sizePt?; color? }`; `serializeSlideXml(slide,
  rels, slidePartPath)`, `serializePage(slide, registry, usedNames, master,
  styleSource)`, and `markSlideEdited(slide, format)` (sets `dirty`, drops
  `sourceXml` for odp) — **every command ends with it**.
- `specs/082-slide-objects-images-shapes-text-boxes.md` — `FrameRect`,
  `MIN_FRAME_SIZE_CM` 0.5, `reorderFrame`/`deleteFrame`/`duplicateFrame`,
  `frameNoun`/`frameAccessibleName`. `specs/086-slide-animations.md` —
  `SlideAnimation { frameId; effect; direction?; trigger; durationMs?;
  delayMs? }`, targeted by `p:cNvPr id` 2,3,… (pptx) / `de-f{n}` (odp).
- `specs/023-table-formatting-and-cell-operations.md` — owner of what is
  already neutral in `src/lib/core/tables.ts`: `BorderLineStyle =
  "none"|"single"|"double"|"dashed"|"dotted"`, `BorderSide { style; widthPt;
  color }`, `CellBorders`, `CellVerticalAlign`, `BorderPreset`, the picker
  bounds `PICKER_MAX_ROWS`/`PICKER_MAX_COLS` and the grid limits/floors the
  ACs name, `CellRange`, `cellRangeBounds`, `isCellInRange`, `OwnerGridTable`,
  `computeOwnerGrid(t): CellOwner[][]`.
- `equalBytes(a: Uint8Array, b: Uint8Array)` is in
  **`src/lib/core/equality.ts`**, not `src/lib/equality.ts`, and compares
  part-level after `unzip`, never container bytes; `equalStructural` is in
  `src/lib/equality.ts`, whose `equalSemantic` takes a `DocumentModel`,
  unusable here.

### The table model (extends SPEC-079's frame union)
```ts
export interface SlideTableCell {
  paragraphs: SlideParagraph[];   // [] on a covered slot, else >= 1
  colSpan?: number; rowSpan?: number; covered?: true;  // absent spans = 1
  fill?: SlideFill; borders?: CellBorders;             // core/tables.ts
  paddingCm?: number; verticalAlign?: CellVerticalAlign; // 0.13cm / "top"
  pptxTcPr?: string; odpCellStyleName?: string;        // preservation
}
export interface SlideTableRow {
  cells: SlideTableCell[];   // exactly columnWidthsCm.length entries
  heightCm?: number; isHeader?: true;  // heightCm is a MINIMUM, not fixed
  odpRowStyleName?: string;
}
export interface TableFrame extends SlideFrameBase {
  kind: "table";
  rows: SlideTableRow[]; columnWidthsCm: number[];
  firstRow?: true; bandRow?: true;   // header-row / banded-row look
  styleId?: string;                  // a:tableStyleId / table:template-name
  pptxTblPr?: string; pptxGraphicFramePr?: string;     // preservation
  odpTableStyleName?: string; odpColumnStyleNames?: (string|undefined)[];
}
```
`firstRow`/`bandRow`/`covered` store only `true`.

### Reuse, and the two-phase move into `src/lib/core`
Neutral, so it moves down: the grid arithmetic private to
`src/lib/tables.ts` — `expandRangeToWholeCells` and the insert-row /
insert-column / delete-row / delete-column / merge / split *plans* —
restated in core over the existing `OwnerGridTable` shape (widened with
`covered?: boolean`), returning span/covered edits plus the slots needing
a fresh cell so each product supplies its own cell factory, plus the ODF
`fo:border` codec from `odf/reader.ts#parseBorderShorthand`/
`odf/writer.ts#borderShorthand`. Per-product, so it stays: the cell
container and its content — core holds no generic `Table<TCell>`. The
move follows `documentModel.ts`'s `MediaPart` pattern: **phase 1** adds
the functions to core, re-exported from their old module so every import
path keeps working; **phase 2** re-points `src/lib/tables.ts` and the odf
reader/writer at core, deleting the private copies.

### pptx wire mapping (`a:tbl`)
`src/lib/presentation/pptx/tbl.ts`, following the `timing.ts` contract:
`writePptxTbl(frame: TableFrame, id: number): string` and
`readPptxTbl(gf: Element, ctx): TableFrame | undefined`, the latter
`undefined` out of subset via a module-private
`const ABORT = Symbol("out-of-subset")` thrown from the walker, caught in the
wrapper. Nesting: `p:graphicFrame` > `p:nvGraphicFramePr`, `p:xfrm` (the **p**
namespace, not `a:`, holding `a:off`/`a:ext`), `a:graphic` >
`a:graphicData@uri="…/drawingml/2006/table"` > `a:tbl` > `a:tblPr` (attrs
`firstRow`/`bandRow`, child `a:tableStyleId`), `a:tblGrid` > `a:gridCol@w`,
`a:tr@h` > `a:tc` > `a:txBody` + `a:tcPr`; widths/heights EMU → `emuToCm`.
Every slot has its own `a:tc`: the owner carries `gridSpan`/`rowSpan`, covered
slots `hMerge`/`vMerge="1"`. `a:tcPr@anchor` `t|ctr|b` → `verticalAlign`;
`marL/marR/marT/marB` → `paddingCm` when all four are equal, else kept in
`pptxTcPr`; `a:lnL/R/T/B` → `CellBorders` (`a:srgbClr`, `w` EMU → pt,
`a:prstDash` → `BorderLineStyle`); fill per SPEC-079's colour resolution.
**Abort → `ForeignFrame`** on a missing `a:tblGrid`, a row whose `a:tc` count
differs from the grid column count, a non-table `graphicData@uri`, or
unparseable `p:xfrm`. Write: `frameXml` in `pptx/serialize.ts` gains
`case "table"` — unlike `group`/`foreign` it **does** consume a minted id,
making it animatable.

### odp wire mapping (`table:table`)
`src/lib/presentation/odp/tableXml.ts`, same module contract. Nesting:
`draw:frame` (geometry per SPEC-079) > `table:table` (attrs
`table:template-name`, `table:use-first-row-styles`,
`table:use-banding-rows-styles`) > `table:table-column` (`table:style-name`,
`table:number-columns-repeated`), optional `table:table-header-rows`, then
`table:table-row` > `table:table-cell` (`table:style-name`,
`table:number-columns-spanned`/`-rows-spanned`, `text:p`) and
`table:covered-table-cell` for covered slots. Namespaces and readers come from
`core/odfXml.ts` (`TABLE_NAMESPACE`, `STYLE_NAMESPACE`, `FO_NAMESPACE`,
`tableAttr`, `styleAttr`, `foAttr`, `odfLengthToCm`, `odfColorToCss`). Widths,
heights and cell formatting live in the automatic styles, so the reader's
`StyleResolver` (today `fillColor`/`stroke`) gains `tableColumnWidthCm`
(`style:table-column-properties/@style:column-width`), `tableRowHeightCm`
(`@style:row-height`) and `tableCellProps` (`style:table-cell-properties` →
`fo:background-color`, `fo:border-*`, `fo:padding`, `style:vertical-align`).
Write: `frameXml` in `odp/serialize.ts` gains `case "table"`, and
`StyleRegistry` gains `tableColumnStyle`/`tableRowStyle`/`tableCellStyle`
minting `de-TC{n}`, `de-TR{n}`, `de-CL{n}` with the `seedFrom` discipline the
existing families use.

### Table styles, and the golden-first process (normative)
A table's look usually lives outside it — pptx in `ppt/tableStyles.xml` (root
`a:tblStyleLst def="{GUID}"`, relationship type `…/relationships/tableStyles`,
referenced from `a:tblPr/a:tableStyleId`), odp in an `office:styles` table
template named by `table:template-name`. The rule, plainly: **resolve the
effective look from that part for rendering; hold only explicit formatting in
the model; never mint or edit a style part.** That part and `styleId` survive
verbatim, so PowerPoint and Impress keep the author's theme while a restyled
cell's explicit formatting overrides it; an absent or unparsed style part
falls back to a built-in look.

What this file states as confirmed is the element and attribute *names*
above. Three details are **not** confirmed from source and must not be
invented: (1) child order and default insets inside `a:tcPr`, and
whether a covered `a:tc` still needs an empty `a:txBody`; (2) the
internals of `a:tblStyleLst` (`a:wholeTbl`, `a:firstRow`, `a:band1H`);
(3) the ODF table-template names behind `table:template-name`. Follow
SPEC-086's process: hand-author canonical bytes into `pptx/tblGolden.ts`
(`TBL_GOLDEN`, `GOLDEN_TABLE`) and `odp/tableGolden.ts` (`TABLE_GOLDEN`,
re-exporting it), assert serialiser output equals them exactly, require
the readers to accept flavoured variants too, and record cross-app
verification as manual QA in the PR.

### UI
`src/views/presentation/useTables.ts` — deps `{ model, currentSlide,
currentIndex, record, commit, selectedFrameId, selectFrame }`, returning
`{ open, toggle, table, cellRange, selectCell, insertRowAbove,
insertRowBelow, insertColumnLeft, insertColumnRight, deleteRow,
deleteColumn, deleteTable, canMerge, mergeCells, canSplit, splitCell,
setHeaderRow, setBandRows, setCellFill, setBorderPreset, setBorder,
setPadding, setVerticalAlign, + the computeds each control binds }`,
obeying the reactivity rule — `record(label)` → mutate → `commit()`, one
undo entry per gesture, derived state as fresh-array computeds; undo
labels in the existing voice ("Insert row", "Merge cells", …). Components:
`TablePanel.vue` (props exactly `{ controller: {type: Object, required:
true}, slideNumber: {type: Number, default: 1} }`, emits nothing,
`section[data-testid="slides-table-panel"]`), `SlideTableView.vue` (from
`SlideFrameView.vue`'s kind switch) and `SlidePanelStack.vue` (the
extraction that pays for the wiring). Typed transforms go in
`components/presentation/tableRender.ts`, never in an SFC script.

### Tests
- `src/lib/core/tables.spec.ts` (new — `core/tables.ts` has no colocated spec
  today) — the hoisted grid plans and the ODF border codec.
- `src/lib/presentation/tables.spec.ts` — insert/delete row and column,
  merge, split, the grid invariant, frame-resize scaling, floors.
- `pptx/tbl.spec.ts` and `odp/tableXml.spec.ts` — golden equality both
  ways, tolerance on the flavoured variants, abort → `undefined`.
- `tablesRoundTrip.spec.ts` (`toEqual` + `equalStructural` on reopen,
  per-part `equalBytes`, byte stability), `tablesFixtures.spec.ts` (both
  decks → `EXPECTED_FIXTURE_TABLE`), `useTables.spec.ts`, the size spec's
  budget row, and `e2e/SPEC-091.spec.ts`.

### Specs this changes
- SPEC-079's frame union gains a sixth member; its "a table is a foreign
  frame" rule is superseded for in-subset tables only.
- `fixtures/build.ts`'s exported `FIXTURE_FRAME_KINDS` slide-4 entry changes
  `"foreign"` → `"table"`; its two consumers, `roundTrip.spec.ts` and
  `slideModel.spec.ts`, re-run unchanged in shape — coverage never weakened.
- SPEC-086's animatable-kind set and both wire-id passes gain `"table"`
  (its model untouched); SPEC-082's frame commands gain a kind and its
  style popover is suppressed for tables; SPEC-023's core types gain a
  second consumer, with no text-side change.
- `specs/000-roadmap.md` gains this spec's Phase 13 row, and its "Tables
  and charts on slides" backlog row narrows to charts.
