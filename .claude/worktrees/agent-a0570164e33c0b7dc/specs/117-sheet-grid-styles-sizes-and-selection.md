---
id: SPEC-117
title: The sheet grid — styles, sizes, merges, frozen panes and selection
status: draft
area: layout
complexity: complex
story_points: 13
depends_on: [SPEC-063, SPEC-077, SPEC-116]
---
## Why
After SPEC-116 a workbook opens as a uniform 64 × 20 px lattice of plain
values. Nothing the author did to the sheet survives the trip to the screen:
the bold header row, the yellow input cells, the boxed totals, the narrow
code column and the wide description column, the title merged across six
columns, the frozen header that keeps the column names visible while you
scroll, the hidden working columns, the 85% zoom it was saved at. All of it
is in the file — `xl/styles.xml`, `<cols>`, `<mergeCells>`, `<sheetViews>`,
the ODS automatic styles and `settings.xml` — sitting in `opaque` and
`sourceXml`, unread. And the grid cannot be navigated: there is no active
cell, no selection, no way to ask what a cell holds.

This spec makes the grid look like the workbook and behave like a
spreadsheet, still without editing anything. It reads every cell style both
formats carry (except number formats, SPEC-118), every size, merge, freeze,
hidden row and column and view setting; renders them in the DOM-virtualised
grid SPEC-116 built; and adds the selection model, the keyboard map, a
read-only formula bar and Name Box, and status-bar aggregates. **It is the
normative source for `CellStyle` and the style tables read from both formats,
the colour mapping (Excel's theme-index swap, tint → luminance transforms,
the indexed palette), grid geometry (`ColumnSpan`, column-width units, the
maximum-digit-width table, `GridMetrics`), merges, frozen panes, hidden rows
and columns, gridlines, zoom, `SheetSelection` and the read-only keyboard
navigation, the formula bar and Name Box, and the status-bar aggregates.**
Every later spreadsheet spec that styles, sizes, selects or merges edits the
shapes defined here.

Like SPEC-116, nothing here marks a sheet dirty. Selection, zoom and the
gridline toggle are view state: remembered across reloads, never a change to
the file.

## Acceptance criteria
- [ ] `spreadsheet/model.ts` gains the types and fields in "Model additions" — `CellStyle`, `CellFont`, `CellFill`, `CellAlignment`, `CellBorderEdge`, `SheetBorderStyle`, `ColumnSpan`, `SheetColorRef`; `SpreadsheetModel.styles` (index 0 = the default "Normal" style, never empty), `defaultStyleId: 0`, `theme?`, `indexedColors?`; `Sheet.columns`, `defaultColWidthPx`, `defaultRowHeightPt`, `merges`, `frozen?`, `showGridLines?`, `showZeros?`, `zoomPercent?`, `tabColor?`, `selection?`, `topLeftCell?`; `Row.heightPt?`, `customHeight?`, `hidden?`, `styleId?`, `sourceAttrs?`; `Cell.styleId?`; `TextRun.font?` — all optional except `styles`, `columns`, `merges` and the two defaults, and the model stays plain JSON (`JSON.parse(JSON.stringify(model))` `toEqual`s it)
- [ ] Style resolution is one function, `effectiveStyleId(sheet, row, col)` in `spreadsheet/styles.ts`: the cell's `styleId`, else its row's `styleId` (present only when the source row had `customFormat`), else the `ColumnSpan` covering `col`, else `0` — a binary search, never a scan — and `cellStyle(model, id)` returns `model.styles[id]` or `styles[0]` for an out-of-range id
- [ ] The DrawingML theme-part reader moves to core per "Theme": `readThemePart(xml): DocumentTheme | undefined` in `src/lib/core/themeDrawingMl.ts`, with `docx/theme.ts`'s `readDocxTheme` delegating to it (two-phase move — every text import unchanged, `docx/themePart.spec.ts` passes untouched); the xlsx reader fills `model.theme` from the workbook's theme relationship, an ODS workbook has none, and every colour resolves through `resolveColor(color, model.theme ?? DEFAULT_THEME)`
- [ ] `spreadsheet/xlsx/styles.ts` reads `xl/styles.xml` per "Reading xlsx styles": `fonts`, `fills` (pattern and gradient), `borders` (all thirteen line styles plus diagonals), `cellXfs` → `model.styles` one-to-one in order (`styles[i].sourceXf === i`, so a `<c s="7">` is `styleId: 7`; never deduplicated at read), `xf@quotePrefix` → `quotePrefix`, `<colors><indexedColors>` → `indexedColors`; `numFmts`, `cellStyleXfs`/`cellStyles`, `dxfs`, `tableStyles`, `extLst` and `colors/mruColors` stay unread in the retained part (SPEC-118, SPEC-123, SPEC-129 and SPEC-133 read them — `numberFormat` and `namedStyle` are declared here, populated there)
- [ ] Colours map per "Colours": `theme="n"` with Excel's swap (0→`lt1`, 1→`dk1`, 2→`lt2`, 3→`dk2`, 4–9→`accent1`–`accent6`, 10→`hlink`, 11→`folHlink`); `tint` t → `lumMod: 1 + t` (t < 0) or `lumMod: 1 − t, lumOff: t` (t > 0), with the source attribute text kept as `xlsxTint`; `indexed` through the 64-entry legacy palette or the file's override (64 and `auto="1"` → automatic, i.e. absent; 65 → `#ffffff`); `rgb` ARGB → lowercase `#rrggbb` with alpha ignored — and a test pins `<color theme="4" tint="0.39997558519241921"/>` ("Accent 1, Lighter 40%") resolved against `DEFAULT_THEME` to Excel's displayed `#8ea9db` ±1 per channel
- [ ] `spreadsheet/ods/styles.ts` reads ODS cell styles per "Reading ODS styles": `style:default-style` (family `table-cell`) → common styles in `styles.xml` → automatic styles in `content.xml`, flattened per property along the `style:parent-style-name` chain into one `CellStyle` per style name actually referenced (`sourceStyleName` set; `namedStyle` left to SPEC-123), with `table-cell-properties`, `paragraph-properties` and `text-properties` mapped by the table there and `fo:border*` strings classified by the pinned "ODS borders" table; `style:data-style-name` is left for SPEC-118
- [ ] `spreadsheet/xlsx/sheetXml.ts` reads each worksheet's layout per "Reading xlsx sheet layout": `<cols>` (1-based `min`/`max`, `width` → `widthPx` through the MDW formula with `sourceWidthChars` kept, `hidden`, `style`, every other attribute verbatim in `sourceAttrs`), `<row>` `ht`/`customHeight`/`hidden`/`s` (+`customFormat`) with unmodelled attributes (`spans`, `x14ac:dyDescent`, `outlineLevel`, `collapsed`, `thickTop`, `thickBot`) in `sourceAttrs`, `sheetFormatPr`, `mergeCells`, `sheetPr/tabColor`, and `sheetViews/sheetView[0]` (`showGridLines`, `showZeros`, `zoomScale`, `topLeftCell`, the `pane` when `state` is `frozen` or `frozenSplit`, and the active pane's `selection`) — a `state="split"` pane is not a freeze and is left in the retained source; a `<c>` carrying only `s` is now kept as a style-only cell (counted against `MAX_WORKBOOK_CELLS`); shared-string and inline `<r>` runs gain `font` from their `<rPr>`
- [ ] `spreadsheet/ods/tableXml.ts` and `ods/settings.ts` read each table's layout per "Reading ODS sheet layout": `table:table-column` widths, `table:visibility` and `table:default-cell-style-name`; `table:table-row` heights, `style:use-optimal-row-height` (false ⇒ `customHeight`), visibility and default cell style; spans (`table:number-columns-spanned`/`-rows-spanned`) → `merges`; covered cells that carry a value or a style are now **stored** (SPEC-116 dropped them; they render hidden under the merge and survive SPEC-122's re-serialisation); `text:span` runs gain `font`; and `settings.xml`'s per-table `HorizontalSplitMode`/`VerticalSplitMode` = 2 (+ positions) → `frozen`, `ShowGrid`, `ZoomValue`, `CursorPositionX`/`Y`, `PositionRight`/`PositionBottom` → view fields, with the view-global item as the fallback for each
- [ ] Repeated styled empty ODS cells never blow up the model: a run whose style equals the effective column/row default stores nothing, a run reaching the last column becomes the row's `styleId`, and every other styled empty cell is stored and counted against `MAX_WORKBOOK_CELLS` exactly as SPEC-116 counts values — asserted by a fixture row of `table:number-columns-repeated="16380"` styled cells that stores no cell
- [ ] Geometry follows "Geometry": widths are px at 100% zoom, heights pt; the xlsx width ↔ px conversion uses the workbook default font's maximum digit width (MDW) from the pinned table — never a measured font — with `xlsxWidthToPx(9.140625, 7) === 64`, `pxToXlsxWidth(64, 7) === 9.140625`, a `baseColWidth`-only sheet defaulting to 64 px for Calibri 11, and every row of the "MDW" test table asserted
- [ ] `spreadsheet/geometry.ts` exports `buildGridMetrics(sheet)` → `GridMetrics` (`colLeft`, `colWidth`, `colAt`, `rowTop`, `rowHeight`, `rowAt`, `totalWidth(cols)`, `totalHeight(rows)`) in integer px, O(log n) per query over the sorted custom sizes and never an array of `MAX_ROWS` entries; hidden rows and columns are zero-sized; a 1,048,576-row sheet with 10,000 custom heights answers 100,000 random `rowAt`/`rowTop` queries in under 50 ms in the unit suite
- [ ] `SheetGrid.vue` renders styles per "Rendering": the cell layer (`SheetCellLayer.vue`, the paper) draws fonts through SPEC-048's `fontStack`, fills (the pinned pattern table; gradients as CSS gradients), alignment, wrap, indent, rotation and vertical text, shrink-to-fit, rich-text runs and a merged range as one element sized to the range with the top-left cell's value and style (`aria-rowspan`/`aria-colspan`); borders are drawn once per edge on a border layer with the pinned shared-edge rule (heavier wins by "Border CSS and weights", a tie goes to the right/bottom cell), diagonals as an inline SVG line
- [ ] Text overflow follows "Overflow": a string spills into empty neighbours — left-aligned to the right, right-aligned to the left, centred to both sides equally — stopping at a cell holding any value (including `""`), a merged range or the frozen boundary, and never for wrapped, filled, justified, rotated or shrink-to-fit text; a string anchored left of the visible window still paints into it; non-strings never spill
- [ ] Frozen panes render as sticky regions per "Frozen panes and hidden rows and columns": the frozen rows and columns (and their corner) stay in place while the rest scrolls, separated by `sheet-freeze-line` lines; hidden rows and columns take no space and their header boundary shows a `sheet-hidden-marker` titled "Hidden rows 4–6" / "Hidden columns C–D"; the initial scroll position is the file's `topLeftCell`
- [ ] Gridlines and zoom are view state: this spec creates `src/components/spreadsheet/SpreadsheetMenuBar.vue` per "The menu bar" (File: Download, Close; View: "Gridlines" `sheet-menu-gridlines` as a `menuitemcheckbox`, and a Zoom submenu), which later specs append to; zoom uses SPEC-063's vocabulary — the status bar's `ZoomControl` (with its fit entries hidden by a new `showFit` prop), the `ZOOM_PRESETS` ladder, Ctrl/Cmd+`=`/`-`/`0` and Ctrl/Cmd+wheel — within 10–400%, applied as a scale transform per "Zoom" so `GridMetrics` never sees it; a sheet opens at its file zoom ("85%" in `zoom-percent` for a `zoomScale="85"` sheet)
- [ ] The cell area is paper in every app theme per "Dark mode": white background, `#d4d4d4` gridlines and the file's own colours in light and dark themes, the SPEC-059 dim-page filter applying to it as to a page (`.sheet-paper`), while headers, formula bar, tabs, selection and status bar follow the app theme; `tokens.spec.ts` exempts `SheetCellLayer.vue` only
- [ ] `spreadsheet/selection.ts` exports `SheetSelection` and the pure navigation functions in "Selection" (`moveActive`, `extendBy`, `extendTo`, `edgeInDirection`, `currentRegion`, `cycleWithin`, `addRange`, `selectRows`, `selectColumns`, `selectAll`, `snapToMerges`), every result snapped to merges, skipping hidden rows and columns, clamped to the grid, and unit-tested against the worked examples there
- [ ] The grid is keyboard-navigable per "Keyboard map" — arrows, Shift+arrows, Ctrl/Cmd+arrows to the data-region edge, Ctrl/Cmd+Shift+arrows, Home, Ctrl/Cmd+Home/End, PageUp/PageDown, Alt+PageUp/PageDown, Ctrl/Cmd+A (current region, then the whole sheet), Shift+Space, Ctrl+Space, Tab/Shift+Tab and Enter/Shift+Enter (cycling inside a multi-cell selection) — with the primary modifier from `core/platform.ts`, the active cell scrolled into view clear of frozen panes, and keys ignored when focus is in an input (`isOutsideDocumentInput`)
- [ ] The mouse selects per "Pointer": click, drag (auto-scrolling past the edge), Shift+click, Ctrl/Cmd+click to add a range, row/column header click and drag (with Shift/Ctrl/Cmd), and the `sheet-select-all` corner; hit-testing goes through `GridMetrics` (`colAt`/`rowAt` on unscaled coordinates), never through which element was under the pointer
- [ ] Selection is drawn as one `sheet-selection` overlay per range plus a `sheet-active-cell` outline, selected headers are highlighted, the grid has `aria-activedescendant` on the active cell's element (`id="sheet-cell-B3"`) and `aria-selected` on selected mounted cells; the file's selection (`sheetView/selection` or the ODS cursor) is restored on open
- [ ] View state persists without touching the document per "View state": selection, zoom, gridlines, scroll position and active sheet are written back to the stored model through a new `storeSpreadsheetViewState(model)` at most once per `VIEW_STATE_PERSIST_DELAY_MS = 1000` and on close — it never changes `lastModified`, never writes a rows record and never marks a sheet dirty — and an untouched workbook's download stays byte-identical after any amount of navigating, zooming and toggling
- [ ] A read-only formula bar (`FormulaBar.vue`, `formula-bar`) sits above the grid: a `name-box` (`<input readonly>`, `aria-label="Name Box"`) showing the active cell's address (a merge's top-left), "3R × 2C" while a drag is in progress, and `formula-bar-content` (`role="textbox"`, `aria-readonly="true"`) showing the active cell's `generalText` with line breaks preserved (SPEC-118 refines this to `editText`)
- [ ] `SpreadsheetStatusBar.vue` (`spreadsheet-status-bar`) shows, while the selection covers more than one cell, "Average: …" (`sheet-status-average`), "Count: …" (`sheet-status-count`) and "Sum: …" (`sheet-status-sum`) per "Aggregates" — computed by `selectionAggregates` in `spreadsheet/aggregates.ts` over stored cells only (each cell once, hidden rows and columns excluded), deferred through the idle scheduler for selections over 100,000 stored cells — with the zoom control at its right
- [ ] Both readers' extensions bump `SPREADSHEET_READER_VERSION` (to the next integer) and append a `ReaderUpgrade` per "Reader version": a workbook stored by SPEC-116's reader reopens with its styles, sizes, merges and view settings, keeping its id, title, starred state and sheet ids
- [ ] Layering holds: everything new under `src/lib/spreadsheet/` passes `boundary.spec.ts`; `SpreadsheetView.vue` gains only `use*` wiring calls and template hookups and its `SpreadsheetView.size.spec.ts` budget is not raised (the grid surfaces live in components, the logic in `useSheetMetrics.ts`, `useSheetSelection.ts`, `useSheetZoom.ts`, `useSheetViewState.ts` and `useSpreadsheetMenu.ts`, each within `composableSize.spec.ts`'s 400 lines); `themeGuard.spec.ts`'s allow-list gains `lib/spreadsheet/(xlsx|ods)/theme*.ts` for the colour codec
- [ ] Fixtures per "Fixtures": `e2e/fixtures/spec117-styles.xlsx` (Excel-shaped) and `spec117-styles.ods` (LibreOffice-shaped) built by `spreadsheet/fixtures/build.ts`, each carrying every construct this spec reads in the shape its producer writes it
- [ ] Unit tests: `core/themeDrawingMl.spec.ts` (`readThemePart`), `spreadsheet/styles.spec.ts`, `xlsx/styles.spec.ts`, `xlsx/sheetXml.spec.ts`, `ods/styles.spec.ts`, `ods/tableXml.spec.ts`, `ods/settings.spec.ts`, `geometry.spec.ts`, `merges.spec.ts`, `selection.spec.ts`, `cellRender.spec.ts`, `aggregates.spec.ts`, `open.spec.ts` (the upgrade), `roundTrip.spec.ts` (still byte-identical for both new fixtures), `views/spreadsheet/useSheetSelection.spec.ts`, `components/spreadsheet/FormulaBar.spec.ts`, `SpreadsheetStatusBar.spec.ts`
- [ ] e2e (`e2e/SPEC-117.spec.ts`): open `spec117-styles.xlsx` → `B2` renders bold, red text on a yellow fill → `C5` is one element as wide as `C5:E5` → the header boundary between F and H shows a `sheet-hidden-marker` titled "Hidden column G" → scroll down 200 rows and rows 1–2 are still visible → `zoom-percent` reads "85%" → click `B3`, press Ctrl/Cmd+Down and `name-box` reads "B12" → click `B3`, Shift+click `B5` → status bar reads "Average: 20", "Count: 3", "Sum: 60" → Ctrl/Cmd+A twice selects the whole sheet → toggle Gridlines off → reload → selection, zoom and gridlines are as left → switch the app to dark mode and `B2`'s computed background is still yellow → download → every part `equalBytes` the fixture's → repeat open-and-check for `spec117-styles.ods` (frozen pane, merge, hidden column, file zoom)

## Out of scope
Editing anything — values, styles, sizes, merges, freezes (SPEC-121
onward; freeze/unfreeze and resizing are SPEC-124). Named cell styles
(`cellStyleXfs`/`cellStyles`, ODS parent styles as names — SPEC-123; the
ODS parent chain is still flattened here for rendering). Number formats: a date
still shows as its serial and `CellStyle.numberFormat` is declared here but
populated by SPEC-118. Formulas (SPEC-119/120). Conditional formats and
`dxfs` (SPEC-129), table styles (SPEC-133), sparklines, outline grouping and
its +/− controls (outline attributes are preserved in `sourceAttrs`, never
rendered). Auto-fitting rows that carry no height — openpyxl-style files
without `ht` render at the default height until an edit refits them
(SPEC-122) or SPEC-124's AutoFit. Split
(non-frozen) panes, custom gridline colours, `rightToLeft` sheets, page
layout and page-break preview views, "Headings" hiding and
`showFormulas` (SPEC-119's view) — each preserved in the source, rendered as
the normal view. Zoom to selection. Touch-specific selection handles
(SPEC-068 territory). Filter-hidden versus user-hidden rows (both are
`hidden` here; SPEC-128 distinguishes them).

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — the model (`SpreadsheetModel {
  kind; id; title; format: "xlsx" | "ods"; packageVariant?; readerVersion;
  sheets; activeSheetIndex; dateSystem; odsNullDate?; odsContentEnvelope?;
  opaque; preview?; properties?; media? }`, `Sheet { id; name; visibility;
  kind: "grid" | "foreign"; rows; partName?; sourceXml? }`, `Row { index;
  cells }`, `Cell { col; value; runs? }`, `TextRun { text }`, `CellValue =
  number | string | boolean | CellError | null`) and its four invariants:
  copy-on-write rows, sparseness, the source is the splice base, dates are
  serials. Addressing in `address.ts` (0-based `CellRef { row; col }`,
  inclusive normalised `GridRange { top; left; bottom; right }`,
  `colToLetters`, `formatA1`, `parseA1`, `formatRange`, `parseRange`,
  `parseSqref`, `findRow`, `getCell`, `usedRange`); `generalText(value)` in
  `general.ts`; `MAX_ROWS = 1_048_576`, `MAX_COLS = 16_384`,
  `MAX_WORKBOOK_CELLS = 2_000_000`; the forward-only scanner
  `core/xmlScanner.ts` (`scanXml`, exact offsets, DTD refused); the reader
  version machinery in `open.ts` (`SPREADSHEET_READER_VERSION`,
  `ReaderUpgrade { from; upgradeDirtySheet(model, sheet, sourceParts) }`,
  `READER_UPGRADES`, `upgradeSpreadsheetModel`); the grid (`SheetGrid.vue`,
  `useSheetViewport.ts`, `GRID_OVERSCAN_ROWS = 8`, `GRID_OVERSCAN_COLS = 4`,
  `MAX_SCROLL_EXTENT_PX = 15_000_000`, scroll extent `max(usedRange.bottom +
  100, 1000)` rows × `max(usedRange.right + 10, 26)` columns, ARIA grid with
  `aria-rowcount = MAX_ROWS + 1`, testids `spreadsheet-grid`, `sheet-cell` +
  `data-ref`, `sheet-col-header`, `sheet-row-header`, `sheet-tabs`,
  `sheet-tab`); the store's row split (`sheetrows:<docId>:<sheetId>` records
  written only when a sheet's `rows` reference changes).
- `specs/063-zoom-and-view-modes.md` — `src/lib/zoom.ts` (`ZOOM_PRESETS =
  [50, 75, 90, 100, 125, 150, 200]`, `stepZoomUp`, `stepZoomDown`,
  `clampZoom(percent, min, max)`), `ZoomControl.vue` (`zoom-control`,
  `zoom-out`, `zoom-percent`, `zoom-in`, `zoom-menu`, `zoom-preset-<n>`,
  `zoom-fit-width`, `zoom-fit-page`), shortcuts `view.zoomIn`/`view.zoomOut`/
  `view.zoomReset` in `src/lib/shortcuts.ts`, and the rule that zoom is a
  view transform: rects measured inside it are scaled, model geometry is not.
- `specs/077-document-themes.md` — `src/lib/core/theme.ts`: `ThemeColorSlot`,
  `DocumentTheme { name; colors: Record<ThemeColorSlot, string>; fonts: {
  major; minor } }`, `ThemeColorRef { theme; lumMod?; lumOff?; shade?; tint?
  }` (fractions 0..1, applied lumMod, lumOff, shade, tint), `ModelColor =
  string | ThemeColorRef`, `ModelFont = string | { theme: "major" | "minor" }`,
  `resolveColor`, `resolveFont`, `DEFAULT_THEME` (Office: minor Calibri). The
  single-resolution-point rule is enforced by `core/themeGuard.spec.ts`
  (no `.lumMod/.lumOff/.shade/.tint` access outside an allow-list). The
  application theme (SPEC-059) and the document's colours never mix.
- Already on main and used as-is: `src/lib/fonts/substitution.ts`'s
  `fontStack(family)` (SPEC-048; app layer — components may import it,
  `lib/spreadsheet` may not), `core/platform.ts`'s `isMacPlatform()`,
  `isOutsideDocumentInput` in `src/lib/shortcuts.ts` (SPEC-021),
  `MenuBar.vue`/`MenuCheckItem.vue`/`MenuSubmenu.vue` (SPEC-049),
  `core/idleScheduler.ts`'s `runWhenIdle`, SPEC-059's dim-page rule in
  `App.vue` (`:root[data-theme="dark"][data-dim-page="1"] .page`), and
  `src/styles/tokens.spec.ts`'s `EXEMPT_FILES` for document-content
  components.

### Model additions
```ts
// src/lib/spreadsheet/model.ts (additions)
import type { ModelColor, ModelFont, ThemeColorRef, DocumentTheme } from "../core/theme";
import type { GridRange, CellRef } from "./address";
import type { SheetSelection } from "./selection";

/** A theme reference read from SpreadsheetML; `xlsxTint` is the source `tint`
 *  attribute text, re-emitted verbatim by SPEC-122 when the colour is unchanged
 *  (resolveColor ignores it — the same move as core/themeDrawingMl.ts's `spelling`). */
export interface SheetColorRef extends ThemeColorRef { xlsxTint?: string }

export type SheetBorderStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair"
  | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
export interface CellBorderEdge { style: SheetBorderStyle; color?: ModelColor }   // color absent = automatic (black)
export interface CellFont {
  family?: ModelFont; sizePt?: number; bold?: true; italic?: true;
  underline?: "single" | "double" | "singleAccounting" | "doubleAccounting"; strike?: true;
  color?: ModelColor;                       // absent = automatic (black)
  vertAlign?: "superscript" | "subscript";
}
export type CellFill =
  | { kind: "solid"; color: ModelColor }
  | { kind: "pattern"; pattern: string; fg?: ModelColor; bg?: ModelColor }  // xlsx patternType other than none/solid
  | { kind: "gradient"; type: "linear" | "path"; degree?: number;
      left?: number; right?: number; top?: number; bottom?: number;         // path focus, 0..1
      stops: { position: number; color: ModelColor }[] };
export interface CellAlignment {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";      // absent = bottom
  wrap?: true; indent?: number; textRotation?: number;                     // 0..180, 255 = vertical stacked
  shrinkToFit?: true;
}
export interface CellStyle {
  font?: CellFont; fill?: CellFill;
  borders?: { top?: CellBorderEdge; right?: CellBorderEdge; bottom?: CellBorderEdge; left?: CellBorderEdge;
              diagonal?: CellBorderEdge & { up?: true; down?: true } };
  alignment?: CellAlignment;
  numberFormat?: string;        // declared here, populated by SPEC-118; absent = "General"
  builtinFormatId?: number;     // SPEC-118: a locale-dependent Excel built-in (14, 22, 5–8, 41–44)
  protection?: { locked?: false; hidden?: true };   // locked defaults to true
  quotePrefix?: true;           // xlsx xf@quotePrefix: the value was entered as text with a leading '
  namedStyle?: string;          // "Good", "Heading 1"… — declared here, read and written by SPEC-123
  sourceXf?: number;            // xlsx: the cellXfs index this was read from
  sourceStyleName?: string;     // ods: the automatic/common style name this was read from
}
export interface ColumnSpan {
  min: number; max: number;     // 0-based inclusive; spans sorted, non-overlapping
  widthPx?: number;             // px at 100% zoom; absent = the sheet default
  hidden?: true; styleId?: number;
  sourceWidthChars?: number;    // xlsx <col width>, re-emitted verbatim while widthPx is unchanged
  sourceAttrs?: string;         // xlsx: unmodelled <col> attributes, verbatim, in source order
  sourceStyleName?: string;     // ods: the table-column style name
}
// TextRun gains:  font?: CellFont   (a full run font in xlsx; the span's own properties in ods)
// Cell gains:     styleId?: number
// Row gains:      heightPt?: number; customHeight?: true; hidden?: true;
//                 styleId?: number   (present only when the source row had customFormat="1" / a default cell style);
//                 sourceAttrs?: string (xlsx unmodelled <row> attributes); sourceStyleName?: string (ods row style)
// Sheet gains:    columns: ColumnSpan[]; defaultColWidthPx: number; defaultRowHeightPt: number;
//                 merges: GridRange[]; frozen?: { rows: number; cols: number };
//                 showGridLines?: false; showZeros?: false; zoomPercent?: number /* 10..400, absent = 100 */;
//                 tabColor?: ModelColor; selection?: SheetSelection; topLeftCell?: CellRef
// SpreadsheetModel gains: styles: CellStyle[]; defaultStyleId: 0; theme?: DocumentTheme;
//                 indexedColors?: string[]   /* the file's 64-entry override palette, lowercase #rrggbb */
```
Sparseness extends to the new fields: a row with no cells is stored only if
it has a height, `hidden` or `styleId`; a column span is stored only if it has
a width different from the default, `hidden` or `styleId`. A style-only cell
(`value: null`, `styleId` set) is stored. `merges` holds only ranges of two or
more cells that overlap no earlier merge (a later overlapping or single-cell
`mergeCell` is not modelled and stays in the source).

### Theme
`readThemePart(xml: string): DocumentTheme | undefined` — `a:theme/
a:themeElements/a:clrScheme` (`a:srgbClr@val`, or `a:sysClr@lastClr` with
the `windowText` → black / `window` → white fallback) and `a:fontScheme`'s
major/minor `a:latin@typeface`; the root `name` is the theme name; undefined
when there is no `clrScheme`. It is `readDocxTheme`'s body moved verbatim;
`readDocxTheme(xml, partName)` becomes `{ ...readThemePart(xml), partName }`.
The xlsx reader finds the part through the workbook's relationship of type
`…/relationships/theme`.

### Colours
Every SpreadsheetML colour element (`color`, `fgColor`, `bgColor`, `tabColor`,
gradient `stop/color`) maps by its first present attribute in this order:

| Attribute | Model |
| --- | --- |
| `auto="1"` | absent (automatic) |
| `theme="n"` (+`tint`) | `SheetColorRef` — slot by the swap below; tint per the rule above |
| `indexed="n"` | `indexedColors[n]` ?? `INDEXED_COLORS[n]` for 0–63; 64 → absent; 65 → `#ffffff` |
| `rgb="AARRGGBB"` | `#rrggbb` (a 6-digit value is accepted as-is) |

Theme swap: 0 `lt1`, 1 `dk1`, 2 `lt2`, 3 `dk2`, 4–9 `accent1`–`accent6`,
10 `hlink`, 11 `folHlink`; anything else is ignored (absent). The swap and
tint code live in `spreadsheet/xlsx/themeColor.ts` (`xlsxColorToModel`,
`modelColorToXlsxAttrs(color, theme)` — the inverse SPEC-121's tab colour and
SPEC-122's style writer call: `lumOff` present → `tint = lumOff`, else
`tint = lumMod − 1`, the source `xlsxTint` text winning when it still
round-trips to the same transforms; a theme reference no tint expresses
exactly — anything but a lone `lumMod` or a `lumMod` + `lumOff` summing to
1 — is written as the `rgb` of `resolveColor(color, theme)`; a hex colour as
`rgb="FFrrggbb"`). `INDEXED_COLORS` (in
`spreadsheet/styles.ts`, ECMA-376 §18.8.27, lowercase):
```
000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff   000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff
800000 008000 000080 808000 800080 008080 c0c0c0 808080   9999ff 993366 ffffcc ccffff 660066 ff8080 0066cc ccccff
000080 ff00ff ffff00 00ffff 800080 800000 008080 0000ff   00ccff ccffff ccffcc ffff99 99ccff ff99cc cc99ff ffcc99
3366ff 33cccc 99cc00 ffcc00 ff9900 ff6600 666699 969696   003366 339966 003300 333300 993300 993366 333399 333333
```
ODS colours are literal `#rrggbb` (`transparent` → absent).

### Reading xlsx styles
- `fonts/font`: `b`, `i`, `strike` (`val` absent or true), `u@val` (absent →
  single; `none` → absent), `vertAlign@val` (`baseline` → absent), `sz@val`,
  `color`, `name@val` → `family`, and `scheme@val` `minor`/`major` →
  `family: { theme }` (the scheme wins over the name, which is what Excel
  renders). `condense`, `extend`, `outline`, `shadow`, `family`, `charset`
  are not modelled (the `sourceXf` reuse in SPEC-122 keeps them).
- `fills/fill`: `patternFill@patternType` `none`/absent → no fill; `solid` →
  `{ kind: "solid", color: fgColor ?? "#000000" }`; any other → `pattern`
  with `fg`/`bg`. `gradientFill@type` (`linear` default, `path`), `degree`,
  `left`/`right`/`top`/`bottom`, `stop@position` + colour.
- `borders/border`: `left`, `right`, `top`, `bottom`, `diagonal` children with
  `@style` (absent or `none` → no edge) and `color`; `border@diagonalUp`/
  `diagonalDown` → the diagonal's `up`/`down`. `start`/`end` are read as
  `left`/`right` (left-to-right sheets).
- `cellXfs/xf`: `fontId`, `fillId`, `borderId` → copies of the parsed
  font/fill/border (plain data, no shared objects), `alignment` child
  (`horizontal`, `vertical`, `wrapText`, `indent`, `textRotation`,
  `shrinkToFit`), `protection` child (`locked="0"` → `locked: false`,
  `hidden="1"`), `quotePrefix`. `xfId` (the named style) is SPEC-123's. The
  `apply*` flags are ignored: Excel writes `cellXfs` fully resolved. A font, fill or border equal to the
  default style's is still copied (no delta encoding).
- The default font (for MDW) is `fonts[cellXfs[0].fontId]` — the Normal
  style's font in every Excel-written file — its family resolved against the
  theme.
- A workbook with no styles part gets `styles: [{}]` and Calibri 11 metrics.

### Reading ODS styles
Flattening walks `style:default-style[@style:family="table-cell"]` →
`style:parent-style-name` ancestors (common styles in `styles.xml`
`office:styles`) → the named style; later definitions win per property.
`styles[0]` is the flattened `Default` style. Each style name referenced by
a cell, a row or column default, or a covered cell gets one entry, in first
reference order. Font names resolve through `office:font-face-decls`
(`style:font-face@style:name` → `svg:font-family`, quotes stripped).

| ODS property | `CellStyle` |
| --- | --- |
| `table-cell-properties@fo:background-color` | `fill: { kind: "solid" }` (`transparent` → none) |
| `@fo:border`, `@fo:border-{top,bottom,left,right}` | `borders.*` per "ODS borders" |
| `@style:diagonal-bl-tr` / `@style:diagonal-tl-br` | `borders.diagonal` with `up` / `down` |
| `@fo:wrap-option="wrap"` | `alignment.wrap` |
| `@style:vertical-align` top / middle / bottom / automatic | `top` / `center` / `bottom` / absent |
| `@style:rotation-angle` a (+`style:direction="ttb"` → 255) | a ≤ 90 → a; a ≥ 270 → 90 + (360 − a); 90 < a ≤ 180 → 90; else 180 |
| `@style:shrink-to-fit="true"` | `alignment.shrinkToFit` |
| `@style:repeat-content="true"` | `horizontal: "fill"` |
| `@style:cell-protect` none / protected / formula-hidden / hidden-and-protected / "protected formula-hidden" | `{ locked: false }` / — / `{ hidden }` / `{ hidden }` / `{ hidden }` |
| `paragraph-properties@fo:text-align` start / end / center / justify (with `style:text-align-source="fix"`) | left / right / center / justify; `value-type` or absent → general |
| `paragraph-properties@fo:margin-left` | `indent = round(px / INDENT_PX_PER_LEVEL)` |
| `text-properties` `style:font-name`, `fo:font-size`, `fo:font-weight="bold"`, `fo:font-style="italic"`, `style:text-underline-style≠none` (+`-type="double"`), `style:text-line-through-style≠none`, `fo:color`, `style:text-position` (`super…`/positive → superscript, `sub…`/negative → subscript) | `font.*` |

### ODS borders
Read — the width w in pt of an `fo:border*` value `"<w> <style> <#color>"`:

| ODS line style | w ≤ 1.0 pt | 1.0 < w ≤ 2.0 pt | w > 2.0 pt |
| --- | --- | --- | --- |
| `solid` | thin | medium | thick |
| `dashed` | dashed | mediumDashed | mediumDashed |
| `dotted` | dotted | dotted | dotted |
| `fine-dashed` | hair | hair | hair |
| `dash-dot` | dashDot | mediumDashDot | mediumDashDot |
| `dash-dot-dot` | dashDotDot | mediumDashDotDot | mediumDashDotDot |
| `double`, `double-thin` | double | double | double |
| `none`, `hidden` | — | — | — |

Write (pinned here for SPEC-122): thin `0.74pt solid`, medium `1.75pt
solid`, thick `2.5pt solid`, dashed `0.74pt dashed`, mediumDashed `1.75pt
dashed`, dotted `0.74pt dotted`, hair `0.74pt fine-dashed`, dashDot `0.74pt
dash-dot`, mediumDashDot and slantDashDot `1.75pt dash-dot`, dashDotDot
`0.74pt dash-dot-dot`, mediumDashDotDot `1.75pt dash-dot-dot`, double
`2.6pt double` — each followed by the colour (`#000000` when automatic).
slantDashDot reads back as mediumDashDot: the one stated loss.

### Reading xlsx sheet layout
- `sheetFormatPr`: `defaultRowHeight` (pt; required by the schema; absent
  element → 15) → `defaultRowHeightPt`; `defaultColWidth` → `defaultColWidthPx
  = xlsxWidthToPx(w, mdw)`; else `baseColWidth` (default 8) →
  `ceil8(trunc(((256·b + trunc(128/mdw)) / 256) · mdw) + 5)` where `ceil8`
  rounds up to a multiple of 8 (Calibri 11: 64). Other attributes
  (`x14ac:dyDescent`, `zeroHeight`, `outlineLevelRow`…) stay in the source.
- `cols/col`: `min`/`max` 1-based → 0-based; `width` → `widthPx`;
  `customWidth` is implied by a width; `hidden="1"`; `style` → `styleId`;
  `bestFit`, `collapsed`, `outlineLevel`, `phonetic` → `sourceAttrs`.
  Overlapping `col` elements: the later one is clipped to the gap.
- `row`: `ht` → `heightPt` (honoured whether or not `customHeight` is set —
  it is the producer's measured height), `customHeight="1"`, `hidden="1"`,
  `s` → `styleId` only with `customFormat="1"`.
- `sheetView[0]`: `showGridLines="0"` → `false`; `showZeros="0"` → `false`;
  `zoomScale` (clamped 10–400) → `zoomPercent` (100 → absent); `topLeftCell`.
  `pane@state` `frozen`/`frozenSplit` → `frozen: { rows: ySplit ?? 0, cols:
  xSplit ?? 0 }` and `pane@topLeftCell` → `topLeftCell`. The `selection`
  whose `@pane` equals `pane@activePane` (default `topLeft`): `activeCell`,
  `sqref` (`parseSqref`), `activeCellId` → `activeRangeIndex`.
- `sheetPr/tabColor` → `tabColor`. `mergeCells/mergeCell@ref` → `merges`.
- SPEC-122 rebuilds `<sheetViews>`, `<sheetFormatPr>` and `<cols>` by
  splicing the modelled attributes into the retained elements (or emitting
  `sourceAttrs`), never by regenerating them blind — unmodelled view
  attributes (`tabSelected`, `view`, `rightToLeft`, `zoomScaleNormal`, split
  panes and the other panes' selections) survive an edit.

### Reading ODS sheet layout
- Column widths: `style:table-column-properties@style:column-width` (any
  ODF length) → px at 96 dpi, unrounded; `table:visibility` `collapse` or
  `filter` → `hidden`; `table:default-cell-style-name` → `styleId`.
  `defaultColWidthPx` is the width of the style used by the table's last
  `table:table-column` run (LibreOffice's trailing repeated column), else
  `style:default-style[@style:family="table-column"]`'s width, else 85.33 px
  (0.8889 in); runs equal to the default with no other property store no span.
- Rows: `style:row-height` → `heightPt`; `use-optimal-row-height="false"` →
  `customHeight`; visibility → `hidden`; `table:default-cell-style-name` →
  `styleId`. `defaultRowHeightPt` is the height of the style on the table's
  last row run, else 12.8 pt. Rows equal to the default store nothing.
- `settings.xml` (`ods/settings.ts`, `parseXml` — the part is small):
  `config:config-item-map-named[@config:name="Tables"]` entry per table name;
  `HorizontalSplitMode` = 2 → `frozen.cols = HorizontalSplitPosition`;
  `VerticalSplitMode` = 2 → `frozen.rows = VerticalSplitPosition`; split mode
  1 (split, not frozen) is ignored; `PositionRight`/`PositionBottom` (frozen)
  or `PositionLeft`/`PositionTop` → `topLeftCell`; `ShowGrid` false →
  `showGridLines: false`; `ZoomType` 0 + `ZoomValue` → `zoomPercent`;
  `CursorPositionX`/`Y` → `selection` (one cell). The view-level items of the
  same names are the fallback. Table style `style:table-properties@tableooo:tab-color`
  → `tabColor`.

### Geometry
```ts
// src/lib/spreadsheet/geometry.ts
export const CELL_PADDING_PX = 2;              // each side; + 1 px gridline = Excel's 5 px
export const INDENT_PX_PER_LEVEL = 9;
export function maxDigitWidthPx(family: string, sizePt: number): number;   // MDW, from the table below
export function xlsxWidthToPx(width: number, mdw: number): number;         // trunc(((256·w + trunc(128/mdw)) / 256) · mdw)
export function pxToXlsxWidth(px: number, mdw: number): number;           // trunc((trunc((px − 5)/mdw · 100 + 0.5)/100 · mdw + 5)/mdw · 256)/256
export interface GridMetrics {
  colLeft(c: number): number; colWidth(c: number): number; colAt(x: number): number;
  rowTop(r: number): number; rowHeight(r: number): number; rowAt(y: number): number;
  totalWidth(cols: number): number; totalHeight(rows: number): number;
}
export function buildGridMetrics(sheet: Sheet): GridMetrics;
```
- **MDW.** `round(ratio × sizePt × 96 / 72)`, at least 1, with `ratio` by
  family (case-insensitive; a metric twin shares its original's row): Calibri,
  Carlito, Aptos Narrow 0.5068; Arial, Helvetica, Liberation Sans, Arimo
  0.5562; Times New Roman, Liberation Serif, Tinos 0.5000; Cambria, Caladea
  0.5557; Courier New, Liberation Mono, Cousine 0.6001; any other family
  0.5068. Aptos Narrow shares Calibri's row because Excel's switch to it kept
  the 8.43-character, 64 px default column. Test table:
  Calibri 10/11/12 → 7/7/8; Aptos Narrow 11 → 7; Arial 10/11 → 7/8;
  Liberation Sans 10 → 7; Helvetica 12 → 9; Times New Roman 11/12 → 7/8;
  Cambria 11 → 8; Courier New 10 → 8; "Foo" 11 → 7. No font is ever
  measured for geometry: a column is as wide on a machine without Calibri as
  on one with it.
- **Integer px.** Queries return integers: `colWidth = round(widthPx)`,
  `rowHeight = round(heightPt × 4/3)` (15 pt → 20 px); hidden → 0.
- **O(log n).** Custom sizes are kept as sorted index runs with prefix sums of
  `(size − default) × length`; `colLeft(c) = c · default + prefix(c)`;
  `colAt`/`rowAt` binary-search the runs. Rebuilt when the sheet's `columns`
  or `rows` reference changes; the build is O(custom runs) (rows with no size
  property are skipped in one pass; SPEC-137 owns finer invalidation).
- SPEC-116's scroll extent and proportional mapping above
  `MAX_SCROLL_EXTENT_PX` now run on `totalHeight`/`totalWidth`.

### Rendering
- **Layers** (bottom to top, inside the scaled content box): gridlines (one
  line element per visible row and column boundary, `#d4d4d4`, absent when
  `showGridLines` is false) → cells (`sheet-cell`, fill and text) → borders →
  selection overlays. Cell elements: every position in the window when it
  holds at most `GRID_MAX_MOUNTED_CELLS = 10_000` positions; beyond that
  (zoomed far out) only stored cells, merge anchors and the active cell are
  mounted. Empty cells are transparent, so a fill hides the gridlines under it.
- **`cellCss(style, theme)`** (`spreadsheet/cellRender.ts`, framework-free)
  returns the CSS declarations: font (`font-size: <pt>pt`, weight, style,
  `text-decoration`, superscript/subscript as `vertical-align` + 0.7 size),
  colour, fill, alignment, padding `0 CELL_PADDING_PX` plus `indent ×
  INDENT_PX_PER_LEVEL` on the aligned side. The family is the resolved name;
  `SheetCellLayer.vue` wraps it in `fontStack()`. Rich runs render as spans
  whose `font` overrides the cell font per property present.
- **Alignment.** `general`: numbers right, strings left, booleans and errors
  centred; vertical default bottom. `wrap`/`justify`/`distributed` →
  `white-space: pre-wrap` clipped to the cell (rows never grow); `fill` →
  the text repeated `ceil(width / textWidth)` times and clipped;
  `centerContinuous` → centred with two-sided overflow; rotation 1–90 →
  `rotate(−θ deg)`, 91–180 → `rotate((θ − 90) deg)`, 255 → `writing-mode:
  vertical-lr; text-orientation: upright`; `shrinkToFit` → font scaled by
  `min(1, available / estimatedTextWidth)`. Estimated text width (here and in
  SPEC-118) = Σ per-character weights × MDW, weight 0.5 for `.` `,` `:` and
  space, 1 otherwise — an estimate, never a measurement.
- **Values** show through `generalText` (SPEC-118 replaces it); `showZeros:
  false` blanks numeric zeros.
- **Fills.** Default colours fg `#000000`, bg `#ffffff`. Pattern table:

| `pattern` | CSS |
| --- | --- |
| `gray0625` / `gray125` / `lightGray` / `mediumGray` / `darkGray` | solid sRGB mix of fg over bg at 6.25 / 12.5 / 25 / 50 / 75 % |
| `lightHorizontal` / `darkHorizontal` | `repeating-linear-gradient(0deg, fg 0 Lpx, bg Lpx 4px)`, L = 1 (light) or 2 (dark) |
| `lightVertical` / `darkVertical` | the same at `90deg` |
| `lightDown` / `darkDown` | the same at `45deg` |
| `lightUp` / `darkUp` | the same at `135deg` |
| `lightGrid` / `darkGrid` | horizontal + vertical layered |
| `lightTrellis` / `darkTrellis` | down + up layered |

  Gradients: linear → `linear-gradient((degree + 90)deg, stops)`; path →
  `radial-gradient(at ((left+right)/2·100)% ((top+bottom)/2·100)%, stops)`.
- **Merges.** `spreadsheet/merges.ts`: `buildMergeIndex(merges)` →
  `MergeIndex`; `mergeAt(index, row, col): GridRange | undefined` and
  `mergesIn(index, range): GridRange[]` in O(log n + k) (merges sorted by top,
  with a running max-bottom). A merge intersecting the window renders once,
  at the anchor's position, sized to the range's visible extent (hidden
  members contribute 0), `data-ref` = the top-left address; member cells
  render nothing; interior edges draw no border; outer edges come from the
  member cells on that edge.

### Border CSS and weights
Heaviest first — the shared-edge rule compares ranks, a tie (same style)
goes to the right or bottom cell's edge, i.e. the neighbour's `left`/`top`:

| Rank | Style | CSS |
| --- | --- | --- |
| 1 | thick | 3px solid |
| 2 | double | 3px double |
| 3 | medium | 2px solid |
| 4 | mediumDashed | 2px dashed |
| 5 | mediumDashDot | 2px dashed |
| 6 | slantDashDot | 2px dashed |
| 7 | mediumDashDotDot | 2px dotted |
| 8 | thin | 1px solid |
| 9 | dashed | 1px dashed |
| 10 | dashDot | 1px dashed |
| 11 | dashDotDot | 1px dotted |
| 12 | dotted | 1px dotted |
| 13 | hair | 1px dotted, 60% opacity |

`resolveEdges(window, …)` in `cellRender.ts` returns one segment per visible
resolved edge (`{ x, y, length, orientation, style, color }`); each renders as
a zero-thickness element with a CSS `border-top`/`border-left`. Automatic
colour is `#000000`.

### Overflow
`overflowExtent(sheet, row, col, align, metrics, merges)` in `cellRender.ts`
scans the row's sorted cells (binary search, at most `OVERFLOW_SCAN_COLS =
256` columns) for the nearest blocking cell on the spill side(s). Blocking:
a stored cell with `value !== null`, a merged cell, or the frozen column
boundary. Centred text extends `min(leftRoom, rightRoom)` both ways. The
overflowing cell is raised above its neighbours (`z-index`) and its text box
widened to the extent with `overflow: hidden`, so `sheet-cell[data-ref]`'s
text content stays the cell's own. For each visible row the pass also takes
the nearest stored string cell left of the window (right of it for
right-aligned and centred), so a long title anchored off-screen still paints.

### Frozen panes and hidden rows and columns
Frozen rows and columns are separate sticky layers in the scroller (as the
headers are): the corner, the top band (scrolls horizontally), the left band
(scrolls vertically) and the main area, whose window starts at the first
unfrozen row and column. `sheet-freeze-line` marks each boundary. A frozen
count beyond the scroll extent is clamped for display only. Hidden runs are
zero-sized in `GridMetrics`; a `sheet-hidden-marker` sits on the header
boundary where the run was, titled with the run ("Hidden rows 4–6",
"Hidden column G").

### Zoom
`useSheetZoom.ts`: initial = the sheet's `zoomPercent ?? 100`; steps with
`stepZoomUp`/`stepZoomDown`; the typed field is clamped with
`clampZoom(value, SHEET_ZOOM_MIN = 10, SHEET_ZOOM_MAX = 400)`. Applied as
`transform: scale(z)` (origin top left) on the content box inside the
scroller, the spacer sized to `total × z`, and the viewport window computed
from `scroll / z` — headers and frozen bands scale with it. `ZoomControl.vue`
gains `showFit` (Boolean, default true); the sheet passes false. Zoom is per
sheet, stored in `zoomPercent` as view state.

### The menu bar
`SpreadsheetMenuBar.vue` follows `PresentationMenuBar.vue`'s pattern: the
product-neutral `MenuBar.vue` with slotted menus, every item emitting a
`command` id routed by `src/views/spreadsheet/useSpreadsheetMenu.ts`, its
state coerced by `src/lib/spreadsheet/menuState.ts`. This spec's items:
File — Download (`sheet-menu-download`, the same action as
`spreadsheet-download`) and Close (`sheet-menu-close`); View — Gridlines
(`sheet-menu-gridlines`, checked from the sheet's `showGridLines`) and Zoom ▸
the `ZOOM_PRESETS` ladder (`sheet-menu-zoom-<n>`, `menuitemradio`), Zoom in
(`sheet-menu-zoom-in`), Zoom out (`sheet-menu-zoom-out`). Later specs append
menus and items; item testids keep the `sheet-menu-` prefix.

### Dark mode
SPEC-059's rule — the page is not a surface — applied to sheets: the cell
layer root has class `sheet-paper`, its background `#ffffff` and gridlines
`#d4d4d4` are literals, every cell colour comes from the model. `App.vue`'s
dim-page selector becomes `.page, .sheet-paper`. Column/row headers, the
corner, freeze lines, selection overlays and everything outside the paper use
tokens. The selection fill is `--accent` at 12% with a 2 px `--accent`
active-cell outline.

### Selection
```ts
// src/lib/spreadsheet/selection.ts
export interface SheetSelection { active: CellRef; ranges: GridRange[]; activeRangeIndex: number }
export type Direction = "up" | "down" | "left" | "right";
export interface NavContext {
  sheet: Sheet; merges: MergeIndex;
  isRowHidden(r: number): boolean; isColHidden(c: number): boolean;
  pageRows: number; pageCols: number;     // visible rows/cols for PageUp/PageDown
}
export function singleCell(ref: CellRef): SheetSelection;
export function moveActive(sel: SheetSelection, dir: Direction, nav: NavContext, steps?: number): SheetSelection;
export function extendBy(sel: SheetSelection, dir: Direction, nav: NavContext, steps?: number): SheetSelection;
export function extendTo(sel: SheetSelection, target: CellRef, nav: NavContext): SheetSelection;
export function edgeInDirection(from: CellRef, dir: Direction, nav: NavContext): CellRef;
export function currentRegion(at: CellRef, nav: NavContext): GridRange;
export function cycleWithin(sel: SheetSelection, order: "row" | "column", back: boolean, nav: NavContext): SheetSelection;
export function addRange(sel: SheetSelection, ref: CellRef, nav: NavContext): SheetSelection;
export function selectRows(sel: SheetSelection, nav: NavContext): SheetSelection;     // Shift+Space
export function selectColumns(sel: SheetSelection, nav: NavContext): SheetSelection;  // Ctrl+Space
export function selectAll(): SheetSelection;
export function snapToMerges(range: GridRange, merges: MergeIndex): GridRange;
```
Rules every function keeps:
- `active` is always inside `ranges[activeRangeIndex]`, visible, and — when
  in a merge — the merge's top-left. Ranges are always merge-snapped
  (expanded until no merge crosses their edge).
- **Moving** collapses to one cell; from a merge, a step leaves from the
  merge's far edge; entering a merge lands on its top-left. Hidden rows and
  columns are skipped; the grid edges clamp.
- **Extending** keeps `active` as the anchor and moves the range's corner
  farthest from it on each axis.
- **`edgeInDirection`**: if the current cell and the next visible cell are
  both non-empty, go to the last non-empty cell before an empty one;
  otherwise go to the next non-empty cell; if there is none, to the grid
  edge. "Non-empty" = a stored cell with `value !== null`. Rows are walked
  through `findRow` per step, cells by binary search — O(k log n).
- **`currentRegion`**: start at the active cell; repeatedly grow each side
  while the row or column just outside (extended one cell diagonally) holds a
  non-empty cell; bounded by `usedRange`.
- **`cycleWithin`** (Tab/Enter in a multi-cell selection) moves `active` to
  the next cell of the current range in row order (Tab) or column order
  (Enter), wrapping to the next range, skipping hidden cells and merge
  members; with a single-cell selection it is `moveActive` right/down.
- Worked examples (tested): in `{A1..C3 filled, E1 filled}` from `A1`,
  Ctrl+Right → `C1`, again → `E1`, again → `XFD1`; Ctrl+A from `B2` → `A1:C3`,
  again → `A1:XFD1048576`; with `C5:E6` merged, Right from `B5` → `C5`, Right
  again → `F5`, and Shift+Right from `B6` → `B6:C6` snapped to `B5:E6`.

### Keyboard map
`Mod` = Cmd on macOS, Ctrl elsewhere (`isMacPlatform()`).

| Keys | Action |
| --- | --- |
| Arrows | `moveActive` |
| Shift+Arrows | `extendBy` |
| Mod+Arrows | move to `edgeInDirection` |
| Mod+Shift+Arrows | extend to `edgeInDirection` of the moving corner |
| Home / Mod+Home | column A of the row / `A1` (first visible) |
| Mod+End | the bottom-right of `usedRange` |
| PageDown / PageUp | ± `pageRows` rows; Alt+PageDown / Alt+PageUp: ± `pageCols` columns |
| Mod+A | `currentRegion`, or `selectAll` when already that region or the active cell is isolated |
| Shift+Space / Ctrl+Space (literal Ctrl on every platform) | `selectRows` / `selectColumns` (macOS keeps Control+Space for its input-source switch when several sources are enabled; a column-header click is the guaranteed path) |
| Tab / Shift+Tab, Enter / Shift+Enter | `cycleWithin` row order / column order |
| Ctrl+PageUp / Ctrl+PageDown | previous / next sheet (SPEC-116 — best effort: ordinary Chrome/Firefox windows keep it for tab switching; the sheet tablist's arrow keys are the guaranteed path) |

The grid element is focusable (`tabindex="0"`); handlers call
`preventDefault` only for keys they consume.

### Pointer
Pointer coordinates are converted to unscaled content px (`/ zoom`, frozen
offsets added) and hit-tested with `colAt`/`rowAt`. Press → `singleCell`
(snapped); drag → `extendTo` the cell under the pointer, auto-scrolling at
`AUTO_SCROLL_PX_PER_FRAME = 16` when the pointer leaves the viewport; Shift+
press → `extendTo`; Mod+press → `addRange`. Headers: press selects the whole
row/column, drag extends across headers, Shift and Mod as for cells.
`sheet-select-all` → `selectAll`. The Name Box shows `<rows>R × <cols>C`
from press until release.

### View state
`useSheetViewState.ts` owns `activeSheetIndex` and each sheet's `selection`,
`zoomPercent`, `showGridLines` and `topLeftCell` (from the scroll position).
A change replaces the sheet object's view fields (sheet objects are never
mutated in place; `rows` keeps its identity, so the row split writes no rows
record) and schedules `storeSpreadsheetViewState(model)` — a new
`documentStore.ts` function that writes the shell record keeping the
existing `lastModified` — debounced by `VIEW_STATE_PERSIST_DELAY_MS`, flushed
on close and `pagehide`. View state never sets `Sheet.dirty` or
`structureDirty` (SPEC-121) and never calls `touchDocument`; SPEC-122's
writer serialises the current view fields only into a sheet it rebuilds for
another reason, so a clean sheet downloads byte-identical.

### Aggregates
```ts
// src/lib/spreadsheet/aggregates.ts
export interface SelectionAggregates { count: number; numericCount: number; sum: number; hasError: boolean }
export function selectionAggregates(sheet: Sheet, ranges: GridRange[], nav: Pick<NavContext, "isRowHidden" | "isColHidden">): SelectionAggregates;
```
Iterates stored cells inside the union of `ranges` (each cell once, hidden
rows and columns excluded). `count` = cells with `value !== null`;
`numericCount`/`sum` over numbers only (booleans and numeric-looking strings
excluded; dates are numbers). Shown only when the selection spans more than
one cell and `count ≥ 1`: "Count: n" always, "Average: …" and "Sum: …" only
when `numericCount ≥ 1` and `!hasError`; numbers through `generalText` here
(SPEC-118 formats them).

### Reader version
Both readers' extensions bump `SPREADSHEET_READER_VERSION` (to the next
integer) and append a `ReaderUpgrade` whose `upgradeDirtySheet` throws
`Error("unreachable: no sheet is dirty before SPEC-121")`. Rationale: a
record behind this version is re-read whole from `opaque` (no sheet can be
dirty before SPEC-121 lands, and every read path upgrades), and a step that
throws keeps `READER_UPGRADES` contiguous while making the impossible case
loud instead of silently merging half a style table.

### Fixtures
- `spec117-styles.xlsx` (Excel 365-shaped, `mc:Ignorable="x14ac xr xr2
  xr3"`): a theme part; `styles.xml` with tinted theme colours, indexed
  colours plus an `indexedColors` override, `gray125` and `darkTrellis`
  patterns, a linear gradient, all thirteen border styles and a diagonal,
  `cellStyles` "Normal"/"Good"/"Heading 1", a `quotePrefix` xf, `dxfs` and
  `extLst` (untouched). Sheet "Styled": `B2` bold red on yellow; `B3:B12`
  numbers (`B3:B5` = 10, 20, 30); a merged `C5:E5` title; column G hidden;
  `customWidth`/`bestFit` columns; `defaultRowHeight="15"
  x14ac:dyDescent="0.25"`; rows with `ht` + `customHeight`, one hidden; a
  style-only cell; a frozen pane (`xSplit="1" ySplit="2" topLeftCell="B3"
  state="frozen"`); `zoomScale="85"`; a two-range `sqref`; a theme `tabColor`;
  a rich shared string with two run fonts. Sheet "Plain": `showGridLines="0"`.
  Sheet "Split": a `state="split"` pane.
- `spec117-styles.ods` (LibreOffice 24-shaped): `ce1…`/`co1…`/`ro1…`/`ta1`
  automatic styles (`ta1` with `tableooo:tab-color`); a `Default`-parented
  bold red-on-yellow cell; borders `0.74pt solid #000000`, `1.75pt solid`,
  `2.5pt solid`, `0.74pt dashed`, `0.74pt fine-dashed`, a `double`; wrap,
  rotation 90, `repeat-content`; a 3×1 span over styled covered cells; column G
  `table:visibility="collapse"`; `table:default-cell-style-name` columns; a
  styled `number-columns-repeated="16380"` trailing run and the trailing
  `1048553`-row run; a `text:span` run; `settings.xml` with both split modes
  2, `ZoomValue` 85, `ShowGrid` false on the second table, a cursor position.

### Tests
Every rule and table in this Reference has a unit test in the files the
Tests AC lists: both style mapping tables row by row, the theme swap and tint
arithmetic both ways, every layout rule (the split-pane exclusion,
covered-cell storage, the styled repeat run), the MDW table and width
conversions, `GridMetrics` against a naive reference on random sheets, every
selection rule and worked example, CSS per style, the pattern and border
tables, the shared-edge rule and overflow extents. Round trip uses part-level
`equalBytes` (`core/equality.ts`), as SPEC-116 does.

### Specs this changes
- `specs/116-…` — style-only cells and ODS covered cells with content are
  now stored (SPEC-116's reader dropped both); `SpreadsheetModel`, `Sheet`,
  `Row`, `Cell` and `TextRun` gain the fields in "Model additions" (`styles`,
  `columns`, `merges` and the two defaults required, the rest optional); the
  uniform 64 × 20 grid becomes `GridMetrics`; the reader version moves on.
- `specs/063-…` — `ZoomControl.vue` gains `showFit`; the zoom vocabulary is
  shared by sheets.
- `specs/059-…` — the dim-page rule covers `.sheet-paper`.
- `specs/058-…` — `tokens.spec.ts` exempts `components/spreadsheet/SheetCellLayer.vue`.
- `specs/077-…` — the theme-part reader moves to `core/themeDrawingMl.ts`
  (`readDocxTheme` delegates); `themeGuard.spec.ts`'s allow-list gains the
  spreadsheet colour codec files.
