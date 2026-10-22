---
id: SPEC-109
title: Combine images into one image
status: draft
area: structure
complexity: normal
story_points: 8
depends_on: [SPEC-058, SPEC-063, SPEC-064, SPEC-107, SPEC-108]
---
## Why
Four screenshots that belong in one image. A before and an after, side by
side. Six photos as a contact sheet. A logo dropped into the corner of a
banner. These are the errands people open a web tool for, paste their files
into a stranger's server for, and get a watermark back from.

SPEC-107 put the images in the library and SPEC-108 defined the rectangle
they land on. What is missing is arrangement: more than one item on one
sheet, in an order the user controls, with a preview that is the output.

This file adds the **assembly workspace** — `/assemble`, a tray of chosen
library items, a layout, a sheet, a live preview and one output. It is the
normative source for the workspace, for `AssemblyPlan`, for the layout
engine (`core/assemblyLayout.ts`) and for raster output; SPEC-110 adds PDF as
an output target and SPEC-111 adds PDF pages as inputs, both into this same
surface rather than beside it.

## Acceptance criteria
- [ ] `/assemble` (`AssembleView.vue`, route name `assemble`) is the workspace: a left **tray** of chosen items in order, a centre **preview** of the sheet at fit-to-window, a right **inspector** holding the layout controls above SPEC-108's `SheetSetupPanel`, and a bottom **output bar** with the format, the estimated size and the two actions; it is reachable from the library selection bar ("Combine…", `bulk-combine`, enabled with two or more image entries selected) and from the library header ("New › Combined image", opening an empty tray)
- [ ] The tray (`assembly-tray`, rows `assembly-item`) holds up to `MAX_ASSEMBLY_ITEMS = 100` entries, each showing its thumbnail, title and pixel size, and supports reorder by drag **and** by keyboard (`Alt+↑`/`Alt+↓` on a focused row, announced through a `role="status"` live region), removal (`item-remove`), duplication (`item-duplicate`) and "Add from library…" (`assembly-add`) opening a picker of image entries; the tray order is the layout order, and nothing else determines it
- [ ] `src/lib/core/assemblyLayout.ts` exports `AssemblyPlan`, `LAYOUT_KINDS` and `planAssembly(items, plan, sheet): PlacedItem[]` per the Reference — pure, framework-free, centimetres in and out, given only each item's pixel size and never its bytes
- [ ] Five layouts, each specified rather than named: `grid` (`columns: number | "auto"`, `rows` derived, `auto` choosing the column count that minimises wasted area for the item count and sheet aspect), `horizontal` (one row, items scaled to a common height), `vertical` (one column, common width), `masonry` (columns of equal width, each item placed in the currently shortest column, preserving aspect) and `overlay` (every item on the whole content box, in tray order back to front, each with its own placement)
- [ ] Spacing and cell behaviour are explicit: `gapCm` between cells (default 0.2, applied between cells only, never at the content-box edge), the SPEC-108 `fit` and `align` applied **within each cell** rather than to the sheet, and `cellBackground` (`"transparent"` by default) drawn behind each item so a `contain` fit can letterbox in a chosen colour
- [ ] `overlay` gives each item its own transform: `xPercent`/`yPercent` of the content box, `scalePercent`, `opacity` (0–100) and `blend` (`normal` | `multiply` | `screen` | `overlay` — the four `globalCompositeOperation` values every target browser implements identically), editable from the inspector when an overlay item is selected and by dragging the item in the preview, with `Shift` constraining to an axis and arrow keys nudging by 1 mm (`Shift`+arrow by 1 cm)
- [ ] The preview **is** the renderer: `src/lib/media/composeCanvas.ts` exports `composeToCanvas(canvas, placed, sheet, sources)`, and both the on-screen preview and the export call it with different pixel sizes — there is no second drawing path that could disagree, and a `data-testid="assembly-preview"` canvas carries the placed rectangles as an ARIA-labelled list beside it for tests and for screen readers
- [ ] Output is PNG or JPEG: `assembly-format` (`png` default, `jpeg` with `assembly-quality` 1–100 defaulting to 90 and shown only for JPEG), the sheet's `transparent` background allowed for PNG and refused for JPEG (falling back to the chosen colour, with the panel saying so before the user commits), the pixel size from SPEC-108's `sheetPixelSize(geometry, dpi)`, and a live `assembly-estimate` of output dimensions and approximate bytes
- [ ] The canvas ceiling is measured, not assumed: `probeMaxCanvasPixels()` allocates and verifies once per session (fill, read back one pixel, halve and retry), caching the result; a plan whose output exceeds it is refused **before** any work with a message naming the limit and the dpi that would fit, and `MAX_OUTPUT_EDGE_PX = 16384` bounds either axis independently
- [ ] Producing the output never blocks the page: sources are decoded once and cached for the session, drawing is chunked row-by-row (or item-by-item for `overlay`) with `yieldToMainThread()` between chunks, `assembly-progress` reports `Item {done} of {total}` with `role="status"`, `assembly-cancel` stops at the next boundary leaving no file and no library record, and every intermediate canvas is released (`canvas.width = 0`) after use
- [ ] The result lands as a new library item or a download: "Save to library" (`assembly-save`) creates a new SPEC-107 media document (format `png`/`jpeg`, its `image` summary read back from the produced bytes with `decodeImageDimensions`, title from `assembly-title` defaulting to "Combined image" under SPEC-003's duplicate numbering), "Download" (`assembly-download`) goes through `core/triggerDownload.ts`, and **no source item is read, moved, renamed or modified by either** — asserted against deep-frozen inputs
- [ ] The plan survives a reload: the workspace's plan (item ids, layout, per-item transforms, sheet setup, output settings — never pixels) is persisted under one key in the existing preferences store and restored on return, with "Start over" (`assembly-reset`) clearing it; an item that has since been trashed or deleted is dropped from the restored tray with a one-line notice naming it
- [ ] Determinism is asserted at the level it exists: `planAssembly` is pure and its rectangles are exact to 3 decimals; the rendered bytes are **not** asserted byte-identical — canvas encoders differ across browsers and platforms, and pretending otherwise would produce a test that fails on someone's machine for no defect — so the render tests assert output dimensions, the alpha of a background sample, and the colour of one pixel at the centre of each placed item, sampled from `getImageData`
- [ ] Tests: `src/lib/core/assemblyLayout.spec.ts` — every layout against 1, 2, 3, 5 and 12 items on a wide and a tall sheet with exact rectangles, `auto` column selection, gap arithmetic at the edges, overlay transforms and z-order, and the empty-tray case (an empty plan places nothing and is not an error); `src/lib/media/composeCanvas.spec.ts` — the pixel-sample assertions above under jsdom's canvas stub where possible and e2e otherwise; `AssembleView.spec.ts` — tray reordering by keyboard, the refusal paths, no source mutation; e2e (`e2e/SPEC-109.spec.ts`): select `spec107-shot.png` and `spec107-photo.jpg` → "Combine…" → grid, 2 columns, A4 landscape, Narrow margins → the preview reports two placed rectangles → Save to library → a new `PNG` row whose details panel reports the expected pixel size → reorder the tray and save again → the second output differs from the first at the sampled centre pixels

## Out of scope
PDF output (SPEC-110) and PDF pages as inputs (SPEC-111) — both land in this
same workspace. Text, captions, labels, arrows or watermarks on the sheet:
annotation belongs to the image editor (SPEC-112), and a combined image can
be opened there afterwards. Per-item cropping and rotation in the tray —
SPEC-112 owns editing, and an edited image arrives here already edited.
Freeform placement outside the `overlay` layout's transform vocabulary, snap
guides, and alignment tools between items. Animated GIF output, APNG, WEBP
output, SVG output, and any vector composition (SVG sources are rasterised at
their placed size, as SPEC-108 states). Feathering, shadows, borders, rounded
corners and any filter on a placed item beyond `opacity` and the four blend
modes. Saving a plan as a reusable named template. Batch mode over folders.

## Reference

### Dependencies restated
- `specs/108-…` — `SheetSetup`/`SheetGeometry`, `resolveSheet(setup,
  itemAspect?)`, `placeOnSheet(itemPx, box, fit, align)`,
  `sheetPixelSize(geometry, dpi)`, `SHEET_DPI_CHOICES`, `MIN_CONTENT_CM`, the
  `r3` rounding rule that keeps preview and output in agreement, and
  `SheetSetupPanel.vue` with its testids. This spec calls `placeOnSheet` **per
  cell**, passing the cell rectangle as the box.
- `specs/107-…` — `MediaModel`/`MediaFormat`, `model.media` holding the source
  `MediaPart`, `mediaDataUri(part)` for decoding, `decodeImageDimensions`,
  the library's kind-aware rows, and the rule that a stored payload is never
  loaded to list a row.
- `src/lib/core/idleScheduler.ts` — `yieldToMainThread()` and `runWhenIdle`,
  whose contract includes not keeping a Node process alive (SPEC-047).
- `src/lib/core/svgToPng.ts` — `convertSvgToPng(bytes)`, the browser-bound
  rasteriser core already permits, used for an SVG source at its placed size.

### The plan
```ts
// src/lib/core/assemblyLayout.ts
export type LayoutKind = "grid" | "horizontal" | "vertical" | "masonry" | "overlay";
export const LAYOUT_KINDS: readonly LayoutKind[];

export interface AssemblyItem {
  /** The library id of the source media document. */
  id: string;
  widthPx: number;
  heightPx: number;
  /** overlay only; ignored by every other layout. */
  transform?: { xPercent: number; yPercent: number; scalePercent: number; opacity: number; blend: BlendMode };
}

export interface AssemblyPlan {
  layout: LayoutKind;
  columns: number | "auto";      // grid and masonry
  gapCm: number;
  fit: SheetFit;                 // applied per cell
  align: SheetAlign;             // applied per cell
  cellBackground: string;        // "#rrggbb" | "transparent"
}

export interface PlacedItem {
  index: number;                 // tray order, and the back-to-front order for overlay
  xCm: number; yCm: number; widthCm: number; heightCm: number;
  /** The source rectangle in image pixels, for the cropping fits. */
  sourcePx: { x: number; y: number; width: number; height: number };
  opacity: number;               // 1 outside overlay
  blend: BlendMode;              // "normal" outside overlay
}

export function planAssembly(
  items: AssemblyItem[], plan: AssemblyPlan, sheet: SheetGeometry,
): PlacedItem[];
```

`auto` columns: choose `c` in `1..items.length` minimising the wasted area of
the content box when `ceil(n / c)` rows of `c` cells each take the box's
aspect, breaking ties towards the more square grid and then towards fewer
columns. It is a closed formula, tested at 1, 2, 3, 5 and 12 items on two
sheet aspects, so "auto" is a specification and not a heuristic that drifts.

`masonry` fills columns left to right, always placing the next item in the
column whose current height is smallest, ties going to the leftmost — a
deterministic rule, not a packing search.

### Rendering
`src/lib/media/composeCanvas.ts` is browser-bound (canvas + image decoding),
the same exception `core/svgToPng.ts` and `pdf/imageEncode.ts` already take,
and holds no Vue import. It draws: the sheet background, then per placed
item — save, set `globalAlpha` and `globalCompositeOperation`, fill the cell
background, `drawImage(source, sourcePx…, destCm→px…)`, restore. Sources are
decoded through one `decodeSource(part)` cache keyed by `mediaId`; an SVG
part goes through `convertSvgToPng` at the placed pixel size, so a vector
logo is sharp at 600 dpi rather than upscaled from 300 × 150.

Export re-runs the same function against an off-screen canvas at
`sheetPixelSize`, then `canvas.toBlob(type, quality)`. The preview runs it at
the on-screen scale. One function, two sizes: what you see is what is
written, and there is no second code path to keep in step.

### The canvas ceiling
Browsers disagree about the largest canvas they will allocate, and the ones
that refuse do it by returning a blank surface rather than throwing — an
export that silently produces a white PNG is the worst possible outcome. So
`probeMaxCanvasPixels()` measures it once per session: allocate a candidate
area, fill one pixel, read it back with `getImageData`, and halve on failure
down to a floor. The result is cached, exposed in the estimate line, and used
to refuse a plan **before** any decoding — with a message naming both the
limit and the dpi that would fit, so the fix is one click.

### Tests
- `src/lib/core/assemblyLayout.spec.ts` as in the acceptance criteria.
- `src/lib/media/composeCanvas.spec.ts` — source-rectangle arithmetic for
  `cover`/`actual` (pure, asserted from `planAssembly` output), the decode
  cache's single-decode property, and SVG rasterisation at the placed size.
- `src/views/AssembleView.spec.ts` + `useAssembly.spec.ts` — tray reordering
  (pointer and keyboard), persistence and restoration including the
  deleted-item notice, the refusal paths (canvas ceiling, JPEG +
  transparency, over `MAX_ASSEMBLY_ITEMS`), cancellation leaving nothing
  behind, and frozen-source assertions.
- `e2e/SPEC-109.spec.ts` as in the acceptance criteria.

### Specs this changes
- `specs/067-…` — the selection bar gains "Combine…" and the header's New
  menu gains "Combined image"; both are inert without image entries.
- `specs/066-…` — one preferences key for the persisted assembly plan.
- `src/router.ts` gains `/assemble`; it takes no id and needs no kind guard.
- `specs/000-roadmap.md` gains this spec's phase row.
