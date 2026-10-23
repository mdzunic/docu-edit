// The assembly layout engine (SPEC-109): where each item of an assembly tray
// lands on SPEC-108's sheet. Five layouts, each a stated rule rather than a
// name — grid, horizontal, vertical, masonry and overlay — given only every
// item's pixel size (never its bytes) and the resolved sheet, and answering
// in centimetres. The rendered image (media/composeCanvas.ts) and the
// workspace's ARIA list both read the same PlacedItem[]; nothing else decides
// where an item goes, and the tray order is the only order there is.
//
// Every cell is a box SPEC-108's `placeOnSheet` fits ONE item into, with the
// plan's fit and align — so contain letterboxes, cover crops and stretch
// fills within the cell exactly as they do on a whole sheet. Gaps separate
// cells and never touch the content box's edge.
//
// Pure maths: no DOM, no canvas, no Date, no randomness, sibling core
// imports only (boundary.spec.ts). Every length goes through SPEC-108's `r3`,
// so the preview and the export, handed the same plan, agree to the last bit.
import {
  DEFAULT_SHEET_DPI,
  placeOnSheet,
  r3,
  TRANSPARENT,
  type ItemPx,
  type RectCm,
  type SheetAlign,
  type SheetFit,
  type SheetGeometry,
} from "./sheetLayout";

export type LayoutKind = "grid" | "horizontal" | "vertical" | "masonry" | "overlay";

export const LAYOUT_KINDS: readonly LayoutKind[] = ["grid", "horizontal", "vertical", "masonry", "overlay"];

/** The four `globalCompositeOperation` values every target browser draws
 * identically ("normal" is source-over). */
export type BlendMode = "normal" | "multiply" | "screen" | "overlay";

export const BLEND_MODES: readonly BlendMode[] = ["normal", "multiply", "screen", "overlay"];

/** The most items one tray holds. */
export const MAX_ASSEMBLY_ITEMS = 100;

export const DEFAULT_GAP_CM = 0.2;

/** The widest gap the inspector accepts. */
export const MAX_GAP_CM = 10;

export const MIN_OVERLAY_SCALE_PERCENT = 1;
export const MAX_OVERLAY_SCALE_PERCENT = 1000;

/** An overlay item's own placement. The item's centre sits at
 * (xPercent, yPercent) of the content box; at scalePercent 100 it is the
 * size the plan's fit gives it on the whole content box. */
export interface OverlayTransform {
  xPercent: number;
  yPercent: number;
  scalePercent: number;
  /** 0 (invisible) – 100 (opaque). */
  opacity: number;
  blend: BlendMode;
}

/** Centred, at the fit's own size, opaque, normal: every item on the whole
 * content box. */
export const DEFAULT_OVERLAY_TRANSFORM: OverlayTransform = {
  xPercent: 50,
  yPercent: 50,
  scalePercent: 100,
  opacity: 100,
  blend: "normal",
};

export interface AssemblyItem {
  /** The library id of the source media document. */
  id: string;
  widthPx: number;
  heightPx: number;
  /** overlay only; ignored by every other layout. */
  transform?: OverlayTransform;
}

export type AssemblyColumns = number | "auto";

export interface AssemblyPlan {
  layout: LayoutKind;
  /** grid and masonry. */
  columns: AssemblyColumns;
  gapCm: number;
  /** Applied per cell. */
  fit: SheetFit;
  /** Applied per cell (overlay positions by its transforms instead). */
  align: SheetAlign;
  /** "#rrggbb" | "transparent". */
  cellBackground: string;
}

export const ASSEMBLY_PLAN_DEFAULTS: AssemblyPlan = {
  layout: "grid",
  columns: "auto",
  gapCm: DEFAULT_GAP_CM,
  fit: "contain",
  align: "center",
  cellBackground: TRANSPARENT,
};

export interface PlacedItem {
  /** Tray order, and the back-to-front order for overlay. */
  index: number;
  /** The visible placed rectangle, clipped to its cell. */
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
  /** The source rectangle in image pixels, for the cropping fits. */
  sourcePx: { x: number; y: number; width: number; height: number };
  /** 0–1; 1 outside overlay. */
  opacity: number;
  /** "normal" outside overlay. */
  blend: BlendMode;
  /** The cell the item was fitted in — what `cellBackground` fills, so a
   * contain fit letterboxes in it. For overlay, the item's own visible
   * rectangle. */
  cell: RectCm;
  cellBackground: string;
  /** The drawing chunk the item belongs to: its grid row, 0 for the one
   * horizontal row, its own index for vertical and overlay, its depth in
   * its column for masonry. Rendering yields between chunks. */
  row: number;
}

function r3Rect(rect: RectCm): RectCm {
  return { xCm: r3(rect.xCm), yCm: r3(rect.yCm), widthCm: r3(rect.widthCm), heightCm: r3(rect.heightCm) };
}

function clampNumber(value: number, min: number, max: number, fallback: number): number {
  return Number.isFinite(value) ? Math.min(max, Math.max(min, value)) : fallback;
}

/** A transform with every field inside its range: x/y 0–100 (the centre
 * stays on the content box), scale 1–1000, opacity 0–100, a known blend. */
export function normaliseTransform(transform: Partial<OverlayTransform> | undefined): OverlayTransform {
  const base = DEFAULT_OVERLAY_TRANSFORM;
  const t = transform ?? {};
  return {
    xPercent: r3(clampNumber(Number(t.xPercent), 0, 100, base.xPercent)),
    yPercent: r3(clampNumber(Number(t.yPercent), 0, 100, base.yPercent)),
    scalePercent: r3(
      clampNumber(Number(t.scalePercent), MIN_OVERLAY_SCALE_PERCENT, MAX_OVERLAY_SCALE_PERCENT, base.scalePercent),
    ),
    opacity: r3(clampNumber(Number(t.opacity), 0, 100, base.opacity)),
    blend: BLEND_MODES.includes(t.blend as BlendMode) ? (t.blend as BlendMode) : base.blend,
  };
}

/** An item's width over its height; 1 for an item whose size is unknown. */
function aspectOf(item: ItemPx): number {
  return item.widthPx > 0 && item.heightPx > 0 ? item.widthPx / item.heightPx : 1;
}

/** "auto" columns: the `c` in 1..count minimising the content box's wasted
 * area when ceil(count / c) rows of c cells fill the box and each cell holds
 * one square item (side min(W / c, H / rows)) — so the answer depends on the
 * item count and the box's aspect only. Ties go to the more square grid
 * (smaller |c − rows|), then to fewer columns. */
export function autoColumns(count: number, box: { widthCm: number; heightCm: number }): number {
  if (count <= 1) {
    return 1;
  }
  const width = Math.max(0, box.widthCm);
  const height = Math.max(0, box.heightCm);
  let best = 1;
  let bestWaste = Number.POSITIVE_INFINITY;
  let bestSkew = Number.POSITIVE_INFINITY;
  for (let columns = 1; columns <= count; columns += 1) {
    const rows = Math.ceil(count / columns);
    const side = Math.min(width / columns, height / rows);
    const waste = r3(width * height - count * side * side);
    const skew = Math.abs(columns - rows);
    if (waste < bestWaste || (waste === bestWaste && skew < bestSkew)) {
      best = columns;
      bestWaste = waste;
      bestSkew = skew;
    }
  }
  return best;
}

/** The column count a grid or masonry plan uses for `count` items: a fixed
 * number (a whole number, 1..MAX_ASSEMBLY_ITEMS), or autoColumns. */
export function resolveColumns(columns: AssemblyColumns, count: number, box: { widthCm: number; heightCm: number }): number {
  if (columns === "auto") {
    return autoColumns(count, box);
  }
  return Math.round(clampNumber(columns, 1, MAX_ASSEMBLY_ITEMS, 1));
}

// The cell builders return exact (unrounded) cells: the item is fitted into
// the exact cell and only the results are rounded, so items that share a
// height or width share it to the last decimal.

/** The grid's cells, row by row from the top-left. */
function gridCells(count: number, columns: number, box: RectCm, gap: number): { cells: RectCm[]; rows: number[] } {
  const rowCount = Math.ceil(count / columns);
  const cellW = Math.max(0, (box.widthCm - (columns - 1) * gap) / columns);
  const cellH = Math.max(0, (box.heightCm - (rowCount - 1) * gap) / rowCount);
  const cells: RectCm[] = [];
  const rows: number[] = [];
  for (let i = 0; i < count; i += 1) {
    const row = Math.floor(i / columns);
    cells.push({ xCm: box.xCm + (i % columns) * (cellW + gap), yCm: box.yCm + row * (cellH + gap), widthCm: cellW, heightCm: cellH });
    rows.push(row);
  }
  return { cells, rows };
}

/** One strip of cells filling the box along one axis, each cell's share in
 * proportion to `weights` — the items' aspects across a row, their inverse
 * aspects down a column — so a contain fit gives every item the same height
 * (horizontal) or width (vertical). */
function stripCells(weights: number[], box: RectCm, gap: number, across: boolean): RectCm[] {
  const total = weights.reduce((sum, weight) => sum + weight, 0);
  const length = across ? box.widthCm : box.heightCm;
  const free = Math.max(0, length - (weights.length - 1) * gap);
  let offset = 0;
  return weights.map((weight) => {
    const share = total > 0 ? (free * weight) / total : 0;
    const cell = across
      ? { xCm: box.xCm + offset, yCm: box.yCm, widthCm: share, heightCm: box.heightCm }
      : { xCm: box.xCm, yCm: box.yCm + offset, widthCm: box.widthCm, heightCm: share };
    offset += share + gap;
    return cell;
  });
}

/** Masonry: `columns` columns of equal width; each item, in tray order, goes
 * to the column whose current height is smallest (ties to the leftmost), at
 * the column width and its own aspect. A stack taller than the box scales
 * down uniformly (gaps keep their size) and centres across the box. */
function masonryCells(items: ItemPx[], columns: number, box: RectCm, gap: number): { cells: RectCm[]; rows: number[] } {
  const colW = Math.max(0, (box.widthCm - (columns - 1) * gap) / columns);
  const heights = new Array<number>(columns).fill(0);
  const counts = new Array<number>(columns).fill(0);
  const sums = new Array<number>(columns).fill(0);
  const slots = items.map((item) => {
    let column = 0;
    for (let c = 1; c < columns; c += 1) {
      if (heights[c] < heights[column]) {
        column = c;
      }
    }
    const height = colW / aspectOf(item);
    const depth = counts[column];
    heights[column] += (depth > 0 ? gap : 0) + height;
    counts[column] += 1;
    sums[column] += height;
    return { column, depth, height };
  });
  let scale = 1;
  for (let c = 0; c < columns; c += 1) {
    if (counts[c] > 0 && heights[c] > box.heightCm && sums[c] > 0) {
      scale = Math.min(scale, Math.max(0, box.heightCm - (counts[c] - 1) * gap) / sums[c]);
    }
  }
  const width = colW * scale;
  const blockW = columns * width + (columns - 1) * gap;
  const left = box.xCm + Math.max(0, (box.widthCm - blockW) / 2);
  const tops = new Array<number>(columns).fill(box.yCm);
  const cells = slots.map((slot) => {
    const height = slot.height * scale;
    const cell = { xCm: left + slot.column * (width + gap), yCm: tops[slot.column], widthCm: width, heightCm: height };
    tops[slot.column] += height + gap;
    return cell;
  });
  return { cells, rows: slots.map((slot) => slot.depth) };
}

/** An item rectangle clipped to `box`: the visible part and the source
 * pixels that survive (the overlay counterpart of placeOnSheet's crop). */
function clipToBox(item: RectCm, box: RectCm, itemPx: ItemPx): { visible: RectCm; sourcePx: PlacedItem["sourcePx"] } {
  const left = Math.max(item.xCm, box.xCm);
  const top = Math.max(item.yCm, box.yCm);
  const right = Math.min(item.xCm + item.widthCm, box.xCm + box.widthCm);
  const bottom = Math.min(item.yCm + item.heightCm, box.yCm + box.heightCm);
  const visibleW = Math.max(0, right - left);
  const visibleH = Math.max(0, bottom - top);
  const pxPerCmX = item.widthCm > 0 ? Math.max(0, itemPx.widthPx) / item.widthCm : 0;
  const pxPerCmY = item.heightCm > 0 ? Math.max(0, itemPx.heightPx) / item.heightCm : 0;
  return {
    visible: r3Rect({ xCm: visibleW > 0 ? left : item.xCm, yCm: visibleH > 0 ? top : item.yCm, widthCm: visibleW, heightCm: visibleH }),
    sourcePx: {
      x: r3(visibleW > 0 ? (left - item.xCm) * pxPerCmX : 0),
      y: r3(visibleH > 0 ? (top - item.yCm) * pxPerCmY : 0),
      width: r3(visibleW * pxPerCmX),
      height: r3(visibleH * pxPerCmY),
    },
  };
}

function overlayItem(item: AssemblyItem, index: number, plan: AssemblyPlan, box: RectCm, dpi: number): PlacedItem {
  const t = normaliseTransform(item.transform);
  // The fit decides the item's size on the whole content box; the
  // transform then scales it and moves its centre.
  const base = placeOnSheet(item, box, plan.fit, "center", dpi).item;
  const widthCm = (base.widthCm * t.scalePercent) / 100;
  const heightCm = (base.heightCm * t.scalePercent) / 100;
  const centreX = box.xCm + (box.widthCm * t.xPercent) / 100;
  const centreY = box.yCm + (box.heightCm * t.yPercent) / 100;
  const { visible, sourcePx } = clipToBox(
    { xCm: centreX - widthCm / 2, yCm: centreY - heightCm / 2, widthCm, heightCm },
    box,
    item,
  );
  return {
    index,
    ...visible,
    sourcePx,
    opacity: r3(t.opacity / 100),
    blend: t.blend,
    cell: visible,
    cellBackground: plan.cellBackground,
    row: index,
  };
}

function cellsFor(items: AssemblyItem[], plan: AssemblyPlan, box: RectCm, gap: number): { cells: RectCm[]; rows: number[] } {
  switch (plan.layout) {
    case "horizontal":
      return { cells: stripCells(items.map(aspectOf), box, gap, true), rows: items.map(() => 0) };
    case "vertical":
      return { cells: stripCells(items.map((item) => 1 / aspectOf(item)), box, gap, false), rows: items.map((_, i) => i) };
    case "masonry":
      return masonryCells(items, resolveColumns(plan.columns, items.length, box), box, gap);
    default:
      return gridCells(items.length, resolveColumns(plan.columns, items.length, box), box, gap);
  }
}

/** Where every item lands, in tray order. An empty tray places nothing. `dpi`
 * matters only to the "actual" fit (one image pixel per dot), as in
 * placeOnSheet. */
export function planAssembly(
  items: AssemblyItem[],
  plan: AssemblyPlan,
  sheet: SheetGeometry,
  dpi: number = DEFAULT_SHEET_DPI,
): PlacedItem[] {
  if (items.length === 0) {
    return [];
  }
  const box = sheet.content;
  if (plan.layout === "overlay") {
    return items.map((item, index) => overlayItem(item, index, plan, box, dpi));
  }
  const gap = clampNumber(plan.gapCm, 0, MAX_GAP_CM, DEFAULT_GAP_CM);
  const { cells, rows } = cellsFor(items, plan, box, gap);
  return items.map((item, index) => {
    const cell = cells[index];
    const placement = placeOnSheet(item, cell, plan.fit, plan.align, dpi);
    return {
      index,
      xCm: placement.xCm,
      yCm: placement.yCm,
      widthCm: placement.widthCm,
      heightCm: placement.heightCm,
      sourcePx: {
        x: placement.crop.xPx,
        y: placement.crop.yPx,
        width: placement.crop.widthPx,
        height: placement.crop.heightPx,
      },
      opacity: 1,
      blend: "normal",
      cell: r3Rect(cell),
      cellBackground: plan.cellBackground,
      row: rows[index],
    };
  });
}

/** The arrangement's own pixel size, for SPEC-108's "Fit to content": the
 * sheet becomes this many pixels at its dpi (plus margins). Grid: columns ×
 * the widest item by rows × the tallest; horizontal: every item at the
 * tallest item's height, side by side; vertical: every item at the widest
 * item's width, stacked; masonry: columns of the widest item's width by the
 * tallest column; overlay: the widest by the tallest. Gaps are not added —
 * they come out of the box at the sheet's scale. "auto" columns decide on a
 * square box, the content's shape being what is being decided. Null for an
 * empty tray. */
export function naturalAssemblyPx(items: AssemblyItem[], plan: AssemblyPlan): ItemPx | null {
  if (items.length === 0) {
    return null;
  }
  const sizes = items.map((item) => ({ widthPx: Math.max(1, item.widthPx), heightPx: Math.max(1, item.heightPx) }));
  const maxW = Math.max(...sizes.map((size) => size.widthPx));
  const maxH = Math.max(...sizes.map((size) => size.heightPx));
  const square = { widthCm: 1, heightCm: 1 };
  let widthPx = maxW;
  let heightPx = maxH;
  if (plan.layout === "grid") {
    const columns = resolveColumns(plan.columns, sizes.length, square);
    widthPx = columns * maxW;
    heightPx = Math.ceil(sizes.length / columns) * maxH;
  } else if (plan.layout === "horizontal") {
    widthPx = sizes.reduce((sum, size) => sum + (size.widthPx * maxH) / size.heightPx, 0);
  } else if (plan.layout === "vertical") {
    heightPx = sizes.reduce((sum, size) => sum + (size.heightPx * maxW) / size.widthPx, 0);
  } else if (plan.layout === "masonry") {
    const columns = resolveColumns(plan.columns, sizes.length, square);
    const heights = new Array<number>(columns).fill(0);
    for (const size of sizes) {
      let column = 0;
      for (let c = 1; c < columns; c += 1) {
        if (heights[c] < heights[column]) {
          column = c;
        }
      }
      heights[column] += (size.heightPx * maxW) / size.widthPx;
    }
    widthPx = columns * maxW;
    heightPx = Math.max(...heights);
  }
  return { widthPx: Math.max(1, Math.round(widthPx)), heightPx: Math.max(1, Math.round(heightPx)) };
}
