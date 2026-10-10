// The assembly sheet (SPEC-108): the one normative geometry every assembled
// output lands on — a converted image's PDF page today, the merged image,
// photo PDF and inserted page of SPEC-109..111 next. Six answers, given once:
// how big the sheet is (SPEC-027's paper catalogue, "Custom", or "Fit to
// content"), which way round, how much margin, how an item whose shape does
// not match the content box is fitted and aligned, what fills the leftover,
// and at what resolution.
//
// Pure maths in centimetres: no DOM, no canvas, no Date, no randomness, and
// only sibling core imports (boundary.spec.ts). Every length this module
// returns goes through `r3`, so the preview and the writer, handed the same
// setup, agree to the last bit rather than each rounding on its own.
import {
  orientationOf,
  pageSizeFor,
  PAPER_SIZE_NAMES,
  swapOrientation,
  type Margins,
  type Orientation,
  type PageSize,
  type PaperSizeName,
} from "./pageSetup";
import { CM_PER_INCH } from "./units";

/** The paper option owned here: a sheet sized to the content itself. */
export const FIT_TO_CONTENT = "Fit to content";

export type SheetPaper = PaperSizeName | typeof FIT_TO_CONTENT;

/** Every paper the setup panel offers, in order: SPEC-027's catalogue (with
 * its "Custom" escape hatch), then "Fit to content". */
export const SHEET_PAPER_NAMES: readonly SheetPaper[] = [...PAPER_SIZE_NAMES, FIT_TO_CONTENT];

export type SheetFit = "contain" | "cover" | "stretch" | "actual";

export const SHEET_FITS: readonly SheetFit[] = ["contain", "cover", "stretch", "actual"];

export type SheetAlign =
  | "top-left" | "top" | "top-right"
  | "left" | "center" | "right"
  | "bottom-left" | "bottom" | "bottom-right";

/** The nine-point grid, row by row from the top-left. */
export const SHEET_ALIGNS: readonly SheetAlign[] = [
  "top-left", "top", "top-right",
  "left", "center", "right",
  "bottom-left", "bottom", "bottom-right",
];

export interface SheetSetup {
  /** A PAPER_SIZE_NAMES entry, "Custom", or "Fit to content". */
  paper: SheetPaper;
  /** Custom only. */
  widthCm?: number;
  /** Custom only. */
  heightCm?: number;
  /** Ignored for "Fit to content": the content's own shape decides. */
  orientation: Orientation;
  margins: Margins;
  fit: SheetFit;
  align: SheetAlign;
  /** "#rrggbb" or "transparent". */
  background: string;
  /** One of SHEET_DPI_CHOICES. */
  dpi: number;
}

/** A rectangle in centimetres, measured from the sheet's top-left. */
export interface RectCm {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

export interface SheetGeometry {
  widthCm: number;
  heightCm: number;
  /** The box inside the margins that items are placed in. */
  content: RectCm;
}

/** An item's own size in pixels — an image's natural dimensions. */
export interface ItemPx {
  widthPx: number;
  heightPx: number;
}

/** Where an item lands. The rectangle itself is the VISIBLE placed box
 * (clipped to the content box); `item` is the item's whole scaled extent,
 * which overflows the content box for a cropping fit; `crop` is the part of
 * the source, in source pixels, that survives — the only part an output may
 * carry. */
export interface SheetPlacement extends RectCm {
  item: RectCm;
  crop: { xPx: number; yPx: number; widthPx: number; heightPx: number };
  /** True when part of the item is cut away (cover/actual overflow). */
  cropped: boolean;
}

/** The output a sheet is assembled into. PNG and JPEG arrive with SPEC-109;
 * the background rules for them are stated here, once. */
export type SheetOutput = "pdf" | "png" | "jpeg";

/** The resolution ladder: no free number, so every surface offers the same
 * four steps. */
export const SHEET_DPI_CHOICES: readonly number[] = [72, 150, 300, 600];

/** The default for raster output, and the default `dpi` of a fresh setup. */
export const DEFAULT_SHEET_DPI = 150;

/** The smallest content box, in both axes, a margin set may leave. */
export const MIN_CONTENT_CM = 1;

/** Refuses a sheet no output path can produce (5 m a side). */
export const MAX_SHEET_CM = 500;

export const TRANSPARENT = "transparent";

/** A4, portrait, Normal margins (SPEC-066's preset, 2.54 cm all round),
 * contain, center, white, 150 dpi. */
export const SHEET_DEFAULTS: SheetSetup = {
  paper: "A4",
  orientation: "portrait",
  margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
  fit: "contain",
  align: "center",
  background: "#ffffff",
  dpi: DEFAULT_SHEET_DPI,
};

/** The one shared rounding rule: three decimal places. Float noise is
 * cleared first (toPrecision), so a value a hair under a half-thousandth
 * through binary arithmetic rounds as its decimal reading does; -0 is 0. */
export function r3(value: number): number {
  if (!Number.isFinite(value)) {
    return 0;
  }
  const rounded = Math.round(Number((value * 1000).toPrecision(12))) / 1000;
  return rounded === 0 ? 0 : rounded;
}

function r3Rect(rect: RectCm): RectCm {
  return { xCm: r3(rect.xCm), yCm: r3(rect.yCm), widthCm: r3(rect.widthCm), heightCm: r3(rect.heightCm) };
}

/** An item's physical size when one of its pixels is one `dpi` dot. */
export function itemSizeCm(itemPx: ItemPx, dpi: number): PageSize {
  const safeDpi = dpi > 0 ? dpi : DEFAULT_SHEET_DPI;
  return {
    widthCm: r3((Math.max(0, itemPx.widthPx) / safeDpi) * CM_PER_INCH),
    heightCm: r3((Math.max(0, itemPx.heightPx) / safeDpi) * CM_PER_INCH),
  };
}

/** The sheet's paper size before margins: a named size or Custom, turned to
 * the setup's orientation through SPEC-027's helpers. Not used for "Fit to
 * content", whose size comes from the item. */
export function paperSizeOf(setup: SheetSetup): PageSize {
  if (setup.paper === "Custom") {
    const custom = { widthCm: setup.widthCm ?? 0, heightCm: setup.heightCm ?? 0 };
    return custom.widthCm !== custom.heightCm && orientationOf(custom) !== setup.orientation
      ? swapOrientation(custom)
      : custom;
  }
  return pageSizeFor(setup.paper === FIT_TO_CONTENT ? "A4" : setup.paper, setup.orientation);
}

/** The sheet and its content box. For "Fit to content" the sheet is the
 * item at `dpi` plus the margins, so `itemPx` is required there — calling
 * without it throws rather than inventing a size. */
export function resolveSheet(setup: SheetSetup, itemPx?: ItemPx): SheetGeometry {
  const { topCm, rightCm, bottomCm, leftCm } = setup.margins;
  if (setup.paper === FIT_TO_CONTENT) {
    if (!itemPx) {
      throw new Error('"Fit to content" needs the item\'s pixel size');
    }
    const item = itemSizeCm(itemPx, setup.dpi);
    return {
      widthCm: r3(item.widthCm + leftCm + rightCm),
      heightCm: r3(item.heightCm + topCm + bottomCm),
      content: r3Rect({ xCm: leftCm, yCm: topCm, widthCm: item.widthCm, heightCm: item.heightCm }),
    };
  }
  const paper = paperSizeOf(setup);
  return {
    widthCm: r3(paper.widthCm),
    heightCm: r3(paper.heightCm),
    content: r3Rect({
      xCm: leftCm,
      yCm: topCm,
      widthCm: paper.widthCm - leftCm - rightCm,
      heightCm: paper.heightCm - topCm - bottomCm,
    }),
  };
}

/** How far along the free space an alignment sits, per axis: 0 takes the
 * left/top edge, 0.5 splits the difference, 1 takes the right/bottom. */
function alignFactors(align: SheetAlign): { fx: number; fy: number } {
  const fx = align.endsWith("left") ? 0 : align.endsWith("right") ? 1 : 0.5;
  const fy = align.startsWith("top") ? 0 : align.startsWith("bottom") ? 1 : 0.5;
  return { fx, fy };
}

/** Places one item in a content box per the fit rules:
 *  - contain: scale min(W/w, H/h) — whole item visible, aspect kept;
 *  - cover: scale max(W/w, H/h), clipped to the box — aspect kept;
 *  - stretch: exactly W × H — aspect broken;
 *  - actual: one pixel per `dpi` dot, clipped to the box when larger.
 * `align` positions the result in the box and, for the cropping fits,
 * chooses which part of the item survives. `dpi` matters only to `actual`:
 * the other three scale by a ratio in which it cancels, so they are
 * computed from pixels directly and are independent of it exactly. */
export function placeOnSheet(
  itemPx: ItemPx,
  box: RectCm,
  fit: SheetFit,
  align: SheetAlign,
  dpi: number = DEFAULT_SHEET_DPI,
): SheetPlacement {
  const w = Math.max(0, itemPx.widthPx);
  const h = Math.max(0, itemPx.heightPx);
  const boxW = Math.max(0, box.widthCm);
  const boxH = Math.max(0, box.heightCm);
  let itemW = 0;
  let itemH = 0;
  if (w > 0 && h > 0) {
    if (fit === "stretch") {
      itemW = boxW;
      itemH = boxH;
    } else if (fit === "actual") {
      const size = itemSizeCm(itemPx, dpi);
      itemW = size.widthCm;
      itemH = size.heightCm;
    } else {
      const scale = fit === "cover" ? Math.max(boxW / w, boxH / h) : Math.min(boxW / w, boxH / h);
      itemW = w * scale;
      itemH = h * scale;
    }
  }
  const { fx, fy } = alignFactors(align);
  const itemX = box.xCm + fx * (boxW - itemW);
  const itemY = box.yCm + fy * (boxH - itemH);
  const left = Math.max(itemX, box.xCm);
  const top = Math.max(itemY, box.yCm);
  const right = Math.min(itemX + itemW, box.xCm + boxW);
  const bottom = Math.min(itemY + itemH, box.yCm + boxH);
  const visibleW = Math.max(0, right - left);
  const visibleH = Math.max(0, bottom - top);
  const pxPerCmX = itemW > 0 ? w / itemW : 0;
  const pxPerCmY = itemH > 0 ? h / itemH : 0;
  const crop = {
    xPx: r3((left - itemX) * pxPerCmX),
    yPx: r3((top - itemY) * pxPerCmY),
    widthPx: r3(visibleW * pxPerCmX),
    heightPx: r3(visibleH * pxPerCmY),
  };
  return {
    ...r3Rect({ xCm: left, yCm: top, widthCm: visibleW, heightCm: visibleH }),
    item: r3Rect({ xCm: itemX, yCm: itemY, widthCm: itemW, heightCm: itemH }),
    crop,
    cropped: w > 0 && h > 0 && (crop.widthPx < r3(w) || crop.heightPx < r3(h)),
  };
}

/** The sheet's size in whole pixels at `dpi` — what a raster output would
 * be, and the estimate the panel shows before anything is written. */
export function sheetPixelSize(geometry: SheetGeometry, dpi: number): { widthPx: number; heightPx: number } {
  return {
    widthPx: Math.round((geometry.widthCm / CM_PER_INCH) * dpi),
    heightPx: Math.round((geometry.heightCm / CM_PER_INCH) * dpi),
  };
}

// --- validation -------------------------------------------------------------
// Reasons, not clamps: the panel rejects an input that fails one of these and
// says why, leaving the setup as it was.

/** Whether `margins` leave a content box of at least MIN_CONTENT_CM in both
 * axes on `paper` (and none of them is negative). */
export function marginsLeaveContent(paper: PageSize, margins: Margins): boolean {
  const { topCm, rightCm, bottomCm, leftCm } = margins;
  if ([topCm, rightCm, bottomCm, leftCm].some((value) => !Number.isFinite(value) || value < 0)) {
    return false;
  }
  return (
    r3(paper.widthCm - leftCm - rightCm) >= MIN_CONTENT_CM &&
    r3(paper.heightCm - topCm - bottomCm) >= MIN_CONTENT_CM
  );
}

/** Whether a sheet dimension is one an output can produce. */
export function sheetSizeAllowed(widthCm: number, heightCm: number): boolean {
  return (
    Number.isFinite(widthCm) &&
    Number.isFinite(heightCm) &&
    widthCm > 0 &&
    heightCm > 0 &&
    widthCm <= MAX_SHEET_CM &&
    heightCm <= MAX_SHEET_CM
  );
}

/** Whether `value` is a background the setup accepts: "#rrggbb" or
 * "transparent". */
export function isSheetBackground(value: string): boolean {
  return value === TRANSPARENT || /^#[0-9a-f]{6}$/i.test(value);
}

/** The background a fresh setup uses for `output`: white for PDF (and
 * JPEG), transparent for PNG. */
export function defaultSheetBackground(output: SheetOutput): string {
  return output === "png" ? TRANSPARENT : "#ffffff";
}

/** Why `background` cannot be used for `output`, or null when it can:
 * JPEG and PDF pages carry no alpha, so "transparent" is refused there. */
export function backgroundRefusal(background: string, output: SheetOutput): string | null {
  if (background !== TRANSPARENT) {
    return null;
  }
  if (output === "pdf") {
    return "PDF pages have no transparency, so the background will be white.";
  }
  if (output === "jpeg") {
    return "JPEG images have no transparency; choose a background colour.";
  }
  return null;
}

/** The colour an output actually paints behind the item: a refused
 * "transparent" falls back to white. */
export function effectiveBackground(background: string, output: SheetOutput): string {
  return backgroundRefusal(background, output) !== null ? "#ffffff" : background;
}
