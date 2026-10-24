// Images to a PDF (SPEC-110): the assembly workspace's multi-page output. A
// tray of images becomes pages — one per page, n-up, or every item on one
// sheet — and the pages become a real PDF written by the SPEC-034 writer this
// repo already owns. THIS FILE is the normative source for when an image is
// passed through and when it is re-encoded, for page planning, for the PDF's
// metadata and outline, and for page numbering.
//
// Page planning is pure: `planPdfPages` cuts the tray into pages and hands
// each page to SPEC-109's `planAssembly` — `single-page` is SPEC-109's plan
// verbatim; `n-up` is its grid of `nUp` cells inside one page; `one-per-page`
// is n-up with one cell, which IS the content box. A short last page keeps
// the full page's cells, so the fourth scan of a 4-up sheet sits where the
// fourth always does.
//
// The quality rule is the point: a baseline JPEG (1 or 3 components) whose
// placement crops nothing, carries no SPEC-112 edit stack, and is opaque with
// a normal blend is embedded VERBATIM as /DCTDecode — the file's own scan
// data. Every other placement is drawn by SPEC-109's renderer
// (media/composeCanvas) at no more than the sheet's dpi and never above the
// image's own resolution, then encoded by `encodeImage` with `maxDpi` set to
// that dpi. A blend can only be drawn against what lies beneath it, so a
// blended placement is flattened with everything under it on its page.
//
// `buildImagePdf` composes one PdfPageInput per page — the background rect,
// each placement's cell background, one PdfImageItem per placement and the
// page number — with a yield between pages, and calls `buildPdf` ONCE. It
// adds no PDF serialisation of its own: a construct the writer cannot emit
// is fixed in the writer. The same image placed on many pages is one key in
// PdfBuildInput.images, which the writer embeds once.
//
// Browser-bound only where it draws (the re-encode step, injectable); no Vue,
// no store. Sources are only ever read.
import {
  naturalAssemblyPx,
  planAssembly,
  type AssemblyItem,
  type AssemblyPlan,
  type PlacedItem,
} from "../core/assemblyLayout";
import {
  buildPdf,
  PDF_PRODUCER,
  type PdfBuildInput,
  type PdfHeading,
  type PdfImageData,
  type PdfImageItem,
  type PdfPageInput,
  type PdfRectItem,
  type PdfTextItem,
} from "../core/pdfBuild";
import { encodeImage, parseJpegHeader, type JpegHeader } from "../core/pdfImages";
import { faceCssFont, type FaceKey } from "../core/pdfFonts";
import type { SfntFont } from "../core/sfnt";
import type { ImageEdit, PixelSize } from "../core/imageEdits";
import type { MediaPart } from "../core/media";
import { parsePdf, type PdfDocument } from "../core/pdfDocument";
import { yieldToMainThread } from "../core/idleScheduler";
import { CM_PER_INCH, cmToPt } from "../core/units";
import {
  DEFAULT_SHEET_DPI,
  effectiveBackground,
  r3,
  TRANSPARENT,
  type ItemPx,
  type RectCm,
  type SheetGeometry,
} from "../core/sheetLayout";
import {
  composeToCanvas,
  decodeSource,
  encodeCanvas,
  prepareSources,
  SVG_CONTENT_TYPE,
  svgRasterSize,
  type ComposeInput,
} from "./composeCanvas";
import { mediaPartBytes } from "./download";
import { releaseCanvas, type DecodedImage } from "./renderEdits";

export type PageMode = "one-per-page" | "n-up" | "single-page";

export const PAGE_MODES: readonly PageMode[] = ["one-per-page", "n-up", "single-page"];

export const DEFAULT_PAGE_MODE: PageMode = "one-per-page";

/** The n-up choices: items per page. */
export const N_UP_CHOICES: readonly number[] = [2, 4, 6, 9, 16];

export const DEFAULT_N_UP = 4;

/** The most pages one PDF assembly produces. */
export const MAX_PDF_ASSEMBLY_PAGES = 500;

/** The largest estimated output one PDF assembly produces. */
export const MAX_PDF_ASSEMBLY_BYTES = 500 * 1024 * 1024;

/** A page number: 9 pt, centred, its baseline 1 cm above the sheet's bottom
 * edge, in #555555, in the bundled sans face SPEC-034 already embeds. */
export const PAGE_NUMBER_SIZE_PT = 9;
export const PAGE_NUMBER_OFFSET_CM = 1;
export const PAGE_NUMBER_COLOR = "#555555";
export const PAGE_NUMBER_FACE: FaceKey = "sans-regular";

const PAGE_NUMBER_RGB: [number, number, number] = [0x55 / 255, 0x55 / 255, 0x55 / 255];

/** How far a page number's digits rise above their baseline, in ems — the
 * bundled face's figures stand 0.716 em; this rounds up. */
const PAGE_NUMBER_RISE_EM = 0.75;

/** The bottom margin page numbers need: the number sits inside the margin,
 * never over a placed image, so the margin must hold its baseline offset
 * plus the digits' height. */
export const PAGE_NUMBER_MIN_MARGIN_CM = r3(
  PAGE_NUMBER_OFFSET_CM + ((PAGE_NUMBER_SIZE_PT * PAGE_NUMBER_RISE_EM) / 72) * CM_PER_INCH,
);

/** The plan a PDF assembly is cut into pages with. */
export type PdfAssemblyPlan = AssemblyPlan & { pageMode: PageMode; nUp: number };

export interface PdfAssemblyPage {
  /** Indices into the tray, in order, that this page carries. */
  itemIndices: number[];
  /** From planAssembly, against this page's content box. Each placement's
   * `index` is its TRAY index, so one source list serves every page. */
  placed: PlacedItem[];
}

/** One tray item as the PDF reads it. */
export interface ImagePdfItem {
  /** The library id: the key into `sources`. */
  id: string;
  /** Its outline entry. */
  title: string;
  /** The pixel size its placements' `sourcePx` is measured in. */
  widthPx: number;
  heightPx: number;
  /** SPEC-112's saved edit stack. Absent or empty for an unedited image —
   * and absence never blocks passthrough. */
  edits?: readonly ImageEdit[];
}

export interface ImagePdfSheet {
  geometry: SheetGeometry;
  /** The sheet's background ("transparent" paints white on a PDF page). */
  background: string;
  /** The sheet's dpi: the most any re-encoded image is sampled at. */
  dpi: number;
}

/** What the re-encode step draws: one placement (or, for a blend, it and
 * everything beneath it on its page) over its visible rectangle. */
export interface RasterJob {
  /** The rectangle the raster covers, in sheet centimetres. */
  region: RectCm;
  widthPx: number;
  heightPx: number;
  /** Back to front, in sheet coordinates. */
  layers: PlacedItem[];
  /** Each tray index's part and pixel size. */
  inputs: readonly (ComposeInput | null)[];
  /** "transparent" for a lone placement; the page's colour when flattened. */
  background: string;
  /** The printed size and the sheet's dpi, for encodeImage. */
  widthPt: number;
  heightPt: number;
  dpi: number;
}

/** The browser-bound and writer steps, injectable for tests. */
export interface ImagePdfSteps {
  /** Draws and encodes a re-encoded placement; null when it cannot. */
  rasterise: (job: RasterJob) => Promise<PdfImageData | null>;
  /** Awaited between pages. */
  between: () => Promise<void>;
  /** The SPEC-034 writer. */
  buildPdf: (input: PdfBuildInput) => Promise<Uint8Array>;
}

export interface ImagePdfOptions {
  title: string;
  author?: string;
  subject?: string;
  pageNumbers: boolean;
  /** /CreationDate and /ModDate — fixed by tests; defaults to now. */
  modified?: Date;
  /** The tray, by index. */
  items: readonly ImagePdfItem[];
  sheet: ImagePdfSheet;
  /** The parsed faces; page numbers need PAGE_NUMBER_FACE. */
  fonts?: Map<FaceKey, SfntFont>;
  steps?: Partial<ImagePdfSteps>;
  /** Checked at every page boundary; true stops the build there. */
  cancelled?: () => boolean;
  /** Called with the pages composed so far. */
  onProgress?: (done: number, total: number) => void;
}

/** Thrown when `cancelled` stopped a build: nothing was written. */
export class ImagePdfCancelled extends Error {
  constructor() {
    super("The PDF was cancelled.");
    this.name = "ImagePdfCancelled";
  }
}

// --- page planning ------------------------------------------------------------

/** An n-up count from the choices (the default for anything else). */
export function normaliseNUp(nUp: number): number {
  return N_UP_CHOICES.includes(nUp) ? nUp : DEFAULT_N_UP;
}

/** How many items one page holds: 1, nUp, or every item. */
export function itemsPerPage(plan: Pick<PdfAssemblyPlan, "pageMode" | "nUp">, count: number): number {
  if (plan.pageMode === "single-page") {
    return Math.max(1, count);
  }
  return plan.pageMode === "n-up" ? normaliseNUp(plan.nUp) : 1;
}

/** How many pages `count` items make (0 for an empty tray). */
export function pdfPageCount(plan: Pick<PdfAssemblyPlan, "pageMode" | "nUp">, count: number): number {
  return count === 0 ? 0 : Math.ceil(count / itemsPerPage(plan, count));
}

/** The column count of an n-up page: SPEC-109's "auto" rule (the least
 * content-box area wasted around nUp square items, ties to the squarer grid,
 * then to fewer columns) over the column counts that divide nUp — so every
 * n-up page is a full grid of exactly nUp cells (16-up is 4 × 4, never 3 × 6
 * with two empty cells). */
export function nUpColumns(nUp: number, box: { widthCm: number; heightCm: number }): number {
  let best = 1;
  let bestWaste = Number.POSITIVE_INFINITY;
  let bestSkew = Number.POSITIVE_INFINITY;
  for (let columns = 1; columns <= nUp; columns += 1) {
    if (nUp % columns !== 0) {
      continue;
    }
    const rows = nUp / columns;
    const side = Math.min(Math.max(0, box.widthCm) / columns, Math.max(0, box.heightCm) / rows);
    const waste = r3(box.widthCm * box.heightCm - nUp * side * side);
    const skew = Math.abs(columns - rows);
    if (waste < bestWaste || (waste === bestWaste && skew < bestSkew)) {
      best = columns;
      bestWaste = waste;
      bestSkew = skew;
    }
  }
  return best;
}

/** Pads a short page: a grid's cells depend only on how many there are. */
const BLANK_ITEM: AssemblyItem = { id: "", widthPx: 1, heightPx: 1 };

/** Where every item lands, page by page, in tray order. Pure. `dpi` matters
 * only to the "actual" fit, as in planAssembly. */
export function planPdfPages(
  items: AssemblyItem[],
  plan: PdfAssemblyPlan,
  sheet: SheetGeometry,
  dpi: number = DEFAULT_SHEET_DPI,
): PdfAssemblyPage[] {
  if (items.length === 0) {
    return [];
  }
  if (plan.pageMode === "single-page") {
    return [{ itemIndices: items.map((_, index) => index), placed: planAssembly(items, plan, sheet, dpi) }];
  }
  const perPage = itemsPerPage(plan, items.length);
  // SPEC-109's grid of perPage cells; one cell is the whole content box.
  const cells: AssemblyPlan = { ...plan, layout: "grid", columns: nUpColumns(perPage, sheet.content) };
  const pages: PdfAssemblyPage[] = [];
  for (let start = 0; start < items.length; start += perPage) {
    const itemIndices = items.slice(start, start + perPage).map((_, offset) => start + offset);
    const padded = [
      ...itemIndices.map((index) => ({ id: items[index].id, widthPx: items[index].widthPx, heightPx: items[index].heightPx })),
      ...Array.from({ length: perPage - itemIndices.length }, () => BLANK_ITEM),
    ];
    const placed = planAssembly(padded, cells, sheet, dpi)
      .slice(0, itemIndices.length)
      .map((item, offset) => ({ ...item, index: itemIndices[offset] }));
    pages.push({ itemIndices, placed });
  }
  return pages;
}

/** "Fit to content" for a PDF: one page's arrangement — SPEC-109's own for
 * single-page, else nUp cells each the size of the largest item. */
export function pdfNaturalPx(items: AssemblyItem[], plan: PdfAssemblyPlan): ItemPx | null {
  if (plan.pageMode === "single-page" || items.length === 0) {
    return naturalAssemblyPx(items, plan);
  }
  const cell = {
    id: "",
    widthPx: Math.max(...items.map((item) => item.widthPx)),
    heightPx: Math.max(...items.map((item) => item.heightPx)),
  };
  const perPage = itemsPerPage(plan, items.length);
  return naturalAssemblyPx(
    Array.from({ length: perPage }, () => cell),
    { ...plan, layout: "grid", columns: nUpColumns(perPage, { widthCm: 1, heightCm: 1 }) },
  );
}

// --- the passthrough decision --------------------------------------------------

/** What the decision reads about a source part — never its pixels. */
export interface SourceFacts {
  mediaId: string;
  contentType: string;
  /** The JPEG frame header; null for anything that is not a readable JPEG. */
  header: JpegHeader | null;
  byteLength: number;
}

export function sourceFacts(part: MediaPart, bytes: Uint8Array = mediaPartBytes(part)): SourceFacts {
  return {
    mediaId: part.mediaId,
    contentType: part.contentType,
    header: part.contentType === "image/jpeg" ? parseJpegHeader(bytes) : null,
    byteLength: bytes.length,
  };
}

/** Why a placement is embedded the way it is: "original" passes the file
 * through; every other reason re-encodes it. */
export type EncodingReason =
  | "original"
  | "edited"
  | "not-jpeg"
  | "not-baseline"
  | "components"
  | "cropped"
  | "opacity"
  | "blend";

/** A crop smaller than this, in source pixels, is rounding, not a crop. */
const CROP_TOLERANCE_PX = 0.01;

/** Whether a placement keeps its whole source rectangle. */
export function coversWholeSource(sourcePx: PlacedItem["sourcePx"], item: ItemPx): boolean {
  return (
    sourcePx.x <= CROP_TOLERANCE_PX &&
    sourcePx.y <= CROP_TOLERANCE_PX &&
    sourcePx.width >= item.widthPx - CROP_TOLERANCE_PX &&
    sourcePx.height >= item.heightPx - CROP_TOLERANCE_PX
  );
}

/** The passthrough decision, exactly: a baseline JPEG of 1 or 3 components,
 * uncropped, with no non-empty edit stack, opaque, normal blend — else the
 * first reason it fails. */
export function encodingReason(
  source: Pick<SourceFacts, "contentType" | "header">,
  item: Pick<ImagePdfItem, "widthPx" | "heightPx" | "edits">,
  placed: Pick<PlacedItem, "sourcePx" | "opacity" | "blend">,
): EncodingReason {
  if (item.edits !== undefined && item.edits.length > 0) {
    return "edited";
  }
  if (source.contentType !== "image/jpeg" || !source.header) {
    return "not-jpeg";
  }
  if (!source.header.baseline) {
    return "not-baseline";
  }
  if (source.header.components !== 1 && source.header.components !== 3) {
    return "components";
  }
  if (!coversWholeSource(placed.sourcePx, item)) {
    return "cropped";
  }
  if (placed.opacity < 1) {
    return "opacity";
  }
  return placed.blend === "normal" ? "original" : "blend";
}

const TYPE_NAMES: Record<string, string> = {
  "image/png": "PNG",
  "image/gif": "GIF",
  "image/webp": "WebP",
  [SVG_CONTENT_TYPE]: "SVG",
};

/** The tray's words for a decision: "original JPEG", "re-encoded: cropped". */
export function encodingLabel(reason: EncodingReason, contentType: string): string {
  switch (reason) {
    case "original":
      return "original JPEG";
    case "not-jpeg":
      return `re-encoded: ${TYPE_NAMES[contentType] ?? "not a JPEG"}`;
    case "not-baseline":
      return "re-encoded: progressive JPEG";
    case "components":
      return "re-encoded: CMYK JPEG";
    case "blend":
      return "re-encoded: blend mode";
    default:
      return `re-encoded: ${reason}`;
  }
}

/** A passthrough image: the file's own bytes, its header's size and colour
 * space. */
export function passthroughImage(header: JpegHeader, bytes: Uint8Array): PdfImageData {
  return {
    widthPx: header.widthPx,
    heightPx: header.heightPx,
    filter: "DCTDecode",
    colorSpace: header.components === 1 ? "DeviceGray" : "DeviceRGB",
    data: bytes,
  };
}

/** The pixel size a re-encoded placement is drawn at: its visible rectangle
 * at the sheet's dpi, never above the source pixels it shows (an SVG has
 * none: it is drawn at the placed size), within the canvas bounds every
 * target browser allocates. */
export function rasterSize(placed: PlacedItem, contentType: string, dpi: number): PixelSize {
  const atDpi = {
    widthPx: (Math.max(0, placed.widthCm) / CM_PER_INCH) * dpi,
    heightPx: (Math.max(0, placed.heightCm) / CM_PER_INCH) * dpi,
  };
  const vector = contentType === SVG_CONTENT_TYPE;
  return svgRasterSize({
    widthPx: Math.max(1, Math.round(vector ? atDpi.widthPx : Math.min(atDpi.widthPx, placed.sourcePx.width))),
    heightPx: Math.max(1, Math.round(vector ? atDpi.heightPx : Math.min(atDpi.heightPx, placed.sourcePx.height))),
  });
}

/** The resolution a placement prints at: the source's own for a passed-
 * through JPEG (which may exceed the sheet's), else its raster's. */
export function placementDpi(placed: PlacedItem, contentType: string, reason: EncodingReason, dpi: number): number {
  const widthIn = placed.widthCm / CM_PER_INCH;
  const heightIn = placed.heightCm / CM_PER_INCH;
  if (widthIn <= 0 || heightIn <= 0) {
    return dpi;
  }
  if (reason === "original") {
    return Math.min(placed.sourcePx.width / widthIn, placed.sourcePx.height / heightIn);
  }
  const size = rasterSize(placed, contentType, dpi);
  return Math.min(size.widthPx / widthIn, size.heightPx / heightIn);
}

/** What a placement's embedded image is keyed by: the source itself when it
 * passes through; the drawing (crop, size, opacity, cell colour) otherwise —
 * so an image placed alike on forty pages is ONE key and one copy. A blend
 * bakes in its page and is its own. */
export function imageKey(placed: PlacedItem, facts: SourceFacts, reason: EncodingReason, dpi: number, pageIndex: number): string {
  if (reason === "original") {
    return `jpeg:${facts.mediaId}`;
  }
  if (reason === "blend") {
    return `flat:${pageIndex}:${placed.index}`;
  }
  const { x, y, width, height } = placed.sourcePx;
  const size = rasterSize(placed, facts.contentType, dpi);
  return (
    `raster:${facts.mediaId}:${x},${y},${width},${height}:${size.widthPx}x${size.heightPx}` +
    `:${placed.opacity}:${placed.cellBackground}`
  );
}

// --- the estimate and the limits ----------------------------------------------

/** The file around the images: header, catalog, xref, info. */
const PDF_BASE_BYTES = 2048;
/** Per page: its object, content stream, outline entries. */
const PDF_PAGE_BYTES = 1024;
/** Page numbers embed the bundled face whole (LiberationSans-Regular.ttf,
 * deflated in the file — this is its uncompressed size, the bound). */
export const PAGE_NUMBER_FONT_BYTES = 410_820;

export interface PdfAssemblyEstimate {
  pageCount: number;
  /** The output's bytes: passed-through files at their own length,
   * re-encoded rasters at their samples' size before compression. */
  bytes: number;
  /** Each tray item's share: an image embedded once counts once, against the
   * first placement that embeds it. */
  itemBytes: number[];
  /** The lowest-resolution placement, or null when nothing is placed. */
  lowest: { index: number; page: number; dpi: number } | null;
}

/** The estimate, from each placement's passthrough decision — a computed
 * number, not a guess. `facts` is by tray index (null: unreadable). */
export function estimateImagePdf(
  pages: readonly PdfAssemblyPage[],
  items: readonly ImagePdfItem[],
  facts: readonly (SourceFacts | null)[],
  dpi: number,
  pageNumbers: boolean,
): PdfAssemblyEstimate {
  const itemBytes = items.map(() => 0);
  const keys = new Set<string>();
  let bytes = PDF_BASE_BYTES + pages.length * PDF_PAGE_BYTES + (pageNumbers ? PAGE_NUMBER_FONT_BYTES : 0);
  let lowest: PdfAssemblyEstimate["lowest"] = null;
  for (const [pageIndex, page] of pages.entries()) {
    for (const placed of page.placed) {
      const source = facts[placed.index];
      const item = items[placed.index];
      if (!source || !item || placed.widthCm <= 0 || placed.heightCm <= 0) {
        continue;
      }
      const reason = encodingReason(source, item, placed);
      const key = imageKey(placed, source, reason, dpi, pageIndex);
      if (!keys.has(key)) {
        keys.add(key);
        const size = rasterSize(placed, source.contentType, dpi);
        const alpha = reason !== "blend" && (source.contentType !== "image/jpeg" || placed.opacity < 1);
        const share = reason === "original" ? source.byteLength : size.widthPx * size.heightPx * (alpha ? 4 : 3);
        itemBytes[placed.index] += share;
        bytes += share;
      }
      const at = placementDpi(placed, source.contentType, reason, dpi);
      if (!lowest || at < lowest.dpi) {
        lowest = { index: placed.index, page: pageIndex, dpi: at };
      }
    }
  }
  return { pageCount: pages.length, bytes, itemBytes, lowest };
}

function megabytes(bytes: number): string {
  return `${Math.round((bytes / (1024 * 1024)) * 10) / 10} MB`;
}

/** "“a”", "“a” and “b”", "“a”, “b” and “c”", "“a”, “b”, “c” and 2 more". */
function quotedList(titles: readonly string[]): string {
  const named = titles.slice(0, 3).map((title) => `“${title || "Untitled"}”`);
  const more = titles.length - named.length;
  if (more > 0) {
    return `${named.join(", ")} and ${more} more`;
  }
  return named.length > 1 ? `${named.slice(0, -1).join(", ")} and ${named[named.length - 1]}` : named.join("");
}

/** Why the PDF cannot be built — naming the limit it breaks and what to
 * remove — or null when it can. Checked before any work starts. */
export function pdfLimitRefusal(
  estimate: PdfAssemblyEstimate,
  titles: readonly string[],
  perPage: number,
): string | null {
  if (estimate.pageCount > MAX_PDF_ASSEMBLY_PAGES) {
    const excess = titles.length - MAX_PDF_ASSEMBLY_PAGES * Math.max(1, perPage);
    return (
      `This PDF would have ${estimate.pageCount} pages; a PDF assembly holds at most ${MAX_PDF_ASSEMBLY_PAGES}. ` +
      `Remove at least ${excess} ${excess === 1 ? "image" : "images"} from the tray, or put more images on each page.`
    );
  }
  if (estimate.bytes > MAX_PDF_ASSEMBLY_BYTES) {
    // The largest contributors, until what is left fits.
    const order = estimate.itemBytes
      .map((bytes, index) => ({ bytes, index }))
      .filter((entry) => entry.bytes > 0)
      .toSorted((a, b) => b.bytes - a.bytes || a.index - b.index);
    const remove: string[] = [];
    let left = estimate.bytes;
    for (const entry of order) {
      if (left <= MAX_PDF_ASSEMBLY_BYTES) {
        break;
      }
      remove.push(titles[entry.index] ?? "");
      left -= entry.bytes;
    }
    return (
      `This PDF would be about ${megabytes(estimate.bytes)}, more than the ${megabytes(MAX_PDF_ASSEMBLY_BYTES)} ` +
      `a PDF assembly can produce. Remove ${quotedList(remove)} from the tray, or choose a lower resolution.`
    );
  }
  return null;
}

// --- page numbers -------------------------------------------------------------

/** The sheet's bottom margin. */
function bottomMarginCm(geometry: SheetGeometry): number {
  return r3(geometry.heightCm - geometry.content.yCm - geometry.content.heightCm);
}

/** Why page numbers cannot be drawn inside this sheet's bottom margin, or
 * null when they can. */
export function pageNumbersRefusal(geometry: SheetGeometry): string | null {
  return bottomMarginCm(geometry) >= PAGE_NUMBER_MIN_MARGIN_CM
    ? null
    : `Page numbers sit ${PAGE_NUMBER_OFFSET_CM} cm above the bottom edge, inside the margin: make the bottom margin ` +
        `at least ${PAGE_NUMBER_MIN_MARGIN_CM} cm, or turn page numbers off.`;
}

/** A string's advance width in points at `sizePt`. */
export function textWidthPt(text: string, font: SfntFont, sizePt: number): number {
  let units = 0;
  for (const char of text) {
    units += font.advanceOf(font.glyphFor(char.codePointAt(0) ?? 32));
  }
  return (units * sizePt) / font.unitsPerEm;
}

/** Page `number` (from 1): centred, its baseline PAGE_NUMBER_OFFSET_CM above
 * the bottom edge. */
export function pageNumberItem(number: number, geometry: SheetGeometry, font: SfntFont): PdfTextItem {
  const text = String(number);
  return {
    text,
    xPt: (cmToPt(geometry.widthCm) - textWidthPt(text, font, PAGE_NUMBER_SIZE_PT)) / 2,
    yPt: cmToPt(geometry.heightCm - PAGE_NUMBER_OFFSET_CM),
    sizePt: PAGE_NUMBER_SIZE_PT,
    face: PAGE_NUMBER_FACE,
    color: PAGE_NUMBER_RGB,
  };
}

/** Draws page `number` on a preview canvas at the sheet's scale: the size,
 * place, colour and bundled face the PDF writes, so the preview shows it
 * where the file will have it. */
export function drawPageNumber(canvas: HTMLCanvasElement, geometry: SheetGeometry, number: number): void {
  const ctx = canvas.getContext("2d");
  if (!ctx || geometry.widthCm <= 0 || geometry.heightCm <= 0) {
    return;
  }
  const sx = canvas.width / geometry.widthCm;
  const sy = canvas.height / geometry.heightCm;
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";
  ctx.fillStyle = PAGE_NUMBER_COLOR;
  ctx.font = faceCssFont(PAGE_NUMBER_FACE, (PAGE_NUMBER_SIZE_PT / 72) * CM_PER_INCH * sy);
  ctx.textAlign = "center";
  ctx.textBaseline = "alphabetic";
  ctx.fillText(String(number), (geometry.widthCm / 2) * sx, (geometry.heightCm - PAGE_NUMBER_OFFSET_CM) * sy);
  ctx.restore();
}

// --- building -----------------------------------------------------------------

export interface RasterSteps {
  decode: (part: MediaPart, placedPx?: PixelSize) => Promise<DecodedImage>;
  createCanvas: (widthPx: number, heightPx: number) => HTMLCanvasElement;
  release: (canvas: HTMLCanvasElement) => void;
  encodeImage: typeof encodeImage;
}

export const BROWSER_RASTER_STEPS: RasterSteps = {
  decode: decodeSource,
  createCanvas: (widthPx, heightPx) => {
    const canvas = document.createElement("canvas");
    canvas.width = widthPx;
    canvas.height = heightPx;
    return canvas;
  },
  release: releaseCanvas,
  encodeImage,
};

/** The canvas path: SPEC-109's renderer draws the job's layers over its
 * region, the canvas is encoded losslessly and released, and encodeImage
 * samples the result at no more than the sheet's dpi. */
export async function rasterisePlacement(job: RasterJob, steps: RasterSteps = BROWSER_RASTER_STEPS): Promise<PdfImageData | null> {
  const { region } = job;
  const shift = (rect: RectCm) => ({ ...rect, xCm: rect.xCm - region.xCm, yCm: rect.yCm - region.yCm });
  const layers = job.layers.map((item) => ({ ...item, ...shift(item), cell: shift(item.cell) }));
  const box = { xCm: 0, yCm: 0, widthCm: region.widthCm, heightCm: region.heightCm };
  const sx = job.widthPx / region.widthCm;
  const sy = job.heightPx / region.heightCm;
  const prepared = await prepareSources(layers, job.inputs, sx, sy, steps.decode);
  if (prepared.failed.length > 0) {
    return null;
  }
  const canvas = steps.createCanvas(job.widthPx, job.heightPx);
  let png: Uint8Array;
  try {
    const sheet = { geometry: { widthCm: box.widthCm, heightCm: box.heightCm, content: box }, background: job.background };
    if (!(await composeToCanvas(canvas, layers, sheet, prepared.sources))) {
      return null;
    }
    png = await encodeCanvas(canvas, "image/png");
  } finally {
    steps.release(canvas);
  }
  const url = URL.createObjectURL(new Blob([png as unknown as BlobPart], { type: "image/png" }));
  try {
    return await steps.encodeImage(url, job.widthPt, job.heightPt, job.dpi);
  } finally {
    URL.revokeObjectURL(url);
  }
}

export const BROWSER_IMAGE_PDF_STEPS: ImagePdfSteps = {
  rasterise: (job) => rasterisePlacement(job),
  between: yieldToMainThread,
  buildPdf,
};

function rectPt(rect: RectCm): Pick<PdfRectItem, "xPt" | "yPt" | "widthPt" | "heightPt"> {
  return { xPt: cmToPt(rect.xCm), yPt: cmToPt(rect.yCm), widthPt: cmToPt(rect.widthCm), heightPt: cmToPt(rect.heightCm) };
}

/** Builds the PDF: pages composed one at a time (a yield and a cancel check
 * between them), then ONE call to the writer. Throws a user-facing Error
 * naming an image that cannot be read or drawn, and ImagePdfCancelled when
 * `cancelled` stopped it. Never touches `sources`. */
export async function buildImagePdf(
  pages: readonly PdfAssemblyPage[],
  sources: Map<string, MediaPart>,
  options: ImagePdfOptions,
): Promise<Uint8Array> {
  const steps = { ...BROWSER_IMAGE_PDF_STEPS, ...options.steps };
  const { geometry, background, dpi } = options.sheet;
  const widthPt = cmToPt(geometry.widthCm);
  const heightPt = cmToPt(geometry.heightCm);
  const font = options.fonts?.get(PAGE_NUMBER_FACE);
  if (options.pageNumbers && !font) {
    throw new Error("page numbers need the bundled sans face, which could not be loaded.");
  }
  const named = (index: number) => `“${options.items[index]?.title || "Untitled"}”`;
  const inputs: (ComposeInput | null)[] = options.items.map((item) => {
    const part = sources.get(item.id);
    return part ? { part, widthPx: item.widthPx, heightPx: item.heightPx } : null;
  });
  // Each part's bytes and facts, read once.
  const read = new Map<string, { bytes: Uint8Array; facts: SourceFacts }>();
  const readPart = (part: MediaPart) => {
    let entry = read.get(part.mediaId);
    if (!entry) {
      const bytes = mediaPartBytes(part);
      entry = { bytes, facts: sourceFacts(part, bytes) };
      read.set(part.mediaId, entry);
    }
    return entry;
  };

  const images = new Map<string, PdfImageData>();
  const out: PdfPageInput[] = [];
  options.onProgress?.(0, pages.length);
  for (const [pageIndex, page] of pages.entries()) {
    if (pageIndex > 0) {
      await steps.between();
    }
    if (options.cancelled?.()) {
      throw new ImagePdfCancelled();
    }
    const rects: PdfRectItem[] = [{ xPt: 0, yPt: 0, widthPt, heightPt, fill: effectiveBackground(background, "pdf") }];
    const placements: PdfImageItem[] = [];
    const headings: PdfHeading[] = [];
    for (const [order, placed] of page.placed.entries()) {
      const item = options.items[placed.index];
      const input = inputs[placed.index];
      if (!item || !input) {
        throw new Error(`${named(placed.index)} could not be read.`);
      }
      // One level-1 outline entry per item, at the top of its placement.
      headings.push({ level: 1, title: item.title || `Image ${placed.index + 1}`, yPt: cmToPt(placed.yCm) });
      if (placed.widthCm <= 0 || placed.heightCm <= 0) {
        continue;
      }
      const { bytes, facts } = readPart(input.part);
      const reason = encodingReason(facts, item, placed);
      const key = imageKey(placed, facts, reason, dpi, pageIndex);
      const box = rectPt(placed);
      if (!images.has(key)) {
        const size = rasterSize(placed, facts.contentType, dpi);
        const data =
          reason === "original" && facts.header
            ? passthroughImage(facts.header, bytes)
            : await steps.rasterise({
                region: { xCm: placed.xCm, yCm: placed.yCm, widthCm: placed.widthCm, heightCm: placed.heightCm },
                ...size,
                layers: reason === "blend" ? page.placed.slice(0, order + 1) : [placed],
                inputs,
                background: reason === "blend" ? effectiveBackground(background, "pdf") : TRANSPARENT,
                widthPt: box.widthPt,
                heightPt: box.heightPt,
                dpi,
              });
        if (!data) {
          throw new Error(`${named(placed.index)} could not be drawn.`);
        }
        images.set(key, data);
      }
      // A placement drawn with its cell's colour baked in needs no rect.
      if (placed.cellBackground !== TRANSPARENT && placed.opacity === 1 && placed.blend === "normal") {
        rects.push({ ...rectPt(placed.cell), fill: placed.cellBackground });
      }
      placements.push({ src: key, ...box });
    }
    out.push({
      widthPt,
      heightPt,
      items: options.pageNumbers && font ? [pageNumberItem(pageIndex + 1, geometry, font)] : [],
      links: [],
      headings,
      images: placements,
      rects,
    });
    options.onProgress?.(pageIndex + 1, pages.length);
  }
  if (options.cancelled?.()) {
    throw new ImagePdfCancelled();
  }
  const modified = options.modified ?? new Date();
  const author = options.author?.trim();
  const subject = options.subject?.trim();
  return steps.buildPdf({
    title: options.title.trim(),
    info: {
      ...(author ? { author } : {}),
      ...(subject ? { subject } : {}),
      creator: PDF_PRODUCER,
      creationDate: modified,
      modDate: modified,
    },
    pages: out,
    fonts: options.pageNumbers && font ? new Map([[PAGE_NUMBER_FACE, font]]) : new Map(),
    images,
  });
}

/** The self-check: the bytes just written, parsed back — the parse a stored
 * document's PdfSummary is built from. A build whose own output will not
 * re-parse, or re-parses to another page count, throws, so nothing is kept. */
export async function checkBuiltPdf(
  bytes: Uint8Array,
  expectedPages: number,
  parse: (bytes: Uint8Array) => Promise<PdfDocument> = parsePdf,
): Promise<PdfDocument> {
  let doc: PdfDocument;
  try {
    doc = await parse(bytes);
  } catch (cause) {
    const reason = cause instanceof Error && cause.message ? cause.message : String(cause);
    throw new Error(`the PDF it built does not read back (${reason}), so nothing was kept.`, { cause });
  }
  if (doc.pages.length !== expectedPages) {
    throw new Error(`the PDF it built has ${doc.pages.length} pages instead of ${expectedPages}, so nothing was kept.`);
  }
  return doc;
}
