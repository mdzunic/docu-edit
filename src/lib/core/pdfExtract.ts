// Reads one realized on-screen page (.page element) into the PDF builder's
// input (SPEC-034): positioned text line fragments, hyperlink rectangles
// and Heading 1–3 outline entries. Extraction happens off the very DOM the
// pagination engine laid out, which is what guarantees the exported pages
// carry exactly the content the on-screen (and printed) pages show — no
// second layout pass exists to disagree with the first.
//
// Framework-free (no Vue imports) but browser-bound: it needs real layout
// geometry (Range.getClientRects), so its coverage lives in the e2e suite;
// the pure helpers are unit-tested.
import type {
  PdfHeading,
  PdfImageItem,
  PdfLinkItem,
  PdfPageInput,
  PdfPathItem,
  PdfPathSubpath,
  PdfRectItem,
  PdfTextItem,
} from "./pdfBuild";
import { faceKeyFor } from "./pdfFonts";
import { parseSvgPath } from "./svgPath";

/** CSS px (96dpi) to PDF points (72dpi). */
export const PT_PER_PX = 72 / 96;

/** Layout centimetres to PDF points (the page geometry's unit). */
export const PT_PER_CM = 72 / 2.54;

/** The fraction of a text fragment's box that hangs below the baseline —
 * the Liberation faces all put ~19% of (ascent+descent) in the descender,
 * so the baseline sits at bottom − 0.19 × height. */
const DESCENT_FRACTION = 0.19;

/** Elements whose text must never reach print output or the PDF (AC-13):
 * anything carrying the shared screen-only marker class. */
const EXCLUDED_SELECTOR = ".print-exclude";

/** Heading elements outside the body flow (header/footer bands, note
 * areas) do not become outline entries. */
const NON_BODY_SELECTOR = ".page-header-band, .page-footer-band, .page-note-area";

/** The word processor's placed images — `ExtractOptions.imageSelector`'s
 * default (SPEC-034). */
export const DEFAULT_IMAGE_SELECTOR = "img.doc-image";

/** The word processor's shape and chart canvases, read back as images —
 * `ExtractOptions.canvasSelector`'s default (SPEC-051/052). */
export const DEFAULT_CANVAS_SELECTOR = "canvas.doc-shape-canvas, canvas.doc-chart-canvas";

export interface ExtractOptions {
  /** Resolves an internal link anchor (href="#name") to its 0-based target
   * page index, or null when the anchor no longer exists. */
  destPageForAnchor?: (anchor: string) => number | null;
  /** The `<img>` elements placed as images (default DEFAULT_IMAGE_SELECTOR). */
  imageSelector?: string;
  /** The `<canvas>` elements read back as images (default
   * DEFAULT_CANVAS_SELECTOR). A comma list is walked one selector at a time,
   * in list order — the shapes-then-charts order SPEC-034's pages paint in. */
  canvasSelector?: string;
  /** The elements whose computed background and uniform border become
   * `PdfRectItem`s (SPEC-104 — slide backgrounds, frame fills and strokes,
   * handout frames, note lines). Absent (the default) extracts no rects and
   * emits no `rects` key at all. */
  rectSelector?: string;
  /** SPEC-136: the elements whose computed background and per-edge borders
   * become cell fill and border paths (default DEFAULT_CELL_SELECTOR, the
   * text product's table cells). */
  cellSelector?: string;
  /** SPEC-136: text inside an element matching this selector is clipped to
   * that element's box — a character whose box extends past it is dropped,
   * never drawn in part. Absent (the default): no clipping. */
  clipSelector?: string;
}

/** The cell-chrome loop's default elements (SPEC-076). */
export const DEFAULT_CELL_SELECTOR = "td.doc-table-cell";

/** Parses a computed CSS color into 0..1 rgb. Only the rgb()/rgba() forms
 * computed styles produce are handled; anything else renders black. */
export function parseCssColor(color: string): [number, number, number] {
  const match = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(color);
  if (!match) {
    return [0, 0, 0];
  }
  return [
    Math.min(255, Number(match[1])) / 255,
    Math.min(255, Number(match[2])) / 255,
    Math.min(255, Number(match[3])) / 255,
  ];
}

/** A computed `fill`/`stroke` as 0..1 rgb, or null when nothing is painted —
 * `none`, or any fully transparent colour. Distinct from parseCssColor,
 * whose "unparseable is black" fallback is right for text and wrong here:
 * an unpainted stroke must not become a black line. */
export function parseCssPaint(paint: string): [number, number, number] | null {
  const value = paint.trim();
  if (value.length === 0 || value === "none" || value === "transparent") {
    return null;
  }
  const alpha = /^rgba\(\s*\d+[,\s]+\d+[,\s]+\d+[,\s/]+([\d.]+)\s*\)$/.exec(value);
  if (alpha && Number(alpha[1]) === 0) {
    return null;
  }
  if (!/^rgba?\(/.test(value)) {
    return null;
  }
  return parseCssColor(value);
}

/** A computed colour as 0..1 rgb only when it is fully opaque, else null.
 * Distinct from parseCssPaint, which admits any non-zero alpha: a table
 * cell's model shading is always an opaque hex, and the one translucent
 * background the canvas paints — the cell-range selection highlight — is
 * screen chrome that must not reach the PDF (SPEC-076 AC-29). */
export function parseOpaqueCssColor(value: string): [number, number, number] | null {
  const trimmed = value.trim();
  if (!/^rgba?\(/.test(trimmed)) {
    return null;
  }
  const alpha = /^rgba\(\s*\d+[,\s]+\d+[,\s]+\d+[,\s/]+([\d.]+)\s*\)$/.exec(trimmed);
  if (alpha && Number(alpha[1]) < 1) {
    return null;
  }
  return parseCssColor(trimmed);
}

/** One side of an element's computed border, as `rectItemsFor` reads it. */
export interface BorderSide {
  /** The painted width, already in points. */
  widthPt: number;
  /** The computed `border-*-style`. */
  style: string;
  /** The computed `border-*-color`. */
  color: string;
}

function paintedSide(side: BorderSide): [number, number, number] | null {
  if (side.style === "none" || side.style === "hidden" || !(side.widthPt > 0)) {
    return null;
  }
  return parseCssPaint(side.color);
}

/** The rect items one matched element contributes (SPEC-104): its computed
 * background (skipped when fully transparent) and its border. A uniform
 * border — four painted sides of one width, style and colour — is a single
 * stroked rect inset by half the width, so PDF's centred stroke lands exactly
 * where CSS painted the border inside the box; a non-uniform one becomes a
 * filled strip per painted side. `box` is the element's border box in points. */
export function rectItemsFor(
  box: PdfRect,
  backgroundColor: string,
  sides: { top: BorderSide; right: BorderSide; bottom: BorderSide; left: BorderSide },
): PdfRectItem[] {
  const fill = parseCssPaint(backgroundColor);
  const painted = [sides.top, sides.right, sides.bottom, sides.left].map(paintedSide);
  const [top, right, bottom, left] = painted;
  const uniform =
    top !== null &&
    painted.every((rgb) => rgb !== null && rgb.every((channel, i) => channel === top[i])) &&
    [sides.right, sides.bottom, sides.left].every(
      (side) => side.widthPt === sides.top.widthPt && side.style === sides.top.style,
    );
  if (uniform) {
    const width = sides.top.widthPt;
    const item: PdfRectItem = {
      xPt: box.xPt + width / 2,
      yPt: box.yPt + width / 2,
      widthPt: Math.max(0, box.widthPt - width),
      heightPt: Math.max(0, box.heightPt - width),
      stroke: top,
      strokeWidthPt: width,
    };
    if (fill) {
      item.fill = fill;
    }
    return [item];
  }
  const items: PdfRectItem[] = [];
  if (fill) {
    items.push({ ...box, fill });
  }
  if (top) {
    items.push({ ...box, heightPt: sides.top.widthPt, fill: top });
  }
  if (bottom) {
    items.push({ ...box, yPt: box.yPt + box.heightPt - sides.bottom.widthPt, heightPt: sides.bottom.widthPt, fill: bottom });
  }
  if (left) {
    items.push({ ...box, widthPt: sides.left.widthPt, fill: left });
  }
  if (right) {
    items.push({ ...box, xPt: box.xPt + box.widthPt - sides.right.widthPt, widthPt: sides.right.widthPt, fill: right });
  }
  return items;
}

/** A rectangle in PDF points, measured from the page's top-left like every
 * other extracted item. */
export interface PdfRect {
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
}

/** One paintable border edge of a table cell. `style` is the computed CSS
 * border-style; dashed, dotted and double get their own geometry below, and
 * anything else paints the plain solid strip. */
export interface CellEdgePaint {
  widthPt: number;
  color: [number, number, number];
  style?: string;
}

/** A rectangle as a closed, filled path item — the shape both a cell's
 * shading and each of its border strips reduce to. */
export function filledRectItem(rect: PdfRect, fill: [number, number, number]): PdfPathItem {
  return {
    subpaths: [
      {
        start: { xPt: rect.xPt, yPt: rect.yPt },
        segments: [
          { toXPt: rect.xPt + rect.widthPt, toYPt: rect.yPt },
          { toXPt: rect.xPt + rect.widthPt, toYPt: rect.yPt + rect.heightPt },
          { toXPt: rect.xPt, toYPt: rect.yPt + rect.heightPt },
        ],
        closed: true,
      },
    ],
    fill,
  };
}

const half = (edge?: CellEdgePaint) => (edge ? edge.widthPt / 2 : 0);

/** A cell's painted border edges as filled strips (SPEC-076 AC-29). Each
 * strip is centred on the cell-rect edge: under border-collapse the painted
 * border straddles the grid line the neighbouring cell rects share, so two
 * adjacent cells overlap-paint the same strip instead of doubling it. Each
 * strip is extended by the perpendicular edges' half-widths so corners
 * close. */
export function cellBorderItems(
  rect: PdfRect,
  edges: {
    top?: CellEdgePaint;
    right?: CellEdgePaint;
    bottom?: CellEdgePaint;
    left?: CellEdgePaint;
  },
): PdfPathItem[] {
  const x0 = rect.xPt - half(edges.left);
  const y0 = rect.yPt - half(edges.top);
  const width = rect.widthPt + half(edges.left) + half(edges.right);
  const height = rect.heightPt + half(edges.top) + half(edges.bottom);
  const items: PdfPathItem[] = [];
  if (edges.top) {
    items.push(...edgePaintItems({ xPt: x0, yPt: y0, widthPt: width, heightPt: edges.top.widthPt }, edges.top));
  }
  if (edges.bottom) {
    items.push(
      ...edgePaintItems(
        { xPt: x0, yPt: rect.yPt + rect.heightPt - half(edges.bottom), widthPt: width, heightPt: edges.bottom.widthPt },
        edges.bottom,
      ),
    );
  }
  if (edges.left) {
    items.push(...edgePaintItems({ xPt: x0, yPt: y0, widthPt: edges.left.widthPt, heightPt: height }, edges.left));
  }
  if (edges.right) {
    items.push(
      ...edgePaintItems(
        { xPt: rect.xPt + rect.widthPt - half(edges.right), yPt: y0, widthPt: edges.right.widthPt, heightPt: height },
        edges.right,
      ),
    );
  }
  return items;
}

/** One border edge's strip as path items honouring its computed style
 * (SPEC-076 AC-29): the screen paints dashed, dotted and double borders via
 * CSS, so a solid strip in the export would not look the same. A dashed or
 * dotted edge becomes a stroked centre line with a dash pattern; a double
 * edge two strips of a third of the width each, CSS's own geometry; anything
 * else is the plain filled strip. The strip rect is already corner-extended
 * by the caller. */
export function edgePaintItems(strip: PdfRect, edge: CellEdgePaint): PdfPathItem[] {
  const horizontal = strip.widthPt >= strip.heightPt;
  const thickness = horizontal ? strip.heightPt : strip.widthPt;
  // Below 3px (2.25pt) a double border has no room for its gap and the
  // browser paints it as one line — so does the fallthrough strip.
  if (edge.style === "double" && thickness >= 2.25) {
    const line = thickness / 3;
    return horizontal
      ? [
          filledRectItem({ ...strip, heightPt: line }, edge.color),
          filledRectItem({ ...strip, yPt: strip.yPt + thickness - line, heightPt: line }, edge.color),
        ]
      : [
          filledRectItem({ ...strip, widthPt: line }, edge.color),
          filledRectItem({ ...strip, xPt: strip.xPt + thickness - line, widthPt: line }, edge.color),
        ];
  }
  if (edge.style === "dashed" || edge.style === "dotted") {
    // Butt-capped dashes: 3×width on / 3×width off for dashed, square
    // width-sized dots for dotted — the proportions CSS borders use.
    const on = edge.style === "dashed" ? thickness * 3 : thickness;
    const start = horizontal
      ? { xPt: strip.xPt, yPt: strip.yPt + thickness / 2 }
      : { xPt: strip.xPt + thickness / 2, yPt: strip.yPt };
    const end = horizontal
      ? { toXPt: strip.xPt + strip.widthPt, toYPt: strip.yPt + thickness / 2 }
      : { toXPt: strip.xPt + thickness / 2, toYPt: strip.yPt + strip.heightPt };
    return [
      {
        subpaths: [{ start, segments: [end] }],
        stroke: { color: edge.color, widthPt: thickness, dashPt: [on, on] },
      },
    ];
  }
  return [filledRectItem(strip, edge.color)];
}

/** SVG stroke cap/join names to PDF's `J`/`j` operand codes. Both default
 * the same way in SVG and PDF (butt, miter), so an unknown value maps to 0. */
const LINE_CAPS: Record<string, 0 | 1 | 2> = { butt: 0, round: 1, square: 2 };
const LINE_JOINS: Record<string, 0 | 1 | 2> = { miter: 0, round: 1, bevel: 2 };

/** A computed `stroke-dasharray` as a px number list; empty for `none`. */
function parseDashArray(value: string): number[] {
  const parts = value
    .split(/[,\s]+/)
    .map((part) => Number.parseFloat(part))
    .filter((part) => Number.isFinite(part) && part >= 0);
  return parts.some((part) => part > 0) ? parts : [];
}

interface LineFragment {
  text: string;
  rect: DOMRect;
  /** The fragment's first UTF-16 offset in its text node. */
  start: number;
}

function visibleRects(range: Range): DOMRect[] {
  return Array.from(range.getClientRects()).filter(
    (rect) => rect.width > 0.1 && rect.height > 0.1,
  );
}

/** Splits a text node into per-rendered-line fragments. Line boundaries are
 * found by binary search on "does [pos, end) still fit one client rect" —
 * O(lines × log length) Range measurements rather than one per character,
 * which is what keeps a 100-page export inside its budget (AC-21). */
function lineFragmentsOf(node: Text): LineFragment[] {
  const textContent = node.nodeValue ?? "";
  const length = textContent.length;
  const range = (node.ownerDocument ?? document).createRange();
  const fragments: LineFragment[] = [];
  let pos = 0;
  while (pos < length) {
    let lo = pos + 1;
    let hi = length;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      range.setStart(node, pos);
      range.setEnd(node, mid);
      if (visibleRects(range).length <= 1) {
        lo = mid;
      } else {
        hi = mid - 1;
      }
    }
    range.setStart(node, pos);
    range.setEnd(node, lo);
    const rects = visibleRects(range);
    if (rects.length > 0) {
      const rect =
        rects.length === 1 ? rects[0] : (range.getBoundingClientRect() as DOMRect);
      fragments.push({ text: textContent.slice(pos, lo), rect, start: pos });
    }
    pos = lo;
  }
  return fragments;
}

/** SPEC-136's clip: the part of a line fragment whose characters lie
 * wholly inside `box` (horizontally, half a pixel of tolerance) — the first
 * such run, measured character by character; null when none does. A
 * fragment inside the box is returned as it is. */
function clippedFragment(node: Text, fragment: LineFragment, box: DOMRect): LineFragment | null {
  const TOLERANCE = 0.5;
  if (fragment.rect.left >= box.left - TOLERANCE && fragment.rect.right <= box.right + TOLERANCE) {
    return fragment;
  }
  const doc = node.ownerDocument;
  const range = doc.createRange();
  let first = -1;
  let last = -1;
  let left = 0;
  let right = 0;
  let index = 0;
  for (const char of Array.from(fragment.text)) {
    const start = fragment.start + index;
    index += char.length;
    range.setStart(node, start);
    range.setEnd(node, start + char.length);
    const rect = range.getBoundingClientRect();
    const inside = rect.left >= box.left - TOLERANCE && rect.right <= box.right + TOLERANCE;
    if (inside && last === -1) {
      if (first === -1) {
        first = start;
        left = rect.left;
      }
      right = rect.right;
    } else if (!inside && first !== -1 && last === -1) {
      last = start;
    }
  }
  range.detach?.();
  if (first === -1) {
    return null;
  }
  const end = last === -1 ? fragment.start + fragment.text.length : last;
  const text = node.data.slice(first, end);
  const rect = new DOMRect(left, fragment.rect.top, Math.max(0, right - left), fragment.rect.height);
  return { text, rect, start: first };
}

/** Extracts everything the PDF needs from one realized page element. The
 * page's on-screen box is mapped onto its PDF media box, so a browser zoom
 * level cannot skew coordinates. */
export function extractPageContent(
  pageEl: HTMLElement,
  widthPt: number,
  heightPt: number,
  options: ExtractOptions = {},
): PdfPageInput {
  const pageRect = pageEl.getBoundingClientRect();
  const scale = pageRect.width > 0 ? widthPt / pageRect.width : PT_PER_PX;
  const items: PdfTextItem[] = [];
  const links: PdfLinkItem[] = [];
  const headings: PdfHeading[] = [];
  const images: PdfImageItem[] = [];
  const paths: PdfPathItem[] = [];
  const rects: PdfRectItem[] = [];
  // The `rects` key exists only when a rectSelector asked for rects, so a
  // SPEC-034 page input is exactly what it always was.
  const withRects = (page: PdfPageInput): PdfPageInput => (options.rectSelector ? { ...page, rects } : page);
  const doc = pageEl.ownerDocument;
  const view = doc.defaultView;
  if (!view) {
    return withRects({ widthPt, heightPt, items, links, headings, images, paths });
  }

  // Backgrounds and borders (SPEC-104), measured like every other box: the
  // on-screen border box mapped into points, the computed border widths
  // scaled by the element's own on-screen/layout ratio (a slide stage is
  // fitted with transform: scale, which offsetWidth does not see).
  if (options.rectSelector) {
    for (const element of Array.from(pageEl.querySelectorAll(options.rectSelector))) {
      if (element.closest(EXCLUDED_SELECTOR)) {
        continue;
      }
      const rect = element.getBoundingClientRect();
      if (rect.width <= 0.1 || rect.height <= 0.1) {
        continue;
      }
      const style = view.getComputedStyle(element);
      const layoutWidth = element instanceof HTMLElement ? element.offsetWidth : 0;
      const unit = (layoutWidth > 0 ? rect.width / layoutWidth : 1) * scale;
      const side = (width: string, borderStyle: string, color: string): BorderSide => ({
        widthPt: (Number.parseFloat(width) || 0) * unit,
        style: borderStyle,
        color,
      });
      rects.push(
        ...rectItemsFor(
          {
            xPt: (rect.left - pageRect.left) * scale,
            yPt: (rect.top - pageRect.top) * scale,
            widthPt: rect.width * scale,
            heightPt: rect.height * scale,
          },
          style.backgroundColor,
          {
            top: side(style.borderTopWidth, style.borderTopStyle, style.borderTopColor),
            right: side(style.borderRightWidth, style.borderRightStyle, style.borderRightColor),
            bottom: side(style.borderBottomWidth, style.borderBottomStyle, style.borderBottomColor),
            left: side(style.borderLeftWidth, style.borderLeftStyle, style.borderLeftColor),
          },
        ),
      );
    }
  }

  const walker = doc.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent || !(node.nodeValue ?? "").length) {
        return NodeFilter.FILTER_REJECT;
      }
      if (parent.closest(EXCLUDED_SELECTOR)) {
        return NodeFilter.FILTER_REJECT;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const parent = (node as Text).parentElement as HTMLElement;
    const style = view.getComputedStyle(parent);
    const fontSizePx = Number.parseFloat(style.fontSize);
    if (!Number.isFinite(fontSizePx) || fontSizePx <= 0) {
      continue;
    }
    const weight = Number.parseInt(style.fontWeight, 10);
    const bold = Number.isFinite(weight) ? weight >= 600 : style.fontWeight === "bold";
    const italic = style.fontStyle.includes("italic") || style.fontStyle.includes("oblique");
    const face = faceKeyFor(style.fontFamily, bold, italic);
    const color = parseCssColor(style.color);
    const clip = options.clipSelector === undefined ? null : parent.closest(options.clipSelector);
    const clipBox = clip === null ? null : clip.getBoundingClientRect();
    for (const whole of lineFragmentsOf(node as Text)) {
      const fragment = clipBox === null ? whole : clippedFragment(node as Text, whole, clipBox);
      if (fragment === null || fragment.text.trim().length === 0) {
        continue;
      }
      const baselinePx =
        fragment.rect.bottom - pageRect.top - fragment.rect.height * DESCENT_FRACTION;
      items.push({
        text: fragment.text,
        xPt: (fragment.rect.left - pageRect.left) * scale,
        yPt: baselinePx * scale,
        sizePt: fontSizePx * scale,
        face,
        color,
      });
    }
  }

  // Images are placed by their rendered border box, so a resized, floated or
  // absolutely positioned (wrap behind/in front) image lands in the PDF
  // exactly where the pagination engine put it on screen.
  for (const element of Array.from(pageEl.querySelectorAll(options.imageSelector ?? DEFAULT_IMAGE_SELECTOR))) {
    if (element.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const image = element as HTMLImageElement;
    const src = image.currentSrc || image.src;
    const rect = image.getBoundingClientRect();
    if (src.length === 0 || rect.width <= 0.1 || rect.height <= 0.1) {
      continue;
    }
    // "In front of text" is the only wrap that paints above the text, and
    // renderFormatting's imageDisplayStyle gives exactly those a positive
    // z-index.
    const zIndex = Number.parseInt(view.getComputedStyle(image).zIndex, 10);
    images.push({
      src,
      xPt: (rect.left - pageRect.left) * scale,
      yPt: (rect.top - pageRect.top) * scale,
      widthPt: rect.width * scale,
      heightPt: rect.height * scale,
      onTop: Number.isFinite(zIndex) && zIndex > 0,
    });
  }

  // Shapes (SPEC-051 AC-62) paint their geometry/fill/outline into a real
  // <canvas> (DocShapeBlock.vue) — reading it back out as a PNG data URI and
  // placing it exactly like an image reuses the entire image pipeline below
  // instead of teaching the PDF builder a second, vector drawing path.
  // Rotation is already baked into the canvas's own pixels (paintShape
  // rotates before painting), so the canvas element's own on-screen
  // (unrotated) bounding box is the right placement rectangle.
  //
  // Charts (SPEC-052 AC-34, AC-35) paint into a real <canvas> (DocChartBlock.vue)
  // exactly like a shape's — reused here the same way, at the canvas's own
  // painted pixel resolution (devicePixelRatio-scaled, not the CSS display
  // size stretched up), never an upscaled bitmap.
  //
  // The selector list is walked one entry at a time, so the default places
  // every shape before every chart, as it always has; a slide sheet
  // (SPEC-104) passes its own frame canvases.
  const canvasSelectors = (options.canvasSelector ?? DEFAULT_CANVAS_SELECTOR)
    .split(",")
    .map((selector) => selector.trim())
    .filter((selector) => selector.length > 0);
  for (const selector of canvasSelectors) {
    for (const element of Array.from(pageEl.querySelectorAll(selector))) {
      if (element.closest(EXCLUDED_SELECTOR) || !(element instanceof HTMLCanvasElement)) {
        continue;
      }
      const canvas = element;
      const rect = canvas.getBoundingClientRect();
      if (rect.width <= 0.1 || rect.height <= 0.1) {
        continue;
      }
      let src = "";
      try {
        src = canvas.toDataURL("image/png");
      } catch {
        continue;
      }
      if (!src || src === "data:,") {
        continue;
      }
      // In front of text when the block hosting the canvas is (its wrap's
      // positive z-index, as for images).
      const hostEl = canvas.closest(".doc-shape, .doc-chart");
      const zIndex = hostEl ? Number.parseInt(view.getComputedStyle(hostEl).zIndex, 10) : NaN;
      images.push({
        src,
        xPt: (rect.left - pageRect.left) * scale,
        yPt: (rect.top - pageRect.top) * scale,
        widthPt: rect.width * scale,
        heightPt: rect.height * scale,
        onTop: Number.isFinite(zIndex) && zIndex > 0,
      });
    }
  }

  // Connectors (SPEC-070) are the one thing the page paints as SVG rather
  // than into a canvas, and they carry a stroke, a dash and arrowheads —
  // rasterizing them would throw all of that away, so they become real PDF
  // vector paths. The layer covers its page box, so its user units are the
  // page's own layout px: `layoutScale` converts those to the on-screen px
  // the rest of this function measures in (they differ under a zoom
  // transform), and `scale` then converts to points.
  const layoutScale = pageEl.offsetWidth > 0 ? pageRect.width / pageEl.offsetWidth : 1;

  // Table cell chrome (SPEC-076 AC-29): the resolved shading and borders the
  // canvas paints as CSS on each cell — tableCellBandedStyle's inline style
  // plus the component's default border — are invisible to a text extractor,
  // so they are re-read from the computed style here and emitted as filled
  // rectangles under the text. All backgrounds first, then all border
  // strips, so no cell's fill can cover a neighbour's border half; pushed
  // before the connector paths so connectors stay painted above tables.
  const cellFills: PdfPathItem[] = [];
  const cellBorders: PdfPathItem[] = [];
  for (const element of Array.from(pageEl.querySelectorAll(options.cellSelector ?? DEFAULT_CELL_SELECTOR))) {
    if (element.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const rect = element.getBoundingClientRect();
    if (rect.width <= 0.1 || rect.height <= 0.1) {
      continue;
    }
    const style = view.getComputedStyle(element);
    const box: PdfRect = {
      xPt: (rect.left - pageRect.left) * scale,
      yPt: (rect.top - pageRect.top) * scale,
      widthPt: rect.width * scale,
      heightPt: rect.height * scale,
    };
    const fill = parseOpaqueCssColor(style.backgroundColor);
    if (fill) {
      cellFills.push(filledRectItem(box, fill));
    }
    const edge = (
      borderStyle: string,
      widthValue: string,
      colorValue: string,
    ): CellEdgePaint | undefined => {
      if (borderStyle === "none" || borderStyle === "hidden") {
        return undefined;
      }
      const widthPx = Number.parseFloat(widthValue);
      const color = parseOpaqueCssColor(colorValue);
      if (!Number.isFinite(widthPx) || widthPx <= 0 || !color) {
        return undefined;
      }
      return { widthPt: widthPx * layoutScale * scale, color, style: borderStyle };
    };
    cellBorders.push(
      ...cellBorderItems(box, {
        top: edge(style.borderTopStyle, style.borderTopWidth, style.borderTopColor),
        right: edge(style.borderRightStyle, style.borderRightWidth, style.borderRightColor),
        bottom: edge(style.borderBottomStyle, style.borderBottomWidth, style.borderBottomColor),
        left: edge(style.borderLeftStyle, style.borderLeftWidth, style.borderLeftColor),
      }),
    );
  }
  paths.push(...cellFills, ...cellBorders);

  for (const layer of Array.from(pageEl.querySelectorAll("svg.doc-connector-layer"))) {
    if (layer.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const layerRect = layer.getBoundingClientRect();
    const originXPt = (layerRect.left - pageRect.left) * scale;
    const originYPt = (layerRect.top - pageRect.top) * scale;
    const unit = layoutScale * scale;
    const zIndex = Number.parseInt(view.getComputedStyle(layer).zIndex, 10);
    const onTop = Number.isFinite(zIndex) && zIndex > 0;
    for (const element of Array.from(layer.querySelectorAll("path"))) {
      if (element.closest(EXCLUDED_SELECTOR)) {
        continue;
      }
      const subpaths = parseSvgPath(element.getAttribute("d") ?? "");
      if (!subpaths || subpaths.length === 0) {
        continue;
      }
      const style = view.getComputedStyle(element);
      const strokeColor = parseCssPaint(style.stroke);
      const fillColor = parseCssPaint(style.fill);
      const strokeWidthPx = Number.parseFloat(style.strokeWidth);
      const item: PdfPathItem = {
        subpaths: subpaths.map(
          (subpath): PdfPathSubpath => ({
            start: { xPt: originXPt + subpath.start.x * unit, yPt: originYPt + subpath.start.y * unit },
            segments: subpath.segments.map((segment) => ({
              toXPt: originXPt + segment.to.x * unit,
              toYPt: originYPt + segment.to.y * unit,
              c1XPt: segment.c1 ? originXPt + segment.c1.x * unit : undefined,
              c1YPt: segment.c1 ? originYPt + segment.c1.y * unit : undefined,
              c2XPt: segment.c2 ? originXPt + segment.c2.x * unit : undefined,
              c2YPt: segment.c2 ? originYPt + segment.c2.y * unit : undefined,
            })),
            closed: subpath.closed,
          }),
        ),
        onTop,
      };
      if (strokeColor && Number.isFinite(strokeWidthPx) && strokeWidthPx > 0) {
        const dashPt = parseDashArray(style.strokeDasharray).map((dash) => dash * unit);
        item.stroke = {
          color: strokeColor,
          widthPt: strokeWidthPx * unit,
          dashPt: dashPt.length > 0 ? dashPt : undefined,
          cap: LINE_CAPS[style.strokeLinecap] ?? 0,
          join: LINE_JOINS[style.strokeLinejoin] ?? 0,
        };
      }
      if (fillColor) {
        item.fill = fillColor;
      }
      if (item.stroke || item.fill) {
        paths.push(item);
      }
    }
  }

  for (const anchor of Array.from(pageEl.querySelectorAll("a[data-doc-link]"))) {
    if (anchor.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const href = anchor.getAttribute("href") ?? "";
    if (href.length === 0) {
      continue;
    }
    let target: Pick<PdfLinkItem, "href" | "destPage">;
    if (href.startsWith("#")) {
      const destPage = options.destPageForAnchor?.(href.slice(1));
      if (destPage === null || destPage === undefined) {
        continue;
      }
      target = { destPage };
    } else {
      target = { href };
    }
    for (const rect of Array.from(anchor.getClientRects())) {
      if (rect.width <= 0.1 || rect.height <= 0.1) {
        continue;
      }
      links.push({
        xPt: (rect.left - pageRect.left) * scale,
        yPt: (rect.top - pageRect.top) * scale,
        widthPt: rect.width * scale,
        heightPt: rect.height * scale,
        ...target,
      });
    }
  }

  for (const heading of Array.from(pageEl.querySelectorAll("h1, h2, h3"))) {
    if (heading.closest(NON_BODY_SELECTOR) || heading.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const title = (heading.textContent ?? "").trim();
    if (title.length === 0) {
      continue;
    }
    const rect = heading.getBoundingClientRect();
    if (rect.width <= 0 && rect.height <= 0) {
      continue;
    }
    headings.push({
      level: Number(heading.tagName.slice(1)),
      title,
      yPt: (rect.top - pageRect.top) * scale,
    });
  }

  return withRects({ widthPt, heightPt, items, links, headings, images, paths });
}
