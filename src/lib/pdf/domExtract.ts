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
  PdfTextItem,
} from "./buildPdf";
import { faceKeyFor } from "./fontMap";
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

export interface ExtractOptions {
  /** Resolves an internal link anchor (href="#name") to its 0-based target
   * page index, or null when the anchor no longer exists. */
  destPageForAnchor?: (anchor: string) => number | null;
}

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
      fragments.push({ text: textContent.slice(pos, lo), rect });
    }
    pos = lo;
  }
  return fragments;
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
  const doc = pageEl.ownerDocument;
  const view = doc.defaultView;
  if (!view) {
    return { widthPt, heightPt, items, links, headings, images, paths };
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
    for (const fragment of lineFragmentsOf(node as Text)) {
      if (fragment.text.trim().length === 0) {
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
  for (const element of Array.from(pageEl.querySelectorAll("img.doc-image"))) {
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
  for (const element of Array.from(pageEl.querySelectorAll("canvas.doc-shape-canvas"))) {
    if (element.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const anchorEl = element.closest(".doc-shape-anchor");
    if (anchorEl && anchorEl.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const canvas = element as HTMLCanvasElement;
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
    const shapeEl = canvas.closest(".doc-shape");
    const zIndex = shapeEl ? Number.parseInt(view.getComputedStyle(shapeEl).zIndex, 10) : NaN;
    images.push({
      src,
      xPt: (rect.left - pageRect.left) * scale,
      yPt: (rect.top - pageRect.top) * scale,
      widthPt: rect.width * scale,
      heightPt: rect.height * scale,
      onTop: Number.isFinite(zIndex) && zIndex > 0,
    });
  }

  // Charts (SPEC-052 AC-34, AC-35) paint into a real <canvas> (DocChartBlock.vue)
  // exactly like a shape's — reused here the same way, at the canvas's own
  // painted pixel resolution (devicePixelRatio-scaled, not the CSS display
  // size stretched up), never an upscaled bitmap.
  for (const element of Array.from(pageEl.querySelectorAll("canvas.doc-chart-canvas"))) {
    if (element.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const anchorEl = element.closest(".doc-chart-anchor");
    if (anchorEl && anchorEl.closest(EXCLUDED_SELECTOR)) {
      continue;
    }
    const canvas = element as HTMLCanvasElement;
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
    const chartEl = canvas.closest(".doc-chart");
    const zIndex = chartEl ? Number.parseInt(view.getComputedStyle(chartEl).zIndex, 10) : NaN;
    images.push({
      src,
      xPt: (rect.left - pageRect.left) * scale,
      yPt: (rect.top - pageRect.top) * scale,
      widthPt: rect.width * scale,
      heightPt: rect.height * scale,
      onTop: Number.isFinite(zIndex) && zIndex > 0,
    });
  }

  // Connectors (SPEC-070) are the one thing the page paints as SVG rather
  // than into a canvas, and they carry a stroke, a dash and arrowheads —
  // rasterizing them would throw all of that away, so they become real PDF
  // vector paths. The layer covers its page box, so its user units are the
  // page's own layout px: `layoutScale` converts those to the on-screen px
  // the rest of this function measures in (they differ under a zoom
  // transform), and `scale` then converts to points.
  const layoutScale = pageEl.offsetWidth > 0 ? pageRect.width / pageEl.offsetWidth : 1;
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

  return { widthPt, heightPt, items, links, headings, images, paths };
}
