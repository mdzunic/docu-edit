// The DOM-bound extraction is covered by the e2e suite (it needs real
// layout geometry); these tests pin the pure helpers — colour parsing and
// the cell-chrome rectangle geometry (SPEC-076 AC-29).
import { describe, expect, it } from "vitest";
import {
  cellBorderItems,
  edgePaintItems,
  filledRectItem,
  parseCssColor,
  parseCssPaint,
  parseOpaqueCssColor,
  type PdfRect,
  DEFAULT_CANVAS_SELECTOR,
  DEFAULT_IMAGE_SELECTOR,
  extractPageContent,
  rectItemsFor,
  DEFAULT_CELL_SELECTOR,
  type BorderSide,
} from "./pdfExtract";

describe("parseOpaqueCssColor", () => {
  it("parses the opaque rgb() form computed styles produce", () => {
    expect(parseOpaqueCssColor("rgb(217, 226, 243)")).toEqual([217 / 255, 226 / 255, 243 / 255]);
  });

  it("accepts an rgba() whose alpha is exactly 1", () => {
    expect(parseOpaqueCssColor("rgba(47, 84, 150, 1)")).toEqual([47 / 255, 84 / 255, 150 / 255]);
  });

  it("rejects any translucent colour — the selection highlight must not print", () => {
    expect(parseOpaqueCssColor("rgba(66, 133, 244, 0.22)")).toBeNull();
    expect(parseOpaqueCssColor("rgba(0, 0, 0, 0)")).toBeNull();
  });

  it("rejects non-rgb forms rather than defaulting to black", () => {
    expect(parseOpaqueCssColor("transparent")).toBeNull();
    expect(parseOpaqueCssColor("none")).toBeNull();
    expect(parseOpaqueCssColor("")).toBeNull();
  });

  it("is stricter than parseCssPaint, which admits partial alpha", () => {
    expect(parseCssPaint("rgba(66, 133, 244, 0.22)")).toEqual([66 / 255, 133 / 255, 244 / 255]);
    expect(parseOpaqueCssColor("rgba(66, 133, 244, 0.22)")).toBeNull();
  });

  it("shares parseCssColor's channel clamping", () => {
    expect(parseCssColor("rgb(300, 0, 0)")).toEqual([1, 0, 0]);
  });
});

describe("filledRectItem", () => {
  it("builds one closed subpath tracing the rectangle, filled and unstroked", () => {
    const item = filledRectItem({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 }, [1, 0, 0]);
    expect(item.fill).toEqual([1, 0, 0]);
    expect(item.stroke).toBeUndefined();
    expect(item.onTop).toBeUndefined();
    expect(item.subpaths).toHaveLength(1);
    const [subpath] = item.subpaths;
    expect(subpath.closed).toBe(true);
    expect(subpath.start).toEqual({ xPt: 10, yPt: 20 });
    expect(subpath.segments.map((s) => [s.toXPt, s.toYPt])).toEqual([
      [110, 20],
      [110, 70],
      [10, 70],
    ]);
    // Straight lines only — no stray curve control points.
    expect(subpath.segments.every((s) => s.c1XPt === undefined && s.c2XPt === undefined)).toBe(true);
  });
});

describe("cellBorderItems", () => {
  const box: PdfRect = { xPt: 100, yPt: 200, widthPt: 60, heightPt: 40 };
  const black: [number, number, number] = [0, 0, 0];
  const grey: [number, number, number] = [0.5, 0.5, 0.5];

  it("returns nothing for a cell with no painted edges", () => {
    expect(cellBorderItems(box, {})).toEqual([]);
  });

  it("skips absent edges and keeps the present ones", () => {
    const items = cellBorderItems(box, { top: { widthPt: 2, color: black } });
    expect(items).toHaveLength(1);
    expect(items[0].fill).toEqual(black);
  });

  it("centres each strip on the cell-rect edge", () => {
    const [top] = cellBorderItems(box, { top: { widthPt: 2, color: black } });
    // The strip straddles y=200: from 199 to 201, spanning the full width.
    expect(top.subpaths[0].start).toEqual({ xPt: 100, yPt: 199 });
    expect(top.subpaths[0].segments[1]).toMatchObject({ toXPt: 160, toYPt: 201 });

    const [right] = cellBorderItems(box, { right: { widthPt: 4, color: grey } });
    // The strip straddles x=160: from 158 to 162, spanning the full height.
    expect(right.subpaths[0].start).toEqual({ xPt: 158, yPt: 200 });
    expect(right.subpaths[0].segments[1]).toMatchObject({ toXPt: 162, toYPt: 240 });
  });

  it("extends strips by the perpendicular edges' half-widths so corners close", () => {
    const edges = {
      top: { widthPt: 2, color: black },
      right: { widthPt: 2, color: black },
      bottom: { widthPt: 2, color: black },
      left: { widthPt: 2, color: black },
    };
    const [top, bottom, left, right] = cellBorderItems(box, edges);
    // Horizontal strips reach the outer face of the vertical ones (99..161)…
    expect(top.subpaths[0].start).toEqual({ xPt: 99, yPt: 199 });
    expect(top.subpaths[0].segments[1]).toMatchObject({ toXPt: 161, toYPt: 201 });
    expect(bottom.subpaths[0].start).toEqual({ xPt: 99, yPt: 239 });
    expect(bottom.subpaths[0].segments[1]).toMatchObject({ toXPt: 161, toYPt: 241 });
    // …and the vertical strips likewise (199..241).
    expect(left.subpaths[0].start).toEqual({ xPt: 99, yPt: 199 });
    expect(left.subpaths[0].segments[1]).toMatchObject({ toXPt: 101, toYPt: 241 });
    expect(right.subpaths[0].start).toEqual({ xPt: 159, yPt: 199 });
    expect(right.subpaths[0].segments[1]).toMatchObject({ toXPt: 161, toYPt: 241 });
  });

  it("carries each edge's own colour", () => {
    const items = cellBorderItems(box, {
      top: { widthPt: 1, color: black },
      bottom: { widthPt: 1, color: grey },
    });
    expect(items.map((item) => item.fill)).toEqual([black, grey]);
  });
});

describe("edgePaintItems — non-solid border styles (AC-29)", () => {
  const black: [number, number, number] = [0, 0, 0];
  // A horizontal top strip as cellBorderItems would build it: 60 wide, 3 thick.
  const strip: PdfRect = { xPt: 100, yPt: 198.5, widthPt: 60, heightPt: 3 };

  it("a dashed edge strokes the strip's centre line with a 3-on/3-off dash", () => {
    const items = edgePaintItems(strip, { widthPt: 3, color: black, style: "dashed" });
    expect(items).toHaveLength(1);
    expect(items[0].fill).toBeUndefined();
    expect(items[0].stroke).toEqual({ color: black, widthPt: 3, dashPt: [9, 9] });
    expect(items[0].subpaths[0].start).toEqual({ xPt: 100, yPt: 200 });
    expect(items[0].subpaths[0].segments[0]).toMatchObject({ toXPt: 160, toYPt: 200 });
  });

  it("a dotted edge uses width-sized dots", () => {
    const [item] = edgePaintItems(strip, { widthPt: 3, color: black, style: "dotted" });
    expect(item.stroke?.dashPt).toEqual([3, 3]);
  });

  it("a vertical dashed edge strokes vertically", () => {
    const vertical: PdfRect = { xPt: 158.5, yPt: 200, widthPt: 3, heightPt: 40 };
    const [item] = edgePaintItems(vertical, { widthPt: 3, color: black, style: "dashed" });
    expect(item.subpaths[0].start).toEqual({ xPt: 160, yPt: 200 });
    expect(item.subpaths[0].segments[0]).toMatchObject({ toXPt: 160, toYPt: 240 });
  });

  it("a double edge paints two strips of a third of the width, at the strip's faces", () => {
    const items = edgePaintItems(strip, { widthPt: 3, color: black, style: "double" });
    expect(items).toHaveLength(2);
    expect(items[0].subpaths[0].start).toEqual({ xPt: 100, yPt: 198.5 });
    expect(items[0].subpaths[0].segments[1]).toMatchObject({ toXPt: 160, toYPt: 199.5 });
    expect(items[1].subpaths[0].start).toEqual({ xPt: 100, yPt: 200.5 });
    expect(items[1].subpaths[0].segments[1]).toMatchObject({ toXPt: 160, toYPt: 201.5 });
  });

  it("a double edge too thin for its gap falls back to the solid strip, as the browser does", () => {
    const thin: PdfRect = { xPt: 100, yPt: 199.5, widthPt: 60, heightPt: 1 };
    const items = edgePaintItems(thin, { widthPt: 1, color: black, style: "double" });
    expect(items).toHaveLength(1);
    expect(items[0].fill).toEqual(black);
  });

  it("a solid or unknown style stays the filled strip", () => {
    for (const style of ["solid", "groove", undefined]) {
      const items = edgePaintItems(strip, { widthPt: 3, color: black, style });
      expect(items).toHaveLength(1);
      expect(items[0].fill).toEqual(black);
      expect(items[0].stroke).toBeUndefined();
    }
  });

  it("cellBorderItems routes each edge's style through", () => {
    const box: PdfRect = { xPt: 100, yPt: 200, widthPt: 60, heightPt: 40 };
    const items = cellBorderItems(box, {
      top: { widthPt: 2, color: black, style: "dashed" },
      bottom: { widthPt: 2, color: black },
    });
    expect(items[0].stroke?.dashPt).toEqual([6, 6]);
    expect(items[1].fill).toEqual(black);
  });
});

// --- SPEC-104: the rect primitive and the configurable selectors -----------

const NO_SIDE: BorderSide = { widthPt: 0, style: "none", color: "rgb(0, 0, 0)" };
const box: PdfRect = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };

describe("rectItemsFor (SPEC-104)", () => {
  it("emits a background fill over the whole border box", () => {
    expect(
      rectItemsFor(box, "rgb(255, 0, 0)", { top: NO_SIDE, right: NO_SIDE, bottom: NO_SIDE, left: NO_SIDE }),
    ).toEqual([{ ...box, fill: [1, 0, 0] }]);
  });

  it("skips a fully transparent background", () => {
    const none = { top: NO_SIDE, right: NO_SIDE, bottom: NO_SIDE, left: NO_SIDE };
    expect(rectItemsFor(box, "rgba(0, 0, 0, 0)", none)).toEqual([]);
    expect(rectItemsFor(box, "transparent", none)).toEqual([]);
  });

  it("strokes a uniform border inset by half its width, carrying the fill too", () => {
    const side: BorderSide = { widthPt: 2, style: "solid", color: "rgb(0, 0, 0)" };
    expect(rectItemsFor(box, "rgb(255, 255, 255)", { top: side, right: side, bottom: side, left: side })).toEqual([
      { xPt: 11, yPt: 21, widthPt: 98, heightPt: 48, stroke: [0, 0, 0], strokeWidthPt: 2, fill: [1, 1, 1] },
    ]);
  });

  it("paints a non-uniform border as one filled strip per painted side", () => {
    const top: BorderSide = { widthPt: 0.5, style: "solid", color: "rgb(128, 128, 128)" };
    const items = rectItemsFor(box, "transparent", { top, right: NO_SIDE, bottom: NO_SIDE, left: NO_SIDE });
    expect(items).toEqual([{ ...box, heightPt: 0.5, fill: [128 / 255, 128 / 255, 128 / 255] }]);
  });
});

/** Stubs an element's on-screen box (jsdom lays nothing out). */
function placeAt(el: Element, left: number, top: number, width: number, height: number): void {
  el.getBoundingClientRect = () =>
    ({ left, top, width, height, right: left + width, bottom: top + height, x: left, y: top, toJSON: () => ({}) }) as DOMRect;
}

function pageWith(html: string): HTMLElement {
  const page = document.createElement("div");
  page.innerHTML = html;
  document.body.append(page);
  placeAt(page, 0, 0, 400, 300);
  return page;
}

describe("extractPageContent selectors (SPEC-104)", () => {
  it("keeps SPEC-034's defaults: doc images and doc shape/chart canvases", () => {
    expect(DEFAULT_IMAGE_SELECTOR).toBe("img.doc-image");
    expect(DEFAULT_CANVAS_SELECTOR).toBe("canvas.doc-shape-canvas, canvas.doc-chart-canvas");
  });

  it("emits no rects key without a rectSelector, so a SPEC-034 page input is unchanged", () => {
    const page = pageWith('<div class="slide-bg" style="background-color: rgb(255, 0, 0)"></div>');
    placeAt(page.querySelector(".slide-bg")!, 0, 0, 400, 300);
    const input = extractPageContent(page, 300, 225);
    expect("rects" in input).toBe(false);
    expect(Object.keys(input).toSorted()).toEqual(["headings", "heightPt", "images", "items", "links", "paths", "widthPt"]);
    page.remove();
  });

  it("extracts a matched element's background and uniform border as rects, in points", () => {
    const page = pageWith(
      '<div class="slide-bg" style="background-color: rgb(0, 0, 255)"></div>' +
        '<div class="frame-box" style="border: 2px solid rgb(0, 0, 0)"></div>' +
        '<div class="slide-bg print-exclude" style="background-color: rgb(255, 0, 0)"></div>',
    );
    const [bg, frame, excluded] = Array.from(page.children);
    placeAt(bg, 0, 0, 400, 300);
    placeAt(frame, 40, 40, 80, 40);
    placeAt(excluded, 0, 0, 400, 300);
    // 400 px wide on screen → 300 pt: 0.75 pt per px.
    const input = extractPageContent(page, 300, 225, { rectSelector: ".slide-bg, .frame-box" });
    expect(input.rects).toEqual([
      { xPt: 0, yPt: 0, widthPt: 300, heightPt: 225, fill: [0, 0, 1] },
      { xPt: 30.75, yPt: 30.75, widthPt: 58.5, heightPt: 28.5, stroke: [0, 0, 0], strokeWidthPt: 1.5 },
    ]);
    page.remove();
  });

  it("places images by the imageSelector it is given, and only those", () => {
    const page = pageWith('<img class="doc-image" src="data:image/png;base64,AA=="><img class="frame-image" src="data:image/png;base64,AB==">');
    const [docImage, frameImage] = Array.from(page.querySelectorAll("img"));
    placeAt(docImage, 0, 0, 40, 40);
    placeAt(frameImage, 100, 100, 40, 40);
    const byDefault = extractPageContent(page, 300, 225);
    expect(byDefault.images.map((image) => image.src)).toEqual(["data:image/png;base64,AA=="]);
    const slides = extractPageContent(page, 300, 225, { imageSelector: "img.frame-image" });
    expect(slides.images).toEqual([
      { src: "data:image/png;base64,AB==", xPt: 75, yPt: 75, widthPt: 30, heightPt: 30, onTop: false },
    ]);
    page.remove();
  });
});

/** A run of 10 px characters' box, [start, end) from x = 0. */
function charRect(start: number, end: number): DOMRect {
  return { left: start * 10, right: end * 10, top: 0, bottom: 16, width: (end - start) * 10, height: 16, x: start * 10, y: 0, toJSON: () => ({}) } as DOMRect;
}

describe("extractPageContent cellSelector and clipSelector (SPEC-136)", () => {
  it("defaults to the text product's table cells; a cellSelector reads another element's fill and borders", () => {
    expect(DEFAULT_CELL_SELECTOR).toBe("td.doc-table-cell");
    const page = pageWith('<div class="sheet-print-cell" style="background-color: rgb(255, 255, 0); border-top: 2px solid rgb(0, 0, 0)"></div>');
    placeAt(page.querySelector(".sheet-print-cell")!, 40, 20, 100, 30);
    expect(extractPageContent(page, 300, 225).paths).toEqual([]);
    const paths = extractPageContent(page, 300, 225, { cellSelector: ".sheet-print-cell" }).paths!;
    expect(paths[0]).toEqual(filledRectItem({ xPt: 30, yPt: 15, widthPt: 75, heightPt: 22.5 }, [1, 1, 0]));
    expect(paths.length).toBeGreaterThan(1);
    page.remove();
  });

  it("clips text inside a clipSelector element to its box — whole characters only — and leaves text alone without one", () => {
    const page = pageWith('<div class="clip"><span style="font-size: 16px">ABCDEF</span></div>');
    const clip = page.querySelector(".clip")!;
    placeAt(clip, 0, 0, 30, 20);
    const text = page.querySelector("span")!.firstChild as Text;
    // Each character is 10 px wide from x = 0 (charRect): D, E and F lie past the 30 px box.
    const ranges = { start: 0, end: 0 };
    const proto = Range.prototype as unknown as Record<string, unknown>;
    const saved = { setStart: proto.setStart, setEnd: proto.setEnd, getClientRects: proto.getClientRects, getBoundingClientRect: proto.getBoundingClientRect };
    proto.setStart = function (this: Range, _node: Node, offset: number) {
      ranges.start = offset;
    };
    proto.setEnd = function (this: Range, _node: Node, offset: number) {
      ranges.end = offset;
    };
    proto.getClientRects = () => [charRect(ranges.start, ranges.end)];
    proto.getBoundingClientRect = () => charRect(ranges.start, ranges.end);
    try {
      expect(text.data).toBe("ABCDEF");
      expect(extractPageContent(page, 300, 225).items.map((item) => item.text)).toEqual(["ABCDEF"]);
      expect(extractPageContent(page, 300, 225, { clipSelector: ".clip" }).items.map((item) => item.text)).toEqual(["ABC"]);
    } finally {
      Object.assign(proto, saved);
      page.remove();
    }
  });
});
