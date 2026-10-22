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
