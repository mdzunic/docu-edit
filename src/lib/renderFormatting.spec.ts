import { describe, it, expect } from "vitest";
import {
  displayParagraphFormatting,
  pageBorderStyle,
  displayRunFormatting,
  drawingWrapSpacerStyle,
  headingTag,
  paragraphHasTrackedChange,
  paragraphStyle,
  runKey,
  runStyle,
  shapeDisplayStyle,
  shapeTextRectStyle,
  tableCellStyle,
  tableColumnWidth,
  tableRowStyle,
  tableStyle,
  trackChangeBlockHidden,
  trackChangeRunHidden,
  trackChangeRunStyle,
} from "./renderFormatting";
import type { Paragraph, Run, Shape, Table, TableCell, TableRow } from "./documentModel";

function paragraph(runs: Run[]): Paragraph {
  return { type: "paragraph", runs };
}

function testShape(overrides: Partial<Shape> = {}): Shape {
  return { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square", ...overrides };
}

describe("headingTag", () => {
  it("returns 'p' for no heading level", () => {
    expect(headingTag(undefined)).toBe("p");
  });

  it("returns hN for a heading level", () => {
    expect(headingTag(1)).toBe("h1");
    expect(headingTag(3)).toBe("h3");
  });

  it("clamps levels above 6 to h6", () => {
    expect(headingTag(9)).toBe("h6");
  });
});

describe("runStyle", () => {
  it("maps bold, italic, underline and strikethrough", () => {
    const style = runStyle({ bold: true, italic: true, underline: true, strikethrough: true });
    expect(style.fontWeight).toBe("bold");
    expect(style.fontStyle).toBe("italic");
    expect(style.textDecorationLine).toBe("underline line-through");
  });

  it("maps font family to its metric-twin stack, size and colour (SPEC-048 AC-2/AC-5)", () => {
    const style = runStyle({ fontFamily: "Arial", fontSizePt: 14, color: "#1155cc" });
    // The run keeps "Arial" in the model; the CSS renders the original if
    // installed, else the bundled metric twin, so layout matches Word.
    expect(style.fontFamily).toBe("Arial, Liberation Sans, sans-serif");
    expect(style.fontSize).toBe("14pt");
    expect(style.color).toBe("#1155cc");
  });

  it("maps a known highlight name to its CSS colour", () => {
    expect(runStyle({ highlight: "yellow" }).backgroundColor).toBe("#ffff00");
  });

  it("falls back to the raw value for an unknown highlight name", () => {
    expect(runStyle({ highlight: "mystery" }).backgroundColor).toBe("mystery");
  });

  it("returns an empty style object for empty formatting", () => {
    expect(runStyle({})).toEqual({});
  });

  it("paints nothing for an explicit 'No colour' highlight (SPEC-013)", () => {
    expect(runStyle({ highlight: false }).backgroundColor).toBeUndefined();
  });

  it("paints an out-of-palette highlight hex as-is (SPEC-013)", () => {
    expect(runStyle({ highlight: "#ff8800" }).backgroundColor).toBe("#ff8800");
  });

  it("maps the white palette entry, which the picker hides (SPEC-013)", () => {
    expect(runStyle({ highlight: "white" }).backgroundColor).toBe("#ffffff");
  });

  it("renders sub/superscript shifted and reduced (SPEC-013)", () => {
    expect(runStyle({ vertAlign: "subscript" }).verticalAlign).toBe("sub");
    expect(runStyle({ vertAlign: "subscript" }).fontSize).toBe("0.65em");
    const sized = runStyle({ vertAlign: "superscript", fontSizePt: 20 });
    expect(sized.verticalAlign).toBe("super");
    expect(sized.fontSize).toBe("13pt");
  });

  it("never paints the shadingFill preservation slot (SPEC-013)", () => {
    expect(runStyle({ highlight: "yellow", shadingFill: "#ff8800" }).backgroundColor).toBe(
      "#ffff00",
    );
  });
});

describe("paragraphStyle", () => {
  it("always sets white-space: pre-wrap so tabs/line breaks render", () => {
    expect(paragraphStyle({ type: "paragraph", runs: [] }).whiteSpace).toBe("pre-wrap");
  });

  it("maps alignment, indents and spacing", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: {
        alignment: "center",
        indentStartCm: 1.27,
        indentEndCm: 1.27,
        indentFirstLineCm: 0.5,
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        lineSpacing: 1.5,
      },
    });
    expect(style.textAlign).toBe("center");
    expect(style.marginInlineStart).toBe("1.27cm");
    expect(style.marginInlineEnd).toBe("1.27cm");
    expect(style.textIndent).toBe("0.5cm");
    expect(style.marginTop).toBe("12pt");
    expect(style.marginBottom).toBe("6pt");
    expect(style.lineHeight).toBe("1.5");
  });

  it("prefers an absolute lineSpacingPt over a lineSpacing multiplier", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { lineSpacing: 1.5, lineSpacingPt: 14 },
    });
    expect(style.lineHeight).toBe("14pt");
  });

  it("sets direction: rtl so logical indents mirror physically (SPEC-014)", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { direction: "rtl", indentStartCm: 2, indentEndCm: 1 },
    });
    expect(style.direction).toBe("rtl");
    // The values stay logical — the browser's inline axis does the mirroring.
    expect(style.marginInlineStart).toBe("2cm");
    expect(style.marginInlineEnd).toBe("1cm");
  });

  it("leaves direction unset for LTR paragraphs", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { indentStartCm: 2 },
    });
    expect(style.direction).toBeUndefined();
  });

  it("indents a list item by its level (SPEC-018 AC-5)", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      list: { type: "bullet", level: 3 },
    });
    expect(style.marginInlineStart).toBe("3.81cm");
  });

  it("prefers an imported per-level indent, and a direct start indent over both (SPEC-018)", () => {
    const imported = paragraphStyle({
      type: "paragraph",
      runs: [],
      list: { type: "number", level: 1, indentCm: 2.5 },
    });
    expect(imported.marginInlineStart).toBe("2.5cm");
    const direct = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { indentStartCm: 4 },
      list: { type: "number", level: 1, indentCm: 2.5 },
    });
    expect(direct.marginInlineStart).toBe("4cm");
  });
});

describe("runKey", () => {
  it("produces a stable key from index and text length", () => {
    expect(runKey({ text: "hello", formatting: {} }, 2)).toBe("2-5");
  });
});

describe("tableCellStyle (SPEC-023 AC-10..17)", () => {
  it("maps each set border side and leaves unset ones alone", () => {
    const cell: TableCell = {
      blocks: [],
      borders: {
        top: { style: "single", widthPt: 1, color: "#ff0000" },
        left: { style: "double", widthPt: 1.5, color: "#00ff00" },
      },
    };
    const style = tableCellStyle(cell);
    expect(style.borderTop).toBe("1pt solid #ff0000");
    expect(style.borderLeft).toBe("1.5pt double #00ff00");
    expect(style.borderBottom).toBeUndefined();
    expect(style.borderRight).toBeUndefined();
  });

  it("renders a 'none' border explicitly so it can clear a default gridline", () => {
    const cell: TableCell = { blocks: [], borders: { top: { style: "none", widthPt: 0, color: "#000000" } } };
    expect(tableCellStyle(cell).borderTop).toBe("none");
  });

  it("maps shading and padding", () => {
    const cell: TableCell = { blocks: [], shadingFill: "#eeeeee", paddingCm: 0.3 };
    const style = tableCellStyle(cell);
    expect(style.backgroundColor).toBe("#eeeeee");
    expect(style.padding).toBe("0.3cm");
  });

  it("maps vertical alignment to its CSS keyword — the model's 'center' is CSS's 'middle', not the invalid 'center'", () => {
    expect(tableCellStyle({ blocks: [], verticalAlign: "center" }).verticalAlign).toBe("middle");
    expect(tableCellStyle({ blocks: [], verticalAlign: "top" }).verticalAlign).toBe("top");
    expect(tableCellStyle({ blocks: [], verticalAlign: "bottom" }).verticalAlign).toBe("bottom");
  });

  it("returns an empty style for a plain cell", () => {
    expect(tableCellStyle({ blocks: [] })).toEqual({});
  });
});

describe("tableRowStyle (AC-4, AC-5)", () => {
  it("maps a minimum height, a floor rather than a fixed size", () => {
    const row: TableRow = { cells: [], minHeightCm: 1.5 };
    expect(tableRowStyle(row)).toEqual({ height: "1.5cm" });
  });

  it("is empty without a minimum height", () => {
    expect(tableRowStyle({ cells: [] })).toEqual({});
  });
});

function table(overrides: Partial<Table> = {}): Table {
  return { type: "table", rows: [], columnWidthsCm: [5, 5], ...overrides };
}

describe("tableStyle (AC-22, AC-23)", () => {
  it("defaults to a fixed width from the column widths' sum", () => {
    const style = tableStyle(table());
    expect(style.tableLayout).toBe("fixed");
    expect(style.width).toBe("10cm");
  });

  it("uses the percentage width when set to percent", () => {
    const style = tableStyle(table({ widthMode: "percent", widthPercent: 60 }));
    expect(style.width).toBe("60%");
  });

  it("hands sizing to the browser for auto width", () => {
    const style = tableStyle(table({ widthMode: "auto" }));
    expect(style.tableLayout).toBe("auto");
    expect(style.width).toBeUndefined();
  });

  it("centres or right-aligns the table on the page", () => {
    expect(tableStyle(table({ alignment: "center" }))).toMatchObject({ marginLeft: "auto", marginRight: "auto" });
    expect(tableStyle(table({ alignment: "right" }))).toMatchObject({ marginLeft: "auto", marginRight: "0" });
    expect(tableStyle(table({ alignment: "left" })).marginLeft).toBeUndefined();
  });
});

describe("tableColumnWidth (AC-23)", () => {
  it("uses the column's own absolute width for a fixed-width table", () => {
    expect(tableColumnWidth(table({ columnWidthsCm: [4, 6] }), 0)).toBe("4cm");
    expect(tableColumnWidth(table({ columnWidthsCm: [4, 6] }), 1)).toBe("6cm");
  });

  it("uses a relative share for a percent-width table, not the absolute width", () => {
    const percentTable = table({ columnWidthsCm: [4, 6], widthMode: "percent", widthPercent: 50 });
    expect(tableColumnWidth(percentTable, 0)).toBe("40%");
    expect(tableColumnWidth(percentTable, 1)).toBe("60%");
  });

  it("uses a relative share for an auto-width table too", () => {
    const autoTable = table({ columnWidthsCm: [1, 1, 2], widthMode: "auto" });
    expect(tableColumnWidth(autoTable, 2)).toBe("50%");
  });
});

describe("trackChangeRunStyle (SPEC-039 AC-4)", () => {
  const insertedRun: Run = {
    text: "x",
    formatting: {},
    changes: { insert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" } },
  };
  const deletedRun: Run = {
    text: "x",
    formatting: {},
    changes: { delete: { changeId: "c2", author: "Grace", date: "2026-01-01T00:00:00.000Z" } },
  };

  it("underlines an insertion in the author's colour in 'all' markup", () => {
    const style = trackChangeRunStyle(insertedRun, "all");
    expect(style.textDecorationLine).toBe("underline");
    expect(style.color).toBeTruthy();
  });

  it("strikes through a deletion in the author's colour in 'all' markup", () => {
    const style = trackChangeRunStyle(deletedRun, "all");
    expect(style.textDecorationLine).toBe("line-through");
  });

  it("gives the same author the same colour every time", () => {
    const a = trackChangeRunStyle(insertedRun, "all");
    const b = trackChangeRunStyle({ ...insertedRun, text: "y" }, "all");
    expect(a.color).toBe(b.color);
  });

  it("shows no inline markup outside 'all' markup mode", () => {
    expect(trackChangeRunStyle(insertedRun, "simple")).toEqual({});
    expect(trackChangeRunStyle(insertedRun, "noMarkup")).toEqual({});
    expect(trackChangeRunStyle(insertedRun, "original")).toEqual({});
  });

  it("returns no style for a run with no tracked change", () => {
    expect(trackChangeRunStyle({ text: "x", formatting: {} }, "all")).toEqual({});
  });
});

describe("trackChangeRunHidden (SPEC-039 AC-9)", () => {
  const insertedRun: Run = {
    text: "x",
    formatting: {},
    changes: { insert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" } },
  };
  const deletedRun: Run = {
    text: "x",
    formatting: {},
    changes: { delete: { changeId: "c2", author: "Grace", date: "2026-01-01T00:00:00.000Z" } },
  };
  const plainRun: Run = { text: "x", formatting: {} };

  it("hides a pending deletion in 'No markup' (as if already accepted)", () => {
    expect(trackChangeRunHidden(deletedRun, "noMarkup")).toBe(true);
    expect(trackChangeRunHidden(insertedRun, "noMarkup")).toBe(false);
  });

  it("hides a pending insertion in 'Original' (as if rejected)", () => {
    expect(trackChangeRunHidden(insertedRun, "original")).toBe(true);
    expect(trackChangeRunHidden(deletedRun, "original")).toBe(false);
  });

  it("hides nothing in 'All markup' or 'Simple markup'", () => {
    expect(trackChangeRunHidden(insertedRun, "all")).toBe(false);
    expect(trackChangeRunHidden(deletedRun, "all")).toBe(false);
    expect(trackChangeRunHidden(insertedRun, "simple")).toBe(false);
    expect(trackChangeRunHidden(deletedRun, "simple")).toBe(false);
  });

  it("never hides a run with no tracked change", () => {
    expect(trackChangeRunHidden(plainRun, "noMarkup")).toBe(false);
    expect(trackChangeRunHidden(plainRun, "original")).toBe(false);
  });
});

describe("displayRunFormatting (SPEC-039 AC-9)", () => {
  it("shows the previous formatting in 'Original' for a pending format change", () => {
    const run: Run = {
      text: "x",
      formatting: { bold: true },
      changes: { formatChange: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z", previous: { bold: false } } },
    };
    expect(displayRunFormatting(run, "original")).toEqual({ bold: false });
    expect(displayRunFormatting(run, "all")).toEqual({ bold: true });
    expect(displayRunFormatting(run, "noMarkup")).toEqual({ bold: true });
  });

  it("returns the run's own formatting with no pending format change", () => {
    const run: Run = { text: "x", formatting: { italic: true } };
    expect(displayRunFormatting(run, "original")).toEqual({ italic: true });
  });
});

describe("displayParagraphFormatting (SPEC-039 AC-9)", () => {
  it("shows the previous formatting in 'Original' for a pending paragraph format change", () => {
    const p = paragraph([{ text: "x", formatting: {} }]);
    p.formatting = { alignment: "center" };
    p.changes = {
      formatChange: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z", previous: { alignment: "left" } },
    };
    expect(displayParagraphFormatting(p, "original")).toEqual({ alignment: "left" });
    expect(displayParagraphFormatting(p, "all")).toEqual({ alignment: "center" });
  });
});

describe("trackChangeBlockHidden (SPEC-039 AC-9)", () => {
  const insert = { insert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" } };
  const del = { delete: { changeId: "c2", author: "Grace", date: "2026-01-01T00:00:00.000Z" } };

  it("hides a tracked-deleted block in 'No markup'", () => {
    expect(trackChangeBlockHidden(del, "noMarkup")).toBe(true);
    expect(trackChangeBlockHidden(insert, "noMarkup")).toBe(false);
  });

  it("hides a tracked-inserted block in 'Original'", () => {
    expect(trackChangeBlockHidden(insert, "original")).toBe(true);
    expect(trackChangeBlockHidden(del, "original")).toBe(false);
  });

  it("hides nothing without pending changes, or outside noMarkup/original", () => {
    expect(trackChangeBlockHidden(undefined, "noMarkup")).toBe(false);
    expect(trackChangeBlockHidden(insert, "all")).toBe(false);
    expect(trackChangeBlockHidden(insert, "simple")).toBe(false);
  });
});

describe("paragraphHasTrackedChange (SPEC-039 AC-5)", () => {
  it("is false for a paragraph with no tracked changes", () => {
    expect(paragraphHasTrackedChange(paragraph([{ text: "x", formatting: {} }]))).toBe(false);
  });

  it("is true when a run carries an insert/delete/formatChange marker", () => {
    expect(
      paragraphHasTrackedChange(
        paragraph([{ text: "x", formatting: {}, changes: { insert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" } } }]),
      ),
    ).toBe(true);
  });

  it("is true when the paragraph's own mark is pending insert/delete", () => {
    const p = paragraph([{ text: "x", formatting: {} }]);
    p.changes = { markInsert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" } };
    expect(paragraphHasTrackedChange(p)).toBe(true);
  });
});

// SPEC-051 regression (AC-29..33): z-index was derived only from wrap mode,
// never from Shape.zOrder, so bring-to-front/forward/send-backward/back had
// no visual effect at all — every shape in the same wrap tier painted at the
// same fixed z-index regardless of reorderShapeZ's model changes.
describe("shapeDisplayStyle — z-order (SPEC-051 AC-29..33)", () => {
  it("breaks ties between same-tier shapes by zOrder, not a shared constant", () => {
    const back = shapeDisplayStyle(testShape({ zOrder: 0 }));
    const front = shapeDisplayStyle(testShape({ zOrder: 2 }));
    expect(Number(front.zIndex)).toBeGreaterThan(Number(back.zIndex));
  });

  it("keeps a 'behind' shape's z-index below every 'infront'/'square'/'topAndBottom' shape regardless of zOrder", () => {
    const behind = shapeDisplayStyle(testShape({ wrap: "behind", zOrder: 99 }));
    const inFront = shapeDisplayStyle(testShape({ wrap: "infront", zOrder: 0 }));
    expect(Number(behind.zIndex)).toBeLessThan(Number(inFront.zIndex));
  });

  it("defaults to zOrder 0 when absent, matching zOrderOf's own fallback", () => {
    expect(shapeDisplayStyle(testShape({ wrap: "square" })).zIndex).toBe(shapeDisplayStyle(testShape({ wrap: "square", zOrder: 0 })).zIndex);
  });
});

describe("shapeTextRectStyle — per-geometry text rectangle (SPEC-069 AC-11)", () => {
  it("returns null for a geometry that fills its box (default)", () => {
    expect(shapeTextRectStyle(testShape({ geometry: "rect" }))).toBeNull();
  });

  it("returns null for a preserved (unsupported) geometry", () => {
    expect(
      shapeTextRectStyle(testShape({ geometry: "rect", unsupportedGeometryXml: "<x/>" })),
    ).toBeNull();
  });

  it("insets text to the declared rectangle for a geometry that has one", () => {
    // The rounded callout lays text in its body (top three-quarters), not the
    // tail — width/height as percentages of the box, right/bottom released.
    const style = shapeTextRectStyle(testShape({ geometry: "wedgeRectCallout" }))!;
    expect(style).not.toBeNull();
    expect(style.left).toBe("0.000%");
    expect(style.top).toBe("0.000%");
    expect(style.width).toBe("100.000%");
    expect(parseFloat(style.height)).toBeCloseTo(75, 3);
    expect(style.right).toBe("auto");
    expect(style.bottom).toBe("auto");
  });
});

describe("drawingWrapSpacerStyle — square/topAndBottom text exclusion", () => {
  it("returns null for the wrap modes that leave text alone", () => {
    for (const wrap of ["inline", "behind", "infront"] as const) {
      expect(drawingWrapSpacerStyle(testShape({ wrap }), 16)).toBeNull();
    }
  });

  it("returns null before the column width has been measured", () => {
    expect(drawingWrapSpacerStyle(testShape(), 0)).toBeNull();
  });

  it("floats left, spanning to the shape's right edge plus the stand-off, when the right side has more room", () => {
    const style = drawingWrapSpacerStyle(testShape({ offsetXCm: 1, offsetYCm: 2 }), 16);
    expect(style).toEqual({
      float: "left",
      width: `${1 + 3 + 0.35}cm`,
      height: `${2 + 2 + 0.35}cm`,
      shapeOutside: "inset(2cm 0 0 0)",
    });
  });

  it("floats right, spanning from the shape's left edge minus the stand-off, when the left side has more room", () => {
    const style = drawingWrapSpacerStyle(testShape({ offsetXCm: 12 }), 16);
    expect(style).toEqual({
      float: "right",
      width: `${16 - 12 + 0.35}cm`,
      height: `${2 + 0.35}cm`,
      shapeOutside: "inset(0cm 0 0 0)",
    });
  });

  it("topAndBottom reserves the full column width for the shape's band only", () => {
    const style = drawingWrapSpacerStyle(testShape({ wrap: "topAndBottom", offsetYCm: 1.5 }), 16);
    expect(style).toEqual({
      float: "left",
      width: "100%",
      height: `${1.5 + 2 + 0.35}cm`,
      shapeOutside: "inset(1.5cm 0 0 0)",
    });
  });

  it("clamps negative offsets to the anchor edge and the spacer to the column width", () => {
    const style = drawingWrapSpacerStyle(
      testShape({ offsetXCm: -2, offsetYCm: -1, widthCm: 30 }),
      16,
    );
    expect(style!.shapeOutside).toBe("inset(0cm 0 0 0)");
    expect(style!.width).toBe("16cm");
  });
});

describe("pageBorderStyle (SPEC-074 AC-21, AC-22)", () => {
  const margins = { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 };
  const line = { style: "single" as const, widthPt: 2, color: "#c00000" };

  it("is null when the section has no border or no edge is set", () => {
    expect(pageBorderStyle(undefined, margins, 0)).toBeNull();
    expect(pageBorderStyle({ offsetPt: 24 }, margins, 0)).toBeNull();
  });

  it("insets from the page edge by the offset when it measures from the page", () => {
    const style = pageBorderStyle({ top: line, bottom: line, offsetPt: 36, measureFrom: "page" }, margins, 0);
    // 36pt = half an inch = 1.27cm.
    expect(style?.top).toBe("1.27cm");
    expect(style?.borderTop).toBe("2pt solid #c00000");
    expect(style?.borderLeft).toBe("none");
  });

  it("insets from the text area outward when it measures from the text", () => {
    const style = pageBorderStyle({ left: line, offsetPt: 36, measureFrom: "text" }, margins, 0);
    // The 2.54cm margin less the 1.27cm offset.
    expect(style?.left).toBe("1.27cm");
  });

  it("never insets past the page edge", () => {
    const style = pageBorderStyle({ top: line, offsetPt: 720, measureFrom: "text" }, margins, 0);
    expect(style?.top).toBe("0cm");
  });

  it("renders on every page of the section by default (AC-22)", () => {
    for (const index of [0, 1, 7]) {
      expect(pageBorderStyle({ top: line }, margins, index)).not.toBeNull();
    }
  });

  it("honours a first-page-only or all-but-first restriction", () => {
    expect(pageBorderStyle({ top: line, appliesTo: "firstPage" }, margins, 0)).not.toBeNull();
    expect(pageBorderStyle({ top: line, appliesTo: "firstPage" }, margins, 1)).toBeNull();
    expect(pageBorderStyle({ top: line, appliesTo: "exceptFirstPage" }, margins, 0)).toBeNull();
    expect(pageBorderStyle({ top: line, appliesTo: "exceptFirstPage" }, margins, 1)).not.toBeNull();
  });

  it("is inert decoration, never a hit target", () => {
    expect(pageBorderStyle({ top: line }, margins, 0)?.pointerEvents).toBe("none");
  });
});

describe("paragraphStyle borders (SPEC-074 AC-13)", () => {
  it("paints a boxed paragraph's edges, padding and shading", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: {
        borders: { top: { style: "single", widthPt: 1, color: "auto" } },
        shadingFill: "#fff2cc",
        indentStartCm: 1,
      },
    });
    expect(style.borderTop).toBe("1pt solid #000000");
    expect(style.backgroundColor).toBe("#fff2cc");
    expect(style.paddingTop).toBe("4pt");
    // The shading follows the indent because the indent is a margin.
    expect(style.marginInlineStart).toBe("1cm");
  });

  it("leaves an unbordered paragraph's style untouched", () => {
    const style = paragraphStyle({ type: "paragraph", runs: [], formatting: { alignment: "center" } });
    expect(style.borderTop).toBeUndefined();
    expect(style.paddingTop).toBeUndefined();
  });
});
