// SPEC-117 "Rendering", "Border CSS and weights" and "Overflow": the
// framework-free half of the cell layer. CSS per style (every alignment,
// General per value kind, wrap, rotation, indent, scripts, decorations,
// colours through a theme), the pinned fill-pattern table row by row and both
// gradient kinds, the pinned border table row by row, the shared-edge rule
// (heavier wins, a tie goes to the right/bottom cell, absent sides, the
// window boundary, merged interiors, hidden rows and columns), diagonals,
// overflow extents and exclusions, the text-width estimate and the zero
// display.
import { describe, expect, it } from "vitest";
import { formatA1, parseRange, type GridRange } from "./address";
import {
  AUTOMATIC_COLOR,
  BORDER_STYLES,
  borderCss,
  borderRank,
  borderShorthand,
  cellBoxCss,
  cellCss,
  cellDisplay,
  cellText,
  diagonalLines,
  displayText,
  edgeElementCss,
  estimateTextWidth,
  fillCss,
  fillRepeat,
  OVERFLOW_SCAN_COLS,
  offscreenSpillCandidate,
  overflowAlignment,
  overflowExtent,
  resolvedHorizontal,
  resolveEdges,
  runCss,
  shrinkScale,
  valueKind,
  type EdgeSegment,
  type ValueKind,
} from "./cellRender";
import { buildGridMetrics, CELL_PADDING_PX, INDENT_PX_PER_LEVEL, type GridMetrics } from "./geometry";
import { buildMergeIndex } from "./merges";
import type {
  Cell,
  CellAlignment,
  CellBorderEdge,
  CellFill,
  CellStyle,
  CellValue,
  ColumnSpan,
  Row,
  Sheet,
  SheetBorderStyle,
} from "./model";
import { DEFAULT_THEME, hexToRgb, resolveColor, type DocumentTheme, type ModelColor } from "../core/theme";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import type { FormatContext } from "./numberFormat/render";

function sheetWith(partial: Partial<Sheet> = {}): Sheet {
  return {
    id: "sheet-1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    ...partial,
  };
}

/** A custom theme: no value from it can pass by coincidence with Office's. */
const CUSTOM_THEME: DocumentTheme = {
  name: "Custom",
  colors: { ...DEFAULT_THEME.colors, dk1: "#202020", accent1: "#123456", accent2: "#a1b2c3" },
  fonts: { major: "Georgia", minor: "Verdana" },
};

const range = (a1: string): GridRange => parseRange(a1)!;

function aligned(alignment: CellAlignment): CellStyle {
  return { alignment };
}

/** The per-channel distance between two #rrggbb colours. */
function channelDistance(a: string, b: string): number {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  return Math.max(...x.map((c, i) => Math.abs(Math.round(c * 255) - Math.round(y[i]! * 255))));
}

/** A font's text-decoration as cellCss emits it. */
function decoration(font: CellStyle["font"]): string | undefined {
  return cellCss({ font }, DEFAULT_THEME).text["text-decoration"];
}

/** One stripe layer of the pinned pattern table, in the default colours. */
function stripe(angle: number, line: number, gap = "#ffffff"): string {
  return `repeating-linear-gradient(${angle}deg, #000000 0 ${line}px, ${gap} ${line}px 4px)`;
}

/** A white → accent1 linear gradient at `degree`. */
function linear(degree: number | undefined): CellFill {
  return {
    kind: "gradient",
    type: "linear",
    ...(degree === undefined ? {} : { degree }),
    stops: [
      { position: 0, color: "#ffffff" },
      { position: 1, color: { theme: "accent1" } },
    ],
  };
}

// --- values and alignment ----------------------------------------------------

describe("SPEC-117 Rendering: value kinds and General alignment", () => {
  it("classifies values; only null is empty", () => {
    const cases: Array<[CellValue, ValueKind]> = [
      [0, "number"],
      [-1.5, "number"],
      ["", "string"],
      ["text", "string"],
      [true, "boolean"],
      [false, "boolean"],
      [{ error: "#N/A" }, "error"],
      [null, "empty"],
    ];
    for (const [value, kind] of cases) {
      expect(valueKind(value)).toBe(kind);
    }
  });

  it("General: numbers right, strings left, booleans and errors centred", () => {
    for (const style of [{}, aligned({ horizontal: "general" })]) {
      expect(resolvedHorizontal(style, "number")).toBe("right");
      expect(resolvedHorizontal(style, "string")).toBe("left");
      expect(resolvedHorizontal(style, "empty")).toBe("left");
      expect(resolvedHorizontal(style, "boolean")).toBe("center");
      expect(resolvedHorizontal(style, "error")).toBe("center");
    }
  });

  it("an explicit alignment wins over the value kind", () => {
    const all = ["left", "center", "right", "fill", "justify", "centerContinuous", "distributed"] as const;
    for (const horizontal of all) {
      for (const kind of ["number", "string", "boolean", "error"] as const) {
        expect(resolvedHorizontal(aligned({ horizontal }), kind)).toBe(horizontal);
      }
    }
  });
});

describe("SPEC-117 Rendering: cellCss alignment", () => {
  it("maps every horizontal alignment to text-align on the box", () => {
    const cases: Array<[CellAlignment["horizontal"], string]> = [
      ["left", "left"],
      ["center", "center"],
      ["right", "right"],
      ["fill", "left"],
      ["justify", "justify"],
      ["centerContinuous", "center"],
      ["distributed", "justify"],
    ];
    for (const [horizontal, textAlign] of cases) {
      expect(cellCss(aligned({ horizontal }), DEFAULT_THEME, "number").box["text-align"]).toBe(textAlign);
    }
    // Distributed spreads every line, the last (or only) one included.
    expect(cellCss(aligned({ horizontal: "distributed" }), DEFAULT_THEME).box["text-align-last"]).toBe("justify");
    expect(cellCss(aligned({ horizontal: "justify" }), DEFAULT_THEME).box["text-align-last"]).toBeUndefined();
  });

  it("aligns General by the value kind", () => {
    expect(cellCss({}, DEFAULT_THEME, "number").box["text-align"]).toBe("right");
    expect(cellCss({}, DEFAULT_THEME, "string").box["text-align"]).toBe("left");
    expect(cellCss({}, DEFAULT_THEME, "boolean").box["text-align"]).toBe("center");
    expect(cellCss({}, DEFAULT_THEME, "error").box["text-align"]).toBe("center");
    expect(cellCss(aligned({ horizontal: "general" }), DEFAULT_THEME, "number").box["text-align"]).toBe("right");
    expect(cellCss({}, DEFAULT_THEME).box["text-align"]).toBe("left"); // kind defaults to string
  });

  it("maps vertical alignment to the column flexbox's justify-content (default bottom)", () => {
    expect(cellCss({}, DEFAULT_THEME).box["justify-content"]).toBe("flex-end");
    const cases: Array<[CellAlignment["vertical"], string]> = [
      ["top", "flex-start"],
      ["center", "center"],
      ["bottom", "flex-end"],
      ["justify", "flex-start"],
      ["distributed", "flex-start"],
    ];
    for (const [vertical, justify] of cases) {
      expect(cellCss(aligned({ vertical }), DEFAULT_THEME).box["justify-content"]).toBe(justify);
    }
  });

  it("wraps (pre-wrap, clipped — rows never grow) for wrap, justify and distributed; pre otherwise", () => {
    expect(cellCss({}, DEFAULT_THEME).text["white-space"]).toBe("pre");
    for (const alignment of [
      { wrap: true },
      { horizontal: "justify" },
      { horizontal: "distributed" },
      { vertical: "justify" },
      { vertical: "distributed" },
    ] as CellAlignment[]) {
      expect(cellCss(aligned(alignment), DEFAULT_THEME).text["white-space"]).toBe("pre-wrap");
    }
    for (const alignment of [
      { horizontal: "fill" },
      { horizontal: "centerContinuous" },
      { horizontal: "right" },
      { shrinkToFit: true },
      { vertical: "top" },
    ] as CellAlignment[]) {
      expect(cellCss(aligned(alignment), DEFAULT_THEME).text["white-space"]).toBe("pre");
    }
  });

  it("rotates 1–90 by −θ, 91–180 by θ − 90, and stacks 255 vertically", () => {
    const rotation = (textRotation: number): Record<string, string> => cellCss(aligned({ textRotation }), DEFAULT_THEME).text;
    expect(rotation(0).transform).toBeUndefined();
    expect(cellCss({}, DEFAULT_THEME).text.transform).toBeUndefined();
    expect(rotation(1).transform).toBe("rotate(-1deg)");
    expect(rotation(45).transform).toBe("rotate(-45deg)");
    expect(rotation(90).transform).toBe("rotate(-90deg)");
    expect(rotation(91).transform).toBe("rotate(1deg)");
    expect(rotation(135).transform).toBe("rotate(45deg)");
    expect(rotation(180).transform).toBe("rotate(90deg)");
    expect(rotation(45)["transform-origin"]).toBe("center");
    const stacked = rotation(255);
    expect(stacked["writing-mode"]).toBe("vertical-lr");
    expect(stacked["text-orientation"]).toBe("upright");
    expect(stacked.transform).toBeUndefined();
    expect(rotation(0)["writing-mode"]).toBeUndefined();
  });

  it("pads 0 CELL_PADDING_PX, plus indent × INDENT_PX_PER_LEVEL on the aligned side", () => {
    const pad = (alignment: CellAlignment, kind: ValueKind = "string"): string | undefined =>
      cellCss(aligned(alignment), DEFAULT_THEME, kind).text.padding;
    expect(cellCss({}, DEFAULT_THEME).text.padding).toBe(`0 ${CELL_PADDING_PX}px`);
    expect(pad({ horizontal: "left", indent: 2 })).toBe(`0 2px 0 ${2 + 2 * INDENT_PX_PER_LEVEL}px`);
    expect(pad({ indent: 1 })).toBe("0 2px 0 11px"); // General → left for a string
    expect(pad({ horizontal: "fill", indent: 1 })).toBe("0 2px 0 11px");
    expect(pad({ horizontal: "right", indent: 3 })).toBe("0 29px 0 2px");
    expect(pad({ indent: 1 }, "number")).toBe("0 11px 0 2px"); // General → right for a number
    expect(pad({ horizontal: "center", indent: 2 })).toBe("0 2px");
    expect(pad({ horizontal: "distributed", indent: 1 })).toBe("0 11px"); // both sides
    expect(pad({ horizontal: "left", indent: 0 })).toBe("0 2px");
  });

  it("leaves shrink-to-fit to the component (it needs the text): the size is untouched", () => {
    const css = cellCss({ font: { sizePt: 12 }, alignment: { shrinkToFit: true } }, DEFAULT_THEME).text;
    expect(css["font-size"]).toBe("12pt");
  });
});

// --- fonts -------------------------------------------------------------------

describe("SPEC-117 Rendering: cellCss fonts and colours", () => {
  it("emits the resolved family name, unquoted, and nothing when the style names none", () => {
    expect(cellCss({ font: { family: "Arial" } }, DEFAULT_THEME).text["font-family"]).toBe("Arial");
    expect(cellCss({ font: { family: "Times New Roman" } }, DEFAULT_THEME).text["font-family"]).toBe("Times New Roman");
    expect(cellCss({ font: { family: { theme: "minor" } } }, DEFAULT_THEME).text["font-family"]).toBe("Calibri");
    expect(cellCss({ font: { family: { theme: "major" } } }, DEFAULT_THEME).text["font-family"]).toBe("Calibri Light");
    expect(cellCss({ font: { family: { theme: "minor" } } }, CUSTOM_THEME).text["font-family"]).toBe("Verdana");
    expect(cellCss({ font: { family: { theme: "major" } } }, CUSTOM_THEME).text["font-family"]).toBe("Georgia");
    expect(cellCss({}, DEFAULT_THEME).text).not.toHaveProperty("font-family");
    expect(cellCss({ font: { bold: true } }, DEFAULT_THEME).text).not.toHaveProperty("font-family");
  });

  it("sizes in pt, weight, style", () => {
    const text = cellCss({ font: { sizePt: 14, bold: true, italic: true } }, DEFAULT_THEME).text;
    expect(text["font-size"]).toBe("14pt");
    expect(text["font-weight"]).toBe("bold");
    expect(text["font-style"]).toBe("italic");
    expect(cellCss({ font: { sizePt: 10.5 } }, DEFAULT_THEME).text["font-size"]).toBe("10.5pt");
    const plain = cellCss({}, DEFAULT_THEME).text;
    expect(plain).not.toHaveProperty("font-size");
    expect(plain).not.toHaveProperty("font-weight");
    expect(plain).not.toHaveProperty("font-style");
  });

  it("decorates underline variants and strike", () => {
    expect(decoration({ underline: "single" })).toBe("underline");
    expect(decoration({ underline: "singleAccounting" })).toBe("underline");
    expect(decoration({ underline: "double" })).toBe("underline double");
    expect(decoration({ underline: "doubleAccounting" })).toBe("underline double");
    expect(decoration({ strike: true })).toBe("line-through");
    expect(decoration({ underline: "single", strike: true })).toBe("underline line-through");
    expect(decoration({})).toBeUndefined();
  });

  it("draws superscript and subscript at 0.7 of the base size (11 pt when the style has none)", () => {
    const sup = cellCss({ font: { vertAlign: "superscript", sizePt: 20 } }, DEFAULT_THEME).text;
    expect(sup["vertical-align"]).toBe("super");
    expect(sup["font-size"]).toBe("14pt");
    const sub = cellCss({ font: { vertAlign: "subscript" } }, DEFAULT_THEME).text;
    expect(sub["vertical-align"]).toBe("sub");
    expect(sub["font-size"]).toBe("7.7pt");
    expect(cellCss({}, DEFAULT_THEME).text).not.toHaveProperty("vertical-align");
  });

  it("colours text through the theme; automatic is black", () => {
    expect(cellCss({}, DEFAULT_THEME).text.color).toBe(AUTOMATIC_COLOR);
    expect(AUTOMATIC_COLOR).toBe("#000000");
    expect(cellCss({ font: { color: "#ff0000" } }, DEFAULT_THEME).text.color).toBe("#ff0000");
    expect(cellCss({ font: { color: { theme: "accent1" } } }, DEFAULT_THEME).text.color).toBe("#4472c4");
    expect(cellCss({ font: { color: { theme: "accent1" } } }, CUSTOM_THEME).text.color).toBe("#123456");
    expect(cellCss({ font: { color: { theme: "dk1" } } }, CUSTOM_THEME).text.color).toBe("#202020");
  });

  it("resolves a lumMod/lumOff reference: Accent 1, Lighter 40% is Excel's #8ea9db ±1 against DEFAULT_THEME", () => {
    const lighter40: ModelColor = { theme: "accent1", lumMod: 0.6, lumOff: 0.4 };
    const color = cellCss({ font: { color: lighter40 } }, DEFAULT_THEME).text.color!;
    expect(color).toBe(resolveColor(lighter40, DEFAULT_THEME));
    expect(channelDistance(color, "#8ea9db")).toBeLessThanOrEqual(1);
    const darker25: ModelColor = { theme: "accent2", lumMod: 0.75 };
    expect(cellCss({ font: { color: darker25 } }, CUSTOM_THEME).text.color).toBe(resolveColor(darker25, CUSTOM_THEME));
    expect(resolveColor(darker25, CUSTOM_THEME)).not.toBe(resolveColor(darker25, DEFAULT_THEME));
  });
});

describe("SPEC-117 Rendering: rich-text runs override the cell font per property present", () => {
  it("emits only the properties the run's font carries", () => {
    expect(runCss({}, DEFAULT_THEME)).toEqual({});
    expect(runCss({ bold: true }, DEFAULT_THEME)).toEqual({ "font-weight": "bold" });
    expect(runCss({ italic: true, sizePt: 9 }, DEFAULT_THEME)).toEqual({ "font-style": "italic", "font-size": "9pt" });
    expect(runCss({ color: { theme: "accent1" } }, CUSTOM_THEME)).toEqual({ color: "#123456" });
    expect(runCss({ family: { theme: "major" } }, DEFAULT_THEME)).toEqual({ "font-family": "Calibri Light" });
    expect(runCss({ underline: "double", strike: true }, DEFAULT_THEME)).toEqual({
      "text-decoration": "underline line-through double",
    });
  });

  it("scales a script run by its own size, else relative to the cell's (0.7em)", () => {
    expect(runCss({ vertAlign: "superscript", sizePt: 10 }, DEFAULT_THEME)).toEqual({
      "vertical-align": "super",
      "font-size": "7pt",
    });
    expect(runCss({ vertAlign: "subscript" }, DEFAULT_THEME)).toEqual({ "vertical-align": "sub", "font-size": "0.7em" });
  });

  it("renders a full xlsx run font completely", () => {
    expect(
      runCss({ family: "Arial", sizePt: 12, bold: true, italic: true, underline: "single", color: "#00ff00" }, DEFAULT_THEME),
    ).toEqual({
      "font-family": "Arial",
      "font-size": "12pt",
      "font-weight": "bold",
      "font-style": "italic",
      "text-decoration": "underline",
      color: "#00ff00",
    });
  });
});

// --- fills -------------------------------------------------------------------

/** fg over bg at p, per channel in 0..255. */
function mixed(fg: string, bg: string, p: number): string {
  const f = hexToRgb(fg).map((c) => c * 255);
  const b = hexToRgb(bg).map((c) => c * 255);
  return `#${f.map((c, i) => Math.round(c * p + b[i]! * (1 - p)).toString(16).padStart(2, "0")).join("")}`;
}

describe("SPEC-117 Rendering: fills", () => {
  it("draws a solid fill as the background colour, through the theme", () => {
    expect(fillCss({ kind: "solid", color: "#ffff00" }, DEFAULT_THEME)).toEqual({ "background-color": "#ffff00" });
    expect(fillCss({ kind: "solid", color: { theme: "accent2" } }, DEFAULT_THEME)).toEqual({ "background-color": "#ed7d31" });
    expect(fillCss({ kind: "solid", color: { theme: "accent2" } }, CUSTOM_THEME)).toEqual({ "background-color": "#a1b2c3" });
    const lighter: ModelColor = { theme: "accent1", lumMod: 0.2, lumOff: 0.8 };
    expect(fillCss({ kind: "solid", color: lighter }, DEFAULT_THEME)).toEqual({
      "background-color": resolveColor(lighter, DEFAULT_THEME),
    });
  });

  it("leaves an unfilled cell transparent (the gridlines show)", () => {
    expect(fillCss(undefined, DEFAULT_THEME)).toEqual({});
    const box = cellCss({}, DEFAULT_THEME).box;
    expect(box).not.toHaveProperty("background-color");
    expect(box).not.toHaveProperty("background-image");
  });

  it("puts the fill on the box, not the text", () => {
    const css = cellCss({ fill: { kind: "solid", color: "#ffff00" }, font: { color: "#ff0000", bold: true } }, DEFAULT_THEME);
    expect(css.box["background-color"]).toBe("#ffff00");
    expect(css.text).not.toHaveProperty("background-color");
    expect(css.text.color).toBe("#ff0000");
  });

  // The pinned pattern table, row by row, with the default colours.
  const PATTERNS: Array<[string, Record<string, string>]> = [
    ["gray0625", { "background-color": "#efefef" }],
    ["gray125", { "background-color": "#dfdfdf" }],
    ["lightGray", { "background-color": "#bfbfbf" }],
    ["mediumGray", { "background-color": "#808080" }],
    ["darkGray", { "background-color": "#404040" }],
    ["lightHorizontal", { "background-color": "#ffffff", "background-image": stripe(0, 1) }],
    ["darkHorizontal", { "background-color": "#ffffff", "background-image": stripe(0, 2) }],
    ["lightVertical", { "background-color": "#ffffff", "background-image": stripe(90, 1) }],
    ["darkVertical", { "background-color": "#ffffff", "background-image": stripe(90, 2) }],
    ["lightDown", { "background-color": "#ffffff", "background-image": stripe(45, 1) }],
    ["darkDown", { "background-color": "#ffffff", "background-image": stripe(45, 2) }],
    ["lightUp", { "background-color": "#ffffff", "background-image": stripe(135, 1) }],
    ["darkUp", { "background-color": "#ffffff", "background-image": stripe(135, 2) }],
    [
      "lightGrid",
      { "background-color": "#ffffff", "background-image": `${stripe(0, 1, "transparent")}, ${stripe(90, 1)}` },
    ],
    [
      "darkGrid",
      { "background-color": "#ffffff", "background-image": `${stripe(0, 2, "transparent")}, ${stripe(90, 2)}` },
    ],
    [
      "lightTrellis",
      { "background-color": "#ffffff", "background-image": `${stripe(45, 1, "transparent")}, ${stripe(135, 1)}` },
    ],
    [
      "darkTrellis",
      { "background-color": "#ffffff", "background-image": `${stripe(45, 2, "transparent")}, ${stripe(135, 2)}` },
    ],
  ];

  it.each(PATTERNS)("pattern %s (fg #000000 over bg #ffffff)", (pattern, css) => {
    expect(fillCss({ kind: "pattern", pattern }, DEFAULT_THEME)).toEqual(css);
  });

  it("mixes the grays as fg over bg at 6.25 / 12.5 / 25 / 50 / 75 %", () => {
    const fractions: Array<[string, number]> = [
      ["gray0625", 0.0625],
      ["gray125", 0.125],
      ["lightGray", 0.25],
      ["mediumGray", 0.5],
      ["darkGray", 0.75],
    ];
    for (const [pattern, p] of fractions) {
      expect(fillCss({ kind: "pattern", pattern }, DEFAULT_THEME)["background-color"]).toBe(mixed("#000000", "#ffffff", p));
      expect(fillCss({ kind: "pattern", pattern, fg: "#ff0000", bg: "#0000ff" }, DEFAULT_THEME)["background-color"]).toBe(
        mixed("#ff0000", "#0000ff", p),
      );
      expect(
        fillCss({ kind: "pattern", pattern, fg: { theme: "accent1" }, bg: "#ffff00" }, CUSTOM_THEME)["background-color"],
      ).toBe(mixed("#123456", "#ffff00", p));
    }
  });

  it("draws the line patterns in the file's colours, resolved through the theme", () => {
    expect(fillCss({ kind: "pattern", pattern: "darkTrellis", fg: { theme: "accent2" }, bg: "#ffff00" }, DEFAULT_THEME)).toEqual({
      "background-color": "#ffff00",
      "background-image":
        "repeating-linear-gradient(45deg, #ed7d31 0 2px, transparent 2px 4px), " +
        "repeating-linear-gradient(135deg, #ed7d31 0 2px, #ffff00 2px 4px)",
    });
  });

  it("treats a solid pattern as its fg, and paints nothing for a pattern it does not know", () => {
    expect(fillCss({ kind: "pattern", pattern: "solid", fg: "#00ff00" }, DEFAULT_THEME)).toEqual({ "background-color": "#00ff00" });
    expect(fillCss({ kind: "pattern", pattern: "none" }, DEFAULT_THEME)).toEqual({});
    expect(fillCss({ kind: "pattern", pattern: "sparkles" }, DEFAULT_THEME)).toEqual({});
  });

  it("draws a linear gradient at (degree + 90)deg with its stops", () => {
    expect(fillCss(linear(90), DEFAULT_THEME)).toEqual({
      "background-image": "linear-gradient(180deg, #ffffff 0%, #4472c4 100%)",
    });
    expect(fillCss(linear(undefined), DEFAULT_THEME)["background-image"]).toBe(
      "linear-gradient(90deg, #ffffff 0%, #4472c4 100%)",
    );
    expect(fillCss(linear(45), CUSTOM_THEME)["background-image"]).toBe("linear-gradient(135deg, #ffffff 0%, #123456 100%)");
    expect(
      fillCss(
        {
          kind: "gradient",
          type: "linear",
          degree: 0,
          stops: [
            { position: 0, color: "#ff0000" },
            { position: 0.25, color: "#00ff00" },
            { position: 1, color: "#0000ff" },
          ],
        },
        DEFAULT_THEME,
      )["background-image"],
    ).toBe("linear-gradient(90deg, #ff0000 0%, #00ff00 25%, #0000ff 100%)");
  });

  it("draws a path gradient as a radial gradient at the focus ((left+right)/2, (top+bottom)/2)", () => {
    const stops = [
      { position: 0, color: "#ffffff" as ModelColor },
      { position: 1, color: "#4472c4" as ModelColor },
    ];
    expect(fillCss({ kind: "gradient", type: "path", left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, stops }, DEFAULT_THEME)).toEqual({
      "background-image": "radial-gradient(at 50% 50%, #ffffff 0%, #4472c4 100%)",
    });
    expect(fillCss({ kind: "gradient", type: "path", left: 0, right: 1, top: 0, bottom: 0, stops }, DEFAULT_THEME)["background-image"]).toBe(
      "radial-gradient(at 50% 0%, #ffffff 0%, #4472c4 100%)",
    );
    expect(fillCss({ kind: "gradient", type: "path", stops }, DEFAULT_THEME)["background-image"]).toBe(
      "radial-gradient(at 0% 0%, #ffffff 0%, #4472c4 100%)",
    );
    expect(fillCss({ kind: "gradient", type: "path", left: 1, right: 1, top: 1, bottom: 1, stops }, DEFAULT_THEME)["background-image"]).toBe(
      "radial-gradient(at 100% 100%, #ffffff 0%, #4472c4 100%)",
    );
  });

  it("degrades a gradient with fewer than two stops", () => {
    expect(fillCss({ kind: "gradient", type: "linear", stops: [{ position: 0, color: "#abcdef" }] }, DEFAULT_THEME)).toEqual({
      "background-color": "#abcdef",
    });
    expect(fillCss({ kind: "gradient", type: "linear", stops: [] }, DEFAULT_THEME)).toEqual({});
  });
});

// --- borders -----------------------------------------------------------------

describe("SPEC-117 Border CSS and weights: the pinned table", () => {
  const TABLE: Array<[number, SheetBorderStyle, number, "solid" | "double" | "dashed" | "dotted", number | undefined]> = [
    [1, "thick", 3, "solid", undefined],
    [2, "double", 3, "double", undefined],
    [3, "medium", 2, "solid", undefined],
    [4, "mediumDashed", 2, "dashed", undefined],
    [5, "mediumDashDot", 2, "dashed", undefined],
    [6, "slantDashDot", 2, "dashed", undefined],
    [7, "mediumDashDotDot", 2, "dotted", undefined],
    [8, "thin", 1, "solid", undefined],
    [9, "dashed", 1, "dashed", undefined],
    [10, "dashDot", 1, "dashed", undefined],
    [11, "dashDotDot", 1, "dotted", undefined],
    [12, "dotted", 1, "dotted", undefined],
    [13, "hair", 1, "dotted", 0.6],
  ];

  it.each(TABLE)("rank %d: %s is %dpx %s", (rank, style, width, lineStyle, opacity) => {
    const expected = opacity === undefined ? { rank, width, lineStyle } : { rank, width, lineStyle, opacity };
    expect(BORDER_STYLES[style]).toEqual(expected);
    expect(borderCss(style)).toEqual(expected);
    expect(borderRank(style)).toBe(rank);
  });

  it("covers all thirteen styles, heaviest first", () => {
    expect(Object.keys(BORDER_STYLES)).toEqual(TABLE.map(([, style]) => style));
  });

  it("writes a border shorthand", () => {
    expect(borderShorthand("double", "#ff0000")).toBe("3px double #ff0000");
    expect(borderShorthand("hair", "#000000")).toBe("1px dotted #000000");
    expect(borderShorthand("mediumDashDotDot", "#123456")).toBe("2px dotted #123456");
  });
});

const edge = (style: SheetBorderStyle, color?: ModelColor): CellBorderEdge => (color === undefined ? { style } : { style, color });

/** A style with the same edge on all four sides. */
function boxed(style: SheetBorderStyle, color?: ModelColor): CellStyle {
  return { borders: { top: edge(style, color), right: edge(style, color), bottom: edge(style, color), left: edge(style, color) } };
}

interface EdgeFixture {
  styles: Record<string, CellStyle>;
  merges?: GridRange[];
  columns?: ColumnSpan[];
  rows?: Row[];
  theme?: DocumentTheme;
}

/** resolveEdges over `view` for styles keyed by A1 address, described as
 * "H B3 thin #000000" (the top edge of B3) / "V C2 …" (the left edge of
 * C2), sorted. */
function edgesIn(view: string, fixture: EdgeFixture): { names: string[]; segments: EdgeSegment[]; metrics: GridMetrics } {
  const sheet = sheetWith({ columns: fixture.columns ?? [], rows: fixture.rows ?? [], merges: fixture.merges ?? [] });
  const metrics = buildGridMetrics(sheet);
  const calls = new Map<string, number>();
  const styleAt = (row: number, col: number): CellStyle => {
    const a1 = formatA1({ row, col });
    calls.set(a1, (calls.get(a1) ?? 0) + 1);
    return fixture.styles[a1] ?? {};
  };
  const segments = resolveEdges(range(view), sheet, styleAt, metrics, buildMergeIndex(sheet.merges), fixture.theme ?? DEFAULT_THEME);
  for (const [a1, count] of calls) {
    expect(count, `styleAt(${a1}) called ${count} times`).toBe(1);
  }
  const names = segments
    .map((s) => {
      const cell = formatA1({ row: metrics.rowAt(s.y), col: metrics.colAt(s.x) });
      return `${s.orientation === "horizontal" ? "H" : "V"} ${cell} ${s.style} ${s.color}`;
    })
    .toSorted();
  return { names, segments, metrics };
}

describe("SPEC-117 Border CSS and weights: resolveEdges and the shared-edge rule", () => {
  it("draws a boxed cell's four edges once each", () => {
    const { names } = edgesIn("A1:C3", { styles: { B2: boxed("thin") } });
    expect(names).toEqual(["H B2 thin #000000", "H B3 thin #000000", "V B2 thin #000000", "V C2 thin #000000"]);
  });

  it("places segments in unscaled sheet px from GridMetrics", () => {
    const { segments } = edgesIn("A1:C3", {
      styles: { B2: boxed("thin") },
      columns: [{ min: 1, max: 1, widthPx: 100 }],
      rows: [{ index: 1, cells: [], heightPt: 30 }],
    });
    const sorted = segments.toSorted((a, b) => a.orientation.localeCompare(b.orientation) || a.y - b.y || a.x - b.x);
    expect(sorted).toEqual([
      { x: 64, y: 20, length: 100, orientation: "horizontal", style: "thin", color: "#000000" },
      { x: 64, y: 60, length: 100, orientation: "horizontal", style: "thin", color: "#000000" },
      { x: 64, y: 20, length: 40, orientation: "vertical", style: "thin", color: "#000000" },
      { x: 164, y: 20, length: 40, orientation: "vertical", style: "thin", color: "#000000" },
    ]);
  });

  it("the heavier style wins a shared edge, with its own colour", () => {
    const below = edgesIn("A1:C4", {
      styles: { B2: { borders: { bottom: edge("thin", "#ff0000") } }, B3: { borders: { top: edge("medium", "#0000ff") } } },
    });
    expect(below.names).toEqual(["H B3 medium #0000ff"]);
    const above = edgesIn("A1:C4", {
      styles: { B2: { borders: { bottom: edge("thick", "#ff0000") } }, B3: { borders: { top: edge("hair", "#0000ff") } } },
    });
    expect(above.names).toEqual(["H B3 thick #ff0000"]);
    const right = edgesIn("A1:D3", {
      styles: { B2: { borders: { right: edge("double", "#ff0000") } }, C2: { borders: { left: edge("thick", "#00ff00") } } },
    });
    expect(right.names).toEqual(["V C2 thick #00ff00"]);
    const left = edgesIn("A1:D3", {
      styles: { B2: { borders: { right: edge("mediumDashDot") } }, C2: { borders: { left: edge("slantDashDot", "#00ff00") } } },
    });
    expect(left.names).toEqual(["V C2 mediumDashDot #000000"]);
  });

  it("a tie (the same style) goes to the right or bottom cell, its colour included", () => {
    const vertical = edgesIn("A1:D3", {
      styles: { B2: { borders: { right: edge("thin", "#ff0000") } }, C2: { borders: { left: edge("thin", "#0000ff") } } },
    });
    expect(vertical.names).toEqual(["V C2 thin #0000ff"]);
    const horizontal = edgesIn("A1:C4", {
      styles: { B2: { borders: { bottom: edge("dashed", "#ff0000") } }, B3: { borders: { top: edge("dashed") } } },
    });
    expect(horizontal.names).toEqual(["H B3 dashed #000000"]);
  });

  it("one side absent → the other side's edge", () => {
    expect(edgesIn("A1:D3", { styles: { B2: { borders: { right: edge("medium", "#ff0000") } } } }).names).toEqual([
      "V C2 medium #ff0000",
    ]);
    expect(edgesIn("A1:D3", { styles: { C2: { borders: { left: edge("dotted") } } } }).names).toEqual(["V C2 dotted #000000"]);
    expect(edgesIn("A1:C4", { styles: { B3: { borders: { top: edge("hair") } } } }).names).toEqual(["H B3 hair #000000"]);
    expect(edgesIn("A1:C3", { styles: {} }).names).toEqual([]);
  });

  it("resolves the edges on the window's boundary against the neighbour just outside it", () => {
    const { names } = edgesIn("B2:B2", {
      styles: {
        A2: { borders: { right: edge("thick") } },
        B1: { borders: { bottom: edge("medium", "#ff0000") } },
        B3: { borders: { top: edge("dashed") } },
        C2: { borders: { left: edge("dotted", "#00ff00") } },
        D2: boxed("thick"), // not adjacent to the window: never drawn
      },
    });
    expect(names).toEqual(["H B2 medium #ff0000", "H B3 dashed #000000", "V B2 thick #000000", "V C2 dotted #00ff00"]);
    // The outside neighbour still competes by rank with the window's cell.
    const competing = edgesIn("B2:B2", {
      styles: { A2: { borders: { right: edge("thin", "#ff0000") } }, B2: { borders: { left: edge("thin", "#0000ff") } } },
    });
    expect(competing.names).toEqual(["V B2 thin #0000ff"]);
  });

  it("draws a grid-edge border with no neighbour", () => {
    expect(edgesIn("A1:A1", { styles: { A1: boxed("medium") } }).names).toEqual([
      "H A1 medium #000000",
      "H A2 medium #000000",
      "V A1 medium #000000",
      "V B1 medium #000000",
    ]);
  });

  it("colours edges through the theme; automatic is black", () => {
    const { names } = edgesIn("A1:C3", {
      styles: { B2: { borders: { top: edge("thin", { theme: "accent1" }), left: edge("thin", { theme: "dk1" }) } } },
      theme: CUSTOM_THEME,
    });
    expect(names).toEqual(["H B2 thin #123456", "V B2 thin #202020"]);
  });

  it("suppresses a merge's interior edges; its outer edges come from the members", () => {
    const styles: Record<string, CellStyle> = { B2: boxed("thin"), C2: boxed("thin"), B3: boxed("thin"), C3: boxed("thin") };
    const { names } = edgesIn("A1:D4", { styles, merges: [range("B2:C3")] });
    expect(names).toEqual([
      "H B2 thin #000000",
      "H B4 thin #000000",
      "H C2 thin #000000",
      "H C4 thin #000000",
      "V B2 thin #000000",
      "V B3 thin #000000",
      "V D2 thin #000000",
      "V D3 thin #000000",
    ]);
  });

  it("takes each outer edge segment from the member cell on that edge", () => {
    // Only C2 of the merged B2:D2 carries a bottom border: only its segment draws.
    const { names } = edgesIn("A1:E3", {
      styles: { C2: { borders: { bottom: edge("medium"), top: edge("thin") } } },
      merges: [range("B2:D2")],
    });
    expect(names).toEqual(["H C2 thin #000000", "H C3 medium #000000"]);
    // A window cutting through a merge: still no interior edge.
    const cut = edgesIn("C1:C3", { styles: { B2: boxed("thin"), C2: boxed("thin"), D2: boxed("thin") }, merges: [range("B2:D2")] });
    expect(cut.names).toEqual(["H C2 thin #000000", "H C3 thin #000000"]);
  });

  it("uses the member on the range's edge even when the rows between are hidden", () => {
    // B2:B4 merged, row 4 hidden: the merge's bottom edge is B4's.
    const { names } = edgesIn("A1:C6", {
      styles: { B4: { borders: { bottom: edge("thick") } }, B3: { borders: { bottom: edge("thin") } } },
      merges: [range("B2:B4")],
      rows: [{ index: 3, cells: [], hidden: true }],
    });
    expect(names).toEqual(["H B5 thick #000000"]);
  });

  it("hidden columns produce no segment; the visible cells either side face each other", () => {
    const { names, segments, metrics } = edgesIn("A1:E3", {
      styles: {
        B2: { borders: { right: edge("thin") } },
        C2: boxed("thick"), // hidden: none of its edges draws
        D2: { borders: { left: edge("medium") } },
      },
      columns: [{ min: 2, max: 2, hidden: true }],
    });
    expect(names).toEqual(["V D2 medium #000000"]);
    expect(segments[0]!.x).toBe(metrics.colLeft(2));
    expect(segments.every((s) => s.length > 0)).toBe(true);
  });

  it("hidden rows produce no segment; the visible rows either side face each other", () => {
    const { names, segments, metrics } = edgesIn("A1:C6", {
      styles: {
        B2: { borders: { bottom: edge("thin", "#ff0000") } },
        B3: boxed("thick"), // hidden row
        B4: { borders: { top: edge("thin", "#0000ff") } },
      },
      rows: [{ index: 2, cells: [], hidden: true }],
    });
    expect(names).toEqual(["H B4 thin #0000ff"]);
    expect(segments[0]!.y).toBe(metrics.rowTop(2));
    // A window of hidden rows only draws nothing.
    expect(
      edgesIn("A3:C3", { styles: { B3: boxed("thick") }, rows: [{ index: 2, cells: [], hidden: true }] }).names,
    ).toEqual([]);
  });

  it("skips a long hidden run inside the window in one step", () => {
    const rows: Row[] = [{ index: 1, cells: [], hidden: true }];
    const sheet = sheetWith({ rows, columns: [{ min: 1, max: 16_000, hidden: true }] });
    const inner = buildGridMetrics(sheet);
    let queries = 0;
    const counted = Object.fromEntries(
      Object.entries(inner).map(([name, query]: [string, (n: number) => number]) => [
        name,
        (n: number): number => {
          queries += 1;
          return query(n);
        },
      ]),
    ) as unknown as GridMetrics;
    let styleCalls = 0;
    const segments = resolveEdges(
      { top: 0, left: 0, bottom: 2, right: 16_383 },
      sheet,
      () => {
        styleCalls += 1;
        return boxed("thin");
      },
      counted,
      buildMergeIndex([]),
      DEFAULT_THEME,
    );
    // Visible: columns A and 16,001–16,383 (384), rows 1 and 3 (+ row 4 below).
    // Stepping through the 16,000 hidden columns one by one would cost
    // tens of thousands of queries.
    expect(queries).toBeLessThan(10_000);
    expect(styleCalls).toBeLessThanOrEqual(3 * 384);
    const vertical = segments.filter((s) => s.orientation === "vertical");
    expect(vertical).toHaveLength(2 * 385); // every visible column boundary on rows 1 and 3
    expect(segments.every((s) => s.length > 0)).toBe(true);
    expect(new Set(vertical.map((s) => s.x)).size).toBe(385);
  });
});

describe("SPEC-117 Rendering: diagonal borders", () => {
  it("draws nothing without a diagonal, or with one that is neither up nor down", () => {
    expect(diagonalLines({}, DEFAULT_THEME)).toBeUndefined();
    expect(diagonalLines({ borders: { diagonal: { style: "thin" } } }, DEFAULT_THEME)).toBeUndefined();
    expect(diagonalLines(boxed("thin"), DEFAULT_THEME)).toBeUndefined();
  });

  it("draws up and/or down with the border table's width and the resolved colour", () => {
    expect(diagonalLines({ borders: { diagonal: { style: "thin", up: true } } }, DEFAULT_THEME)).toEqual({
      up: true,
      down: false,
      width: 1,
      lineStyle: "solid",
      color: "#000000",
    });
    expect(
      diagonalLines({ borders: { diagonal: { style: "medium", up: true, down: true, color: { theme: "accent1" } } } }, CUSTOM_THEME),
    ).toEqual({ up: true, down: true, width: 2, lineStyle: "solid", color: "#123456" });
  });

  it("dashes and dots dashed and dotted styles; hair at 60% opacity; double as 3 px", () => {
    expect(diagonalLines({ borders: { diagonal: { style: "dashed", down: true } } }, DEFAULT_THEME)).toMatchObject({
      down: true,
      up: false,
      dasharray: "3 1",
    });
    expect(diagonalLines({ borders: { diagonal: { style: "mediumDashed", down: true } } }, DEFAULT_THEME)).toMatchObject({
      width: 2,
      dasharray: "6 2",
    });
    expect(diagonalLines({ borders: { diagonal: { style: "dotted", down: true } } }, DEFAULT_THEME)).toMatchObject({
      dasharray: "1 1",
    });
    expect(diagonalLines({ borders: { diagonal: { style: "hair", up: true } } }, DEFAULT_THEME)).toMatchObject({
      lineStyle: "dotted",
      opacity: 0.6,
    });
    const double = diagonalLines({ borders: { diagonal: { style: "double", up: true } } }, DEFAULT_THEME)!;
    expect(double).toMatchObject({ width: 3, lineStyle: "double" });
    expect(double.dasharray).toBeUndefined();
  });
});

// --- overflow ----------------------------------------------------------------

/** A sheet whose row 0 holds `cells` (col → value). */
function rowSheet(cells: Record<number, CellValue | Cell>, partial: Partial<Sheet> = {}): Sheet {
  const stored: Cell[] = Object.entries(cells)
    .map(([col, value]): Cell =>
      typeof value === "object" && value !== null && "col" in value ? value : { col: Number(col), value: value as CellValue },
    )
    .toSorted((a, b) => a.col - b.col);
  return sheetWith({ rows: [{ index: 0, cells: stored }], ...partial });
}

function extent(
  sheet: Sheet,
  col: number,
  align: "left" | "right" | "center",
  frozenCols = 0,
): { left: number; right: number } {
  return overflowExtent(sheet, 0, col, align, buildGridMetrics(sheet), buildMergeIndex(sheet.merges), frozenCols);
}

describe("SPEC-117 Overflow: overflowExtent", () => {
  const A = 0;
  const B = 1;
  const C = 2;
  const D = 3;
  const E = 4;
  const F = 5;

  it("left-aligned text spills right into empty neighbours up to the first value", () => {
    const sheet = rowSheet({ [B]: "a long title", [F]: 5 });
    expect(extent(sheet, B, "left")).toEqual({ left: 0, right: 3 * 64 });
  });

  it("right-aligned text spills left only", () => {
    const sheet = rowSheet({ [A]: "x", [E]: "a long label" });
    expect(extent(sheet, E, "right")).toEqual({ left: 3 * 64, right: 0 });
  });

  it("centred text extends min(leftRoom, rightRoom) both ways", () => {
    expect(extent(rowSheet({ [B]: 1, [D]: "centre", [F]: 1 }), D, "center")).toEqual({ left: 64, right: 64 });
    expect(extent(rowSheet({ [A]: 1, [D]: "centre", [F]: 1 }), D, "center")).toEqual({ left: 64, right: 64 });
    expect(extent(rowSheet({ [C]: 1, [D]: "centre" }), D, "center")).toEqual({ left: 0, right: 0 });
    expect(extent(rowSheet({ [D]: "centre", [E]: true }), D, "center")).toEqual({ left: 0, right: 0 });
  });

  it("stops at any value: an empty string, a boolean, an error", () => {
    expect(extent(rowSheet({ [B]: "text", [E]: "" }), B, "left")).toEqual({ left: 0, right: 2 * 64 });
    expect(extent(rowSheet({ [B]: "text", [D]: false }), B, "left")).toEqual({ left: 0, right: 64 });
    expect(extent(rowSheet({ [B]: "text", [C]: { error: "#N/A" } }), B, "left")).toEqual({ left: 0, right: 0 });
  });

  it("passes style-only cells (value null): they are empty", () => {
    const sheet = rowSheet({ [B]: "text", [C]: { col: C, value: null, styleId: 3 }, [E]: 1 });
    expect(extent(sheet, B, "left")).toEqual({ left: 0, right: 2 * 64 });
  });

  it("stops at a merged range, and a merged cell never spills", () => {
    const merged = rowSheet({ [B]: "text" }, { merges: [range("E1:F2")] });
    expect(extent(merged, B, "left")).toEqual({ left: 0, right: 2 * 64 });
    const leftOf = rowSheet({ [F]: "text" }, { merges: [range("A1:B1")] });
    expect(extent(leftOf, F, "right")).toEqual({ left: 3 * 64, right: 0 });
    // A merge on another row does not block this one.
    expect(extent(rowSheet({ [B]: "text", [F]: 1 }, { merges: [range("D2:E3")] }), B, "left")).toEqual({
      left: 0,
      right: 3 * 64,
    });
    expect(extent(rowSheet({ [B]: "text" }, { merges: [range("B1:C1")] }), B, "left")).toEqual({ left: 0, right: 0 });
  });

  it("stops at the frozen column boundary, in both directions", () => {
    const sheet = rowSheet({ [A]: "frozen text", [E]: "scrolling text" });
    // Columns A–B frozen: A's text reaches B only; E's text reaches back to C only.
    expect(extent(sheet, A, "left", 2)).toEqual({ left: 0, right: 64 });
    expect(extent(sheet, E, "right", 2)).toEqual({ left: 2 * 64, right: 0 });
    expect(extent(sheet, B, "center", 2)).toEqual({ left: 0, right: 0 });
    // Without a freeze both run on.
    expect(extent(sheet, A, "left")).toEqual({ left: 0, right: 3 * 64 });
    expect(extent(sheet, E, "right")).toEqual({ left: 3 * 64, right: 0 });
  });

  it(`scans at most OVERFLOW_SCAN_COLS (${OVERFLOW_SCAN_COLS}) columns`, () => {
    expect(OVERFLOW_SCAN_COLS).toBe(256);
    expect(extent(rowSheet({ [A]: "text" }), A, "left")).toEqual({ left: 0, right: 256 * 64 });
    expect(extent(rowSheet({ [A]: "text", 400: 1 }), A, "left")).toEqual({ left: 0, right: 256 * 64 });
    expect(extent(rowSheet({ [A]: "text", 200: 1 }), A, "left")).toEqual({ left: 0, right: 199 * 64 });
    expect(extent(rowSheet({ 300: "text" }), 300, "right")).toEqual({ left: 256 * 64, right: 0 });
  });

  it("stops at the grid's edges", () => {
    expect(extent(rowSheet({ [A]: "text" }), A, "right")).toEqual({ left: 0, right: 0 });
    expect(extent(rowSheet({ 16_383: "text" }), 16_383, "left")).toEqual({ left: 0, right: 0 });
    expect(extent(rowSheet({ 16_380: "text" }), 16_380, "left")).toEqual({ left: 0, right: 3 * 64 });
  });

  it("adds each column's own width; hidden columns add 0 px and do not block", () => {
    const columns: ColumnSpan[] = [
      { min: C, max: C, widthPx: 100 },
      { min: D, max: D, hidden: true },
      { min: E, max: E, widthPx: 30 },
    ];
    expect(extent(rowSheet({ [B]: "text", [F]: 1 }, { columns }), B, "left")).toEqual({ left: 0, right: 130 });
    // A value stored in a hidden column is still a value: it blocks.
    expect(extent(rowSheet({ [B]: "text", [D]: 1 }, { columns }), B, "left")).toEqual({ left: 0, right: 100 });
    // A hidden cell paints nothing, so it never spills.
    expect(extent(rowSheet({ [D]: "hidden text" }, { columns }), D, "left")).toEqual({ left: 0, right: 0 });
  });
});

describe("SPEC-117 Overflow: overflowAlignment", () => {
  it("never spills a non-string (or an empty string)", () => {
    for (const value of [5, 0, true, { error: "#DIV/0!" }, null, ""] as CellValue[]) {
      expect(overflowAlignment({}, value)).toBeNull();
    }
  });

  it("maps general/left → left, right → right, center/centerContinuous → center", () => {
    expect(overflowAlignment({}, "text")).toBe("left");
    expect(overflowAlignment(aligned({ horizontal: "general" }), "text")).toBe("left");
    expect(overflowAlignment(aligned({ horizontal: "left", indent: 2 }), "text")).toBe("left");
    expect(overflowAlignment(aligned({ horizontal: "right" }), "text")).toBe("right");
    expect(overflowAlignment(aligned({ horizontal: "center" }), "text")).toBe("center");
    expect(overflowAlignment(aligned({ horizontal: "centerContinuous" }), "text")).toBe("center");
    expect(overflowAlignment(aligned({ textRotation: 0, vertical: "top" }), "text")).toBe("left");
  });

  it("never spills wrapped, filled, justified, distributed, rotated or shrink-to-fit text", () => {
    for (const alignment of [
      { wrap: true },
      { horizontal: "fill" },
      { horizontal: "justify" },
      { horizontal: "distributed" },
      { vertical: "justify" },
      { vertical: "distributed" },
      { textRotation: 45 },
      { textRotation: 135 },
      { textRotation: 255 },
      { shrinkToFit: true },
      { horizontal: "center", wrap: true },
    ] as CellAlignment[]) {
      expect(overflowAlignment(aligned(alignment), "long text")).toBeNull();
    }
  });
});

describe("SPEC-117 Overflow: strings anchored outside the visible window", () => {
  it("finds the nearest stored string left of the window", () => {
    const sheet = rowSheet({ 0: "A long title anchored at A1", 10: "inside" });
    expect(offscreenSpillCandidate(sheet, 0, 5, 12).left?.col).toBe(0);
    expect(offscreenSpillCandidate(sheet, 0, 5, 12).right).toBeUndefined();
  });

  it("finds the nearest stored string right of the window (for right-aligned and centred text)", () => {
    const sheet = rowSheet({ 1: "inside", 14: "a right-aligned label", 20: "further" });
    expect(offscreenSpillCandidate(sheet, 0, 0, 10).right?.col).toBe(14);
    expect(offscreenSpillCandidate(sheet, 0, 0, 10).left).toBeUndefined();
  });

  it("returns nothing when a nearer value would block it, but looks past style-only cells", () => {
    expect(offscreenSpillCandidate(rowSheet({ 0: "title", 3: 42 }), 0, 5, 12).left).toBeUndefined();
    expect(offscreenSpillCandidate(rowSheet({ 0: "title", 3: "" }), 0, 5, 12).left).toBeUndefined();
    const styled = rowSheet({ 0: "title", 3: { col: 3, value: null, styleId: 2 } });
    expect(offscreenSpillCandidate(styled, 0, 5, 12).left?.col).toBe(0);
    expect(offscreenSpillCandidate(rowSheet({ 20: 7, 25: "label" }), 0, 0, 10).right).toBeUndefined();
  });

  it("looks no further than OVERFLOW_SCAN_COLS columns", () => {
    const sheet = rowSheet({ 0: "title" });
    expect(offscreenSpillCandidate(sheet, 0, 256, 300).left?.col).toBe(0);
    expect(offscreenSpillCandidate(sheet, 0, 257, 300).left).toBeUndefined();
    const right = rowSheet({ 266: "label" });
    expect(offscreenSpillCandidate(right, 0, 0, 10).right?.col).toBe(266);
    expect(offscreenSpillCandidate(right, 0, 0, 9).right).toBeUndefined();
  });

  it("returns nothing for an unstored row", () => {
    expect(offscreenSpillCandidate(sheetWith(), 4, 5, 9)).toEqual({});
  });

  it("feeds overflowExtent: an off-screen title reaches into the window", () => {
    const sheet = rowSheet({ 0: "A title much wider than its column" });
    const candidate = offscreenSpillCandidate(sheet, 0, 3, 8).left!;
    expect(overflowAlignment({}, candidate.value)).toBe("left");
    const room = extent(sheet, candidate.col, "left");
    expect(room.right).toBeGreaterThan(buildGridMetrics(sheet).colLeft(3) - 64);
  });
});

// --- text width, shrink, fill, display -----------------------------------------

describe("SPEC-117 Rendering: the text-width estimate", () => {
  it("weighs . , : and space 0.5 and everything else 1, times the MDW", () => {
    expect(estimateTextWidth("Hello", 7)).toBe(35);
    expect(estimateTextWidth("1,234.50", 7)).toBe(7 * (6 + 2 * 0.5));
    expect(estimateTextWidth("a b", 8)).toBe(8 * 2.5);
    expect(estimateTextWidth("12:30", 7)).toBe(7 * 4.5);
    expect(estimateTextWidth("", 7)).toBe(0);
    expect(estimateTextWidth("😀", 7)).toBe(7); // per character, not per UTF-16 unit
  });

  it("shrinks by min(1, available / estimate)", () => {
    expect(shrinkScale("Hello", 35, 7)).toBe(1);
    expect(shrinkScale("Hello", 100, 7)).toBe(1);
    expect(shrinkScale("Hello", 17.5, 7)).toBe(0.5);
    expect(shrinkScale("", 10, 7)).toBe(1);
    expect(shrinkScale("Hello", 0, 7)).toBe(0);
  });

  it("repeats fill text ceil(width / text width) times", () => {
    expect(fillRepeat("-", 64, 7)).toBe("-".repeat(10));
    expect(fillRepeat("ab", 28, 7)).toBe("abab");
    expect(fillRepeat("ab", 29, 7)).toBe("ababab");
    expect(fillRepeat("*", 7, 7)).toBe("*");
    expect(fillRepeat("ab", 0, 7)).toBe("ab");
    expect(fillRepeat("", 50, 7)).toBe("");
  });
});

describe("SPEC-117 Rendering: values show through generalText; showZeros: false blanks numeric zeros", () => {
  it("shows General text", () => {
    expect(displayText(5, {})).toBe("5");
    expect(displayText(0.5, {})).toBe("0.5");
    expect(displayText("text", {})).toBe("text");
    expect(displayText(true, {})).toBe("TRUE");
    expect(displayText({ error: "#N/A" }, {})).toBe("#N/A");
    expect(displayText(null, {})).toBe("");
    expect(displayText(0, {})).toBe("0");
  });

  it("blanks numeric zeros only", () => {
    const hidden = { showZeros: false } as const;
    expect(displayText(0, hidden)).toBe("");
    expect(displayText(-0, hidden)).toBe("");
    expect(displayText("0", hidden)).toBe("0");
    expect(displayText(false, hidden)).toBe("FALSE");
    expect(displayText(0.0001, hidden)).toBe("0.0001");
  });
});

describe("SPEC-118: a cell's formatted display and its layout", () => {
  const ctx = (widthPx?: number, tag: keyof typeof SPREADSHEET_LOCALES = "en-US"): FormatContext => ({
    locale: SPREADSHEET_LOCALES[tag],
    dateSystem: "1900",
    digitWidthPx: 7,
    ...(widthPx === undefined ? {} : { widthPx }),
  });

  it("renders a value through formatCodeOf(style, locale), fitted to the width; showZeros: false still blanks zeros", () => {
    expect(cellDisplay(46303, { builtinFormatId: 14, numberFormat: "m/d/yyyy" }, {}, ctx(100))).toMatchObject({ text: "10/8/2026", alignHint: "right" });
    expect(cellDisplay(46303, { builtinFormatId: 14, numberFormat: "m/d/yyyy" }, {}, ctx(100, "de-DE")).text).toBe("08.10.2026");
    expect(cellDisplay(46303, { numberFormat: "m/d/yyyy" }, {}, ctx(30))).toMatchObject({ text: "###", overflow: "hash" });
    expect(cellDisplay(-5, { numberFormat: "0;[Red]-0" }, {}, ctx())).toMatchObject({ text: "-5", color: "#ff0000" });
    expect(cellDisplay(0, { numberFormat: "0.00" }, { showZeros: false }, ctx())).toEqual({ text: "" });
    expect(cellDisplay("text", { numberFormat: "0.00" }, {}, ctx(10))).toMatchObject({ text: "text", alignHint: "left" });
  });

  it("lays a format fill out as the text either side of a run long enough for the cell", () => {
    const formatted = cellDisplay(5, { numberFormat: '_("$"* #,##0.00_)' }, {}, ctx(100));
    const laid = cellText(formatted, {}, "number", 100, 7);
    expect(laid.text).toBe(" $ 5.00 ");
    expect(laid.fill).toEqual({ before: " $", run: " ".repeat(Math.ceil(100 / 3.5) + 1), after: "5.00 " });
    expect(laid.horizontal).toBe("right");
    expect(cellText({ text: "x", alignHint: "left" }, {}, "number", 64, 7)).toEqual({ text: "x", fill: null, horizontal: "left" });
  });

  it("keeps SPEC-117's fill alignment and shrink-to-fit", () => {
    expect(cellText({ text: "ab" }, { alignment: { horizontal: "fill" } }, "string", 64, 7).text).toBe("ab".repeat(Math.ceil(64 / 14)));
    const shrunk = cellText({ text: "a long text that cannot fit" }, { alignment: { shrinkToFit: true }, font: { sizePt: 11 } }, "string", 64, 7);
    expect(shrunk.shrunkFontSize).toMatch(/^\d+\.\d{2}pt$/);
    expect(cellText({ text: "ok" }, { alignment: { shrinkToFit: true } }, "string", 64, 7).shrunkFontSize).toBeUndefined();
  });

  it("places a cell's box, a merge without a fill painting over its interior gridlines", () => {
    const place = { left: 10, top: 0, width: 64, height: 20 };
    expect(cellBoxCss({ "text-align": "left" }, place, false)).toEqual({ "text-align": "left", left: "10px", top: "0px", width: "64px", height: "20px" });
    expect(cellBoxCss({}, place, true)).toMatchObject({ "background-image": "linear-gradient(#ffffff, #ffffff)", "background-size": "calc(100% - 1px) calc(100% - 1px)" });
    expect(cellBoxCss({ "background-color": "#ffff00" }, place, true)["background-image"]).toBeUndefined();
  });

  it("centres an edge's line on it, the hair style at 60%", () => {
    const thick = edgeElementCss({ x: 64, y: 20, length: 64, orientation: "horizontal", style: "thick", color: "#000000" }, 0, 0);
    expect(thick).toEqual({ left: "64px", top: "18px", width: "64px", "border-top": "3px solid #000000" });
    const hair = edgeElementCss({ x: 64, y: 20, length: 20, orientation: "vertical", style: "hair", color: "#ff0000" }, 10, 5);
    expect(hair).toEqual({ left: "53px", top: "15px", height: "20px", "border-left": "1px dotted #ff0000", opacity: "0.6" });
  });
});
