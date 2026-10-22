import { describe, expect, it } from "vitest";
import { MARGIN_PRESETS } from "../preferences";
import { PAPER_SIZES } from "./pageSetup";
import {
  backgroundRefusal,
  defaultSheetBackground,
  effectiveBackground,
  FIT_TO_CONTENT,
  isSheetBackground,
  itemSizeCm,
  MAX_SHEET_CM,
  marginsLeaveContent,
  MIN_CONTENT_CM,
  placeOnSheet,
  r3,
  resolveSheet,
  SHEET_ALIGNS,
  SHEET_DEFAULTS,
  SHEET_DPI_CHOICES,
  SHEET_PAPER_NAMES,
  sheetPixelSize,
  sheetSizeAllowed,
  type SheetFit,
  type SheetSetup,
} from "./sheetLayout";

const ONE_CM = { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 };
const NO_MARGINS = { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 };

function setup(patch: Partial<SheetSetup>): SheetSetup {
  return { ...SHEET_DEFAULTS, ...patch };
}

// A 30 × 20 cm landscape sheet and a 20 × 30 cm portrait one, 1 cm margins:
// content boxes 28 × 18 and 18 × 28 at (1, 1).
const WIDE_SHEET = resolveSheet(setup({ paper: "Custom", widthCm: 30, heightCm: 20, orientation: "landscape", margins: ONE_CM }));
const TALL_SHEET = resolveSheet(setup({ paper: "Custom", widthCm: 20, heightCm: 30, orientation: "portrait", margins: ONE_CM }));
// At 150 dpi: 25.4 × 12.7 cm and 12.7 × 25.4 cm.
const WIDE_ITEM = { widthPx: 1500, heightPx: 750 };
const TALL_ITEM = { widthPx: 750, heightPx: 1500 };

type Expected = {
  rect: [number, number, number, number];
  item: [number, number, number, number];
  crop: [number, number, number, number];
};

const MATRIX: Array<{ sheet: "wide" | "tall"; item: "wide" | "tall"; fit: SheetFit; expected: Expected }> = [
  // --- wide sheet (content 28 × 18 at 1,1) ---
  { sheet: "wide", item: "wide", fit: "contain", expected: { rect: [1, 3, 28, 14], item: [1, 3, 28, 14], crop: [0, 0, 1500, 750] } },
  { sheet: "wide", item: "wide", fit: "cover", expected: { rect: [1, 1, 28, 18], item: [-3, 1, 36, 18], crop: [166.667, 0, 1166.667, 750] } },
  { sheet: "wide", item: "wide", fit: "stretch", expected: { rect: [1, 1, 28, 18], item: [1, 1, 28, 18], crop: [0, 0, 1500, 750] } },
  { sheet: "wide", item: "wide", fit: "actual", expected: { rect: [2.3, 3.65, 25.4, 12.7], item: [2.3, 3.65, 25.4, 12.7], crop: [0, 0, 1500, 750] } },
  { sheet: "wide", item: "tall", fit: "contain", expected: { rect: [10.5, 1, 9, 18], item: [10.5, 1, 9, 18], crop: [0, 0, 750, 1500] } },
  { sheet: "wide", item: "tall", fit: "cover", expected: { rect: [1, 1, 28, 18], item: [1, -18, 28, 56], crop: [0, 508.929, 750, 482.143] } },
  { sheet: "wide", item: "tall", fit: "stretch", expected: { rect: [1, 1, 28, 18], item: [1, 1, 28, 18], crop: [0, 0, 750, 1500] } },
  { sheet: "wide", item: "tall", fit: "actual", expected: { rect: [8.65, 1, 12.7, 18], item: [8.65, -2.7, 12.7, 25.4], crop: [0, 218.504, 750, 1062.992] } },
  // --- tall sheet (content 18 × 28 at 1,1) ---
  { sheet: "tall", item: "wide", fit: "contain", expected: { rect: [1, 10.5, 18, 9], item: [1, 10.5, 18, 9], crop: [0, 0, 1500, 750] } },
  { sheet: "tall", item: "wide", fit: "cover", expected: { rect: [1, 1, 18, 28], item: [-18, 1, 56, 28], crop: [508.929, 0, 482.143, 750] } },
  { sheet: "tall", item: "wide", fit: "stretch", expected: { rect: [1, 1, 18, 28], item: [1, 1, 18, 28], crop: [0, 0, 1500, 750] } },
  { sheet: "tall", item: "wide", fit: "actual", expected: { rect: [1, 8.65, 18, 12.7], item: [-2.7, 8.65, 25.4, 12.7], crop: [218.504, 0, 1062.992, 750] } },
  { sheet: "tall", item: "tall", fit: "contain", expected: { rect: [3, 1, 14, 28], item: [3, 1, 14, 28], crop: [0, 0, 750, 1500] } },
  { sheet: "tall", item: "tall", fit: "cover", expected: { rect: [1, 1, 18, 28], item: [1, -3, 18, 36], crop: [0, 166.667, 750, 1166.667] } },
  { sheet: "tall", item: "tall", fit: "stretch", expected: { rect: [1, 1, 18, 28], item: [1, 1, 18, 28], crop: [0, 0, 750, 1500] } },
  { sheet: "tall", item: "tall", fit: "actual", expected: { rect: [3.65, 2.3, 12.7, 25.4], item: [3.65, 2.3, 12.7, 25.4], crop: [0, 0, 750, 1500] } },
];

describe("sheetLayout — the fit matrix (4 fits × wide/tall item × wide/tall sheet)", () => {
  it("has all sixteen cases", () => {
    expect(MATRIX).toHaveLength(16);
  });

  it.each(MATRIX)("$fit: $item item on a $sheet sheet", ({ sheet, item, fit, expected }) => {
    const geometry = sheet === "wide" ? WIDE_SHEET : TALL_SHEET;
    const placed = placeOnSheet(item === "wide" ? WIDE_ITEM : TALL_ITEM, geometry.content, fit, "center", 150);
    expect([placed.xCm, placed.yCm, placed.widthCm, placed.heightCm]).toEqual(expected.rect);
    expect([placed.item.xCm, placed.item.yCm, placed.item.widthCm, placed.item.heightCm]).toEqual(expected.item);
    expect([placed.crop.xPx, placed.crop.yPx, placed.crop.widthPx, placed.crop.heightPx]).toEqual(expected.crop);
    const whole = expected.crop[2] === (item === "wide" ? 1500 : 750) && expected.crop[3] === (item === "wide" ? 750 : 1500);
    expect(placed.cropped).toBe(!whole);
  });

  it("contain keeps the aspect and never crops; stretch fills the box exactly", () => {
    const contain = placeOnSheet(WIDE_ITEM, TALL_SHEET.content, "contain", "center");
    expect(contain.widthCm / contain.heightCm).toBe(2);
    expect(contain.cropped).toBe(false);
    const stretch = placeOnSheet(WIDE_ITEM, TALL_SHEET.content, "stretch", "center");
    expect(stretch).toMatchObject(TALL_SHEET.content);
  });

  it("the sheet results themselves", () => {
    expect(WIDE_SHEET).toEqual({ widthCm: 30, heightCm: 20, content: { xCm: 1, yCm: 1, widthCm: 28, heightCm: 18 } });
    expect(TALL_SHEET).toEqual({ widthCm: 20, heightCm: 30, content: { xCm: 1, yCm: 1, widthCm: 18, heightCm: 28 } });
  });
});

describe("sheetLayout — the nine alignments", () => {
  // A 12.7 × 6.35 cm item (actual size at 150 dpi) in the 28 × 18 box leaves
  // 15.3 cm of slack across and 11.65 cm down.
  const item = { widthPx: 750, heightPx: 375 };
  const expected: Record<string, [number, number]> = {
    "top-left": [1, 1],
    top: [8.65, 1],
    "top-right": [16.3, 1],
    left: [1, 6.825],
    center: [8.65, 6.825],
    right: [16.3, 6.825],
    "bottom-left": [1, 12.65],
    bottom: [8.65, 12.65],
    "bottom-right": [16.3, 12.65],
  };

  it("lists the grid row by row", () => {
    expect(SHEET_ALIGNS).toEqual(Object.keys(expected));
  });

  it.each(SHEET_ALIGNS)("%s", (align) => {
    const placed = placeOnSheet(item, WIDE_SHEET.content, "actual", align, 150);
    expect([placed.xCm, placed.yCm]).toEqual(expected[align]);
    expect([placed.widthCm, placed.heightCm]).toEqual([12.7, 6.35]);
  });

  it("for a cropping fit, the alignment chooses which part of the item survives", () => {
    const left = placeOnSheet(WIDE_ITEM, WIDE_SHEET.content, "cover", "left");
    const right = placeOnSheet(WIDE_ITEM, WIDE_SHEET.content, "cover", "right");
    expect(left.crop).toEqual({ xPx: 0, yPx: 0, widthPx: 1166.667, heightPx: 750 });
    expect(right.crop).toEqual({ xPx: 333.333, yPx: 0, widthPx: 1166.667, heightPx: 750 });
    // The visible box is the content box either way.
    expect([left.xCm, left.widthCm, right.xCm, right.widthCm]).toEqual([1, 28, 1, 28]);
  });

  it("defaults to center", () => {
    expect(SHEET_DEFAULTS.align).toBe("center");
    expect(SHEET_DEFAULTS.fit).toBe("contain");
  });
});

describe("sheetLayout — paper", () => {
  it("offers SPEC-027's catalogue plus Fit to content, resolved through PAPER_SIZES", () => {
    expect(SHEET_PAPER_NAMES).toEqual(["A4", "Letter", "Legal", "A3", "A5", "Tabloid", "Custom", FIT_TO_CONTENT]);
    const letter = resolveSheet(setup({ paper: "Letter", margins: NO_MARGINS }));
    expect([letter.widthCm, letter.heightCm]).toEqual([PAPER_SIZES.Letter.widthCm, PAPER_SIZES.Letter.heightCm]);
  });

  it("applies orientation to a named paper and to Custom", () => {
    const a4 = resolveSheet(setup({ paper: "A4", orientation: "landscape" }));
    expect([a4.widthCm, a4.heightCm]).toEqual([29.7, 21]);
    const custom = resolveSheet(setup({ paper: "Custom", widthCm: 10, heightCm: 15, orientation: "landscape" }));
    expect([custom.widthCm, custom.heightCm]).toEqual([15, 10]);
  });

  it("the defaults: A4 portrait, Normal margins, white, 150 dpi", () => {
    expect(SHEET_DEFAULTS).toMatchObject({ paper: "A4", orientation: "portrait", background: "#ffffff", dpi: 150 });
    expect(SHEET_DEFAULTS.margins).toEqual(MARGIN_PRESETS.Normal);
    expect(resolveSheet(SHEET_DEFAULTS)).toEqual({
      widthCm: 21,
      heightCm: 29.7,
      content: { xCm: 2.54, yCm: 2.54, widthCm: 15.92, heightCm: 24.62 },
    });
  });
});

describe('sheetLayout — "Fit to content"', () => {
  it("a wide item makes a sheet of exactly its shape plus margins", () => {
    // 1200 × 800 px at 72 dpi is 42.333 × 28.222 cm.
    const geometry = resolveSheet(setup({ paper: FIT_TO_CONTENT, dpi: 72, margins: NO_MARGINS }), { widthPx: 1200, heightPx: 800 });
    expect(geometry).toEqual({ widthCm: 42.333, heightCm: 28.222, content: { xCm: 0, yCm: 0, widthCm: 42.333, heightCm: 28.222 } });
    const placed = placeOnSheet({ widthPx: 1200, heightPx: 800 }, geometry.content, "contain", "center", 72);
    expect([placed.xCm, placed.yCm, placed.widthCm, placed.heightCm]).toEqual([0, 0, 42.333, 28.222]);
  });

  it("a tall item makes a tall sheet, whatever the orientation says", () => {
    const geometry = resolveSheet(
      setup({ paper: FIT_TO_CONTENT, dpi: 150, orientation: "landscape", margins: ONE_CM }),
      TALL_ITEM,
    );
    expect(geometry).toEqual({ widthCm: 14.7, heightCm: 27.4, content: { xCm: 1, yCm: 1, widthCm: 12.7, heightCm: 25.4 } });
  });

  it("throws without the item's size rather than inventing one", () => {
    expect(() => resolveSheet(setup({ paper: FIT_TO_CONTENT }))).toThrow();
  });
});

describe("sheetLayout — margin validation", () => {
  const a4 = PAPER_SIZES.A4;

  it(`accepts margins that leave exactly ${MIN_CONTENT_CM} cm, refuses a thousandth less`, () => {
    expect(marginsLeaveContent(a4, { topCm: 1, rightCm: 10, bottomCm: 1, leftCm: 10 })).toBe(true);
    expect(marginsLeaveContent(a4, { topCm: 1, rightCm: 10.001, bottomCm: 1, leftCm: 10 })).toBe(false);
    expect(marginsLeaveContent(a4, { topCm: 14.35, rightCm: 1, bottomCm: 14.35, leftCm: 1 })).toBe(true);
    expect(marginsLeaveContent(a4, { topCm: 14.35, rightCm: 1, bottomCm: 14.351, leftCm: 1 })).toBe(false);
  });

  it("refuses a negative margin", () => {
    expect(marginsLeaveContent(a4, { topCm: -0.1, rightCm: 1, bottomCm: 1, leftCm: 1 })).toBe(false);
  });

  it(`refuses a sheet over ${MAX_SHEET_CM} cm or without area`, () => {
    expect(sheetSizeAllowed(MAX_SHEET_CM, 10)).toBe(true);
    expect(sheetSizeAllowed(MAX_SHEET_CM + 0.001, 10)).toBe(false);
    expect(sheetSizeAllowed(0, 10)).toBe(false);
  });
});

describe("sheetLayout — resolution", () => {
  it("is a fixed ladder with 150 the default", () => {
    expect(SHEET_DPI_CHOICES).toEqual([72, 150, 300, 600]);
    expect(SHEET_DEFAULTS.dpi).toBe(150);
  });

  it("sheetPixelSize rounds to whole pixels at every rung", () => {
    const a3 = resolveSheet(setup({ paper: "A3" }));
    expect(SHEET_DPI_CHOICES.map((dpi) => sheetPixelSize(a3, dpi))).toEqual([
      { widthPx: 842, heightPx: 1191 },
      { widthPx: 1754, heightPx: 2480 },
      { widthPx: 3508, heightPx: 4961 },
      { widthPx: 7016, heightPx: 9921 },
    ]);
  });

  it("placement for contain/cover/stretch does not depend on dpi; actual does", () => {
    for (const fit of ["contain", "cover", "stretch"] as const) {
      const [first, ...rest] = SHEET_DPI_CHOICES.map((dpi) => placeOnSheet({ widthPx: 1201, heightPx: 799 }, WIDE_SHEET.content, fit, "center", dpi));
      for (const other of rest) {
        expect(other).toEqual(first);
      }
    }
    const at72 = placeOnSheet({ widthPx: 100, heightPx: 100 }, WIDE_SHEET.content, "actual", "center", 72);
    const at600 = placeOnSheet({ widthPx: 100, heightPx: 100 }, WIDE_SHEET.content, "actual", "center", 600);
    expect([at72.widthCm, at600.widthCm]).toEqual([3.528, 0.423]);
    expect(itemSizeCm({ widthPx: 600, heightPx: 300 }, 600)).toEqual({ widthCm: 2.54, heightCm: 1.27 });
  });
});

describe("sheetLayout — background", () => {
  it("accepts #rrggbb or transparent", () => {
    expect(isSheetBackground("#ff00Aa")).toBe(true);
    expect(isSheetBackground("transparent")).toBe(true);
    expect(isSheetBackground("red")).toBe(false);
    expect(isSheetBackground("#fff")).toBe(false);
  });

  it("defaults to white for PDF and transparent for PNG", () => {
    expect(defaultSheetBackground("pdf")).toBe("#ffffff");
    expect(defaultSheetBackground("png")).toBe("transparent");
  });

  it("refuses transparent for PDF (falling back to white) and JPEG, with a reason", () => {
    expect(backgroundRefusal("transparent", "pdf")).toMatch(/white/);
    expect(backgroundRefusal("transparent", "jpeg")).toMatch(/transparency/);
    expect(backgroundRefusal("transparent", "png")).toBeNull();
    expect(backgroundRefusal("#123456", "pdf")).toBeNull();
    expect(effectiveBackground("transparent", "pdf")).toBe("#ffffff");
    expect(effectiveBackground("#123456", "pdf")).toBe("#123456");
  });
});

describe("sheetLayout — rounding and purity", () => {
  it("r3 rounds to three decimals, decimal-wise, and never returns -0", () => {
    expect(r3(11.5375)).toBe(11.538);
    expect(r3(1.0005)).toBe(1.001);
    expect(r3(2.6500000000000004)).toBe(2.65);
    expect(Object.is(r3(-0.0001), 0)).toBe(true);
    expect(r3(Number.NaN)).toBe(0);
  });

  it("every returned length has at most three decimals", () => {
    const geometry = resolveSheet(setup({ paper: "Letter", margins: { topCm: 1.111, rightCm: 2.2222, bottomCm: 0.3333, leftCm: 1.7777 } }));
    const placed = placeOnSheet({ widthPx: 333, heightPx: 777 }, geometry.content, "cover", "bottom-right", 300);
    const numbers = [
      geometry.widthCm, geometry.heightCm, ...Object.values(geometry.content),
      placed.xCm, placed.yCm, placed.widthCm, placed.heightCm,
      ...Object.values(placed.item), ...Object.values(placed.crop),
    ];
    for (const value of numbers) {
      expect(r3(value)).toBe(value);
    }
  });

  it("the preview and the writer, given the same setup, get identical numbers", () => {
    const run = () => {
      const s = setup({ paper: "A4", margins: MARGIN_PRESETS.Narrow, fit: "contain" });
      const geometry = resolveSheet(s, { widthPx: 64, heightPx: 40 });
      return { geometry, placed: placeOnSheet({ widthPx: 64, heightPx: 40 }, geometry.content, s.fit, s.align, s.dpi) };
    };
    const preview = run();
    const writer = run();
    expect(writer).toEqual(preview);
    expect([preview.placed.widthCm, preview.placed.heightCm]).toEqual([18.46, 11.538]);
  });

  it("is total: a zero-size item places nothing rather than NaN", () => {
    const placed = placeOnSheet({ widthPx: 0, heightPx: 40 }, WIDE_SHEET.content, "contain", "center");
    expect(Object.values(placed.crop).every(Number.isFinite)).toBe(true);
    expect([placed.widthCm, placed.heightCm, placed.cropped]).toEqual([0, 0, false]);
  });
});
