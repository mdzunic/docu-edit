import { describe, expect, it } from "vitest";
import {
  initialSheetSetup,
  marginPresetOf,
  sheetGeometryFor,
  sheetOutputEstimate,
  sheetSetupRefusal,
} from "./assemblySheet";
import { MARGIN_PRESETS, PREFERENCE_DEFAULTS } from "./preferences";
import { FIT_TO_CONTENT, placeOnSheet, resolveSheet, SHEET_DEFAULTS, type SheetSetup } from "./core/sheetLayout";

function setup(patch: Partial<SheetSetup>): SheetSetup {
  return { ...SHEET_DEFAULTS, ...patch };
}

describe("initialSheetSetup", () => {
  it("starts from the page-setup defaults, with the output's background", () => {
    const pdf = initialSheetSetup({ ...PREFERENCE_DEFAULTS, defaultPageSizeId: "Letter", defaultMarginsId: "Wide" }, "pdf");
    expect(pdf.setup).toMatchObject({ paper: "Letter", margins: MARGIN_PRESETS.Wide, background: "#ffffff", dpi: 150 });
    expect(pdf.linked).toBe(false);
    expect(initialSheetSetup(PREFERENCE_DEFAULTS, "png").setup.background).toBe("transparent");
  });

  it("prefers the user's assembly choices, but never forces margins a paper cannot hold", () => {
    const chosen = initialSheetSetup({ ...PREFERENCE_DEFAULTS, assemblyPaper: "A5", assemblyDpi: 600, assemblyMargins: { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 } }, "pdf");
    expect(chosen.setup).toMatchObject({ paper: "A5", dpi: 600, margins: { topCm: 1 } });
    expect(chosen.linked).toBe(true);
    const tooBig = initialSheetSetup({ ...PREFERENCE_DEFAULTS, assemblyPaper: "A5", assemblyMargins: { topCm: 8, rightCm: 8, bottomCm: 8, leftCm: 8 } }, "pdf");
    expect(tooBig.setup.margins).toEqual(MARGIN_PRESETS.Normal);
  });
});

describe("sheetSetupRefusal", () => {
  const item = { widthPx: 64, heightPx: 40 };

  it("accepts a valid setup", () => {
    expect(sheetSetupRefusal(SHEET_DEFAULTS, item, "cm")).toBeNull();
  });

  it("words the margin refusal in the user's unit", () => {
    const margins = { topCm: 2.54, rightCm: 10, bottomCm: 2.54, leftCm: 10.5 };
    expect(sheetSetupRefusal(setup({ margins }), item, "in")).toBe(
      "These margins leave 0.2 × 9.69 in for the content on a 8.27 × 11.69 in sheet; it needs at least 0.39 in each way.",
    );
  });

  it("refuses negative margins, an oversized sheet and Fit to content without a size", () => {
    expect(sheetSetupRefusal(setup({ margins: { ...SHEET_DEFAULTS.margins, topCm: -1 } }), item, "cm")).toBe("Margins cannot be negative.");
    expect(sheetSetupRefusal(setup({ paper: "Custom", widthCm: 600, heightCm: 10 }), item, "cm")).toContain("at most 500 cm");
    expect(sheetSetupRefusal(setup({ paper: FIT_TO_CONTENT }), null, "cm")).toContain("needs the image's size");
    expect(sheetSetupRefusal(setup({ paper: FIT_TO_CONTENT, dpi: 72 }), { widthPx: 20000, heightPx: 10 }, "cm")).toContain(
      "That sheet would be 710.64 × 5.43 cm",
    );
  });

  it("does not hold Fit to content to the content minimum: the content is the item itself", () => {
    expect(sheetSetupRefusal(setup({ paper: FIT_TO_CONTENT, margins: { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 } }), item, "cm")).toBeNull();
  });
});

describe("marginPresetOf", () => {
  it("names a preset, else Custom", () => {
    expect(marginPresetOf(MARGIN_PRESETS.Moderate)).toBe("Moderate");
    expect(marginPresetOf({ ...MARGIN_PRESETS.Narrow, leftCm: 2 })).toBe("Custom");
  });
});

describe("sheetOutputEstimate", () => {
  it("a raster output is the whole sheet; a PDF carries only the placed samples", () => {
    const a3 = resolveSheet(setup({ paper: "A3" }));
    expect(sheetOutputEstimate(a3, null, 600, "png")).toBe("7016 × 9921 px at 600 dpi · up to about 265.5 MB");
    const placed = placeOnSheet({ widthPx: 64, heightPx: 40 }, a3.content, "contain", "center", 600);
    expect(sheetOutputEstimate(a3, placed, 600, "pdf")).toBe("7016 × 9921 px at 600 dpi · up to about 12 KB");
    expect(sheetGeometryFor(setup({ paper: FIT_TO_CONTENT }), null)).toBeNull();
  });
});
