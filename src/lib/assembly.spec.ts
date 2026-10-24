// SPEC-109: the workspace's framework-free rules — the output's two bounds
// (MAX_OUTPUT_EDGE_PX per axis, the canvas ceiling on the area) refused with
// the limit and the dpi that would fit; the estimate; JPEG's fallback for a
// transparent sheet; the tray's announcements and the restored-tray notice;
// and the persisted plan's tolerant reader (lib/assemblyPersist.ts).
import { describe, expect, it } from "vitest";
import {
  droppedItemsNotice,
  estimateOutputBytes,
  fittingDpi,
  jpegBackgroundNotice,
  MAX_OUTPUT_EDGE_PX,
  moveEntry,
  outputBackground,
  outputEstimate,
  outputFileName,
  outputFits,
  outputRefusal,
  pageCountLabel,
  pdfDpiWarning,
  pdfEstimateLine,
  trayMoveAnnouncement,
} from "./assembly";
import {
  ASSEMBLY_LAYOUT_DEFAULTS,
  ASSEMBLY_OUTPUT_DEFAULTS,
  capTrayItems,
  DEFAULT_ASSEMBLY_TITLE,
  freshAssembly,
  isPageItem,
  parsePersistedAssembly,
  type PersistedAssemblyItem,
} from "./assemblyPersist";
import { resolveSheet, SHEET_DEFAULTS, type SheetSetup } from "./core/sheetLayout";
import { GUARANTEED_CANVAS_PIXELS } from "./media/canvasCeiling";

const A4 = resolveSheet({ ...SHEET_DEFAULTS, orientation: "landscape" });

function custom(widthCm: number, heightCm: number) {
  return resolveSheet({ ...SHEET_DEFAULTS, paper: "Custom", widthCm, heightCm, orientation: widthCm >= heightCm ? "landscape" : "portrait" });
}

/** 16200 × 100 px of content with 2 cm margins, fitted at `dpi`. */
function framedAt(dpi: number) {
  const margins = { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 };
  return resolveSheet({ ...SHEET_DEFAULTS, paper: "Fit to content", margins, dpi }, { widthPx: 16200, heightPx: 100 });
}

describe("the output's bounds", () => {
  it("bounds either axis independently at 16384 px", () => {
    expect(MAX_OUTPUT_EDGE_PX).toBe(16384);
    const huge = Number.MAX_SAFE_INTEGER;
    expect(outputFits({ widthPx: 16384, heightPx: 10 }, huge)).toBe(true);
    expect(outputFits({ widthPx: 16385, heightPx: 10 }, huge)).toBe(false);
    expect(outputFits({ widthPx: 10, heightPx: 16385 }, huge)).toBe(false);
  });

  it("holds the area to the canvas ceiling", () => {
    expect(outputFits({ widthPx: 4000, heightPx: 4000 }, 16_000_000)).toBe(true);
    expect(outputFits({ widthPx: 4001, heightPx: 4000 }, 16_000_000)).toBe(false);
  });

  it("accepts A4 at every dpi under a generous ceiling", () => {
    for (const dpi of [72, 150, 300, 600]) {
      expect(outputRefusal(A4, dpi, 268_435_456)).toBeNull();
    }
  });

  it("refuses a side over 16384 px, naming the limit and the dpi that would fit", () => {
    // 200 cm at 150 dpi is 11811 px a side; at 300 dpi, 23622.
    const sheet = custom(200, 100);
    expect(outputRefusal(sheet, 150, 268_435_456)).toBeNull();
    expect(outputRefusal(sheet, 300, 268_435_456)).toBe(
      "At 300 dpi the image would be 23622 × 11811 px; either side can be at most 16384 px. At 150 dpi it would fit.",
    );
    expect(fittingDpi(() => sheet, 268_435_456)).toBe(150);
  });

  it("refuses an area over the measured ceiling, naming it and the dpi that would fit", () => {
    // A4 landscape at 300 dpi is 3508 × 2480 = 8.7 megapixels.
    expect(outputRefusal(A4, 300, 4_000_000)).toBe(
      "At 300 dpi the image would be 3508 × 2480 px (8.7 megapixels), more than this browser's canvas limit of 4.0 megapixels. At 150 dpi it would fit.",
    );
  });

  it("tries each dpi on the sheet at that dpi — a sheet fitted to its content does not shrink with it", () => {
    // 20000 × 3000 px of content, no margins: the sheet is the content's
    // pixels at every dpi, so no lower dpi brings it under 16384 px.
    const content = { widthPx: 20000, heightPx: 3000 };
    const at = (dpi: number) =>
      resolveSheet({ ...SHEET_DEFAULTS, paper: "Fit to content", margins: { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 }, dpi }, content);
    expect(fittingDpi(at, 268_435_456)).toBeNull();
    expect(outputRefusal(at(150), 150, 268_435_456, at)).toBe(
      "At 150 dpi the image would be 20000 × 3000 px; either side can be at most 16384 px. It would not fit even at 72 dpi; choose a smaller sheet.",
    );
    // With 2 cm margins, the margins' pixels do shrink with the dpi.
    expect(fittingDpi(framedAt, 268_435_456)).toBe(72);
  });

  it("says so when not even 72 dpi fits", () => {
    const sheet = custom(500, 500);
    expect(fittingDpi(() => sheet, GUARANTEED_CANVAS_PIXELS)).toBeNull();
    expect(outputRefusal(sheet, 72, GUARANTEED_CANVAS_PIXELS)).toContain("would not fit even at 72 dpi; choose a smaller sheet.");
  });
});

describe("the estimate", () => {
  it("states the output's dimensions and approximate bytes", () => {
    expect(outputEstimate({ widthPx: 1754, heightPx: 1240 }, 150, "png", 90, null)).toBe(
      "1754 × 1240 px at 150 dpi · PNG up to about 8.3 MB",
    );
    expect(outputEstimate({ widthPx: 1754, heightPx: 1240 }, 150, "jpeg", 90, null)).toBe(
      "1754 × 1240 px at 150 dpi · JPEG about 1.1 MB",
    );
  });

  it("appends the measured canvas limit once it is known", () => {
    expect(outputEstimate({ widthPx: 100, heightPx: 100 }, 72, "png", 90, 268_435_456)).toContain("· canvas limit 268.4 megapixels");
  });

  it("grows with JPEG quality", () => {
    const pixels = { widthPx: 1000, heightPx: 1000 };
    expect(estimateOutputBytes(pixels, "jpeg", 10)).toBeLessThan(estimateOutputBytes(pixels, "jpeg", 90));
    expect(estimateOutputBytes(pixels, "jpeg", 100)).toBeLessThan(estimateOutputBytes(pixels, "png", 100));
  });
});

describe("backgrounds and names", () => {
  it("keeps a transparent sheet for PNG and falls back to the chosen colour for JPEG", () => {
    expect(outputBackground("transparent", "png", "#ffffff")).toBe("transparent");
    expect(outputBackground("transparent", "jpeg", "#336699")).toBe("#336699");
    expect(outputBackground("#000000", "jpeg", "#336699")).toBe("#000000");
    // A PDF page has no transparency either: white, whatever the matte.
    expect(outputBackground("transparent", "pdf", "#336699")).toBe("#ffffff");
    expect(outputBackground("#000000", "pdf", "#336699")).toBe("#000000");
    expect(jpegBackgroundNotice("transparent", "png")).toBe("");
    expect(jpegBackgroundNotice("#ffffff", "jpeg")).toBe("");
    expect(jpegBackgroundNotice("transparent", "jpeg")).toContain("JPEG has no transparency");
  });

  it("names the file after the title", () => {
    expect(outputFileName("Before and after", "png")).toBe("Before and after.png");
    expect(outputFileName("  ", "jpeg")).toBe(`${DEFAULT_ASSEMBLY_TITLE}.jpeg`);
  });
});

describe("the tray's words", () => {
  it("moves an entry and announces where it went", () => {
    expect(moveEntry(["a", "b", "c"], 0, 1)).toEqual(["b", "a", "c"]);
    expect(moveEntry(["a", "b", "c"], 2, 0)).toEqual(["c", "a", "b"]);
    expect(moveEntry(["a", "b", "c"], 1, 9)).toEqual(["a", "c", "b"]);
    expect(moveEntry(["a"], 3, 0)).toEqual(["a"]);
    expect(trayMoveAnnouncement("shot", 2, 3)).toBe("“shot” moved to position 2 of 3.");
  });

  it("names every dropped item in one line", () => {
    expect(droppedItemsNotice([])).toBe("");
    expect(droppedItemsNotice(["shot"])).toBe("“shot” is no longer in the library and was removed from the tray.");
    expect(droppedItemsNotice(["a", "b"])).toBe("“a”, “b” are no longer in the library and were removed from the tray.");
  });
});

describe("the persisted plan", () => {
  const sheet: SheetSetup = { ...SHEET_DEFAULTS, orientation: "landscape", margins: { topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 } };

  it("round-trips a whole plan", () => {
    const plan = {
      items: [
        { id: "a", title: "shot" },
        { id: "b", title: "photo", transform: { xPercent: 10, yPercent: 20, scalePercent: 50, opacity: 40, blend: "screen" as const } },
      ],
      layout: { layout: "overlay" as const, columns: 2, gapCm: 0.5, cellBackground: "#ff0000" },
      sheet,
      marginsLinked: true,
      output: {
        format: "jpeg" as const,
        quality: 75,
        title: "Pair",
        matte: "#000000",
        // SPEC-110's PDF settings travel with the rest.
        pageMode: "n-up" as const,
        nUp: 6,
        pageNumbers: true,
        author: "Ana",
        subject: "Receipts",
      },
    };
    expect(parsePersistedAssembly(JSON.parse(JSON.stringify(plan)))).toEqual(plan);
  });

  it("reads a plan stored before PDF output with the PDF settings at their defaults (SPEC-110)", () => {
    const stored = { items: [], output: { format: "png", quality: 90, title: "Pair", matte: "#ffffff" } };
    expect(parsePersistedAssembly(stored)?.output).toEqual({ ...ASSEMBLY_OUTPUT_DEFAULTS, title: "Pair" });
    expect(ASSEMBLY_OUTPUT_DEFAULTS).toMatchObject({ pageMode: "one-per-page", nUp: 4, pageNumbers: false, author: "", subject: "" });
    expect(parsePersistedAssembly({ output: { format: "pdf" } })?.output.format).toBe("pdf");
  });

  it("starts a fresh plan from items only — never pixels", () => {
    const fresh = freshAssembly([{ id: "a", title: "shot" }]);
    expect(fresh).toEqual({
      items: [{ id: "a", title: "shot" }],
      layout: ASSEMBLY_LAYOUT_DEFAULTS,
      sheet: null,
      marginsLinked: true,
      output: ASSEMBLY_OUTPUT_DEFAULTS,
    });
    expect(JSON.stringify(fresh)).not.toMatch(/Px|data/);
  });

  it("repairs each bad field to its default and drops bad items, never the whole plan", () => {
    const repaired = parsePersistedAssembly({
      items: [{ id: "a", title: "kept" }, { title: "no id" }, "junk", { id: "b", transform: { xPercent: 500 } }],
      layout: { layout: "spiral", columns: 0, gapCm: -1, cellBackground: "red" },
      sheet: { ...sheet, dpi: 123 },
      output: { format: "webp", quality: 500, title: 7, matte: "white", pageMode: "booklet", nUp: 5, pageNumbers: "yes", author: 1, subject: null },
    });
    expect(repaired).toEqual({
      items: [
        { id: "a", title: "kept" },
        { id: "b", title: "", transform: { xPercent: 100, yPercent: 50, scalePercent: 100, opacity: 100, blend: "normal" } },
      ],
      layout: ASSEMBLY_LAYOUT_DEFAULTS,
      sheet: null,
      marginsLinked: true,
      output: ASSEMBLY_OUTPUT_DEFAULTS,
    });
    expect(parsePersistedAssembly(null)).toBeUndefined();
    expect(parsePersistedAssembly("plan")).toBeUndefined();
  });

  it("keeps at most 100 items", () => {
    const items = Array.from({ length: 140 }, (_, i) => ({ id: `i${i}`, title: `${i}` }));
    expect(parsePersistedAssembly({ items })?.items).toHaveLength(100);
    expect(freshAssembly(items).items).toHaveLength(100);
  });

  it("persists PDF pages with their turns and blank pages, dropping a malformed page item (SPEC-111)", () => {
    const plan = parsePersistedAssembly({
      items: [
        { id: "doc", title: "contract", page: 2, rotateBy: 450 },
        { id: "doc", title: "contract", page: 0 },
        { title: "Blank page", blank: true, rotateBy: -90 },
        { id: "doc", title: "contract", page: -1 },
        { id: "doc", title: "contract", page: 1.5 },
      ],
    });
    expect(plan?.items).toEqual([
      { id: "doc", title: "contract", page: 2, rotateBy: 90 },
      { id: "doc", title: "contract", page: 0 },
      { id: "", title: "Blank page", blank: true, rotateBy: 270 },
    ]);
    expect(freshAssembly(plan?.items ?? [], { format: "pdf", title: "Combined PDF" })).toMatchObject({
      items: plan?.items,
      output: { ...ASSEMBLY_OUTPUT_DEFAULTS, format: "pdf", title: "Combined PDF" },
    });
  });

  it("holds up to 100 images and 2000 entries in all (SPEC-111)", () => {
    const images: PersistedAssemblyItem[] = Array.from({ length: 140 }, (_, i) => ({ id: `i${i}`, title: `${i}` }));
    const pages: PersistedAssemblyItem[] = Array.from({ length: 2100 }, (_, i) => ({ id: "doc", title: "big", page: i }));
    const kept = capTrayItems([...images, ...pages], isPageItem);
    expect(kept.filter((item) => !isPageItem(item))).toHaveLength(100);
    expect(kept).toHaveLength(2000);
    expect(kept[100]).toEqual({ id: "doc", title: "big", page: 0 });
  });

  it("requires a Custom sheet's own size", () => {
    expect(parsePersistedAssembly({ sheet: { ...sheet, paper: "Custom" } })?.sheet).toBeNull();
    expect(parsePersistedAssembly({ sheet: { ...sheet, paper: "Custom", widthCm: 30, heightCm: 20 } })?.sheet).toMatchObject({
      paper: "Custom",
      widthCm: 30,
      heightCm: 20,
    });
  });
});

describe("the PDF output's words (SPEC-110)", () => {
  const estimate = { pageCount: 3, bytes: 1_258_291, images: [], lowest: { index: 1, page: 1, dpi: 95.6 } };

  it("states the page count, the bytes and the lowest-resolution placement", () => {
    expect(pageCountLabel(1)).toBe("1 page");
    expect(pageCountLabel(3)).toBe("3 pages");
    expect(pdfEstimateLine(estimate, ["photo", "scan", "shot"])).toBe(
      "3 pages · PDF up to about 1.2 MB · lowest resolution 96 dpi: “scan” on page 2",
    );
    expect(pdfEstimateLine({ ...estimate, pageCount: 0, lowest: null }, [])).toBe("0 pages · PDF up to about 1.2 MB");
    expect(outputFileName("Receipts", "pdf")).toBe("Receipts.pdf");
  });

  it("warns, before anything is exported, when a placement prints below the sheet's dpi", () => {
    expect(pdfDpiWarning(estimate, ["photo", "scan", "shot"], 300)).toBe(
      "“scan” on page 2 prints at 96 dpi, below the sheet's 300 dpi: it will print soft.",
    );
    expect(pdfDpiWarning(estimate, ["photo", "scan", "shot"], 72)).toBe("");
    expect(pdfDpiWarning({ ...estimate, lowest: { index: 0, page: 0, dpi: 299.7 } }, ["photo"], 300)).toBe("");
    expect(pdfDpiWarning({ ...estimate, lowest: null }, [], 300)).toBe("");
  });
});
