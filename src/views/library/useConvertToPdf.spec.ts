// SPEC-108's Convert to PDF pipeline: the PdfPageInput the writer receives
// for three setups, maxDpi pass-through, the crop applied before encoding,
// SVG rasterised at the placed size, no mutation of the source, byte-
// identical output on a repeat run, and the save/download paths (a new
// "<title>.pdf" document with its parsed summary; failures write nothing).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope } from "vue";
import { equalBytes } from "../../lib/core/equality";
import { parsePdf } from "../../lib/core/pdfDocument";
import { cmToPt } from "../../lib/core/units";
import { FIT_TO_CONTENT, SHEET_DEFAULTS, type SheetSetup } from "../../lib/core/sheetLayout";
import { MARGIN_PRESETS } from "../../lib/preferences";
import { resetSheetSetupSessions } from "../../composables/useSheetSetup";
import { assembleMediaModel } from "../../lib/media/open";
import { buildPngFixtureBytes, buildSvgFixtureBytes } from "../../lib/media/fixtures/build";
import type { MediaModel } from "../../lib/media/model";
import * as store from "../../lib/documentStore";
import {
  convertImageToPdf,
  saveConvertedPdf,
  SHEET_IMAGE_SRC,
  sheetPdfPage,
  useConvertToPdf,
  type ConvertSteps,
} from "./useConvertToPdf";
import { placeOnSheet, resolveSheet } from "../../lib/core/sheetLayout";

vi.mock("../../lib/idb", () => ({
  loadAllRecords: vi.fn().mockResolvedValue([]),
  putRecord: vi.fn().mockResolvedValue(undefined),
  deleteRecord: vi.fn().mockResolvedValue(undefined),
  getRecord: vi.fn().mockResolvedValue(undefined),
}));

const triggerBlobDownload = vi.fn();
vi.mock("../../lib/core/triggerDownload", () => ({
  triggerBlobDownload: (...args: unknown[]) => triggerBlobDownload(...args),
}));

const CROPPED_SRC = "data:image/png;base64,Q1JPUA==";

/** Stand-ins for the canvas-bound steps: a fixed encoded image, a crop that
 * returns a marker URI, and an SVG "raster" at exactly the requested size. */
function fakeSteps(overrides: Partial<ConvertSteps> = {}) {
  const steps = {
    encode: vi.fn(async () => ({
      widthPx: 2,
      heightPx: 1,
      filter: "FlateDecode" as const,
      colorSpace: "DeviceRGB" as const,
      data: new Uint8Array([1, 2, 3, 4]),
    })),
    crop: vi.fn(async () => CROPPED_SRC),
    rasteriseSvg: vi.fn(async (_bytes: Uint8Array, size: { widthPx: number; heightPx: number }) => ({
      data: new Uint8Array([1, 2, 3]),
      widthPx: size.widthPx,
      heightPx: size.heightPx,
    })),
    ...overrides,
  };
  return steps;
}

function setup(patch: Partial<SheetSetup>): SheetSetup {
  return { ...SHEET_DEFAULTS, ...patch, margins: { ...(patch.margins ?? SHEET_DEFAULTS.margins) } };
}

async function pngModel(title = "Shot"): Promise<MediaModel> {
  return assembleMediaModel(title, "png", await buildPngFixtureBytes());
}

const PNG_PX = { widthPx: 64, heightPx: 40 };

beforeEach(() => {
  localStorage.clear();
  resetSheetSetupSessions();
  triggerBlobDownload.mockReset();
});

afterEach(() => {
  for (const entry of store.listDocuments()) {
    store.deleteDocument(entry.id);
  }
});

describe("the page the writer receives", () => {
  it("A4, Narrow, contain: the page is the sheet and the image sits at the placed rectangle", () => {
    const s = setup({ paper: "A4", margins: MARGIN_PRESETS.Narrow, fit: "contain" });
    const geometry = resolveSheet(s, PNG_PX);
    const placement = placeOnSheet(PNG_PX, geometry.content, s.fit, s.align, s.dpi);
    expect(sheetPdfPage(geometry, placement, s.background)).toEqual({
      widthPt: cmToPt(21),
      heightPt: cmToPt(29.7),
      items: [],
      links: [],
      headings: [],
      rects: [{ xPt: 0, yPt: 0, widthPt: cmToPt(21), heightPt: cmToPt(29.7), fill: "#ffffff" }],
      images: [
        { src: SHEET_IMAGE_SRC, xPt: cmToPt(1.27), yPt: cmToPt(9.081), widthPt: cmToPt(18.46), heightPt: cmToPt(11.538) },
      ],
    });
    expect(cmToPt(21)).toBeCloseTo(595.276, 3);
  });

  it("Letter landscape, cover, transparent: the visible box only, on a white page", () => {
    const s = setup({ paper: "Letter", orientation: "landscape", fit: "cover", background: "transparent" });
    const geometry = resolveSheet(s, PNG_PX);
    const placement = placeOnSheet(PNG_PX, geometry.content, s.fit, s.align, s.dpi);
    const page = sheetPdfPage(geometry, placement, s.background);
    expect([page.widthPt, page.heightPt]).toEqual([cmToPt(27.94), cmToPt(21.59)]);
    expect(page.rects).toEqual([{ xPt: 0, yPt: 0, widthPt: cmToPt(27.94), heightPt: cmToPt(21.59), fill: "#ffffff" }]);
    expect(page.images).toEqual([
      { src: SHEET_IMAGE_SRC, xPt: cmToPt(2.54), yPt: cmToPt(2.54), widthPt: cmToPt(22.86), heightPt: cmToPt(16.51) },
    ]);
    expect(placement.crop).toEqual({ xPx: 4.308, yPx: 0, widthPx: 55.385, heightPx: 40 });
  });

  it("Fit to content at 72 dpi with no margins: one image pixel per point", () => {
    const s = setup({ paper: FIT_TO_CONTENT, dpi: 72, margins: { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 }, background: "#102030" });
    const geometry = resolveSheet(s, PNG_PX);
    const placement = placeOnSheet(PNG_PX, geometry.content, s.fit, s.align, s.dpi);
    const page = sheetPdfPage(geometry, placement, s.background);
    expect(page.widthPt).toBeCloseTo(64, 1);
    expect(page.heightPt).toBeCloseTo(40, 1);
    expect(page.images[0]).toMatchObject({ xPt: 0, yPt: 0, widthPt: page.widthPt, heightPt: page.heightPt });
    expect(page.rects?.[0].fill).toBe("#102030");
  });
});

describe("convertImageToPdf", () => {
  it("encodes the uncropped source at the placed size and hands it to buildPdf", async () => {
    const source = await pngModel();
    const steps = fakeSteps();
    const bytes = await convertImageToPdf(source, setup({ margins: MARGIN_PRESETS.Narrow }), "Shot", steps);
    expect(steps.crop).not.toHaveBeenCalled();
    expect(steps.encode).toHaveBeenCalledTimes(1);
    const [src, widthPt, heightPt, maxDpi] = steps.encode.mock.calls[0] as unknown as [string, number, number, number];
    expect(src.startsWith("data:image/png;base64,")).toBe(true);
    expect([widthPt, heightPt, maxDpi]).toEqual([cmToPt(18.46), cmToPt(11.538), 150]);
    const pdf = await parsePdf(bytes);
    expect(pdf.pages).toHaveLength(1);
    expect(pdf.pages[0].widthPt).toBeCloseTo(595.28, 2);
    expect(pdf.pages[0].heightPt).toBeCloseTo(841.89, 2);
  });

  it("passes the setup's dpi through as encodeImage's maxDpi, 600 included", async () => {
    const source = await pngModel();
    const steps = fakeSteps();
    await convertImageToPdf(source, setup({ dpi: 600 }), "Shot", steps);
    expect(steps.encode.mock.calls[0][3]).toBe(600);
  });

  it("crops a cover/actual overflow BEFORE encoding, so the cut-away region never reaches the file", async () => {
    const source = await pngModel();
    const steps = fakeSteps();
    await convertImageToPdf(source, setup({ fit: "cover", align: "left" }), "Shot", steps);
    expect(steps.crop).toHaveBeenCalledTimes(1);
    const [, crop] = steps.crop.mock.calls[0] as unknown as [string, unknown];
    // A4 Normal's 15.92 × 24.62 cm box, covered by a 64 × 40 image: the left
    // edge survives, 25.865 of 64 pixels across.
    expect(crop).toEqual({ xPx: 0, yPx: 0, widthPx: 25.865, heightPx: 40 });
    expect(steps.encode.mock.calls[0][0]).toBe(CROPPED_SRC);
  });

  it("rasterises an SVG at its placed size first", async () => {
    const source = await assembleMediaModel("Mark", "svg", buildSvgFixtureBytes());
    const steps = fakeSteps();
    await convertImageToPdf(source, SHEET_DEFAULTS, "Mark", steps);
    // 96 × 72 px contained in 15.92 × 24.62 cm → 15.92 × 11.94 cm → 940 × 705 px at 150 dpi.
    expect(steps.rasteriseSvg.mock.calls[0][1]).toEqual({ widthPx: 940, heightPx: 705 });
    expect(steps.encode.mock.calls[0][0]).toBe("data:image/png;base64,AQID");
  });

  it("yields byte-identical PDFs for the same image and setup", async () => {
    const source = await pngModel();
    const s = setup({ fit: "cover", margins: MARGIN_PRESETS.Wide, background: "#336699" });
    const first = await convertImageToPdf(source, s, "Shot", fakeSteps());
    const second = await convertImageToPdf(source, s, "Shot", fakeSteps());
    expect(equalBytes(first, second)).toBe(true);
  });

  it("never mutates the source model or its MediaPart", async () => {
    const source = await pngModel();
    const part = source.media[0];
    const before = structuredClone(source);
    await convertImageToPdf(source, setup({ fit: "actual", align: "bottom-right" }), "Shot", fakeSteps());
    await convertImageToPdf(source, setup({ paper: FIT_TO_CONTENT, dpi: 72 }), "Shot", fakeSteps());
    expect(source).toEqual(before);
    expect(source.media[0]).toBe(part);
  });

  it("fails with a reason when the image cannot be encoded", async () => {
    const source = await pngModel();
    await expect(
      convertImageToPdf(source, SHEET_DEFAULTS, "Shot", fakeSteps({ encode: vi.fn(async () => null) })),
    ).rejects.toThrow("could not be read");
  });
});

describe("saving and downloading", () => {
  it("stores a NEW pdf media document titled <title>.pdf, numbered past an existing one", async () => {
    const source = await pngModel("Shot");
    store.storeMedia(source);
    const before = structuredClone(source);
    const bytes = await convertImageToPdf(source, setup({ margins: MARGIN_PRESETS.Narrow }), "Shot", fakeSteps());
    const first = await saveConvertedPdf(bytes, "Shot");
    const second = await saveConvertedPdf(bytes, "Shot");
    expect([first.title, second.title]).toEqual(["Shot.pdf", "Shot.pdf (2)"]);
    expect(first).toMatchObject({ kind: "media", format: "pdf", contentType: "application/pdf", byteLength: bytes.length });
    expect(first.pdf?.pageCount).toBe(1);
    expect(first.pdf?.pages[0].widthPt).toBeCloseTo(595.28, 2);
    expect(store.getMedia(first.id)).toBe(first);
    expect(store.getMedia(source.id)).toEqual(before);
    expect(store.listDocuments()).toHaveLength(3);
  });

  it("useConvertToPdf: saves through the store, and a failure writes nothing", async () => {
    const source = await pngModel("Shot");
    store.storeMedia(source);
    const scope = effectScope();
    const ok = scope.run(() => useConvertToPdf({ sourceId: source.id, steps: fakeSteps() }))!;
    expect(ok.title.value).toBe("Shot");
    expect(ok.canConvert.value).toBe(true);
    const saved = await ok.save();
    expect(saved?.title).toBe("Shot.pdf");
    expect(store.listDocuments()).toHaveLength(2);

    const failing = scope.run(() =>
      useConvertToPdf({ sourceId: source.id, steps: fakeSteps({ encode: vi.fn(async () => null) }) }),
    )!;
    expect(await failing.save()).toBeNull();
    expect(failing.failure.value).toMatch(/^Could not save the PDF: /);
    expect(await failing.download()).toBe(false);
    expect(failing.failure.value).toMatch(/^Could not download the PDF: /);
    expect(triggerBlobDownload).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(2);
    scope.stop();
  });

  it("useConvertToPdf: downloads <title>.pdf as application/pdf without storing anything", async () => {
    const source = await pngModel("Shot");
    store.storeMedia(source);
    const scope = effectScope();
    const conv = scope.run(() => useConvertToPdf({ sourceId: source.id, steps: fakeSteps() }))!;
    conv.title.value = "Form scan";
    expect(await conv.download()).toBe(true);
    expect(triggerBlobDownload).toHaveBeenCalledTimes(1);
    const [bytes, name, mime] = triggerBlobDownload.mock.calls[0] as [Uint8Array, string, string];
    expect([name, mime]).toEqual(["Form scan.pdf", "application/pdf"]);
    expect((await parsePdf(bytes)).pages).toHaveLength(1);
    expect(store.listDocuments()).toHaveLength(1);
    scope.stop();
  });

  it("useConvertToPdf: blocks the actions without a title, or without the source", async () => {
    const source = await pngModel("Shot");
    store.storeMedia(source);
    const scope = effectScope();
    const conv = scope.run(() => useConvertToPdf({ sourceId: source.id, steps: fakeSteps() }))!;
    conv.title.value = "  ";
    expect(conv.canConvert.value).toBe(false);
    expect(conv.blockedReason.value).toBe("Give the PDF a title.");
    const missing = scope.run(() => useConvertToPdf({ sourceId: "nope", steps: fakeSteps() }))!;
    expect(missing.canConvert.value).toBe(false);
    expect(await missing.save()).toBeNull();
    scope.stop();
  });

  it("remembers the sheet for the session", async () => {
    const source = await pngModel("Shot");
    store.storeMedia(source);
    const scope = effectScope();
    const first = scope.run(() => useConvertToPdf({ sourceId: source.id, steps: fakeSteps() }))!;
    first.sheet.setPaper("A3");
    first.sheet.setFit("cover");
    const second = scope.run(() => useConvertToPdf({ sourceId: source.id, steps: fakeSteps() }))!;
    expect(second.sheet.setup.value).toMatchObject({ paper: "A3", fit: "cover" });
    scope.stop();
  });
});
