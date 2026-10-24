// SPEC-110: images to a PDF. Page planning for the three page modes over 1, 5
// and 13 items; the passthrough decision table as a data-driven case list
// (every row, both outcomes — a new source format cannot be added without
// deciding its row); image dedupe across pages, down to a 40-page letterhead
// carrying one copy; the outline tree; page-number geometry; the MAX_*
// refusals; and the twice-identical property, with a passthrough item's
// embedded stream asserted byte-for-byte against the source file's own JPEG
// data. The built files are read back with core/pdfDocument's parser.
import { readFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
import { describe, expect, it, vi } from "vitest";
import {
  ASSEMBLY_PLAN_DEFAULTS,
  planAssembly,
  type AssemblyItem,
  type PlacedItem,
} from "../core/assemblyLayout";
import type { PdfBuildInput, PdfImageData } from "../core/pdfBuild";
import { parsePdf, pdfTextOf, resolve, type PdfDocument, type PdfValue } from "../core/pdfDocument";
import { parseJpegHeader } from "../core/pdfImages";
import type { FaceKey } from "../core/pdfFonts";
import type { MediaPart } from "../core/media";
import { parseSfnt, type SfntFont } from "../core/sfnt";
import { equalBytes } from "../core/equality";
import { cmToPt } from "../core/units";
import { resolveSheet, SHEET_DEFAULTS, type SheetGeometry } from "../core/sheetLayout";
import { bytesToBase64, buildJpegFixtureBytes, buildPngFixtureBytes } from "./fixtures/build";
import {
  buildImagePdf,
  coversWholeSource,
  drawPageNumber,
  encodingLabel,
  encodingReason,
  estimateImagePdf,
  ImagePdfCancelled,
  imageKey,
  itemsPerPage,
  MAX_PDF_ASSEMBLY_BYTES,
  MAX_PDF_ASSEMBLY_PAGES,
  N_UP_CHOICES,
  nUpColumns,
  PAGE_NUMBER_MIN_MARGIN_CM,
  pageNumberItem,
  pageNumbersRefusal,
  pdfLimitRefusal,
  pdfNaturalPx,
  pdfPageCount,
  placementDpi,
  planPdfPages,
  rasterisePlacement,
  rasterSize,
  reencodedBytes,
  sourceFacts,
  textWidthPt,
  type EncodingReason,
  type ImagePdfItem,
  type ImagePdfOptions,
  type PageMode,
  type PdfAssemblyPlan,
  type RasterJob,
  type SourceFacts,
} from "./imagesToPdf";

const sans: SfntFont = parseSfnt(
  new Uint8Array(readFileSync(resolvePath(__dirname, "../../../public/fonts/LiberationSans-Regular.ttf"))),
);
const fonts = () => new Map<FaceKey, SfntFont>([["sans-regular", sans]]);

/** A4 portrait, Normal margins: a 15.92 × 24.62 cm content box. */
const A4: SheetGeometry = resolveSheet(SHEET_DEFAULTS);
const MODIFIED = new Date(Date.UTC(2026, 9, 24, 12, 0, 0));

function plan(pageMode: PageMode, overrides: Partial<PdfAssemblyPlan> = {}): PdfAssemblyPlan {
  return { ...ASSEMBLY_PLAN_DEFAULTS, pageMode, nUp: 4, ...overrides };
}

function trayOf(count: number, size = { widthPx: 160, heightPx: 120 }): AssemblyItem[] {
  return Array.from({ length: count }, (_, index) => ({ id: `item-${index}`, ...size }));
}

function part(mediaId: string, contentType: string, bytes: Uint8Array): MediaPart {
  return { mediaId, contentType, data: bytesToBase64(bytes) };
}

/** A JPEG that is only a frame header: SOI, one SOFn segment, EOI. */
function jpegHeaderBytes(marker: number, components: number, widthPx = 160, heightPx = 120): Uint8Array {
  const length = 8 + components * 3;
  const segment = [0xff, marker, length >> 8, length & 0xff, 8, heightPx >> 8, heightPx & 0xff, widthPx >> 8, widthPx & 0xff, components];
  for (let c = 0; c < components; c += 1) {
    segment.push(c + 1, 0x11, 0);
  }
  return new Uint8Array([0xff, 0xd8, ...segment, 0xff, 0xd9]);
}

/** The file's image XObjects (masks excluded), in object order. */
function imageStreams(doc: PdfDocument): Array<{ filter: string; raw: Uint8Array; width: number }> {
  const out: Array<{ filter: string; raw: Uint8Array; width: number }> = [];
  for (const value of doc.objects.values()) {
    if (value.t !== "stream" || resolve(doc, value.dict.get("Subtype") ?? { t: "null" }).t !== "name") {
      continue;
    }
    const subtype = value.dict.get("Subtype") as { t: "name"; v: string };
    const isMask = [...doc.objects.values()].some(
      (other) => other.t === "stream" && other !== value && other.dict.get("SMask")?.t === "ref" &&
        doc.objects.get((other.dict.get("SMask") as { num: number }).num) === value,
    );
    if (subtype.v === "Image" && !isMask) {
      const filter = value.dict.get("Filter") as { t: "name"; v: string };
      const width = value.dict.get("Width") as { t: "num"; v: number };
      out.push({ filter: filter.v, raw: value.raw, width: width.v });
    }
  }
  return out;
}

function dict(doc: PdfDocument, value: PdfValue | undefined): Map<string, PdfValue> {
  const resolved = resolve(doc, value ?? { t: "null" });
  if (resolved.t !== "dict") {
    throw new Error(`not a dictionary: ${resolved.t}`);
  }
  return resolved.v;
}

/** The /Outlines entries in order: title, level (1 = a root child) and the
 * page index its destination names. */
function outline(doc: PdfDocument): Array<{ title: string; page: number; depth: number }> {
  const catalog = dict(doc, doc.trailer.get("Root"));
  const root = dict(doc, catalog.get("Outlines"));
  const pageNums = doc.pages.map((page) => page.ref.num);
  const entries: Array<{ title: string; page: number; depth: number }> = [];
  const walk = (first: PdfValue | undefined, depth: number) => {
    let node = first;
    while (node && node.t === "ref") {
      const entry = dict(doc, node);
      const dest = resolve(doc, entry.get("Dest") ?? { t: "null" });
      const target = dest.t === "array" ? dest.v[0] : undefined;
      entries.push({
        title: pdfTextOf(entry.get("Title")),
        page: target && target.t === "ref" ? pageNums.indexOf(target.num) : -1,
        depth,
      });
      walk(entry.get("First"), depth + 1);
      node = entry.get("Next");
    }
  };
  walk(root.get("First"), 1);
  return entries;
}

function latin1(bytes: Uint8Array): string {
  let text = "";
  for (const byte of bytes) {
    text += String.fromCharCode(byte);
  }
  return text;
}

/** A deterministic stand-in for the canvas path: a tiny Flate image whose
 * samples name the job's size. */
function fakeRasterise(calls: RasterJob[] = []) {
  return vi.fn(async (job: RasterJob): Promise<PdfImageData> => {
    calls.push(job);
    return {
      widthPx: job.widthPx,
      heightPx: job.heightPx,
      filter: "FlateDecode",
      colorSpace: "DeviceRGB",
      data: new Uint8Array([0x78, 0x9c, job.widthPx & 0xff, job.heightPx & 0xff]),
    };
  });
}

async function sources() {
  const jpeg = buildJpegFixtureBytes();
  const png = await buildPngFixtureBytes();
  return {
    jpeg,
    png,
    jpegPart: part("m-jpeg", "image/jpeg", jpeg),
    pngPart: part("m-png", "image/png", png),
  };
}

function options(items: readonly ImagePdfItem[], overrides: Partial<ImagePdfOptions> = {}): ImagePdfOptions {
  return {
    title: "Receipts",
    pageNumbers: false,
    modified: MODIFIED,
    items,
    sheet: { geometry: A4, background: "#ffffff", dpi: 150 },
    fonts: fonts(),
    steps: { rasterise: fakeRasterise(), between: async () => {} },
    ...overrides,
  };
}

// --- page planning -------------------------------------------------------------

describe("planPdfPages", () => {
  const counts = [1, 5, 13];

  it.each(counts)("one-per-page puts each of %i items alone on its page, fitted to the whole content box", (count) => {
    const items = trayOf(count);
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    expect(pages).toHaveLength(count);
    expect(pages.map((page) => page.itemIndices)).toEqual(items.map((_, index) => [index]));
    // Its cell IS the content box: exactly SPEC-109's single-item plan.
    const alone = planAssembly([items[0]], { ...ASSEMBLY_PLAN_DEFAULTS, columns: 1 }, A4)[0];
    for (const [index, page] of pages.entries()) {
      expect(page.placed).toEqual([{ ...alone, index }]);
      expect(page.placed[0].cell).toEqual(A4.content);
    }
  });

  it.each(counts)("n-up 4 puts %i items four to a page in SPEC-109's 2 × 2 grid", (count) => {
    const items = trayOf(count);
    const pages = planPdfPages(items, plan("n-up", { nUp: 4 }), A4);
    expect(pages).toHaveLength(Math.ceil(count / 4));
    expect(pages.flatMap((page) => page.itemIndices)).toEqual(items.map((_, index) => index));
    const full = planAssembly(trayOf(4), { ...ASSEMBLY_PLAN_DEFAULTS, columns: 2 }, A4);
    for (const page of pages) {
      // Every page, the short last one too, uses the same four cells.
      expect(page.placed.map((item) => item.cell)).toEqual(full.slice(0, page.placed.length).map((item) => item.cell));
      expect(page.placed.map((item) => item.index)).toEqual(page.itemIndices);
    }
  });

  it.each(counts)("single-page puts all %i items on one sheet — SPEC-109's own plan", (count) => {
    const items = trayOf(count);
    const pages = planPdfPages(items, plan("single-page"), A4);
    expect(pages).toHaveLength(1);
    expect(pages[0].itemIndices).toEqual(items.map((_, index) => index));
    expect(pages[0].placed).toEqual(planAssembly(items, ASSEMBLY_PLAN_DEFAULTS, A4));
  });

  it("counts pages and items per page for every mode", () => {
    expect([1, 5, 13].map((n) => pdfPageCount(plan("one-per-page"), n))).toEqual([1, 5, 13]);
    expect([1, 5, 13].map((n) => pdfPageCount(plan("n-up", { nUp: 4 }), n))).toEqual([1, 2, 4]);
    expect([1, 5, 13].map((n) => pdfPageCount(plan("single-page"), n))).toEqual([1, 1, 1]);
    expect(pdfPageCount(plan("one-per-page"), 0)).toBe(0);
    expect(itemsPerPage(plan("n-up", { nUp: 7 }), 13)).toBe(4);
    expect(planPdfPages([], plan("one-per-page"), A4)).toEqual([]);
  });

  it("lays n-up out as a full grid of exactly n cells, turned to the sheet", () => {
    const portrait = A4.content;
    const landscape = { widthCm: portrait.heightCm, heightCm: portrait.widthCm };
    expect(N_UP_CHOICES.map((n) => [n, nUpColumns(n, portrait)])).toEqual([[2, 1], [4, 2], [6, 2], [9, 3], [16, 4]]);
    expect(N_UP_CHOICES.map((n) => [n, nUpColumns(n, landscape)])).toEqual([[2, 2], [4, 2], [6, 3], [9, 3], [16, 4]]);
    for (const n of N_UP_CHOICES) {
      const [page] = planPdfPages(trayOf(n), plan("n-up", { nUp: n }), A4);
      const columns = new Set(page.placed.map((item) => item.cell.xCm)).size;
      const rows = new Set(page.placed.map((item) => item.cell.yCm)).size;
      expect(columns * rows).toBe(n);
    }
  });

  it("keeps the tray order across every mode, and ignores overlay transforms outside single-page", () => {
    const items = trayOf(5).map((item) => ({
      ...item,
      transform: { xPercent: 10, yPercent: 10, scalePercent: 50, opacity: 40, blend: "multiply" as const },
    }));
    for (const mode of ["one-per-page", "n-up"] as const) {
      const pages = planPdfPages(items, plan(mode, { layout: "overlay" }), A4);
      expect(pages.flatMap((page) => page.placed.map((item) => item.index))).toEqual([0, 1, 2, 3, 4]);
      expect(pages.flatMap((page) => page.placed.map((item) => [item.opacity, item.blend]))).toEqual(
        items.map(() => [1, "normal"]),
      );
    }
    const [single] = planPdfPages(items, plan("single-page", { layout: "overlay" }), A4);
    expect(single.placed.map((item) => item.opacity)).toEqual([0.4, 0.4, 0.4, 0.4, 0.4]);
  });

  it("sizes 'Fit to content' to one page's arrangement", () => {
    const items = [
      { id: "a", widthPx: 100, heightPx: 50 },
      { id: "b", widthPx: 40, heightPx: 80 },
    ];
    expect(pdfNaturalPx(items, plan("one-per-page"))).toEqual({ widthPx: 100, heightPx: 80 });
    expect(pdfNaturalPx(items, plan("n-up", { nUp: 4 }))).toEqual({ widthPx: 200, heightPx: 160 });
    expect(pdfNaturalPx(items, plan("single-page", { columns: 2 }))).toEqual({ widthPx: 200, heightPx: 80 });
    expect(pdfNaturalPx([], plan("one-per-page"))).toBeNull();
  });
});

// --- the passthrough decision ---------------------------------------------------

describe("the passthrough decision", () => {
  const baseline = jpegHeaderBytes(0xc0, 3);
  const whole = { x: 0, y: 0, width: 160, height: 120 };
  const item = { widthPx: 160, heightPx: 120 };

  interface Row {
    row: string;
    contentType?: string;
    bytes?: Uint8Array;
    sourcePx?: PlacedItem["sourcePx"];
    opacity?: number;
    blend?: PlacedItem["blend"];
    edits?: ImagePdfItem["edits"];
    expected: EncodingReason;
  }

  // Every row of the table, with the case that passes it and the case that
  // fails it. A source format not listed here has no decided row.
  const rows: Row[] = [
    { row: "baseline JPEG, 3 components", bytes: baseline, expected: "original" },
    { row: "baseline JPEG, 1 component", bytes: jpegHeaderBytes(0xc0, 1), expected: "original" },
    { row: "extended sequential JPEG (SOF1)", bytes: jpegHeaderBytes(0xc1, 3), expected: "original" },
    { row: "progressive JPEG", bytes: jpegHeaderBytes(0xc2, 3), expected: "not-baseline" },
    { row: "arithmetic-coded JPEG", bytes: jpegHeaderBytes(0xc9, 3), expected: "not-baseline" },
    { row: "CMYK JPEG", bytes: jpegHeaderBytes(0xc0, 4), expected: "components" },
    { row: "an unreadable JPEG", bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]), expected: "not-jpeg" },
    { row: "PNG", contentType: "image/png", expected: "not-jpeg" },
    { row: "GIF", contentType: "image/gif", expected: "not-jpeg" },
    { row: "WEBP", contentType: "image/webp", expected: "not-jpeg" },
    { row: "SVG", contentType: "image/svg+xml", expected: "not-jpeg" },
    { row: "sourcePx covers the whole image", sourcePx: whole, expected: "original" },
    { row: "sourcePx is a cover crop", sourcePx: { x: 20, y: 0, width: 120, height: 120 }, expected: "cropped" },
    { row: "sourcePx is an actual-size crop", sourcePx: { x: 0, y: 0, width: 160, height: 90 }, expected: "cropped" },
    { row: "opacity 1", opacity: 1, expected: "original" },
    { row: "opacity < 1", opacity: 0.5, expected: "opacity" },
    { row: "blend normal", blend: "normal", expected: "original" },
    { row: "blend multiply", blend: "multiply", expected: "blend" },
    { row: "blend screen", blend: "screen", expected: "blend" },
    { row: "no edit stack (field absent)", edits: undefined, expected: "original" },
    { row: "an empty edit stack", edits: [], expected: "original" },
    { row: "a non-empty edit stack", edits: [{ kind: "rotate", quarterTurns: 1 }], expected: "edited" },
  ];

  it.each(rows)("$row → $expected", (row) => {
    const contentType = row.contentType ?? "image/jpeg";
    const bytes = row.bytes ?? (contentType === "image/jpeg" ? baseline : new Uint8Array([1, 2, 3]));
    const facts = sourceFacts(part("m", contentType, bytes));
    const placed = { sourcePx: row.sourcePx ?? whole, opacity: row.opacity ?? 1, blend: row.blend ?? "normal" };
    expect(encodingReason(facts, { ...item, edits: row.edits }, placed)).toBe(row.expected);
  });

  it("reads the fits through the plan: contain and stretch never crop; cover and actual crop only when they cut", () => {
    const facts = sourceFacts(part("m", "image/jpeg", baseline));
    const reasonFor = (fit: "contain" | "cover" | "stretch" | "actual", sheet: SheetGeometry, dpi = 150) => {
      const [page] = planPdfPages([{ id: "a", ...item }], plan("one-per-page", { fit }), sheet, dpi);
      return encodingReason(facts, item, page.placed[0]);
    };
    // A box of the image's own aspect (4 : 3) at a size where "actual" fits.
    const sameAspect: SheetGeometry = { widthCm: 40, heightCm: 30, content: { xCm: 0, yCm: 0, widthCm: 40, heightCm: 30 } };
    expect(reasonFor("contain", A4)).toBe("original");
    expect(reasonFor("stretch", A4)).toBe("original");
    expect(reasonFor("cover", A4)).toBe("cropped");
    expect(reasonFor("cover", sameAspect)).toBe("original");
    expect(reasonFor("actual", A4)).toBe("original");
    // 160 px at 10 dpi is 40.6 cm: wider than A4's content box.
    expect(reasonFor("actual", A4, 10)).toBe("cropped");
    expect(coversWholeSource({ x: 0, y: 0, width: 159.995, height: 120 }, item)).toBe(true);
  });

  it("says in the tray which happened", () => {
    expect(encodingLabel("original", "image/jpeg")).toBe("original JPEG");
    expect(encodingLabel("cropped", "image/jpeg")).toBe("re-encoded: cropped");
    expect(encodingLabel("not-jpeg", "image/png")).toBe("re-encoded: PNG");
    expect(encodingLabel("not-jpeg", "image/jpeg")).toBe("re-encoded: unsupported JPEG");
    expect(encodingLabel("not-baseline", "image/jpeg")).toBe("re-encoded: progressive JPEG");
    expect(encodingLabel("components", "image/jpeg")).toBe("re-encoded: CMYK JPEG");
    expect(encodingLabel("edited", "image/png")).toBe("re-encoded: edited");
    expect(encodingLabel("opacity", "image/jpeg")).toBe("re-encoded: opacity");
    expect(encodingLabel("blend", "image/jpeg")).toBe("re-encoded: blend mode");
  });

  it("re-encodes at the sheet's dpi, never above the image's own pixels", () => {
    const [small] = planPdfPages([{ id: "a", widthPx: 64, heightPx: 40 }], plan("one-per-page"), A4, 600);
    // 64 × 40 px on 15.92 cm: far below 600 dpi — embedded at its own size.
    expect(rasterSize(small.placed[0], "image/png", 600)).toEqual({ widthPx: 64, heightPx: 40 });
    expect(Math.round(placementDpi(small.placed[0], "image/png", "not-jpeg", 600))).toBe(10);
    const [large] = planPdfPages([{ id: "b", widthPx: 8000, heightPx: 6000 }], plan("one-per-page"), A4, 600);
    // 15.92 cm at 600 dpi is 3761 px: sampled down to the sheet's dpi, not 300.
    expect(rasterSize(large.placed[0], "image/png", 600)).toEqual({ widthPx: 3761, heightPx: 2820 });
    expect(Math.round(placementDpi(large.placed[0], "image/png", "not-jpeg", 600))).toBe(600);
    // A passed-through JPEG keeps its own resolution, above the sheet's.
    expect(Math.round(placementDpi(large.placed[0], "image/jpeg", "original", 300))).toBe(1276);
    // An SVG has no pixels of its own: drawn at the placed size.
    expect(rasterSize(small.placed[0], "image/svg+xml", 300)).toEqual({ widthPx: 1880, heightPx: 1175 });
  });
});

// --- building -------------------------------------------------------------------

describe("buildImagePdf", () => {
  it("embeds a passthrough JPEG's own bytes verbatim as /DCTDecode", async () => {
    const { jpeg, jpegPart } = await sources();
    const items = [{ id: "photo", title: "photo", widthPx: 160, heightPx: 120 }];
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    const rasterise = fakeRasterise();
    const bytes = await buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, { steps: { rasterise } }));
    const doc = await parsePdf(bytes);
    const streams = imageStreams(doc);
    expect(streams).toHaveLength(1);
    expect(streams[0].filter).toBe("DCTDecode");
    expect(equalBytes(streams[0].raw, jpeg)).toBe(true);
    expect(rasterise).not.toHaveBeenCalled();
    expect(parseJpegHeader(streams[0].raw)?.baseline).toBe(true);
  });

  it("composes each page from a background rect, the placed images and the page number, and calls the writer once", async () => {
    const { jpegPart, pngPart } = await sources();
    const items = [
      { id: "photo", title: "photo", widthPx: 160, heightPx: 120 },
      { id: "shot", title: "shot", widthPx: 64, heightPx: 40 },
    ];
    const writes: PdfBuildInput[] = [];
    const buildPdf = vi.fn(async (input: PdfBuildInput) => {
      writes.push(input);
      return new Uint8Array([1]);
    });
    const pages = planPdfPages(items, plan("one-per-page", { cellBackground: "#ff0000" }), A4);
    await buildImagePdf(
      pages,
      new Map([["photo", jpegPart], ["shot", pngPart]]),
      options(items, {
        pageNumbers: true,
        sheet: { geometry: A4, background: "transparent", dpi: 300 },
        steps: { rasterise: fakeRasterise(), between: async () => {}, buildPdf },
      }),
    );
    expect(buildPdf).toHaveBeenCalledTimes(1);
    const [input] = writes;
    expect(input.pages).toHaveLength(2);
    for (const [index, page] of input.pages.entries()) {
      expect(page.widthPt).toBeCloseTo(595.28, 1);
      expect(page.heightPt).toBeCloseTo(841.89, 1);
      // The sheet background, white where "transparent" was asked for, then
      // the cell's colour behind its image.
      expect(page.rects?.[0]).toEqual({ xPt: 0, yPt: 0, widthPt: page.widthPt, heightPt: page.heightPt, fill: "#ffffff" });
      expect(page.rects?.[1]?.fill).toBe("#ff0000");
      expect(page.images).toHaveLength(1);
      const placed = pages[index].placed[0];
      expect(page.images[0]).toMatchObject({
        xPt: cmToPt(placed.xCm),
        yPt: cmToPt(placed.yCm),
        widthPt: cmToPt(placed.widthCm),
        heightPt: cmToPt(placed.heightCm),
      });
      expect(page.items.map((item) => item.text)).toEqual([String(index + 1)]);
    }
    expect([...(input.images?.keys() ?? [])]).toEqual([input.pages[0].images[0].src, input.pages[1].images[0].src]);
    expect(input.images?.get(input.pages[0].images[0].src)?.filter).toBe("DCTDecode");
    expect(input.images?.get(input.pages[1].images[0].src)?.filter).toBe("FlateDecode");
    // The one bundled face, nothing new.
    expect([...input.fonts.keys()]).toEqual(["sans-regular"]);
  });

  it("re-encodes through the canvas path at the sheet's dpi, never upsampling", async () => {
    const { pngPart, jpegPart } = await sources();
    const items = [
      { id: "shot", title: "shot", widthPx: 64, heightPx: 40 },
      { id: "photo", title: "photo", widthPx: 160, heightPx: 120 },
    ];
    const jobs: RasterJob[] = [];
    const pages = planPdfPages(items, plan("one-per-page", { fit: "cover" }), A4, 600);
    await buildImagePdf(
      pages,
      new Map([["shot", pngPart], ["photo", jpegPart]]),
      options(items, { sheet: { geometry: A4, background: "#ffffff", dpi: 600 }, steps: { rasterise: fakeRasterise(jobs), between: async () => {} } }),
    );
    // The PNG, and the cover-cropped JPEG: both drawn, at 600 dpi at most.
    expect(jobs).toHaveLength(2);
    expect(jobs.map((job) => job.dpi)).toEqual([600, 600]);
    expect(jobs.map((job) => [job.widthPx, job.heightPx])).toEqual([
      [Math.round(pages[0].placed[0].sourcePx.width), 40],
      [Math.round(pages[1].placed[0].sourcePx.width), 120],
    ]);
    expect(jobs.every((job) => job.layers.length === 1 && job.background === "transparent")).toBe(true);
  });

  it("flattens a blended placement with what lies beneath it on its page", async () => {
    const { jpegPart } = await sources();
    const items = [
      { id: "photo", title: "under", widthPx: 160, heightPx: 120 },
      { id: "photo", title: "over", widthPx: 160, heightPx: 120 },
    ];
    const tray = items.map((item, index) => ({
      ...item,
      transform: { xPercent: 50, yPercent: 50, scalePercent: index === 0 ? 100 : 50, opacity: 100, blend: index === 0 ? ("normal" as const) : ("multiply" as const) },
    }));
    const jobs: RasterJob[] = [];
    const pages = planPdfPages(tray, plan("single-page", { layout: "overlay" }), A4);
    await buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, { steps: { rasterise: fakeRasterise(jobs), between: async () => {} } }));
    expect(jobs).toHaveLength(1);
    expect(jobs[0].layers.map((layer) => layer.index)).toEqual([0, 1]);
    expect(jobs[0].background).toBe("#ffffff");
    // At the sheet's dpi, not the blended image's own 160 px: what lies
    // beneath must not print at the overlay's resolution.
    const over = pages[0].placed[1];
    expect(jobs[0].widthPx).toBe(Math.round((over.widthCm / 2.54) * 150));
    expect(jobs[0].widthPx).toBeGreaterThan(160);
    expect(rasterSize(over, "image/jpeg", 150)).toMatchObject({ widthPx: 160 });
  });

  it("embeds the same source once however many pages place it: a 40-page letterhead carries one copy", async () => {
    const { jpeg, jpegPart, pngPart } = await sources();
    for (const [source, contentType] of [[jpegPart, "image/jpeg"], [pngPart, "image/png"]] as const) {
      const items = Array.from({ length: 40 }, (_, index) => ({
        id: "letterhead",
        title: `letterhead ${index + 1}`,
        widthPx: contentType === "image/jpeg" ? 160 : 64,
        heightPx: contentType === "image/jpeg" ? 120 : 40,
      }));
      const rasterise = fakeRasterise();
      const pages = planPdfPages(items, plan("one-per-page"), A4);
      const bytes = await buildImagePdf(pages, new Map([["letterhead", source]]), options(items, { steps: { rasterise, between: async () => {} } }));
      const doc = await parsePdf(bytes);
      expect(doc.pages).toHaveLength(40);
      const streams = imageStreams(doc);
      expect(streams).toHaveLength(1);
      if (contentType === "image/jpeg") {
        expect(equalBytes(streams[0].raw, jpeg)).toBe(true);
        expect(rasterise).not.toHaveBeenCalled();
        // One copy of the JPEG in the whole file.
        expect(latin1(bytes).split(latin1(jpeg)).length - 1).toBe(1);
      } else {
        expect(rasterise).toHaveBeenCalledTimes(1);
      }
    }
  });

  it("writes one level-1 outline entry per item, titled by its source, pointing at its page", async () => {
    const { jpegPart } = await sources();
    const items = ["contract p1", "contract p2", "contract p3", "passport", "utility bill"].map((title) => ({
      id: "photo",
      title,
      widthPx: 160,
      heightPx: 120,
    }));
    for (const [mode, expectedPages] of [["one-per-page", [0, 1, 2, 3, 4]], ["n-up", [0, 0, 0, 0, 1]], ["single-page", [0, 0, 0, 0, 0]]] as const) {
      const pages = planPdfPages(items, plan(mode, { nUp: 4 }), A4);
      const doc = await parsePdf(await buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items)));
      expect(outline(doc)).toEqual(items.map((item, index) => ({ title: item.title, page: expectedPages[index], depth: 1 })));
    }
  });

  it("fills the metadata: Title, Author and Subject as given, Creator and Producer naming the application", async () => {
    const { jpegPart } = await sources();
    const items = [{ id: "photo", title: "photo", widthPx: 160, heightPx: 120 }];
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    const map = new Map([["photo", jpegPart]]);
    const filled = await parsePdf(
      await buildImagePdf(pages, map, options(items, { title: "Expense claim", author: " Ana ", subject: "May receipts" })),
    );
    const info = filled.info!;
    expect(["Title", "Author", "Subject", "Creator", "Producer"].map((key) => pdfTextOf(info.get(key)))).toEqual([
      "Expense claim",
      "Ana",
      "May receipts",
      "docu-edit",
      "docu-edit",
    ]);
    expect(pdfTextOf(info.get("ModDate"))).toBe("D:20261024120000Z");
    // Empty by default: never filled in from anywhere else.
    const empty = await parsePdf(await buildImagePdf(pages, map, options(items, { author: "", subject: "  " })));
    expect(empty.info?.has("Author")).toBe(false);
    expect(empty.info?.has("Subject")).toBe(false);
  });

  it("is deterministic: the same plan and sources twice give equalBytes-identical files", async () => {
    const { jpegPart, pngPart } = await sources();
    const items = [
      { id: "photo", title: "photo", widthPx: 160, heightPx: 120 },
      { id: "shot", title: "shot", widthPx: 64, heightPx: 40 },
      { id: "photo", title: "photo again", widthPx: 160, heightPx: 120 },
    ];
    const pages = planPdfPages(items, plan("n-up", { nUp: 2 }), A4);
    const build = () =>
      buildImagePdf(pages, new Map([["photo", jpegPart], ["shot", pngPart]]), options(items, { pageNumbers: true, author: "Ana" }));
    const [first, second] = [await build(), await build()];
    expect(first.length).toBeGreaterThan(0);
    expect(equalBytes(first, second)).toBe(true);
  });

  it("composes pages one at a time, yielding and reporting between them, and stops at a page boundary on cancel", async () => {
    const { jpegPart } = await sources();
    const items = trayOf(3).map((item) => ({ ...item, id: "photo", title: item.id }));
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    const between = vi.fn(async () => {});
    const progress: string[] = [];
    await buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, {
      steps: { rasterise: fakeRasterise(), between },
      onProgress: (done, total) => progress.push(`${done}/${total}`),
    }));
    expect(between).toHaveBeenCalledTimes(2);
    expect(progress).toEqual(["0/3", "1/3", "2/3", "3/3"]);

    const buildPdf = vi.fn(async () => new Uint8Array([1]));
    let calls = 0;
    await expect(
      buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, {
        steps: { rasterise: fakeRasterise(), between, buildPdf },
        cancelled: () => (calls += 1) > 1,
      })),
    ).rejects.toBeInstanceOf(ImagePdfCancelled);
    expect(buildPdf).not.toHaveBeenCalled();
  });

  it("fails loudly, naming the image, when a source is missing or cannot be drawn", async () => {
    const { pngPart } = await sources();
    const items = [{ id: "shot", title: "shot", widthPx: 64, heightPx: 40 }];
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    await expect(buildImagePdf(pages, new Map(), options(items))).rejects.toThrow("“shot” could not be read.");
    await expect(
      buildImagePdf(pages, new Map([["shot", pngPart]]), options(items, { steps: { rasterise: async () => null, between: async () => {} } })),
    ).rejects.toThrow("“shot” could not be drawn.");
  });

  it("never touches its sources", async () => {
    const { jpegPart, pngPart } = await sources();
    const frozen = [jpegPart, pngPart].map((source) => Object.freeze({ ...source }));
    const items = [
      { id: "photo", title: "photo", widthPx: 160, heightPx: 120 },
      { id: "shot", title: "shot", widthPx: 64, heightPx: 40 },
    ];
    const map = new Map([["photo", frozen[0]], ["shot", frozen[1]]]);
    const pages = planPdfPages(items, plan("n-up", { nUp: 2 }), A4);
    await buildImagePdf(pages, map, options(Object.freeze(items.map((item) => Object.freeze(item)))));
    expect(map.get("photo")).toEqual(jpegPart);
    expect(map.get("shot")).toEqual(pngPart);
  });
});

// --- page numbers ---------------------------------------------------------------

describe("page numbers", () => {
  it("are 9 pt, centred, 1 cm above the bottom edge, #555555, in the bundled sans face", () => {
    const item = pageNumberItem(12, A4, sans);
    expect(item.text).toBe("12");
    expect(item.sizePt).toBe(9);
    expect(item.face).toBe("sans-regular");
    expect(item.color).toEqual([0x55 / 255, 0x55 / 255, 0x55 / 255]);
    expect(item.yPt).toBeCloseTo(cmToPt(29.7) - cmToPt(1), 6);
    const width = textWidthPt("12", sans, 9);
    expect(width).toBeGreaterThan(0);
    expect(item.xPt + width / 2).toBeCloseTo(cmToPt(21) / 2, 6);
  });

  it("sit inside the bottom margin, never over a placed image", () => {
    const pages = planPdfPages(trayOf(4), plan("n-up", { nUp: 4, fit: "stretch" }), A4);
    const lowestImage = Math.max(...pages[0].placed.map((item) => cmToPt(item.yCm + item.heightCm)));
    const item = pageNumberItem(1, A4, sans);
    // The digits' tops are below the content box's bottom edge.
    expect(item.yPt - (sans.capHeight / sans.unitsPerEm) * item.sizePt).toBeGreaterThan(lowestImage);
    expect(pageNumbersRefusal(A4)).toBeNull();
    // Narrow margins (1.27 cm) still hold it.
    expect(pageNumbersRefusal(resolveSheet({ ...SHEET_DEFAULTS, margins: { topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 } }))).toBeNull();
    const tight = resolveSheet({ ...SHEET_DEFAULTS, margins: { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 } });
    expect(pageNumbersRefusal(tight)).toBe(
      `Page numbers sit 1 cm above the bottom edge, inside the margin: make the bottom margin at least ${PAGE_NUMBER_MIN_MARGIN_CM} cm, or turn page numbers off.`,
    );
    expect(1 + (sans.capHeight / sans.unitsPerEm) * (9 / 72) * 2.54).toBeLessThan(PAGE_NUMBER_MIN_MARGIN_CM);
  });

  it("are drawn on the preview where the file has them: centred, 1 cm up, 9 pt, #555555, the bundled face", () => {
    const calls: unknown[][] = [];
    const ctx = {
      save() {},
      restore() {},
      setTransform() {},
      fillStyle: "",
      font: "",
      textAlign: "",
      textBaseline: "",
      fillText: (...args: unknown[]) => calls.push([ctx.fillStyle, ctx.font, ctx.textAlign, ctx.textBaseline, ...args]),
    };
    // A4 at 10 px per cm.
    const canvas = { width: 210, height: 297, getContext: () => ctx } as unknown as HTMLCanvasElement;
    drawPageNumber(canvas, A4, 3);
    expect(calls).toHaveLength(1);
    const [fill, font, align, baseline, text, x, y] = calls[0] as [string, string, string, string, string, number, number];
    expect([fill, align, baseline, text]).toEqual(["#555555", "center", "alphabetic", "3"]);
    expect(font).toBe('3.18px "Liberation Sans", sans-serif');
    expect(x).toBeCloseTo(105, 6);
    expect(y).toBeCloseTo(287, 6);
  });

  it("number the produced pages from 1, and are absent when off", async () => {
    const { jpegPart } = await sources();
    const items = trayOf(5).map((item) => ({ ...item, id: "photo", title: item.id }));
    const writes: PdfBuildInput[] = [];
    const buildPdf = async (input: PdfBuildInput) => {
      writes.push(input);
      return new Uint8Array([1]);
    };
    const pages = planPdfPages(items, plan("n-up", { nUp: 2 }), A4);
    const steps = { rasterise: fakeRasterise(), between: async () => {}, buildPdf };
    await buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, { pageNumbers: true, steps }));
    await buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, { pageNumbers: false, steps }));
    expect(writes[0].pages.map((page) => page.items.map((item) => item.text))).toEqual([["1"], ["2"], ["3"]]);
    expect(writes[1].pages.map((page) => page.items)).toEqual([[], [], []]);
    expect(writes[1].fonts.size).toBe(0);
    await expect(
      buildImagePdf(pages, new Map([["photo", jpegPart]]), options(items, { pageNumbers: true, fonts: new Map(), steps })),
    ).rejects.toThrow("page numbers need the bundled sans face");
  });
});

// --- the estimate and the limits ------------------------------------------------

function factsFor(count: number, facts: SourceFacts): SourceFacts[] {
  return Array.from({ length: count }, () => facts);
}

const scanTitles = (count: number) => Array.from({ length: count }, (_, index) => `scan ${index + 1}`);

describe("the estimate and the MAX_* refusals", () => {

  it("counts a passed-through file at its own length and a re-encode from its raster, each image once however often it is placed", () => {
    const jpeg: SourceFacts = { mediaId: "m", contentType: "image/jpeg", header: parseJpegHeader(jpegHeaderBytes(0xc0, 3)), byteLength: 4000 };
    const png: SourceFacts = { mediaId: "p", contentType: "image/png", header: null, byteLength: 900 };
    const items = [
      { id: "a", title: "a", widthPx: 160, heightPx: 120 },
      { id: "a", title: "a again", widthPx: 160, heightPx: 120 },
      { id: "b", title: "b", widthPx: 64, heightPx: 40 },
    ];
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    const estimate = estimateImagePdf(pages, items, [jpeg, jpeg, png], 150, false);
    expect(estimate.pageCount).toBe(3);
    // The PNG's raster is its own 64 × 40 px, at twice its file's 0.70 bytes
    // per pixel (under the 4 raw RGBA bytes).
    const png64 = Math.ceil(64 * 40 * ((2 * 900) / (64 * 40)));
    expect(estimate.images).toEqual([
      { bytes: 4000, indices: [0, 1] },
      { bytes: png64, indices: [2] },
    ]);
    expect(estimate.bytes).toBe(2048 + 3 * 1024 + 4000 + png64);
    // The 64 × 40 screenshot on 15.92 cm is the softest placement: page 3.
    expect(estimate.lowest?.index).toBe(2);
    expect(estimate.lowest?.page).toBe(2);
    expect(Math.round(estimate.lowest?.dpi ?? 0)).toBe(10);
    expect(estimateImagePdf(pages, items, [jpeg, jpeg, png], 150, true).bytes - estimate.bytes).toBe(410_820);
    expect(imageKey(pages[0].placed[0], jpeg, "original", 150, 0)).toBe(imageKey(pages[1].placed[0], jpeg, "original", 150, 1));
  });

  it(`refuses more than ${MAX_PDF_ASSEMBLY_PAGES} pages before any work, naming the limit and what to remove`, () => {
    const jpeg: SourceFacts = { mediaId: "m", contentType: "image/jpeg", header: parseJpegHeader(jpegHeaderBytes(0xc0, 3)), byteLength: 10 };
    const at = (count: number, mode: PageMode) => {
      const items = trayOf(count).map((item, index) => ({ ...item, title: `scan ${index + 1}` }));
      const pages = planPdfPages(items, plan(mode, { nUp: 4 }), A4);
      return pdfLimitRefusal(estimateImagePdf(pages, items, factsFor(count, jpeg), 150, false), scanTitles(count), itemsPerPage(plan(mode, { nUp: 4 }), count));
    };
    expect(at(500, "one-per-page")).toBeNull();
    expect(at(501, "one-per-page")).toBe(
      "This PDF would have 501 pages; a PDF assembly holds at most 500. Remove at least 1 image from the tray, or put more images on each page.",
    );
    expect(at(2000, "n-up")).toBeNull();
    expect(at(2010, "n-up")).toBe(
      "This PDF would have 503 pages; a PDF assembly holds at most 500. Remove at least 10 images from the tray, or put more images on each page.",
    );
    expect(at(2010, "single-page")).toBeNull();
  });

  it(`refuses an estimated output over ${MAX_PDF_ASSEMBLY_BYTES / 1024 / 1024} MB, naming the limit and the largest images to remove`, () => {
    const big = (mediaId: string, mb: number): SourceFacts => ({
      mediaId,
      contentType: "image/jpeg",
      header: parseJpegHeader(jpegHeaderBytes(0xc0, 3)),
      byteLength: mb * 1024 * 1024,
    });
    const items = ["small", "huge", "large", "medium"].map((title) => ({ id: title, title, widthPx: 160, heightPx: 120 }));
    const facts = [big("s", 10), big("h", 300), big("l", 200), big("m", 100)];
    const pages = planPdfPages(items, plan("one-per-page"), A4);
    const estimate = estimateImagePdf(pages, items, facts, 150, false);
    expect(pdfLimitRefusal(estimate, items.map((item) => item.title), 1)).toBe(
      "This PDF would be about 610 MB, more than the 500 MB a PDF assembly can produce. Remove “huge” from the tray, or choose a lower resolution.",
    );
    const fits = estimateImagePdf(pages, items, [big("s", 10), big("h", 300), big("l", 100), big("m", 50)], 150, false);
    expect(fits.bytes).toBeLessThan(MAX_PDF_ASSEMBLY_BYTES);
    expect(pdfLimitRefusal(fits, items.map((item) => item.title), 1)).toBeNull();
    const three = estimateImagePdf(pages, items, [big("s", 270), big("h", 300), big("l", 290), big("m", 280)], 150, false);
    expect(pdfLimitRefusal(three, items.map((item) => item.title), 1)).toContain("Remove “huge”, “large” and “medium” from the tray");
    const tray = Array.from({ length: 10 }, (_, index) => ({ id: `i${index + 1}`, title: `i${index + 1}`, widthPx: 160, heightPx: 120 }));
    const ten = estimateImagePdf(
      planPdfPages(tray, plan("one-per-page"), A4),
      tray,
      tray.map((item) => big(item.id, 100)),
      150,
      false,
    );
    // 1000 MB and the file's own overhead: six of them must go.
    expect(pdfLimitRefusal(ten, tray.map((item) => item.title), 1)).toContain("Remove “i1”, “i2”, “i3” and 3 more from the tray");
    // One image in three tray entries: removing one copy would leave two.
    const copies = [...Array.from({ length: 3 }, () => ({ id: "scan", title: "scan", widthPx: 160, heightPx: 120 })), items[0]];
    const shared = estimateImagePdf(
      planPdfPages(copies, plan("one-per-page"), A4),
      copies,
      [big("scan", 450), big("scan", 450), big("scan", 450), big("s", 100)],
      150,
      false,
    );
    expect(shared.images.map((image) => image.indices)).toEqual([[0, 1, 2], [3]]);
    expect(pdfLimitRefusal(shared, copies.map((item) => item.title), 1)).toContain("Remove “scan” (all 3 copies) from the tray");
  });

  it("estimates a re-encode from its raster: a lossless source at its own density (doubled), anything else at its raw samples", () => {
    const size = { widthPx: 1000, heightPx: 500 };
    const item = { widthPx: 2000, heightPx: 1000 };
    // A screenshot PNG of 2 MP in 400 KB: 0.2 bytes a pixel, so 0.4 here.
    expect(reencodedBytes(size, { contentType: "image/png", byteLength: 400_000 }, item, true)).toBe(200_000);
    // A dense PNG never counts above its raw samples.
    expect(reencodedBytes(size, { contentType: "image/png", byteLength: 9_000_000 }, item, true)).toBe(2_000_000);
    expect(reencodedBytes(size, { contentType: "image/gif", byteLength: 9_000_000 }, item, false)).toBe(1_500_000);
    // A JPEG's density says nothing about Flate; nor has an SVG pixels.
    expect(reencodedBytes(size, { contentType: "image/jpeg", byteLength: 100_000 }, item, false)).toBe(1_500_000);
    expect(reencodedBytes(size, { contentType: "image/svg+xml", byteLength: 300 }, item, true)).toBe(2_000_000);
  });
});

// --- the browser canvas path, with stub steps ----------------------------------

describe("rasterisePlacement", () => {
  it("draws through SPEC-109's renderer, releases its canvas, and encodes at the sheet's dpi", async () => {
    const { pngPart } = await sources();
    const draws: unknown[][] = [];
    const ctx = {
      save() {},
      restore() {},
      setTransform() {},
      clearRect() {},
      fillRect: (...args: unknown[]) => draws.push(["fillRect", ...args]),
      drawImage: (...args: unknown[]) => draws.push(["drawImage", ...args.slice(1)]),
    };
    const canvas = {
      width: 0,
      height: 0,
      getContext: () => ctx,
      // jsdom's Blob has no arrayBuffer(); the PNG's bytes are what matter.
      toBlob: (done: (blob: Blob) => void) =>
        done({ arrayBuffer: async () => new Uint8Array([137, 80, 78, 71]).buffer } as unknown as Blob),
    };
    const released: number[] = [];
    const encodeImage = vi.fn(async () => ({ widthPx: 32, heightPx: 20, filter: "FlateDecode" as const, colorSpace: "DeviceRGB" as const, data: new Uint8Array([1]) }));
    const createObjectURL = vi.fn(() => "blob:raster");
    const revokeObjectURL = vi.fn();
    vi.stubGlobal("URL", Object.assign(URL, { createObjectURL, revokeObjectURL }));
    const [page] = planPdfPages([{ id: "shot", widthPx: 64, heightPx: 40 }], plan("one-per-page", { cellBackground: "#00ff00" }), A4, 600);
    const placed = page.placed[0];
    const result = await rasterisePlacement(
      {
        region: placed,
        widthPx: 64,
        heightPx: 40,
        layers: [placed],
        inputs: [{ part: pngPart, widthPx: 64, heightPx: 40 }],
        background: "transparent",
        widthPt: 400,
        heightPt: 250,
        dpi: 600,
      },
      {
        decode: async () => ({ image: {} as CanvasImageSource, widthPx: 64, heightPx: 40, release: () => {} }),
        createCanvas: (width, height) => Object.assign(canvas, { width, height }) as unknown as HTMLCanvasElement,
        release: (target) => {
          released.push(target.width);
          target.width = 0;
          target.height = 0;
        },
        encodeImage,
      },
    );
    expect(result?.widthPx).toBe(32);
    // The image fills the raster exactly, over its cell's colour (the cell
    // is the whole content box; the canvas keeps the part under the image).
    expect(draws).toHaveLength(2);
    const [fill, image] = draws as [[string, number, number, number, number], unknown[]];
    expect(fill[0]).toBe("fillRect");
    expect(fill[1] <= 0 && fill[2] <= 0 && fill[1] + fill[3] >= 64 && fill[2] + fill[4] >= 40).toBe(true);
    expect(image).toEqual(["drawImage", 0, 0, 64, 40, 0, 0, 64, 40]);
    expect(released).toEqual([64]);
    expect(canvas.width).toBe(0);
    expect(encodeImage).toHaveBeenCalledWith("blob:raster", 400, 250, 600);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:raster");
    vi.unstubAllGlobals();
  });
});
