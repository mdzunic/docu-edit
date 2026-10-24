// SPEC-110: the assembly workspace's PDF output, run through the whole
// useAssembly controller in an effect scope over a real (IndexedDB-mocked)
// document store. Page-mode switching re-cuts the same tray in its order;
// PageUp/PageDown, the strip and the tray's selection move between pages
// under one sheet setup; each item says how it will be embedded; the
// estimate names the lowest-resolution placement and warns before export;
// the MAX_* and page-number refusals come before any work; cancellation and
// leaving stop at a page boundary with nothing behind; the parse-back
// self-check fails loudly on a corrupted build and stores nothing; and every
// source is deep-frozen throughout.
import { readFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope } from "vue";
import type { MediaModel } from "../../lib/media/model";
import type { PdfBuildInput, PdfImageData } from "../../lib/core/pdfBuild";
import type { RasterJob } from "../../lib/media/imagesToPdf";
import type { FaceKey } from "../../lib/core/pdfFonts";
import { parseSfnt, type SfntFont } from "../../lib/core/sfnt";

const sans: SfntFont = parseSfnt(
  new Uint8Array(readFileSync(resolvePath(__dirname, "../../../public/fonts/LiberationSans-Regular.ttf"))),
);

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

// Every workspace a test opens is disposed after it, so its debounced plan
// write cannot land in the next test.
const scopes: (() => void)[] = [];

afterEach(() => {
  scopes.splice(0).forEach((stop) => stop());
  vi.doUnmock("../../lib/idb");
  vi.doUnmock("../../lib/core/assemblyLayout");
  vi.restoreAllMocks();
});

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const inner of Object.values(value)) {
      deepFreeze(inner);
    }
  }
  return value;
}

function latin1(bytes: Uint8Array): string {
  let text = "";
  for (const byte of bytes) {
    text += String.fromCharCode(byte);
  }
  return text;
}

const key = (name: string) => new KeyboardEvent("keydown", { key: name, cancelable: true });

async function setup(pdfOverrides: Record<string, unknown> = {}) {
  const store = await import("../../lib/documentStore");
  const { assembleMediaModel } = await import("../../lib/media/open");
  const { mediaSourceBytes } = await import("../../lib/media/download");
  const { withSavedStack } = await import("../../lib/media/editModel");
  const fixtures = await import("../../lib/media/fixtures/build");
  const { buildPdf } = await import("../../lib/core/pdfBuild");
  const { useAssembly } = await import("../assembly/useAssembly");
  const png = await fixtures.buildPngFixtureBytes();
  const jpeg = fixtures.buildJpegFixtureBytes();

  async function seed(title: string, format: "png" | "jpeg" = "jpeg", image?: { widthPx: number; heightPx: number }): Promise<MediaModel> {
    const model = await assembleMediaModel(title, format, format === "png" ? png : jpeg);
    const stored = image ? { ...model, image } : model;
    store.storeMedia(stored);
    return stored;
  }

  const jobs: RasterJob[] = [];
  const writes: PdfBuildInput[] = [];
  const steps = {
    // The canvas path: a tiny deterministic stand-in.
    rasterise: vi.fn(async (job: RasterJob): Promise<PdfImageData> => {
      jobs.push(job);
      return { widthPx: job.widthPx, heightPx: job.heightPx, filter: "FlateDecode", colorSpace: "DeviceRGB", data: new Uint8Array([0x78, 0x9c, 1, 2]) };
    }),
    between: vi.fn(async () => {}),
    // The real writer, recorded.
    buildPdf: vi.fn(async (input: PdfBuildInput) => {
      writes.push(input);
      return buildPdf(input);
    }),
    loadFonts: vi.fn(async () => new Map<FaceKey, SfntFont>([["sans-regular", sans]])),
    download: vi.fn(),
    ...pdfOverrides,
  };

  function open() {
    const scope = effectScope();
    scopes.push(() => scope.stop());
    const assembly = scope.run(() => useAssembly({ pdfSteps: steps }))!;
    return { assembly, stop: () => scope.stop() };
  }

  return { store, seed, open, steps, jobs, writes, jpeg, png, mediaSourceBytes, withSavedStack };
}

type Assembly = ReturnType<Awaited<ReturnType<typeof setup>>["open"]>["assembly"];

const titles = (assembly: Assembly) => assembly.tray.entries.value.map((entry) => entry.title);
const pageIndices = (assembly: Assembly) => assembly.pdf.pages.value.map((page) => page.itemIndices);

describe("pages", () => {
  it("re-cuts the same tray into pages as the page mode switches, keeping the tray order", async () => {
    const { seed, open } = await setup();
    const ids = [];
    for (const title of ["p1", "p2", "p3", "p4", "p5"]) {
      ids.push((await seed(title)).id);
    }
    const { assembly } = open();
    assembly.tray.add(ids);
    expect(assembly.setFormat("pdf")).toBe(true);
    expect(assembly.isPdf.value).toBe(true);
    // One per page is the default.
    expect(assembly.output.pageMode).toBe("one-per-page");
    expect(pageIndices(assembly)).toEqual([[0], [1], [2], [3], [4]]);
    expect(assembly.pdf.pageCountLabel.value).toBe("5 pages");

    expect(assembly.pdf.setPageMode("n-up")).toBe(true);
    expect(assembly.pdf.setNUp("4")).toBe(true);
    expect(pageIndices(assembly)).toEqual([[0, 1, 2, 3], [4]]);
    expect(assembly.pdf.pageCountLabel.value).toBe("2 pages");

    expect(assembly.pdf.setPageMode("single-page")).toBe(true);
    expect(pageIndices(assembly)).toEqual([[0, 1, 2, 3, 4]]);
    expect(assembly.pdf.pageCountLabel.value).toBe("1 page");
    // The tray never moved.
    expect(titles(assembly)).toEqual(["p1", "p2", "p3", "p4", "p5"]);

    // And the tray order is still the page order.
    assembly.pdf.setPageMode("one-per-page");
    assembly.tray.move(assembly.tray.entries.value[0].key, 2);
    expect(titles(assembly)).toEqual(["p2", "p3", "p1", "p4", "p5"]);
    expect(assembly.pdf.pages.value.map((page) => titles(assembly)[page.itemIndices[0]])).toEqual(["p2", "p3", "p1", "p4", "p5"]);

    expect(assembly.pdf.setPageMode("booklet")).toBe(false);
    expect(assembly.pdf.error.value).toBe('"booklet" is not a page mode.');
    expect(assembly.pdf.setNUp("5")).toBe(false);
    expect(assembly.output.nUp).toBe(4);
  });

  it("moves between pages with PageUp/PageDown and the strip, and shows the page a tray selection is on", async () => {
    const { seed, open } = await setup();
    const ids = [];
    for (const title of ["a", "b", "c", "d", "e"]) {
      ids.push((await seed(title)).id);
    }
    const { assembly } = open();
    assembly.tray.add(ids);
    assembly.setFormat("pdf");
    const pdf = assembly.pdf;
    expect(pdf.current.value).toBe(0);
    // The preview's placements are the current page's.
    expect(assembly.placed.value).toEqual(pdf.pages.value[0].placed);

    const down = key("PageDown");
    pdf.onPageKey(down);
    expect(down.defaultPrevented).toBe(true);
    expect(pdf.current.value).toBe(1);
    expect(assembly.placed.value.map((item) => item.index)).toEqual([1]);
    pdf.onPageKey(key("PageUp"));
    pdf.onPageKey(key("PageUp"));
    expect(pdf.current.value).toBe(0);
    const other = key("ArrowDown");
    pdf.onPageKey(other);
    expect(other.defaultPrevented).toBe(false);

    // The strip.
    pdf.goTo(4);
    expect(pdf.current.value).toBe(4);
    pdf.goTo(9);
    expect(pdf.current.value).toBe(4);
    // Fewer pages: the current page stays on the last one there is.
    pdf.setPageMode("n-up");
    expect(pdf.pageCount.value).toBe(2);
    expect(pdf.current.value).toBe(1);

    pdf.setPageMode("one-per-page");
    assembly.select(assembly.tray.entries.value[2].key);
    await Promise.resolve();
    expect(pdf.current.value).toBe(2);
  });

  it("sets up every page with the one sheet", async () => {
    const { seed, open } = await setup();
    const ids = [(await seed("a")).id, (await seed("b")).id, (await seed("c")).id];
    const { assembly } = open();
    assembly.tray.add(ids);
    assembly.setFormat("pdf");
    assembly.sheet.setOrientation("landscape");
    assembly.sheet.setMarginPreset("Narrow");
    const content = assembly.sheet.geometry.value!.content;
    expect(content).toEqual({ xCm: 1.27, yCm: 1.27, widthCm: 27.16, heightCm: 18.46 });
    for (const page of assembly.pdf.pages.value) {
      expect(page.placed[0].cell).toEqual(content);
    }
  });

  it("keeps the page settings in the persisted plan", async () => {
    const { seed, open } = await setup();
    const a = await seed("a");
    const first = open();
    first.assembly.tray.add([a.id]);
    first.assembly.setFormat("pdf");
    first.assembly.pdf.setPageMode("n-up");
    first.assembly.pdf.setNUp("6");
    first.assembly.pdf.setPageNumbers(true);
    first.assembly.pdf.setAuthor("Ana");
    first.stop();
    const { assembly } = open();
    expect(assembly.output).toMatchObject({ format: "pdf", pageMode: "n-up", nUp: 6, pageNumbers: true, author: "Ana", subject: "" });
  });
});

describe("before anything is built", () => {
  it("says per item whether its file is embedded as the original JPEG or re-encoded, and why", async () => {
    const { seed, open, store, withSavedStack, jpeg } = await setup();
    const photo = await seed("photo");
    const shot = await seed("shot", "png");
    const edited = withSavedStack(
      await seed("edited"),
      [{ kind: "rotate", quarterTurns: 1 }],
      { bytes: jpeg, contentType: "image/jpeg", widthPx: 120, heightPx: 160 },
      null,
      () => crypto.randomUUID(),
    );
    store.storeMedia(edited);
    const { assembly } = open();
    assembly.tray.add([photo.id, shot.id, edited.id]);
    assembly.setFormat("pdf");
    expect(assembly.pdf.encodings.value).toEqual(["original JPEG", "re-encoded: PNG", "re-encoded: edited"]);
    assembly.sheet.setFit("cover");
    expect(assembly.pdf.encodings.value[0]).toBe("re-encoded: cropped");
    assembly.sheet.setFit("stretch");
    expect(assembly.pdf.encodings.value[0]).toBe("original JPEG");
    // An overlay's opacity on the one page.
    assembly.sheet.setFit("contain");
    assembly.pdf.setPageMode("single-page");
    assembly.setLayout("overlay");
    assembly.select(assembly.tray.entries.value[0].key);
    assembly.setTransformField("opacity", "50");
    expect(assembly.pdf.encodings.value[0]).toBe("re-encoded: opacity");
  });

  it("reports the lowest-resolution placement in the estimate, and warns before export when it prints soft", async () => {
    const { seed, open } = await setup();
    const photo = await seed("photo");
    const shot = await seed("shot", "png");
    const { assembly } = open();
    assembly.tray.add([photo.id, shot.id]);
    assembly.setFormat("pdf");
    assembly.sheet.setDpi(300);
    // 64 px across 15.92 cm: 10 dpi, on page 2.
    expect(assembly.pdf.estimate.value).toMatch(/^2 pages · PDF up to about .+ · lowest resolution 10 dpi: “shot” on page 2$/);
    expect(assembly.pdf.dpiWarning.value).toBe("“shot” on page 2 prints at 10 dpi, below the sheet's 300 dpi: it will print soft.");
    // At actual size on a 72 dpi sheet, every pixel is a dot: nothing to warn of.
    assembly.sheet.setDpi(72);
    assembly.sheet.setFit("actual");
    expect(assembly.pdf.estimate.value).toContain("lowest resolution 72 dpi");
    expect(assembly.pdf.dpiWarning.value).toBe("");
  });

  it("refuses an estimate over MAX_PDF_ASSEMBLY_BYTES before any work, naming the limit and what to remove", async () => {
    const { seed, open, store, steps } = await setup();
    const ids = [];
    for (let i = 1; i <= 11; i += 1) {
      ids.push(deepFreeze(await seed(`big ${i}`, "jpeg", { widthPx: 10000, heightPx: 10000 })).id);
    }
    const { assembly } = open();
    assembly.tray.add(ids);
    assembly.setFormat("pdf");
    // Cropped, the 10000 px squares are re-encoded: each fills A4's content
    // box with RGB samples at 600 dpi — held to the canvas area every browser
    // allocates, about 525 dpi and 48 MB a page.
    assembly.sheet.setFit("cover");
    assembly.sheet.setDpi(600);
    expect(assembly.pdf.encodings.value[0]).toBe("re-encoded: cropped");
    expect(assembly.pdf.estimate.value).toMatch(/^11 pages · PDF up to about 52\d\.\d MB · lowest resolution 525 dpi/);
    expect(assembly.pdf.refusal.value).toMatch(
      /^This PDF would be about 52\d\.\d MB, more than the 500 MB a PDF assembly can produce\. Remove “big 1” from the tray, or choose a lower resolution\.$/,
    );
    expect(assembly.pdf.canProduce.value).toBe(false);
    const before = store.listDocuments().length;
    expect(await assembly.pdf.save()).toBeNull();
    expect(assembly.pdf.failure.value).toBe(assembly.pdf.refusal.value);
    expect(steps.rasterise).not.toHaveBeenCalled();
    expect(steps.buildPdf).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(before);
    // One fewer fits.
    assembly.tray.remove(assembly.tray.entries.value[0].key);
    expect(assembly.pdf.refusal.value).toBe("");
  });

  it("refuses more than MAX_PDF_ASSEMBLY_PAGES before any work, naming the limit and what to remove", async () => {
    // A tray holds 100 images; let this one hold more, to reach 501 pages.
    vi.doMock("../../lib/core/assemblyLayout", async (importOriginal) => ({
      ...(await importOriginal<typeof import("../../lib/core/assemblyLayout")>()),
      MAX_ASSEMBLY_ITEMS: 600,
    }));
    const { seed, open, steps } = await setup();
    const scan = await seed("scan");
    const { assembly } = open();
    expect(assembly.tray.add(Array.from({ length: 501 }, () => scan.id))).toBe(501);
    assembly.setFormat("pdf");
    expect(assembly.pdf.pageCount.value).toBe(501);
    expect(assembly.pdf.refusal.value).toBe(
      "This PDF would have 501 pages; a PDF assembly holds at most 500. Remove at least 1 image from the tray, or put more images on each page.",
    );
    expect(await assembly.pdf.download()).toBe(false);
    expect(steps.buildPdf).not.toHaveBeenCalled();
    expect(steps.download).not.toHaveBeenCalled();
    assembly.pdf.setPageMode("n-up");
    assembly.pdf.setNUp("2");
    expect(assembly.pdf.refusal.value).toBe("");
  });

  it("refuses page numbers the bottom margin cannot hold, and an empty tray", async () => {
    const { seed, open } = await setup();
    const { assembly } = open();
    assembly.setFormat("pdf");
    expect(assembly.pdf.refusal.value).toBe("Add at least one image to the tray.");
    assembly.tray.add([(await seed("a")).id]);
    expect(assembly.pdf.refusal.value).toBe("");
    assembly.sheet.setMargin("top", "1");
    assembly.pdf.setPageNumbers(true);
    expect(assembly.pdf.refusal.value).toMatch(/^Page numbers sit 1 cm above the bottom edge, inside the margin/);
    assembly.pdf.setPageNumbers(false);
    expect(assembly.pdf.refusal.value).toBe("");
  });
});

describe("building", () => {
  it("saves a NEW PDF whose summary is the parse of its own bytes, with the JPEG embedded verbatim", async () => {
    const { seed, open, store, steps, jobs, writes, jpeg, mediaSourceBytes } = await setup();
    const photo = deepFreeze(await seed("photo"));
    const shot = deepFreeze(await seed("shot", "png"));
    const recordsBefore = [photo.id, shot.id].map((id) => ({ ...store.getStoredRecord(id) }));
    const { assembly } = open();
    assembly.tray.add([photo.id, shot.id]);
    assembly.setFormat("pdf");
    assembly.pdf.setPageNumbers(true);
    assembly.pdf.setAuthor("Ana");
    const progress: string[] = [];
    steps.between.mockImplementation(async () => {
      progress.push(assembly.pdf.progress.value);
    });
    const saved = await assembly.pdf.save();
    expect(saved).toMatchObject({ kind: "media", format: "pdf", contentType: "application/pdf", title: "Combined image" });
    expect(saved?.pdf?.pageCount).toBe(2);
    expect(saved?.pdf?.pages[0].widthPt).toBeCloseTo(595.28, 1);
    expect(saved?.properties).toMatchObject({ title: "Combined image", author: "Ana" });
    expect(store.getMedia(saved!.id)).toBe(saved);
    // The writer ran once, with the one bundled face.
    expect(steps.buildPdf).toHaveBeenCalledTimes(1);
    expect([...writes[0].fonts.keys()]).toEqual(["sans-regular"]);
    expect(steps.loadFonts).toHaveBeenCalledTimes(1);
    // A page at a time, reported.
    expect(progress).toEqual(["Page 1 of 2"]);
    expect(assembly.pdf.progress.value).toBe("");
    // The PNG went through the canvas path at the sheet's dpi; the JPEG did not.
    expect(jobs.map((job) => job.dpi)).toEqual([150]);
    expect(latin1(mediaSourceBytes(saved!)).includes(latin1(jpeg))).toBe(true);
    // The sources: untouched.
    expect([photo.id, shot.id].map((id) => ({ ...store.getStoredRecord(id) }))).toEqual(recordsBefore);
    expect((await assembly.pdf.save())?.title).toBe("Combined image (2)");
  });

  it("downloads as <title>.pdf, storing nothing", async () => {
    const { seed, open, store, steps } = await setup();
    const a = deepFreeze(await seed("a"));
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.setFormat("pdf");
    assembly.setTitle("Receipts");
    const before = store.listDocuments().length;
    expect(await assembly.pdf.download()).toBe(true);
    expect(steps.download).toHaveBeenCalledWith(expect.any(Uint8Array), "Receipts.pdf", "application/pdf");
    expect(store.listDocuments()).toHaveLength(before);
  });

  it("stops at the next page boundary when cancelled, leaving no download, no record and no mutation", async () => {
    const { seed, open, store, steps } = await setup();
    const ids = [deepFreeze(await seed("a")).id, deepFreeze(await seed("b")).id, deepFreeze(await seed("c")).id];
    const records = ids.map((id) => ({ ...store.getStoredRecord(id) }));
    const { assembly } = open();
    assembly.tray.add(ids);
    assembly.setFormat("pdf");
    steps.between.mockImplementation(async () => {
      expect(assembly.pdf.busy.value).toBe(true);
      expect(assembly.pdf.progress.value).toBe("Page 1 of 3");
      assembly.pdf.cancel();
    });
    const before = store.listDocuments().length;
    expect(await assembly.pdf.save()).toBeNull();
    expect(steps.buildPdf).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(before);
    expect(assembly.pdf.failure.value).toBe("");
    expect(assembly.pdf.busy.value).toBe(false);
    expect(assembly.pdf.progress.value).toBe("");
    expect(await assembly.pdf.download()).toBe(false);
    expect(steps.download).not.toHaveBeenCalled();
    expect(ids.map((id) => ({ ...store.getStoredRecord(id) }))).toEqual(records);
    expect(titles(assembly)).toEqual(["a", "b", "c"]);
  });

  it("stops when the workspace is left mid-build, writing nothing", async () => {
    const { seed, open, store, steps } = await setup();
    const ids = [(await seed("a")).id, (await seed("b")).id];
    const { assembly, stop } = open();
    assembly.tray.add(ids);
    assembly.setFormat("pdf");
    steps.between.mockImplementation(async () => {
      stop();
    });
    const before = store.listDocuments().length;
    expect(await assembly.pdf.save()).toBeNull();
    expect(steps.buildPdf).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(before);
  });

  it("fails loudly and stores nothing when its own output will not parse back", async () => {
    const junk = vi.fn(async () => new TextEncoder().encode("%PDF-1.7 this is not a PDF"));
    const { seed, open, store, steps } = await setup({ buildPdf: junk });
    const { assembly } = open();
    assembly.tray.add([(await seed("a")).id, (await seed("b")).id]);
    assembly.setFormat("pdf");
    const before = store.listDocuments().length;
    expect(await assembly.pdf.save()).toBeNull();
    expect(junk).toHaveBeenCalledTimes(1);
    expect(assembly.pdf.failure.value).toMatch(
      /^Could not save the PDF: the PDF it built does not read back \(.+\), so nothing was kept\.$/,
    );
    expect(store.listDocuments()).toHaveLength(before);
    expect(await assembly.pdf.download()).toBe(false);
    expect(steps.download).not.toHaveBeenCalled();
    expect(assembly.pdf.failure.value).toMatch(/^Could not download the PDF: the PDF it built does not read back/);
  });

  it("fails loudly when its own output parses back to the wrong page count", async () => {
    const { buildPdf } = await import("../../lib/core/pdfBuild");
    const short = vi.fn((input: PdfBuildInput) => buildPdf({ ...input, pages: input.pages.slice(0, 1) }));
    const { seed, open, store } = await setup({ buildPdf: short });
    const { assembly } = open();
    assembly.tray.add([(await seed("a")).id, (await seed("b")).id]);
    assembly.setFormat("pdf");
    const before = store.listDocuments().length;
    expect(await assembly.pdf.save()).toBeNull();
    expect(assembly.pdf.failure.value).toBe("Could not save the PDF: the PDF it built has 1 pages instead of 2, so nothing was kept.");
    expect(store.listDocuments()).toHaveLength(before);
  });

  it("names an image that cannot be drawn, storing nothing", async () => {
    const { seed, open, store } = await setup({ rasterise: vi.fn(async () => null) });
    const { assembly } = open();
    assembly.tray.add([(await seed("shot", "png")).id]);
    assembly.setFormat("pdf");
    const before = store.listDocuments().length;
    expect(await assembly.pdf.save()).toBeNull();
    expect(assembly.pdf.failure.value).toBe("Could not save the PDF: “shot” could not be drawn.");
    expect(store.listDocuments()).toHaveLength(before);
  });
});
