// SPEC-111: the page organiser's output, run through the whole useAssembly
// controller in an effect scope over a real (IndexedDB-mocked) document
// store with the real reader, writer and zip. Page expansion from a source;
// rotation arithmetic through the tray; the three destinations — merge to
// one PDF, extract the selection, separate PDFs zipped with stored entries;
// an image and a blank page merged as ordinary sources through SPEC-110's
// builder; the refusals (encrypted, unreadable, bounds, an empty selection)
// before any work; cancellation at a page boundary leaving nothing behind;
// the re-parse self-check failing loudly and storing nothing; the result
// preview's debounce, page ceiling and Refresh; the persisted plan; and every
// source deep-frozen throughout.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref } from "vue";
import type { MediaModel } from "../../lib/media/model";
import type { PdfBuildInput, PdfImageData } from "../../lib/core/pdfBuild";
import type { PdfAssemblyHooks, PdfAssemblyPlan, PdfAssemblyResult } from "../../lib/core/pdfDocument";
import type { TrayEntry } from "../assembly/useAssemblyTray";

const FIXTURES = join(__dirname, "../../../e2e/fixtures");
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(join(FIXTURES, name)));
const PAGES = fixture("spec107-pages.pdf");
const OUTLINES = fixture("spec111-outlines.pdf");
const ENCRYPTED = fixture("spec107-encrypted.pdf");

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
  vi.doMock("../../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

// Every workspace a test opens is disposed after it, so neither its plan
// write nor its preview timer lands in the next test.
const scopes: (() => void)[] = [];

afterEach(() => {
  scopes.splice(0).forEach((stop) => stop());
  vi.doUnmock("../../lib/idb");
  vi.restoreAllMocks();
  vi.useRealTimers();
});

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value) && !ArrayBuffer.isView(value)) {
    Object.freeze(value);
    for (const inner of Object.values(value)) {
      deepFreeze(inner);
    }
  }
  return value;
}

const latin1 = (bytes: Uint8Array): string => Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");

async function setup(mergeOverrides: Record<string, unknown> = {}) {
  const store = await import("../../lib/documentStore");
  const { assembleMediaModel } = await import("../../lib/media/open");
  const { mediaSourceBytes } = await import("../../lib/media/download");
  const { bytesToBase64 } = await import("../../lib/core/packageParts");
  const fixtures = await import("../../lib/media/fixtures/build");
  const pdf = await import("../../lib/core/pdfDocument");
  const { buildPdf } = await import("../../lib/core/pdfBuild");
  const { zip } = await import("../../lib/core/zipWriter");
  const { unzip } = await import("../../lib/core/zip");
  const { useAssembly } = await import("../assembly/useAssembly");
  const persistence = await import("../assembly/useAssemblyPersistence");
  const { usePdfMerge, PREVIEW_DEBOUNCE_MS, PREVIEW_MAX_PAGES } = await import("./usePdfMerge");
  const jpeg = fixtures.buildJpegFixtureBytes();

  /** A library PDF, deep-frozen: nothing may change it. */
  async function seedPdf(title: string, bytes: Uint8Array, override: (model: MediaModel) => MediaModel = (model) => model): Promise<MediaModel> {
    const model = override(await assembleMediaModel(title, "pdf", bytes));
    store.storeMedia(model);
    return deepFreeze(model);
  }

  async function seedImage(title: string): Promise<MediaModel> {
    const model = await assembleMediaModel(title, "jpeg", jpeg);
    store.storeMedia(model);
    return deepFreeze(model);
  }

  let clock = 0;
  const writes: PdfBuildInput[] = [];
  const steps = {
    parse: vi.fn(pdf.parsePdf),
    assemble: vi.fn((plan: PdfAssemblyPlan, hooks?: PdfAssemblyHooks) => pdf.assemblePdf(plan, hooks)),
    imagePdf: {
      rasterise: vi.fn(async (): Promise<PdfImageData> => ({
        widthPx: 1,
        heightPx: 1,
        filter: "FlateDecode",
        colorSpace: "DeviceRGB",
        data: new Uint8Array([0x78, 0x9c, 0x63, 0x60, 0x60, 0x60, 0, 0, 0, 4, 0, 1]),
      })),
      between: vi.fn(async () => {}),
      buildPdf: vi.fn(async (input: PdfBuildInput) => {
        writes.push(input);
        return buildPdf(input);
      }),
    },
    between: vi.fn(async () => {}),
    download: vi.fn(),
    zip: vi.fn(zip),
    previewUrl: vi.fn(() => `blob:preview-${steps.previewUrl.mock.calls.length}`),
    revokeUrl: vi.fn(),
    now: () => clock,
    ...mergeOverrides,
  };

  function open() {
    const scope = effectScope();
    scopes.push(() => scope.stop());
    const assembly = scope.run(() => useAssembly({ mergeSteps: steps }))!;
    return { assembly, stop: () => scope.stop() };
  }

  /** Lets the clock and the timers run `ms` on. */
  async function wait(ms: number) {
    clock += ms;
    await vi.advanceTimersByTimeAsync(ms);
  }

  return {
    store,
    pdf,
    steps,
    writes,
    jpeg,
    seedPdf,
    seedImage,
    open,
    wait,
    unzip,
    bytesToBase64,
    mediaSourceBytes,
    persistence,
    usePdfMerge,
    PREVIEW_DEBOUNCE_MS,
    PREVIEW_MAX_PAGES,
  };
}

type Setup = Awaited<ReturnType<typeof setup>>;
type Assembly = ReturnType<Setup["open"]>["assembly"];

const pages = (assembly: Assembly) => assembly.tray.entries.value.map((entry) => [entry.title, entry.page?.index ?? -1]);
const keyAt = (assembly: Assembly, index: number) => assembly.tray.entries.value[index].key;

/** The stored document a save made, read back from the store. */
function pdfSummary(model: MediaModel | null) {
  return model?.pdf?.pages.map((page) => [Math.round(page.widthPt), Math.round(page.heightPt), page.rotate]);
}

describe("page expansion", () => {
  it("expands each PDF of a 'Combine…' selection into its pages, in order, as a PDF plan", async () => {
    const { seedPdf, open, persistence } = await setup();
    const outlines = await seedPdf("outlines", OUTLINES);
    const book = await seedPdf("pages", PAGES);
    persistence.startAssembly([outlines.id, book.id]);
    const { assembly } = open();
    expect(pages(assembly)).toEqual([
      ["outlines", 0],
      ["outlines", 1],
      ["pages", 0],
      ["pages", 1],
      ["pages", 2],
    ]);
    const second = assembly.tray.entries.value[3].page;
    expect(second).toEqual({ index: 1, count: 3, widthPt: 595.28, heightPt: 841.89, rotate: 90 });
    expect(assembly.merging.value).toBe(true);
    expect(assembly.merge.active.value).toBe(true);
    expect(assembly.output.format).toBe("pdf");
    expect(assembly.output.title).toBe("Combined PDF");
    expect(assembly.merge.estimate.value).toBe("5 pages from 2 files");
    // Only a PDF can carry PDF pages.
    expect(assembly.setFormat("png")).toBe(false);
    expect(assembly.error.value).toBe("The tray holds PDF pages, which only a PDF can carry: remove them to make an image.");
    expect(assembly.output.format).toBe("pdf");
  });

  it("adds a PDF from the library picker as its pages, turning the output to PDF", async () => {
    const { seedPdf, seedImage, open } = await setup();
    const book = await seedPdf("pages", PAGES);
    const photo = await seedImage("photo");
    const { assembly } = open();
    expect(assembly.output.format).toBe("png");
    expect(assembly.tray.add([photo.id, book.id])).toBe(4);
    expect(assembly.tray.announcement.value).toBe("Added 4 pages.");
    expect(pages(assembly)).toEqual([
      ["photo", -1],
      ["pages", 0],
      ["pages", 1],
      ["pages", 2],
    ]);
    expect(assembly.output.format).toBe("pdf");
    expect(assembly.imageEntries.value.map((entry) => entry.title)).toEqual(["photo"]);
  });
});

describe("page operations through the tray", () => {
  it("turns pages by quarter turns, modulo 360, onto each page's own /Rotate — and the output carries it", async () => {
    const { seedPdf, open, persistence } = await setup();
    const book = await seedPdf("pages", PAGES);
    persistence.startAssembly([book.id]);
    const { assembly } = open();
    const tray = assembly.tray;
    const rotation = (index: number) => {
      const entry = tray.entries.value[index];
      return ((entry.page?.rotate ?? 0) + (entry.rotateBy ?? 0)) % 360;
    };
    // Page 2 is /Rotate 90 in its source.
    tray.rotate(keyAt(assembly, 1), 90);
    expect(rotation(1)).toBe(180);
    expect(tray.announcement.value).toBe("Turned page 2 of “pages” right: rotate 180.");
    // Three left turns: 90 → 0 → 270 → 180 added, so 270 in all.
    tray.rotate(keyAt(assembly, 1), -90);
    tray.rotate(keyAt(assembly, 1), -90);
    tray.rotate(keyAt(assembly, 1), -90);
    expect(tray.entries.value[1].rotateBy).toBe(180);
    expect(rotation(1)).toBe(270);
    // Two right turns wrap past 360 back to the page's own rotation.
    tray.rotate(keyAt(assembly, 1), 90);
    tray.rotate(keyAt(assembly, 1), 90);
    expect(tray.entries.value[1].rotateBy).toBe(0);
    expect(rotation(1)).toBe(90);
    // On a selection, every selected page turns.
    tray.pick(keyAt(assembly, 0));
    tray.pick(keyAt(assembly, 2), { toggle: true });
    tray.rotate(keyAt(assembly, 2), -90);
    expect(tray.entries.value.map((entry) => entry.rotateBy ?? 0)).toEqual([270, 0, 270]);
    expect(tray.announcement.value).toBe("Turned 2 pages left.");
    const saved = await assembly.merge.save();
    expect(pdfSummary(saved)).toEqual([
      [612, 792, 270],
      [595, 842, 90],
      [720, 540, 270],
    ]);
  });

  it("selects by range, by toggle, everything, and every page of one source", async () => {
    const { seedPdf, open } = await setup();
    const outlines = await seedPdf("outlines", OUTLINES);
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const tray = assembly.tray;
    tray.add([book.id, outlines.id]);
    tray.pick(keyAt(assembly, 1));
    tray.pick(keyAt(assembly, 3), { range: true });
    expect(tray.selectedKeys.value).toEqual([1, 2, 3].map((i) => keyAt(assembly, i)));
    tray.pick(keyAt(assembly, 2), { toggle: true });
    expect(tray.selectedKeys.value).toEqual([1, 3].map((i) => keyAt(assembly, i)));
    tray.selectAll();
    expect(tray.selectedKeys.value).toHaveLength(5);
    tray.selectSource(keyAt(assembly, 4));
    expect(tray.selectedKeys.value).toEqual([3, 4].map((i) => keyAt(assembly, i)));
    expect(tray.announcement.value).toBe("Selected 2 pages of “outlines”.");
    // A plain click starts again.
    tray.pick(keyAt(assembly, 0));
    expect(tray.selectedKeys.value).toEqual([keyAt(assembly, 0)]);
  });

  it("deletes and duplicates pages in the plan only, and inserts a blank page after the selection", async () => {
    const { seedPdf, open, store, mediaSourceBytes } = await setup();
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const tray = assembly.tray;
    tray.add([book.id]);
    tray.duplicatePages(keyAt(assembly, 0));
    expect(pages(assembly)).toEqual([
      ["pages", 0],
      ["pages", 0],
      ["pages", 1],
      ["pages", 2],
    ]);
    tray.removePages(keyAt(assembly, 2));
    expect(tray.announcement.value).toBe("Removed page 2 of “pages” from the plan.");
    tray.pick(keyAt(assembly, 0));
    expect(tray.addBlank()).toBe(true);
    expect(tray.entries.value.map((entry) => (entry.blank ? "blank" : entry.page?.index))).toEqual([0, "blank", 0, 2]);
    // The stored source still has all three pages, byte for byte.
    expect(store.getMedia(book.id)?.pdf?.pageCount).toBe(3);
    expect(mediaSourceBytes(store.getMedia(book.id) as MediaModel)).toEqual(PAGES);
  });
});

describe("the three destinations", () => {
  it("merges the whole tray in order into a NEW library document, every source untouched", async () => {
    const { seedPdf, open, persistence, store, steps, pdf } = await setup();
    const outlines = await seedPdf("outlines", OUTLINES);
    const book = await seedPdf("pages", PAGES);
    const before = [structuredClone(store.getMedia(outlines.id)), structuredClone(store.getMedia(book.id))];
    persistence.startAssembly([outlines.id, book.id]);
    const { assembly } = open();
    const tray = assembly.tray;
    // Rotate page 2 right, delete page 4, drag page 5 to the front.
    tray.rotate(keyAt(assembly, 1), 90);
    tray.removePages(keyAt(assembly, 3));
    tray.moveTo(keyAt(assembly, 3), 0);
    expect(tray.announcement.value).toBe("Page 3 of “pages” moved to position 1 of 4.");
    const saved = await assembly.merge.save();
    expect(assembly.merge.failure.value).toBe("");
    expect(saved).toMatchObject({ kind: "media", format: "pdf", title: "Combined PDF" });
    expect(saved?.pdf?.pageCount).toBe(4);
    expect(pdfSummary(saved)).toEqual([
      [720, 540, 0],
      [420, 595, 0],
      [595, 420, 90],
      [612, 792, 0],
    ]);
    expect(assembly.merge.savedText.value).toBe("Saved “Combined PDF” to the library.");
    // It is in the library, and re-reads; the sources are as they were.
    expect(store.listDocuments().map((entry) => entry.title).toSorted()).toEqual(["Combined PDF", "outlines", "pages"]);
    expect((await pdf.parsePdf(new Uint8Array(readFileSync(join(FIXTURES, "spec107-pages.pdf"))))).pages).toHaveLength(3);
    expect([store.getMedia(outlines.id), store.getMedia(book.id)]).toEqual(before);
    // What changed: the tagged source's structure tree.
    expect(assembly.merge.dropGroups.value.map((group) => [group.kind, group.records.length])).toEqual([["structureTree", 1]]);
    expect(steps.assemble).toHaveBeenCalledTimes(1);
  });

  it("extracts only the selected pages, in tray order, refusing an empty selection first", async () => {
    const { seedPdf, open, steps, pdf } = await setup();
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id]);
    expect(merge.setDestination("extract")).toBe(true);
    expect(merge.refusal.value).toBe("Select the pages to extract: click a page, then Shift-click or Ctrl/Cmd-click to add more.");
    expect(merge.canProduce.value).toBe(false);
    expect(await merge.save()).toBeNull();
    expect(merge.failure.value).toBe(merge.refusal.value);
    expect(steps.assemble).not.toHaveBeenCalled();
    tray.pick(keyAt(assembly, 2));
    tray.pick(keyAt(assembly, 0), { toggle: true });
    expect(merge.estimate.value).toBe("2 pages of 3 selected");
    expect(await merge.download()).toBe(true);
    expect(steps.download).toHaveBeenCalledTimes(1);
    const [bytes, name, type] = steps.download.mock.calls[0] as [Uint8Array, string, string];
    expect([name, type]).toEqual(["Combined image.pdf", "application/pdf"]);
    const out = await pdf.parsePdf(bytes);
    expect(out.pages.map((page) => [page.widthPt, page.heightPt])).toEqual([
      [612, 792],
      [720, 540],
    ]);
    expect(merge.setDestination("booklet")).toBe(false);
  });

  it("saves separate PDFs — one per source-contiguous run — downloaded as a .zip of stored entries, or saved one by one", async () => {
    const { seedPdf, open, steps, unzip, pdf, store } = await setup();
    const outlines = await seedPdf("outlines", OUTLINES);
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id, outlines.id]);
    // pages 1, pages 2, outlines 1, pages 3 (outlines 2 removed).
    tray.removePages(keyAt(assembly, 4));
    tray.moveTo(keyAt(assembly, 2), 3);
    expect(pages(assembly)).toEqual([
      ["pages", 0],
      ["pages", 1],
      ["outlines", 0],
      ["pages", 2],
    ]);
    merge.setDestination("split");
    expect(merge.estimate.value).toBe("3 PDFs, one per run of pages from the same file · downloads as a .zip");
    expect(await merge.download()).toBe(true);
    const [entries] = steps.zip.mock.calls[0] as [Array<{ name: string; data: Uint8Array; store?: boolean }>];
    expect(entries.map((entry) => [entry.name, entry.store])).toEqual([
      ["01 pages.pdf", true],
      ["02 outlines.pdf", true],
      ["03 pages.pdf", true],
    ]);
    const [archive, name, type] = steps.download.mock.calls[0] as [Uint8Array, string, string];
    expect([name, type]).toEqual(["Combined image.zip", "application/zip"]);
    const files = await unzip(archive);
    const counts = await Promise.all(Array.from(files.values(), async (bytes) => (await pdf.parsePdf(bytes)).pages.length));
    expect(counts).toEqual([2, 1, 1]);
    // Each file was parsed back before it was zipped.
    expect(steps.parse.mock.calls.length).toBeGreaterThanOrEqual(5);

    // Saved one by one, each under its source's title (numbered: the
    // sources themselves hold the plain titles).
    const first = await merge.save();
    expect(first?.title).toBe("pages (2)");
    expect(merge.savedText.value).toBe("Saved 3 PDFs to the library.");
    expect(store.listDocuments().map((entry) => entry.title).toSorted()).toEqual(["outlines", "outlines (2)", "pages", "pages (2)", "pages (3)"]);
  });

  it("builds an image into a one-page PDF on the sheet through SPEC-110's builder, and a blank page the same way", async () => {
    const { seedPdf, seedImage, open, writes, jpeg, pdf } = await setup();
    const book = await seedPdf("pages", PAGES);
    const photo = await seedImage("photo");
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id, photo.id]);
    tray.addBlank();
    const saved = await merge.save();
    expect(merge.failure.value).toBe("");
    // Two one-page builds: the photo, then the blank page.
    expect(writes.map((input) => [input.pages.length, input.pages[0].images.length])).toEqual([
      [1, 1],
      [1, 0],
    ]);
    const sheet = assembly.geometry.value!;
    const sheetPt = [Math.round((sheet.widthCm / 2.54) * 72), Math.round((sheet.heightCm / 2.54) * 72), 0];
    expect(pdfSummary(saved)?.slice(3)).toEqual([sheetPt, sheetPt]);
    // The photo is a baseline JPEG: it rides through both writers verbatim.
    const bytes = await (await import("../../lib/media/download")).mediaSourceBytes(saved as MediaModel);
    expect(latin1(bytes).includes(latin1(jpeg))).toBe(true);
    expect((await pdf.parsePdf(bytes)).pages).toHaveLength(5);
  });
});

describe("refusals", () => {
  it("refuses the whole plan when a source is encrypted or unreadable, naming each file, before anything is built", async () => {
    const { seedPdf, open, steps, store, bytesToBase64 } = await setup();
    const book = await seedPdf("pages", PAGES);
    // Stored sources whose bytes turn out not to be a readable PDF.
    const locked = await seedPdf("locked", PAGES, (model) => ({ ...model, media: [{ ...model.media[0], data: bytesToBase64(ENCRYPTED) }] }));
    const junk = await seedPdf("junk", PAGES, (model) => ({ ...model, media: [{ ...model.media[0], data: bytesToBase64(new Uint8Array([1, 2, 3])) }] }));
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id, locked.id, junk.id]);
    const count = store.listDocuments().length;
    expect(await merge.save()).toBeNull();
    expect(merge.failure.value).toMatch(/^Could not save the PDF: “locked” is password-protected, so it cannot be merged\. “junk” could not be read as a PDF \(.+\), so it cannot be merged\. Remove them from the plan or replace them, then try again\.$/);
    expect(steps.assemble).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(count);
    expect(await merge.download()).toBe(false);
    expect(steps.download).not.toHaveBeenCalled();
    // Removing the offenders, the plan goes through.
    tray.selectSource(keyAt(assembly, 3));
    tray.removePages(keyAt(assembly, 3));
    tray.selectSource(keyAt(assembly, 3));
    tray.removePages(keyAt(assembly, 3));
    expect((await merge.save())?.pdf?.pageCount).toBe(3);
  });

  it("refuses plans over MAX_MERGE_PAGES and MAX_MERGE_SOURCE_BYTES before any work, naming the limit", async () => {
    const { seedPdf, steps, usePdfMerge, pdf } = await setup();
    const book = await seedPdf("pages", PAGES);
    const huge = await seedPdf("huge", PAGES, (model) => ({ ...model, byteLength: pdf.MAX_MERGE_SOURCE_BYTES + 1 }));
    const { libraryPdf, pageEntries } = await import("../assembly/useAssemblyTray");
    const entries = ref<TrayEntry[]>([]);
    const scope = effectScope();
    scopes.push(() => scope.stop());
    const merge = scope.run(() =>
      usePdfMerge({
        entries,
        selectedKeys: computed(() => []),
        active: computed(() => entries.value.length > 0),
        sheet: { setup: ref({ background: "#ffffff", dpi: 150 }), geometry: ref(null) } as never,
        plan: computed(() => ({ layout: "grid", columns: "auto", gapCm: 0, fit: "contain", align: "center", cellBackground: "transparent" }) as never),
        output: { title: "Merged", author: "", subject: "" } as never,
        steps: steps as never,
      }),
    )!;
    const one = pageEntries(libraryPdf(book.id)!)[0];
    entries.value = Array.from({ length: pdf.MAX_MERGE_PAGES + 1 }, (_, i) => ({ ...one, key: `k${i}` }));
    expect(merge.refusal.value).toBe("A merged PDF can have at most 2000 pages; this one would have 2001. Remove 1 and try again.");
    expect(await merge.save()).toBeNull();
    expect(merge.failure.value).toBe(merge.refusal.value);
    entries.value = [one, ...pageEntries(libraryPdf(huge.id)!)];
    expect(merge.refusal.value).toMatch(/^The PDFs in one merge can add up to at most 500 MB; these add up to 501 MB\./);
    expect(steps.assemble).not.toHaveBeenCalled();
    expect(steps.parse).not.toHaveBeenCalled();
  });

  it("refuses a source that has left the library, naming it", async () => {
    const { seedPdf, open, store } = await setup();
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    assembly.tray.add([book.id]);
    store.deleteDocument(book.id);
    expect(assembly.merge.refusal.value).toBe("“pages” is no longer in the library or could not be read; remove it from the tray.");
  });
});

describe("cancellation and the self-check", () => {
  it("stops at the next page boundary when cancelled, leaving no file, no record and no mutation", async () => {
    const between = vi.fn(async () => {});
    const { seedPdf, open, store, steps } = await setup({ between });
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id]);
    tray.selectAll();
    for (let i = 0; i < 4; i += 1) {
      tray.duplicatePages(keyAt(assembly, 0));
    }
    expect(tray.entries.value.length).toBeGreaterThan(8);
    const progress: string[] = [];
    between.mockImplementation(async () => {
      progress.push(merge.progress.value);
      merge.cancel();
    });
    const count = store.listDocuments().length;
    expect(await merge.save()).toBeNull();
    expect(merge.failure.value).toBe("");
    expect(progress).toEqual([`Page 8 of ${tray.entries.value.length}`]);
    expect(store.listDocuments()).toHaveLength(count);
    expect(merge.busy.value).toBe(false);
    expect(merge.progress.value).toBe("");
    expect(await merge.download()).toBe(false);
    expect(steps.download).not.toHaveBeenCalled();
    // A later run is not cancelled by the earlier one.
    between.mockImplementation(async () => {});
    expect((await merge.save())?.pdf?.pageCount).toBe(tray.entries.value.length);
  });

  it("fails loudly when its own output does not re-parse, or re-parses to other pages, and keeps nothing", async () => {
    const assemble = vi.fn<(plan: PdfAssemblyPlan, hooks?: PdfAssemblyHooks) => Promise<PdfAssemblyResult>>(async () => ({
      bytes: new Uint8Array([0x25, 0x50, 0x44, 0x46]),
      dropped: [],
    }));
    const { seedPdf, open, store, steps, pdf } = await setup({ assemble });
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id]);
    const count = store.listDocuments().length;
    expect(await merge.save()).toBeNull();
    expect(merge.failure.value).toMatch(/^Could not save the PDF: the merged PDF does not read back \(.+\), so nothing was kept\.$/);
    expect(store.listDocuments()).toHaveLength(count);
    expect(await merge.download()).toBe(false);
    expect(steps.download).not.toHaveBeenCalled();
    // A writer that loses a turn is caught by the geometry check.
    assemble.mockImplementation(async (plan: PdfAssemblyPlan) =>
      pdf.assemblePdf({ ...plan, pages: plan.pages.map((page) => ({ ...page, rotateBy: 0 as const })) }, { between: async () => {} }),
    );
    tray.rotate(keyAt(assembly, 0), 90);
    expect(await merge.save()).toBeNull();
    expect(merge.failure.value).toBe(
      "Could not save the PDF: page 1 of the merged PDF is 612 × 792 pt, rotated 0° instead of 612 × 792 pt, rotated 90°, so nothing was kept.",
    );
    expect(store.listDocuments()).toHaveLength(count);
  });
});

describe("the result preview", () => {
  it("rebuilds after a change at most once every PREVIEW_DEBOUNCE_MS, and on Refresh", async () => {
    const { seedPdf, open, steps, wait, PREVIEW_DEBOUNCE_MS } = await setup();
    const book = await seedPdf("pages", PAGES);
    const { assembly, stop } = open();
    const { tray, merge } = assembly;
    tray.add([book.id]);
    await nextTick();
    expect(merge.preview.stale.value).toBe(true);
    await wait(PREVIEW_DEBOUNCE_MS - 1);
    expect(steps.assemble).not.toHaveBeenCalled();
    await wait(1);
    await vi.waitFor(() => expect(merge.preview.url.value).toBe("blob:preview-1"));
    expect(steps.assemble).toHaveBeenCalledTimes(1);
    expect(merge.preview.stale.value).toBe(false);
    // Changes in quick succession: one rebuild, a debounce after the last.
    tray.rotate(keyAt(assembly, 0), 90);
    await nextTick();
    await wait(300);
    tray.rotate(keyAt(assembly, 0), 90);
    await nextTick();
    await wait(300);
    expect(steps.assemble).toHaveBeenCalledTimes(1);
    await wait(PREVIEW_DEBOUNCE_MS);
    await vi.waitFor(() => expect(merge.preview.url.value).toBe("blob:preview-2"));
    expect(steps.assemble).toHaveBeenCalledTimes(2);
    expect(steps.revokeUrl).toHaveBeenCalledWith("blob:preview-1");
    // Refresh at once after a build waits out the debounce.
    merge.preview.refresh();
    await wait(PREVIEW_DEBOUNCE_MS - 1);
    expect(steps.assemble).toHaveBeenCalledTimes(2);
    await wait(1);
    await vi.waitFor(() => expect(steps.assemble).toHaveBeenCalledTimes(3));
    // Leaving frees the URL.
    await vi.waitFor(() => expect(merge.preview.url.value).toBe("blob:preview-3"));
    stop();
    expect(steps.revokeUrl).toHaveBeenCalledWith("blob:preview-3");
  });

  it("never rebuilds by itself past PREVIEW_MAX_PAGES — only on Refresh", async () => {
    const { seedPdf, open, steps, wait, PREVIEW_MAX_PAGES } = await setup();
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id]);
    while (tray.entries.value.length <= PREVIEW_MAX_PAGES) {
      tray.selectAll();
      tray.duplicatePages(keyAt(assembly, 0));
    }
    expect(merge.preview.manual.value).toBe(true);
    await wait(10_000);
    expect(steps.assemble).not.toHaveBeenCalled();
    expect(merge.preview.stale.value).toBe(true);
    merge.preview.refresh();
    await wait(0);
    await vi.waitFor(() => expect(merge.preview.url.value).not.toBe(""));
    expect(steps.assemble).toHaveBeenCalledTimes(1);
  });

  it("shows what changed for the planned output, and nothing when nothing was dropped", async () => {
    const { seedPdf, open, wait, PREVIEW_DEBOUNCE_MS } = await setup();
    const outlines = await seedPdf("outlines", OUTLINES);
    const book = await seedPdf("pages", PAGES);
    const { assembly } = open();
    const { tray, merge } = assembly;
    tray.add([book.id]);
    await wait(PREVIEW_DEBOUNCE_MS);
    await vi.waitFor(() => expect(merge.preview.url.value).not.toBe(""));
    expect(merge.dropGroups.value).toEqual([]);
    // Outlines' page 1 alone: its link to page 2 and more are dropped.
    tray.add([outlines.id]);
    tray.removePages(keyAt(assembly, 4));
    await wait(PREVIEW_DEBOUNCE_MS);
    await vi.waitFor(() => expect(merge.dropGroups.value.length).toBeGreaterThan(0));
    expect(merge.dropGroups.value.map((group) => [group.label, group.records.length])).toEqual([
      ["Structure tree (accessibility tags)", 1],
      ["Named destinations", 1],
      ["Bookmarks", 1],
      ["Links and annotations", 2],
    ]);
  });
});

describe("the persisted plan", () => {
  it("restores the pages, their turns and blank pages, and drops a PDF that has left the library, naming it once", async () => {
    const { seedPdf, open, store } = await setup();
    const outlines = await seedPdf("outlines", OUTLINES);
    const book = await seedPdf("pages", PAGES);
    const first = open();
    const tray = first.assembly.tray;
    tray.add([book.id, outlines.id]);
    tray.rotate(keyAt(first.assembly, 1), -90);
    tray.removePages(keyAt(first.assembly, 0));
    tray.addBlank();
    first.assembly.flush();
    first.stop();

    const second = open();
    expect(second.assembly.tray.entries.value.map((entry) => [entry.title, entry.page?.index ?? null, entry.rotateBy ?? 0, entry.blank ?? false])).toEqual([
      ["pages", 1, 270, false],
      ["pages", 2, 0, false],
      ["outlines", 0, 0, false],
      ["outlines", 1, 0, false],
      ["Blank page", null, 0, true],
    ]);
    expect(second.assembly.output.format).toBe("pdf");
    second.assembly.flush();
    second.stop();

    store.trashDocument(book.id);
    const third = open();
    expect(third.assembly.tray.entries.value.map((entry) => entry.title)).toEqual(["outlines", "outlines", "Blank page"]);
    expect(third.assembly.notice.value).toBe("“pages” is no longer in the library and was removed from the tray.");
  });
});
