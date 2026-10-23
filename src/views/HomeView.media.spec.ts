// SPEC-107 in the library: media entries are first-class rows and cards —
// badges, the kind-grouped "Images"/"PDFs" chips, the size column, honest
// per-kind thumbnails, and the shared search/sort/star/select/trash/download
// machinery over MIXED kinds with no new code path. Same harness as
// HomeView.spec067.spec.ts.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { equalBytes } from "../lib/core/equality";

// The checked-in spec107-pages.pdf (the app's own buildPdf output).
const PAGES_PDF = new Uint8Array(readFileSync(join(__dirname, "../../e2e/fixtures/spec107-pages.pdf")));

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

afterEach(() => {
  vi.doUnmock("../lib/idb");
  vi.doUnmock("../lib/core/triggerDownload");
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

async function loadModules() {
  const store = await import("../lib/documentStore");
  const { assembleMediaModel } = await import("../lib/media/open");
  const fixtures = await import("../lib/media/fixtures/build");
  const { default: HomeView } = await import("./HomeView.vue");
  async function mountHome() {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push("/");
    const wrapper = mount(HomeView, { global: { plugins: [router] }, attachTo: document.body });
    // HomeView preloads the editor chunk on mount (SPEC-047 AC-3) and never
    // awaits it. vitest 4 fails the run when an import settles after the
    // file's environment is torn down — which the full suite's load makes
    // likely for this file's last mount — so join that same in-flight import
    // here. Here rather than in afterEach: the cold DocumentView import is
    // 3-4 s idle and far slower under load (HomeView.library.spec.ts), too
    // close to the 10 s hook timeout, while a test gets 30 s.
    await import("./DocumentView.vue");
    await flushPromises();
    return { wrapper, router };
  }
  async function seedMedia(title: string, format: "png" | "pdf", at: number) {
    const bytes =
      format === "pdf" ? PAGES_PDF : await fixtures.buildPngFixtureBytes();
    const model = await assembleMediaModel(title, format, bytes);
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.storeMedia(model);
    return model;
  }
  function seedText(id: string, title: string, at: number) {
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.getOrCreateDocument(id).title = title;
  }
  return { store, fixtures, seedMedia, seedText, mountHome };
}

async function selectAllRows(wrapper: { findAll: (selector: string) => Array<{ setValue: (value: boolean) => Promise<void> }> }) {
  for (const check of wrapper.findAll('[data-testid="row-select"]')) {
    await check.setValue(true);
  }
}

describe("media rows and cards (AC-9)", () => {
  it("lists media beside documents with format badges and sizes", async () => {
    const { seedMedia, seedText, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    const png = await seedMedia("Shot", "png", 2000);
    await seedMedia("Pages", "pdf", 1000);
    const { wrapper } = await mountHome();

    const badges = wrapper.findAll('[data-testid="row-format"]').map((c) => c.text());
    expect(badges).toContain("PNG");
    expect(badges).toContain("PDF");
    // The size column reads formatFileSize(byteLength) for media rows.
    const rows = wrapper.findAll('[data-testid="library-row"]');
    const pngRow = rows.find((row) => row.attributes("data-doc-id") === png.id);
    expect(pngRow?.get('[data-testid="row-size"]').text()).toBe(`${Math.round(png.byteLength / 1024)} KB`);
  });

  it('groups the media chips into one-click "Images" and "PDFs" filters', async () => {
    const { seedMedia, seedText, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    await seedMedia("Shot", "png", 2000);
    await seedMedia("Pages", "pdf", 1000);
    const { wrapper } = await mountHome();

    const imagesChip = wrapper.get('[data-testid="filter-images"]');
    expect(imagesChip.text()).toContain("Images");
    expect(imagesChip.text()).toContain("1");
    await imagesChip.trigger("click");
    let titles = wrapper.findAll('[data-testid="row-title"]').map((c) => c.text());
    expect(titles).toEqual(["Shot"]);

    await imagesChip.trigger("click"); // off again
    await wrapper.get('[data-testid="filter-pdf"]').trigger("click");
    titles = wrapper.findAll('[data-testid="row-title"]').map((c) => c.text());
    expect(titles).toEqual(["Pages"]);
  });

  it("search, starring and multi-select span mixed kinds with no special-casing", async () => {
    const { store, seedMedia, seedText, mountHome } = await loadModules();
    seedText("doc-1", "Quarterly Report", 3000);
    const png = await seedMedia("Quarterly Shot", "png", 2000);
    const { wrapper } = await mountHome();

    // Subsequence search matches the media title too.
    await wrapper.get('[data-testid="library-search"]').setValue("qshot");
    expect(wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())).toEqual([
      "Quarterly Shot",
    ]);
    await wrapper.get('[data-testid="library-search"]').setValue("");

    // Star the media row: it sorts first, the flag persists in the store.
    const rows = wrapper.findAll('[data-testid="library-row"]');
    const pngRow = rows.find((row) => row.attributes("data-doc-id") === png.id);
    await pngRow?.get('[data-testid="row-star"]').trigger("click");
    expect(store.getStoredRecord(png.id)?.starred).toBe(true);
    expect(
      wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())[0],
    ).toBe("Quarterly Shot");

    // Select both rows: the bulk bar counts them.
    for (const row of wrapper.findAll('[data-testid="row-select"]')) {
      await row.setValue(true);
    }
    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("2");
  });

  it("trash and restore work on a media entry, model intact", async () => {
    const { store, seedMedia, mountHome } = await loadModules();
    const png = await seedMedia("Shot", "png", 2000);
    const { wrapper } = await mountHome();

    store.trashDocument(png.id);
    await flushPromises();
    expect(store.listDocuments()).toHaveLength(0);
    store.restoreDocument(png.id);
    expect(store.getMedia(png.id)?.byteLength).toBe(png.byteLength);
    expect(store.listDocuments()).toHaveLength(1);
    wrapper.unmount();
  });
});

describe("media never enters document paths (AC-13)", () => {
  it("getOrCreateDocument and saveDraft throw loudly on a media id", async () => {
    const { store, seedMedia } = await loadModules();
    const png = await seedMedia("Shot", "png", 2000);
    expect(() => store.getOrCreateDocument(png.id)).toThrow(/media file/);
    await expect(store.saveDraft(png.id, { caret: null, scroll: 0 })).rejects.toThrow(/media file/);
    expect(store.getDocument(png.id)).toBeUndefined();
    expect(store.getPresentation(png.id)).toBeUndefined();
    expect(store.getMedia(png.id)).toBeDefined();
  });
});

describe("honest thumbnails per kind (AC-28, AC-29)", () => {
  it("an image card renders a real raster, downscaled to MEDIA_THUMBNAIL_MAX_EDGE_PX, produced once per record", async () => {
    // jsdom decodes no images and has no 2D canvas: stand both in, with a
    // source larger than the cap.
    const drawn: number[][] = [];
    class FakeImage extends EventTarget {
      naturalWidth = 2048;
      naturalHeight = 1024;
      set src(_value: string) {
        queueMicrotask(() => this.dispatchEvent(new Event("load")));
      }
    }
    vi.stubGlobal("Image", FakeImage);
    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
      () => ({ drawImage: (...args: number[]) => drawn.push(args.slice(1)) }) as never,
    );
    vi.spyOn(HTMLCanvasElement.prototype, "toDataURL").mockReturnValue("data:image/png;base64,UkFTVEVS");
    const { seedMedia, mountHome } = await loadModules();
    const { MEDIA_THUMBNAIL_MAX_EDGE_PX } = await import("../lib/media/thumbnail");
    await seedMedia("Shot", "png", 2000);
    const first = await mountHome();
    await first.wrapper.get('[data-testid="view-grid"]').trigger("click");

    await vi.waitFor(() => expect(first.wrapper.find('[data-testid="media-thumbnail"]').exists()).toBe(true));
    expect(first.wrapper.get('[data-testid="media-thumbnail"]').attributes("src")).toBe("data:image/png;base64,UkFTVEVS");
    expect(drawn).toEqual([[0, 0, MEDIA_THUMBNAIL_MAX_EDGE_PX, MEDIA_THUMBNAIL_MAX_EDGE_PX / 2]]);
    first.wrapper.unmount();

    // Remounting the library reuses the cached raster (id + lastModified).
    const second = await mountHome();
    expect(second.wrapper.get('[data-testid="media-thumbnail"]').attributes("src")).toBe("data:image/png;base64,UkFTVEVS");
    expect(drawn).toHaveLength(1);
    second.wrapper.unmount();
  });

  it("a PDF card is a paper-shaped placeholder at its first page's real aspect, with page count and size", async () => {
    const { seedMedia, mountHome } = await loadModules();
    await seedMedia("Pages", "pdf", 1000);
    const { wrapper } = await mountHome();
    await wrapper.get('[data-testid="view-grid"]').trigger("click");

    // jsdom has no Worker, so no raster can be produced: the placeholder is
    // all the card shows — it never fakes a preview.
    await vi.waitFor(() => expect(wrapper.find('[data-testid="pdf-placeholder"]').exists()).toBe(true));
    const placeholder = wrapper.get('[data-testid="pdf-placeholder"]');
    expect(placeholder.text()).toContain("3 pages");
    expect(placeholder.text()).toContain("612 × 792 pt");
    // The ratio the paper is sized from (jsdom has no layout; e2e measures
    // the rendered box).
    const paper = placeholder.get('[data-testid="pdf-paper"]').element as HTMLElement;
    expect(Number(paper.style.getPropertyValue("--paper-aspect"))).toBeCloseTo(612 / 792);
    expect(wrapper.find('[data-testid="media-thumbnail"]').exists()).toBe(false);
    await vi.waitFor(() => expect(wrapper.get('[data-testid="row-counts"]').text()).toContain("3 pages"));
    wrapper.unmount();
  });
});

describe("bulk actions over mixed kinds (AC-27)", () => {
  it("bulk download sends each kind down its own path, media byte-for-byte", async () => {
    const downloads: Array<{ bytes: Uint8Array; name: string; type: string }> = [];
    vi.doMock("../lib/core/triggerDownload", () => ({
      triggerBlobDownload: (bytes: Uint8Array, name: string, type: string) => downloads.push({ bytes, name, type }),
    }));
    const { fixtures, seedMedia, seedText, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    await seedMedia("Shot", "png", 2000);
    await seedMedia("Pages", "pdf", 1000);
    const { wrapper } = await mountHome();

    await selectAllRows(wrapper);
    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("3");
    await wrapper.get('[data-testid="bulk-download"]').trigger("click");
    await vi.waitFor(() => expect(downloads).toHaveLength(3));

    const byName = new Map(downloads.map((download) => [download.name, download]));
    expect(byName.get("Pages.pdf")?.type).toBe("application/pdf");
    expect(equalBytes(byName.get("Pages.pdf")?.bytes ?? new Uint8Array(), PAGES_PDF)).toBe(true);
    expect(byName.get("Shot.png")?.type).toBe("image/png");
    expect(equalBytes(byName.get("Shot.png")?.bytes ?? new Uint8Array(), await fixtures.buildPngFixtureBytes())).toBe(true);
    expect(byName.has("Report.docx")).toBe(true);
    wrapper.unmount();
  });

  it("bulk Move to trash and its Undo take every kind in one batch", async () => {
    const { store, seedMedia, seedText, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    const png = await seedMedia("Shot", "png", 2000);
    const pdf = await seedMedia("Pages", "pdf", 1000);
    const { wrapper } = await mountHome();

    await selectAllRows(wrapper);
    await wrapper.get('[data-testid="bulk-trash"]').trigger("click");
    expect(store.listDocuments()).toEqual([]);
    expect(store.listTrashedDocuments().map((entry) => entry.kind).toSorted()).toEqual(["media", "media", "text"]);

    await wrapper.get('[data-testid="toast-action"]').trigger("click");
    expect(store.listDocuments().map((entry) => entry.id).toSorted()).toEqual(["doc-1", png.id, pdf.id].toSorted());
    expect(store.getMedia(pdf.id)?.pdf?.pageCount).toBe(3);
    wrapper.unmount();
  });
});
