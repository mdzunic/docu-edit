// SPEC-107 in the library (AC-9): media entries are first-class rows and
// cards — badges, the kind-grouped "Images"/"PDFs" chips, the size column,
// and the shared search/sort/star/select/trash machinery over MIXED kinds
// with no new code path. Same harness as HomeView.spec067.spec.ts.
import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";

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
  vi.restoreAllMocks();
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
    await flushPromises();
    return { wrapper, router };
  }
  async function seedMedia(title: string, format: "png" | "pdf", at: number) {
    const bytes =
      format === "pdf" ? fixtures.buildPagesPdfBytes() : await fixtures.buildPngFixtureBytes();
    const model = await assembleMediaModel(title, format, bytes);
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.storeMedia(model);
    return model;
  }
  function seedText(id: string, title: string, at: number) {
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.getOrCreateDocument(id).title = title;
  }
  return { store, seedMedia, seedText, mountHome };
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
    await expect(store.saveDraft(png.id, null, 0)).rejects.toThrow(/media file/);
    expect(store.getDocument(png.id)).toBeUndefined();
    expect(store.getPresentation(png.id)).toBeUndefined();
    expect(store.getMedia(png.id)).toBeDefined();
  });
});
