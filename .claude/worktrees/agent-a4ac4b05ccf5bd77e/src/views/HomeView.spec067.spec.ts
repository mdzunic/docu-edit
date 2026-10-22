import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";

// SPEC-067 — the redesigned document library: search, sort, filters, starring,
// multi-select + bulk actions, trash, storage usage and the empty/no-results
// states. Fresh modules per test (the store is module state); idb mocked so
// persistence never interferes; localStorage cleared so preferences start fresh.

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
  const prefs = await import("../lib/libraryPrefs");
  const { default: HomeView } = await import("./HomeView.vue");
  async function mountHome() {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push("/");
    const wrapper = mount(HomeView, { global: { plugins: [router] }, attachTo: document.body });
    await flushPromises();
    return { wrapper, router };
  }
  return { store, prefs, mountHome };
}

function seed(
  store: Awaited<ReturnType<typeof loadModules>>["store"],
  id: string,
  title: string,
  at: number,
) {
  vi.spyOn(Date, "now").mockReturnValue(at);
  store.getOrCreateDocument(id).title = title;
  return id;
}

describe("search (AC-8, AC-9, AC-10)", () => {
  it("filters by subsequence, highlights matches, and restores on clear", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "Quarterly Report", 3000);
    seed(store, "b", "Budget", 2000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="library-search"]').setValue("qrt");
    const titles = wrapper.findAll('[data-testid="row-title"]').map((c) => c.text());
    expect(titles).toEqual(["Quarterly Report"]);
    // Matched characters are highlighted.
    expect(wrapper.findAll('[data-testid="title-match"]').length).toBeGreaterThan(0);

    await wrapper.get('[data-testid="library-search"]').setValue("");
    expect(wrapper.findAll('[data-testid="row-title"]').length).toBe(2);
  });

  it("shows a no-results state that clears the filter (AC-54, AC-55)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "Report", 1000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="library-search"]').setValue("zzzzz");
    expect(wrapper.find('[data-testid="no-results"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="library-list"]').exists()).toBe(false);

    await wrapper.get('[data-testid="clear-filters"]').trigger("click");
    expect(wrapper.find('[data-testid="no-results"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="library-list"]').exists()).toBe(true);
  });
});

describe("filters (AC-16, AC-17)", () => {
  it("filters by format chip with a live count", async () => {
    const { store, mountHome } = await loadModules();
    const a = seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    store.getStoredRecord(a)!.format = "docx";
    const { wrapper } = await mountHome();

    const docxChip = wrapper.get('[data-testid="filter-docx"]');
    expect(docxChip.text()).toContain("1");
    await docxChip.trigger("click");
    expect(wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())).toEqual(["One"]);
  });

  it("filters by starred chip", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "Starred one", 3000);
    seed(store, "b", "Plain", 2000);
    store.setStarred("a", true);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="filter-starred"]').trigger("click");
    expect(wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())).toEqual(["Starred one"]);
  });
});

describe("starring (AC-18, AC-21, AC-22)", () => {
  it("stars from the row and sorts starred first, persisting the flag", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "Alpha", 3000);
    seed(store, "b", "Bravo", 2000);
    const { wrapper } = await mountHome();

    // Star the older document (row 1) from its row.
    const stars = wrapper.findAll('[data-testid="row-star"]');
    await stars[1]!.trigger("click");

    expect(store.getStoredRecord("b")?.starred).toBe(true);
    // Starred sorts first within the active (newest-first) sort.
    expect(wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())).toEqual(["Bravo", "Alpha"]);
  });
});

describe("sort + view persistence (AC-2, AC-13, AC-14, AC-15)", () => {
  it("sorts by size, toggles direction, and persists", async () => {
    const { store, prefs, mountHome } = await loadModules();
    const a = seed(store, "a", "Small", 3000);
    const b = seed(store, "b", "Large", 2000);
    store.getDocument(b)!.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "x".repeat(4000), formatting: {} }] },
    ] as never;
    void a;
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="sort-key"]').setValue("size");
    // Default size direction is descending → largest first.
    expect(wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())).toEqual(["Large", "Small"]);
    expect(prefs.loadLibrarySort().key).toBe("size");

    await wrapper.get('[data-testid="sort-direction"]').trigger("click");
    expect(wrapper.findAll('[data-testid="row-title"]').map((c) => c.text())).toEqual(["Small", "Large"]);
    expect(prefs.loadLibrarySort().direction).toBe("asc");
  });

  it("toggles to the grid view and persists the choice", async () => {
    const { store, prefs, mountHome } = await loadModules();
    seed(store, "a", "One", 1000);
    const { wrapper } = await mountHome();

    expect(wrapper.find('[data-testid="library-list"]').exists()).toBe(true);
    await wrapper.get('[data-testid="view-grid"]').trigger("click");
    expect(wrapper.find('[data-testid="library-grid"]').exists()).toBe(true);
    expect(prefs.loadLibraryView()).toBe("grid");
  });
});

describe("grid click previews in quick look; double-click opens", () => {
  it("a plain card click selects the card into the quick-look pane without navigating", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    const { wrapper, router } = await mountHome();
    await wrapper.get('[data-testid="view-grid"]').trigger("click");

    await wrapper.get('[data-doc-id="a"]').trigger("click");
    await flushPromises();
    expect(router.currentRoute.value.path).toBe("/");
    expect(wrapper.get('[data-testid="quick-look-title"]').text()).toBe("One");

    // Clicking another card moves the preview — the selection is replaced,
    // never accumulated.
    await wrapper.get('[data-doc-id="b"]').trigger("click");
    expect(wrapper.get('[data-testid="quick-look-title"]').text()).toBe("Two");
    expect(wrapper.find('[data-testid="quick-look-multiple"]').exists()).toBe(false);
  });

  it("a double-click opens the document", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    const { wrapper, router } = await mountHome();
    await wrapper.get('[data-testid="view-grid"]').trigger("click");

    // Await the navigation the dblclick started — the push promise is the
    // deterministic "navigation settled" signal.
    const pushSpy = vi.spyOn(router, "push");
    await wrapper.get('[data-doc-id="a"]').trigger("dblclick");
    await pushSpy.mock.results[0]?.value;
    expect(router.currentRoute.value.path).toBe("/doc/a");
  });

  it("a multi-selection shows the no-preview message instead of one entry", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    const { wrapper } = await mountHome();
    await wrapper.get('[data-testid="view-grid"]').trigger("click");

    const checks = wrapper.findAll('[data-testid="row-select"]');
    await checks[0]!.setValue(true);
    await checks[1]!.setValue(true);
    const message = wrapper.get('[data-testid="quick-look-multiple"]');
    expect(message.text()).toContain("2 documents selected");
    expect(wrapper.find('[data-testid="quick-look-title"]').exists()).toBe(false);
  });
});

describe("multi-select + bulk actions (AC-23, AC-26, AC-27..32)", () => {
  it("selects rows and shows a bulk bar with a count and Clear", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    const { wrapper } = await mountHome();

    expect(wrapper.find('[data-testid="selection-bar"]').exists()).toBe(false);
    const checks = wrapper.findAll('[data-testid="row-select"]');
    await checks[0]!.setValue(true);
    await checks[1]!.setValue(true);

    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("2");
    for (const id of ["bulk-download", "bulk-duplicate", "bulk-star", "bulk-trash", "clear-selection"]) {
      expect(wrapper.find(`[data-testid="${id}"]`).exists()).toBe(true);
    }

    await wrapper.get('[data-testid="clear-selection"]').trigger("click");
    expect(wrapper.find('[data-testid="selection-bar"]').exists()).toBe(false);
  });

  it("Ctrl/Cmd+A selects everything matching the filter (AC-26)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="library-list"]').trigger("keydown", { key: "a", ctrlKey: true });
    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("2");
  });

  it("bulk Move to trash trashes the selection", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    const { wrapper } = await mountHome();

    const checks = wrapper.findAll('[data-testid="row-select"]');
    await checks[0]!.setValue(true);
    await checks[1]!.setValue(true);
    await wrapper.get('[data-testid="bulk-trash"]').trigger("click");

    expect(store.listDocuments()).toEqual([]);
    expect(store.listTrashedDocuments().length).toBe(2);
  });
});

describe("trash view (AC-36..39, AC-44, AC-45, AC-46)", () => {
  it("restores a document from the trash", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "Trashed", 1000);
    store.trashDocument("a");
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="trash-link"]').trigger("click");
    expect(wrapper.get('[data-testid="trash-view"]').text()).toContain("Trashed");
    await wrapper.get('[data-testid="row-restore"]').trigger("click");

    expect(store.listDocuments().map((e) => e.id)).toEqual(["a"]);
    expect(store.listTrashedDocuments()).toEqual([]);
  });

  it("permanently deletes after a single confirmation naming the document", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "Doomed", 1000);
    store.trashDocument("a");
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="trash-link"]').trigger("click");
    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-delete-permanently"]').trigger("click");
    expect(wrapper.get('[data-testid="delete-message"]').text()).toContain("Doomed");
    await wrapper.get('[data-testid="delete-confirm"]').trigger("click");

    expect(store.getDocument("a")).toBeUndefined();
    expect(store.listTrashedDocuments()).toEqual([]);
  });

  it("empties the trash after one confirmation", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 2000);
    seed(store, "b", "Two", 1000);
    store.trashDocument("a");
    store.trashDocument("b");
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="trash-link"]').trigger("click");
    await wrapper.get('[data-testid="empty-trash"]').trigger("click");
    await wrapper.get('[data-testid="delete-confirm"]').trigger("click");

    expect(store.getDocument("a")).toBeUndefined();
    expect(store.getDocument("b")).toBeUndefined();
  });
});

describe("storage usage (AC-47)", () => {
  it("shows how much storage the library occupies", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 1000);
    const { wrapper } = await mountHome();
    expect(wrapper.find('[data-testid="storage-panel"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="storage-usage"]').text()).toMatch(/B|KB|MB|GB/);
  });

  it("meters the quota share with a progress bar", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 1000);
    // jsdom ships no Storage API — stub the estimate the panel reads on mount.
    const totalBytes = store.libraryStorageUsage().totalBytes;
    Object.defineProperty(navigator, "storage", {
      configurable: true,
      value: { estimate: () => Promise.resolve({ usage: totalBytes, quota: totalBytes * 10 }) },
    });
    try {
      const { wrapper } = await mountHome();
      const bar = wrapper.get('[data-testid="storage-bar"]');
      expect(bar.attributes("role")).toBe("progressbar");
      expect(bar.attributes("aria-valuenow")).toBe("10");
      expect(bar.attributes("aria-valuemax")).toBe("100");
      expect(wrapper.get('[data-testid="storage-usage"]').text()).toContain("(10%) in use");
      expect(wrapper.get('[data-testid="storage-bar-fill"]').attributes("style")).toContain("width: 10%");
    } finally {
      Reflect.deleteProperty(navigator, "storage");
    }
  });
});

describe("keyboard navigation (AC-58, AC-59, AC-60)", () => {
  it("Space selects, Delete trashes the focused row", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "a", "One", 3000);
    seed(store, "b", "Two", 2000);
    const { wrapper } = await mountHome();

    const rows = wrapper.findAll('[data-testid="library-row"]');
    await rows[0]!.trigger("focus");
    await wrapper.get('[data-testid="library-list"]').trigger("keydown", { key: " " });
    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("1");

    await wrapper.get('[data-testid="library-list"]').trigger("keydown", { key: "Delete" });
    expect(store.listTrashedDocuments().length).toBe(1);
  });
});

// SPEC-068 AC-37/38/39 — the compact tier. The tier keys off window width at
// module-import time (jsdom has no layout for the ResizeObserver to measure),
// so we seed innerWidth to a compact value before loading the fresh modules.
describe("compact tier (SPEC-068)", () => {
  it("collapses the header actions to a primary + overflow and keeps search visible", async () => {
    const original = window.innerWidth;
    Object.defineProperty(window, "innerWidth", { value: 390, configurable: true, writable: true });
    try {
      const { store, mountHome } = await loadModules();
      seed(store, "a", "One", 3000);
      const { wrapper } = await mountHome();

      // The single primary action and the search box remain (AC-39).
      expect(wrapper.find('[data-testid="new-document"]').exists()).toBe(true);
      expect(wrapper.find('[data-testid="library-search"]').exists()).toBe(true);
      // The other two actions collapse behind an overflow menu (AC-38): not inline.
      expect(wrapper.find('[data-testid="library-actions-overflow"]').exists()).toBe(true);
      expect(wrapper.find('[data-testid="new-from-template"]').exists()).toBe(false);
      expect(wrapper.find('[data-testid="open-document"]').exists()).toBe(false);

      // Opening the overflow exposes them, still reachable (AC-38, AC-11).
      await wrapper.get('[data-testid="library-actions-overflow"]').trigger("click");
      expect(wrapper.find('[data-testid="new-from-template"]').exists()).toBe(true);
      expect(wrapper.find('[data-testid="open-document"]').exists()).toBe(true);
    } finally {
      Object.defineProperty(window, "innerWidth", { value: original, configurable: true, writable: true });
    }
  });
});
