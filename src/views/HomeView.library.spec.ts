import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";

// SPEC-003 — the library list at `/`. Fresh modules per test (the store is
// module state); the idb seam is mocked so persistence never interferes.
// The pre-existing empty-state behaviour is covered by HomeView.spec.ts.

beforeEach(() => {
  vi.resetModules();
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

// Navigations triggered by clicks lazy-load the target route component,
// which needs real event-loop time (module fetch), not just promise flushes;
// poll with a real delay (bounded) until the router settles.
async function untilPath(
  router: ReturnType<typeof createRouter>,
  path: string,
) {
  for (let i = 0; i < 200; i += 1) {
    if (router.currentRoute.value.path === path) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

async function loadModules() {
  const store = await import("../lib/documentStore");
  const { default: HomeView } = await import("./HomeView.vue");
  async function mountHome() {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push("/");
    const wrapper = mount(HomeView, { global: { plugins: [router] } });
    await flushPromises();
    return { wrapper, router };
  }
  return { store, mountHome };
}

/** Seeds a document at a fixed timestamp and returns its id. */
function seed(
  store: Awaited<ReturnType<typeof loadModules>>["store"],
  id: string,
  title: string,
  at: number,
) {
  vi.spyOn(Date, "now").mockReturnValue(at);
  const doc = store.getOrCreateDocument(id);
  doc.title = title;
  return id;
}

describe("HomeView library list", () => {
  it("shows the empty state, not the library, when no documents exist", async () => {
    const { mountHome } = await loadModules();
    const { wrapper } = await mountHome();
    expect(wrapper.find('[data-testid="empty-state"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="library-list"]').exists()).toBe(false);
  });

  it("lists stored documents with title, relative time and badge (AC-3..AC-6)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-a", "Quarterly report", 1_000_000);
    vi.spyOn(Date, "now").mockReturnValue(1_000_000 + 2 * 60 * 60 * 1000);
    const { wrapper } = await mountHome();

    expect(wrapper.find('[data-testid="empty-state"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="library-list"]').exists()).toBe(true);
    const headers = wrapper
      .findAll("th")
      .map((header) => header.text());
    expect(headers).toContain("Title");
    expect(headers).toContain("Last modified");
    expect(headers).toContain("Format");
    expect(wrapper.get('[data-testid="row-title"]').text()).toBe(
      "Quarterly report",
    );
    expect(wrapper.get('[data-testid="row-modified"]').text()).toBe(
      "2 hours ago",
    );
    expect(wrapper.get('[data-testid="row-format"]').text()).toBe("New");
  });

  it("sorts rows by last modified, newest first (AC-7)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-old", "Older", 1000);
    seed(store, "doc-new", "Newer", 2000);
    const { wrapper } = await mountHome();

    const titles = wrapper
      .findAll('[data-testid="row-title"]')
      .map((cell) => cell.text());
    expect(titles).toEqual(["Newer", "Older"]);
  });

  it("clicking a row navigates to /doc/:id (AC-8)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-open", "Openable", 1000);
    const { wrapper, router } = await mountHome();

    await wrapper.get('[data-testid="library-row"]').trigger("click");
    await untilPath(router, "/doc/doc-open");
    expect(router.currentRoute.value.path).toBe("/doc/doc-open");
  });

  it("the overflow menu offers Rename, Duplicate and Delete (AC-10..AC-13)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-menu", "With menu", 1000);
    const { wrapper } = await mountHome();

    expect(wrapper.find('[data-testid="row-menu-list"]').exists()).toBe(false);
    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    expect(wrapper.get('[data-testid="menu-rename"]').text()).toBe("Rename");
    expect(wrapper.get('[data-testid="menu-duplicate"]').text()).toBe(
      "Duplicate",
    );
    expect(wrapper.get('[data-testid="menu-delete"]').text()).toBe("Delete");
  });

  it("renames via the modal, prefilled with the current title", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-rename", "Old name", 1000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-rename"]').trigger("click");

    const input = wrapper.get('[data-testid="rename-input"]');
    expect((input.element as HTMLInputElement).value).toBe("Old name");
    await input.setValue("New name");
    await wrapper.get('[data-testid="rename-save"]').trigger("click");

    expect(store.getDocument("doc-rename")?.title).toBe("New name");
    expect(wrapper.get('[data-testid="row-title"]').text()).toBe("New name");
    expect(wrapper.find('[data-testid="rename-dialog"]').exists()).toBe(false);
  });

  it("appends a count when renaming to a taken title", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-taken", "Report", 1000);
    seed(store, "doc-rename", "Something else", 2000);
    const { wrapper } = await mountHome();

    // Row 0 is "Something else" (newer); rename it to the taken title.
    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-rename"]').trigger("click");
    await wrapper.get('[data-testid="rename-input"]').setValue("Report");
    await wrapper.get('[data-testid="rename-save"]').trigger("click");

    expect(store.getDocument("doc-rename")?.title).toBe("Report (2)");
  });

  it("blocks an empty title in the rename modal with an error", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-rename", "Keep me", 1000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-rename"]').trigger("click");
    const input = wrapper.get('[data-testid="rename-input"]');
    await input.setValue("   ");

    expect(wrapper.get('[data-testid="rename-error"]').text()).toBe(
      "Title cannot be empty",
    );
    const save = wrapper.get('[data-testid="rename-save"]');
    expect(save.attributes("disabled")).toBeDefined();
    await save.trigger("click");
    expect(store.getDocument("doc-rename")?.title).toBe("Keep me");
  });

  it("renaming bumps the row to the top of the list", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-old", "Renamed later", 1000);
    seed(store, "doc-new", "Newer", 2000);
    const { wrapper } = await mountHome();

    // Row 1 is the older document; open its menu and rename it.
    vi.spyOn(Date, "now").mockReturnValue(3000);
    const rows = wrapper.findAll('[data-testid="library-row"]');
    await rows[1]!.get('[data-testid="row-menu"]').trigger("click");
    await rows[1]!.get('[data-testid="menu-rename"]').trigger("click");
    await wrapper.get('[data-testid="rename-input"]').setValue("Front now");
    await wrapper.get('[data-testid="rename-save"]').trigger("click");

    const titles = wrapper
      .findAll('[data-testid="row-title"]')
      .map((cell) => cell.text());
    expect(titles).toEqual(["Front now", "Newer"]);
  });

  it("duplicates a document as '<title> (copy)' at the top (AC-16, AC-17)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-dup", "Original", 1000);
    const { wrapper } = await mountHome();

    vi.spyOn(Date, "now").mockReturnValue(2000);
    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-duplicate"]').trigger("click");

    const titles = wrapper
      .findAll('[data-testid="row-title"]')
      .map((cell) => cell.text());
    expect(titles).toEqual(["Original (copy)", "Original"]);
    const copy = store
      .listDocuments()
      .find((entry) => entry.title === "Original (copy)");
    expect(copy).toBeDefined();
    expect(copy!.id).not.toBe("doc-dup");
  });

  // SPEC-067 amends SPEC-003's delete AC: deleting from the library moves the
  // document to the trash (not a confirm-then-erase), with an Undo toast. The
  // document is kept intact, so it can be restored. Permanent deletion is a
  // separate, explicit step from the Trash view (covered in HomeView.trash.spec).
  it("deleting moves the document to trash rather than erasing it (AC-33)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-del", "Doomed doc", 1000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-delete"]').trigger("click");

    // Gone from the library list, but kept in the store and now in the trash.
    expect(store.listDocuments()).toEqual([]);
    expect(store.getDocument("doc-del")?.title).toBe("Doomed doc");
    expect(store.listTrashedDocuments().map((e) => e.id)).toEqual(["doc-del"]);
    // No confirm dialog for a library delete — the trash is the safety net.
    expect(wrapper.find('[data-testid="delete-dialog"]').exists()).toBe(false);
  });

  it("offers an Undo toast that restores the deleted document (AC-34, AC-35)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-del", "Bring me back", 1000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="row-menu"]').trigger("click");
    await wrapper.get('[data-testid="menu-delete"]').trigger("click");

    const toast = wrapper.get('[data-testid="toast"]');
    expect(toast.text()).toContain("Bring me back");
    await wrapper.get('[data-testid="toast-action"]').trigger("click");

    expect(store.listTrashedDocuments()).toEqual([]);
    expect(store.listDocuments().map((e) => e.id)).toEqual(["doc-del"]);
    expect(wrapper.get('[data-testid="row-title"]').text()).toBe("Bring me back");
  });

  // SPEC-060: returning from a closed document carries its id in router state;
  // the library flashes that row so a user sees where they came from (AC-22).
  it("highlights the just-closed row on return, tagging every row with its id (AC-21, AC-22)", async () => {
    const { store, mountHome } = await loadModules();
    seed(store, "doc-a", "Alpha", 2000);
    seed(store, "doc-b", "Bravo", 1000);
    window.history.replaceState({ closedDocId: "doc-b" }, "");
    try {
      const { wrapper } = await mountHome();
      await flushPromises();
      const rows = wrapper.findAll("[data-doc-id]");
      // Every row is addressable by its document id (AC-21's scroll target).
      expect(rows.map((row) => row.attributes("data-doc-id"))).toEqual([
        "doc-a",
        "doc-b",
      ]);
      // Only the closed row is flashed.
      expect(wrapper.get('[data-doc-id="doc-b"]').classes()).toContain(
        "row-highlight",
      );
      expect(wrapper.get('[data-doc-id="doc-a"]').classes()).not.toContain(
        "row-highlight",
      );
    } finally {
      window.history.replaceState(null, "");
    }
  });
});
