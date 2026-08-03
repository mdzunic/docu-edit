import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";

// SPEC-003 — deleting the currently open document (AC-19, AC-20) and
// persisting title edits. Fresh modules per test; the idb seam is mocked so
// write-through calls can be asserted.

const putRecord = vi.fn();

beforeEach(() => {
  vi.resetModules();
  putRecord.mockReset();
  putRecord.mockResolvedValue(undefined);
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord,
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  // Every routable id resolves to a document (SPEC-001 decision); the stub
  // itself is tested in ../lib/documents.spec.ts.
  vi.doMock("../lib/documents", () => ({ documentExists: () => true }));
});

afterEach(() => {
  vi.doUnmock("../lib/idb");
  vi.doUnmock("../lib/documents");
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

async function mountAt(path: string) {
  const store = await import("../lib/documentStore");
  const { default: DocumentView } = await import("./DocumentView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  const wrapper = mount(DocumentView, { global: { plugins: [router] } });
  await flushPromises();
  return { wrapper, router, store };
}

describe("DocumentView document deletion", () => {
  it("offers a Delete action that confirms naming the document (AC-14)", async () => {
    const { wrapper } = await mountAt("/doc/open-doc");
    expect(wrapper.find('[data-testid="delete-dialog"]').exists()).toBe(false);

    await wrapper.get('[data-testid="delete-document"]').trigger("click");
    expect(wrapper.get('[data-testid="delete-message"]').text()).toContain(
      "Untitled document",
    );
  });

  it("confirming removes the doc, navigates to / and toasts (AC-15, AC-19, AC-20)", async () => {
    const { wrapper, router, store } = await mountAt("/doc/open-doc");

    await wrapper.get('[data-testid="delete-document"]').trigger("click");
    await wrapper.get('[data-testid="delete-confirm"]').trigger("click");
    await untilPath(router, "/");

    expect(store.getDocument("open-doc")).toBeUndefined();
    expect(router.currentRoute.value.path).toBe("/");
    expect(wrapper.get('[data-testid="toast"]').text()).toBe(
      "Deleted Untitled document",
    );
  });

  it("the toast names the document's current title", async () => {
    const { wrapper } = await mountAt("/doc/named-doc");

    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Meeting Notes");
    await input.trigger("keydown.enter");

    await wrapper.get('[data-testid="delete-document"]').trigger("click");
    await wrapper.get('[data-testid="delete-confirm"]').trigger("click");
    await flushPromises();

    expect(wrapper.get('[data-testid="toast"]').text()).toBe(
      "Deleted Meeting Notes",
    );
  });

  it("cancelling keeps the document open and stored", async () => {
    const { wrapper, router, store } = await mountAt("/doc/open-doc");

    await wrapper.get('[data-testid="delete-document"]').trigger("click");
    await wrapper.get('[data-testid="delete-cancel"]').trigger("click");

    expect(wrapper.find('[data-testid="delete-dialog"]').exists()).toBe(false);
    expect(store.getDocument("open-doc")).toBeDefined();
    expect(router.currentRoute.value.path).toBe("/doc/open-doc");
  });
});

describe("DocumentView title persistence", () => {
  it("persists a title edit with a bumped lastModified", async () => {
    vi.spyOn(Date, "now").mockReturnValue(1000);
    const { wrapper } = await mountAt("/doc/title-doc");

    vi.spyOn(Date, "now").mockReturnValue(2000);
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Persisted title");
    await input.trigger("keydown.enter");
    await flushPromises();

    expect(putRecord).toHaveBeenLastCalledWith(
      expect.objectContaining({
        id: "title-doc",
        lastModified: 2000,
        model: expect.objectContaining({ title: "Persisted title" }),
      }),
    );
  });
});
