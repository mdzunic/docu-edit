import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import type { Paragraph } from "../lib/documentModel";

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
    loadAllTemplateRecords: vi.fn().mockResolvedValue([]),
    putTemplateRecord: vi.fn().mockResolvedValue(undefined),
    deleteTemplateRecord: vi.fn().mockResolvedValue(undefined),
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

  // SPEC-067: File → Delete moves the document to the trash rather than
  // erasing it, keeping its toast and its navigation. The record is kept
  // intact (so Trash can restore it) but drops out of the library list.
  it("confirming moves the doc to trash, navigates to / and toasts (AC-33, AC-19, AC-20)", async () => {
    const { wrapper, router, store } = await mountAt("/doc/open-doc");

    await wrapper.get('[data-testid="delete-document"]').trigger("click");
    await wrapper.get('[data-testid="delete-confirm"]').trigger("click");
    await untilPath(router, "/");

    // Kept intact but out of the library and now in the trash.
    expect(store.getDocument("open-doc")).toBeDefined();
    expect(store.listDocuments().find((e) => e.id === "open-doc")).toBeUndefined();
    expect(store.listTrashedDocuments().map((e) => e.id)).toContain("open-doc");
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

describe("DocumentView save as template (SPEC-046 AC-5)", () => {
  it("prompts for a name, prefilled with the document's title, and stores a user template", async () => {
    const { wrapper } = await mountAt("/doc/save-template-doc");

    expect(wrapper.find('[data-testid="rename-dialog"]').exists()).toBe(false);
    await wrapper.get('[data-testid="file-menu-save-as-template"]').trigger("click");

    const input = wrapper.get('[data-testid="rename-input"]');
    expect((input.element as HTMLInputElement).value).toBe("Untitled document");
    await input.setValue("My Invoice Template");
    await wrapper.get('[data-testid="rename-save"]').trigger("click");

    const { listUserTemplates } = await import("../lib/templateStore");
    expect(listUserTemplates().some((t) => t.title === "My Invoice Template")).toBe(true);
    expect(wrapper.get('[data-testid="toast"]').text()).toBe('Saved template "My Invoice Template"');
    expect(wrapper.find('[data-testid="rename-dialog"]').exists()).toBe(false);
  });

  it("cancelling saves nothing", async () => {
    const { wrapper } = await mountAt("/doc/save-template-doc-2");
    const { listUserTemplates } = await import("../lib/templateStore");
    const before = listUserTemplates().length;

    await wrapper.get('[data-testid="file-menu-save-as-template"]').trigger("click");
    await wrapper.get('[data-testid="rename-cancel"]').trigger("click");

    expect(wrapper.find('[data-testid="rename-dialog"]').exists()).toBe(false);
    expect(listUserTemplates().length).toBe(before);
  });
});

describe("DocumentView insert merge field (SPEC-046 AC-12, AC-13, AC-14, AC-15)", () => {
  it("opens a picker listing names already used by a chip, inserting one at the caret", async () => {
    const { wrapper, store } = await mountAt("/doc/insert-merge-doc");
    const model = store.getOrCreateDocument("insert-merge-doc");
    model.sections = [
      { blocks: [{ type: "paragraph", runs: [{ text: "Dear ", formatting: {} }, { text: "1", formatting: {}, field: { kind: "merge", name: "existing_field" } }] }] },
    ];
    await wrapper.vm.$nextTick();

    expect(wrapper.find('[data-testid="merge-field-picker-dialog"]').exists()).toBe(false);
    await wrapper.get('[data-testid="insert-menu-merge-field"]').trigger("click");

    const options = wrapper.findAll('[data-testid="merge-field-picker-option"]').map((el) => el.text());
    expect(options).toContain("existing_field");

    await wrapper.get('[data-testid="merge-field-picker-custom-input"]').setValue("client_name");
    await wrapper.get('[data-testid="merge-field-picker-custom-insert"]').trigger("click");

    expect(wrapper.find('[data-testid="merge-field-picker-dialog"]').exists()).toBe(false);
    const runs = (store.getDocument("insert-merge-doc")!.sections[0].blocks[0] as { runs: { field?: unknown }[] }).runs;
    expect(runs.some((r) => r.field && (r.field as { name?: string }).name === "client_name")).toBe(true);
  });

  it("rejects a malformed field name with an inline error", async () => {
    const { wrapper } = await mountAt("/doc/insert-merge-invalid-doc");
    await wrapper.get('[data-testid="insert-menu-merge-field"]').trigger("click");

    await wrapper.get('[data-testid="merge-field-picker-custom-input"]').setValue("not a valid name");
    expect(wrapper.get('[data-testid="merge-field-picker-error"]').text().length).toBeGreaterThan(0);
    expect(wrapper.get('[data-testid="merge-field-picker-custom-insert"]').attributes("disabled")).toBeDefined();
  });

  it("cancelling inserts nothing", async () => {
    const { wrapper, store } = await mountAt("/doc/insert-merge-cancel-doc");
    await wrapper.get('[data-testid="insert-menu-merge-field"]').trigger("click");
    await wrapper.get('[data-testid="merge-field-picker-cancel"]').trigger("click");

    expect(wrapper.find('[data-testid="merge-field-picker-dialog"]').exists()).toBe(false);
    const runs = (store.getDocument("insert-merge-cancel-doc")!.sections[0].blocks[0] as { runs: { field?: unknown }[] }).runs;
    expect(runs.every((r) => !r.field)).toBe(true);
  });
});

describe("DocumentView convert literal {{...}} text to merge fields (SPEC-046 AC-20)", () => {
  async function mountWithBody(id: string, text: string) {
    const store = await import("../lib/documentStore");
    const model = store.getOrCreateDocument(id);
    model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text, formatting: {} }] }] }];
    const { wrapper, router } = await mountAt(`/doc/${id}`);
    return { wrapper, router, store };
  }

  it("is disabled with no candidates, enabled once a well-formed {{...}} literal exists", async () => {
    const { wrapper: noneWrapper } = await mountWithBody("convert-none", "Dear client,");
    expect(noneWrapper.get('[data-testid="tools-menu-convert-merge-fields"]').attributes("disabled")).toBeDefined();

    const { wrapper } = await mountWithBody("convert-some", "Dear {{client_name}},");
    expect(wrapper.get('[data-testid="tools-menu-convert-merge-fields"]').attributes("disabled")).toBeUndefined();
  });

  it("converts every literal match in one step and toasts a count", async () => {
    const { wrapper, store } = await mountWithBody("convert-doc", "Dear {{client_name}}, re {{invoice_no}}.");

    await wrapper.get('[data-testid="tools-menu-convert-merge-fields"]').trigger("click");

    expect(wrapper.get('[data-testid="toast"]').text()).toBe("Converted 2 merge fields");
    const model = store.getDocument("convert-doc");
    const run = (model!.sections[0].blocks[0] as Paragraph).runs[1];
    expect(run.field).toEqual({ kind: "merge", name: "client_name" });
  });

  it("is undoable in one step", async () => {
    const { wrapper, store } = await mountWithBody("convert-undo-doc", "Dear {{client_name}},");

    await wrapper.get('[data-testid="tools-menu-convert-merge-fields"]').trigger("click");
    expect((store.getDocument("convert-undo-doc")!.sections[0].blocks[0] as Paragraph).runs[1].field).toBeDefined();

    await wrapper.get('[data-testid="edit-menu-undo"]').trigger("click");
    const undone = (store.getDocument("convert-undo-doc")!.sections[0].blocks[0] as Paragraph).runs;
    expect(undone.every((r) => !r.field)).toBe(true);
    expect(undone.map((r) => r.text).join("")).toBe("Dear {{client_name}},");
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
