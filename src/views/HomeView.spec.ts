import { describe, it, expect, onTestFinished } from "vitest";
import { mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import HomeView from "./HomeView.vue";
import { deleteDocument, getDocument, getPresentation, getSpreadsheet, listDocuments } from "../lib/documentStore";
import { DEFAULT_TITLE } from "../lib/documentModel";
import { DEFAULT_PRESENTATION_TITLE } from "../lib/presentation/model";
import { DEFAULT_SPREADSHEET_TITLE } from "../lib/spreadsheet/model";
import { routes } from "../router";

function mountHome() {
  const router: Router = createRouter({ history: createMemoryHistory(), routes });
  return { wrapper: mount(HomeView, { global: { plugins: [router] } }), router };
}

describe("HomeView", () => {
  it("shows no document title in the app bar — the library is not a document", () => {
    const { wrapper } = mountHome();
    expect(wrapper.find('[data-testid="document-title"]').exists()).toBe(false);
  });

  // SPEC-067 AC-51..53: the empty state is a first-run screen describing what
  // the product does, the three ways to start, and a note that documents are
  // stored in this browser.
  it("shows a first-run empty state describing the product and storage", () => {
    const { wrapper } = mountHome();
    const canvas = wrapper.get('[data-testid="canvas"]').text();
    expect(wrapper.find('[data-testid="empty-state"]').exists()).toBe(true);
    expect(canvas).toContain("docu-edit");
    expect(canvas).toContain("stored in this browser");
  });

  it("shows the three ways to start a document (AC-52)", () => {
    const { wrapper } = mountHome();
    const newButton = wrapper.get('[data-testid="new-document"]');
    const templateButton = wrapper.get('[data-testid="new-from-template"]');
    const openButton = wrapper.get('[data-testid="open-document"]');
    expect(newButton.text()).toBe("New document");
    expect(templateButton.isVisible()).toBe(true);
    expect(newButton.isVisible()).toBe(true);
    expect(openButton.isVisible()).toBe(true);
  });

  // SPEC-121 AC-1: "New spreadsheet" sits right after "New presentation" on
  // the first-run empty state. Like the 'New presentation' test below, it
  // needs a genuinely empty library, so it runs before any doc-creating test
  // and permanently deletes what it created when it finishes — the library is
  // genuinely empty again for that test.
  it("offers 'New spreadsheet' on the genuinely empty first-run state, storing a blank workbook (SPEC-121 AC-1)", async () => {
    onTestFinished(() => {
      for (const entry of listDocuments()) {
        deleteDocument(entry.id);
      }
    });
    const { wrapper, router } = mountHome();
    expect(wrapper.find('[data-testid="empty-state"]').exists()).toBe(true);
    const emptyState = wrapper.get('[data-testid="empty-state"]');
    const actions = emptyState.findAll("button").map((button) => button.attributes("data-testid"));
    expect(actions.indexOf("new-spreadsheet")).toBe(actions.indexOf("new-presentation") + 1);
    const button = emptyState.get('[data-testid="new-spreadsheet"]');
    expect(button.text()).toBe("New spreadsheet");
    await button.trigger("click");
    await router.isReady();

    expect(router.currentRoute.value.path).toMatch(/^\/sheets\/.+/);
    const id = router.currentRoute.value.params.id as string;
    const model = getSpreadsheet(id);
    expect(model?.title).toBe(DEFAULT_SPREADSHEET_TITLE);
    expect(model?.format).toBe("xlsx");
    expect(model?.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1"]);
  });

  // SPEC-080 AC-1: "New presentation" sits beside "New document" on the
  // first-run empty state, not only once the library already holds a document.
  // This runs before any doc-creating test, so the library is genuinely empty
  // and the empty state (not the library controls) is what renders.
  it("offers 'New presentation' on the genuinely empty first-run state (AC-1)", async () => {
    const { wrapper, router } = mountHome();
    expect(wrapper.find('[data-testid="empty-state"]').exists()).toBe(true);
    const emptyState = wrapper.get('[data-testid="empty-state"]');
    const button = emptyState.get('[data-testid="new-presentation"]');
    expect(button.text()).toBe("New presentation");
    await button.trigger("click");
    await router.isReady();
    expect(router.currentRoute.value.path).toMatch(/^\/slides\/.+/);
  });

  it("creates a document and navigates to /doc/:id on click (AC-1)", async () => {
    const { wrapper, router } = mountHome();
    await wrapper.get('[data-testid="new-document"]').trigger("click");
    await router.isReady();

    expect(router.currentRoute.value.path).toMatch(/^\/doc\/.+/);
    const id = router.currentRoute.value.params.id as string;
    expect(getDocument(id)?.title).toBe(DEFAULT_TITLE);
  });

  it("stores a blank deck and navigates to /slides/:id on 'New presentation' (SPEC-080 AC-1)", async () => {
    const { wrapper, router } = mountHome();
    const button = wrapper.get('[data-testid="new-presentation"]');
    expect(button.text()).toBe("New presentation");
    await button.trigger("click");
    await router.isReady();

    expect(router.currentRoute.value.path).toMatch(/^\/slides\/.+/);
    const id = router.currentRoute.value.params.id as string;
    const model = getPresentation(id);
    expect(model?.title).toBe(DEFAULT_PRESENTATION_TITLE);
    expect(model?.format).toBe("pptx");
    expect(model?.slides).toHaveLength(1);
  });

  // SPEC-121 AC-1/AC-2: the library header's "New spreadsheet" (the library
  // holds documents by now, so the header renders rather than the empty
  // state); the stored workbook is badged XLSX — no new library format.
  it("stores a blank workbook and navigates to /sheets/:id on the library header's 'New spreadsheet' (SPEC-121 AC-1)", async () => {
    const { wrapper, router } = mountHome();
    expect(wrapper.find('[data-testid="empty-state"]').exists()).toBe(false);
    const button = wrapper.get('[data-testid="library"]').get('[data-testid="new-spreadsheet"]');
    expect(button.text()).toBe("New spreadsheet");
    await button.trigger("click");
    await router.isReady();

    expect(router.currentRoute.value.path).toMatch(/^\/sheets\/.+/);
    const id = router.currentRoute.value.params.id as string;
    const model = getSpreadsheet(id);
    expect(model?.title).toBe(DEFAULT_SPREADSHEET_TITLE);
    expect(model?.format).toBe("xlsx");
    expect(model?.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1"]);

    const { wrapper: library } = mountHome();
    const row = library.get(`[data-testid="library-row"][data-doc-id="${id}"]`);
    expect(row.get('[data-testid="row-title"]').text()).toBe(DEFAULT_SPREADSHEET_TITLE);
    expect(row.get('[data-testid="row-format"]').text()).toBe("XLSX");
  });

  it("'New from template…' opens the template gallery, offering 'Blank document' alongside templates (SPEC-046 AC-1)", async () => {
    const { wrapper } = mountHome();
    expect(wrapper.find('[data-testid="template-gallery-dialog"]').exists()).toBe(false);
    await wrapper.get('[data-testid="new-from-template"]').trigger("click");

    expect(wrapper.find('[data-testid="template-gallery-dialog"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="template-card-blank"]').text()).toContain("Blank document");
    expect(wrapper.findAll('[data-testid="template-card"]').length).toBeGreaterThanOrEqual(6);
  });

  it("choosing 'Blank document' from the gallery creates a document and navigates to /doc/:id (AC-1)", async () => {
    const { wrapper, router } = mountHome();
    await wrapper.get('[data-testid="new-from-template"]').trigger("click");
    await wrapper.get('[data-testid="template-card-blank"]').trigger("click");
    await router.isReady();

    expect(router.currentRoute.value.path).toMatch(/^\/doc\/.+/);
    const id = router.currentRoute.value.params.id as string;
    expect(getDocument(id)?.title).toBe(DEFAULT_TITLE);
  });
});
