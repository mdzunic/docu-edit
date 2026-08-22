import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import HomeView from "./HomeView.vue";
import { getDocument } from "../lib/documentStore";
import { DEFAULT_TITLE } from "../lib/documentModel";
import { routes } from "../router";

function mountHome() {
  const router: Router = createRouter({ history: createMemoryHistory(), routes });
  return { wrapper: mount(HomeView, { global: { plugins: [router] } }), router };
}

describe("HomeView", () => {
  it("titles the app bar 'No document'", () => {
    const { wrapper } = mountHome();
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      "No document",
    );
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

  it("creates a document and navigates to /doc/:id on click (AC-1)", async () => {
    const { wrapper, router } = mountHome();
    await wrapper.get('[data-testid="new-document"]').trigger("click");
    await router.isReady();

    expect(router.currentRoute.value.path).toMatch(/^\/doc\/.+/);
    const id = router.currentRoute.value.params.id as string;
    expect(getDocument(id)?.title).toBe(DEFAULT_TITLE);
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
