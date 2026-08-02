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

  it("shows the 'No document open' empty state in the canvas", () => {
    const { wrapper } = mountHome();
    expect(wrapper.get('[data-testid="canvas"]').text()).toContain(
      "No document open",
    );
  });

  it("shows the 'New document' and 'Open document' buttons", () => {
    const { wrapper } = mountHome();
    const newButton = wrapper.get('[data-testid="new-document"]');
    const openButton = wrapper.get('[data-testid="open-document"]');
    expect(newButton.text()).toBe("New document");
    expect(openButton.text()).toBe("Open document");
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
});
