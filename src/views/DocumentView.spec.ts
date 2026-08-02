import { describe, it, expect, vi, beforeEach } from "vitest";
import { mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import DocumentView from "./DocumentView.vue";
import { documentExists } from "../lib/documents";
import { routes } from "../router";

// documentExists is a stub that accepts every routable id (SPEC-001 decision),
// so the not-found branch is unreachable through a real URL. Mock it to cover
// both branches; the stub itself is tested in ../lib/documents.spec.ts.
vi.mock("../lib/documents");

async function mountAt(path: string) {
  const router: Router = createRouter({
    history: createMemoryHistory(),
    routes,
  });
  await router.push(path);
  return mount(DocumentView, { global: { plugins: [router] } });
}

describe("DocumentView", () => {
  beforeEach(() => {
    vi.mocked(documentExists).mockReset();
  });

  it("titles the app bar with the client-side timestamp for a valid id", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    const wrapper = await mountAt("/doc/some-id");
    expect(wrapper.get('[data-testid="document-title"]').text()).toMatch(
      /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/,
    );
    expect(wrapper.find('[data-testid="document-page"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="not-found"]').exists()).toBe(false);
  });

  it("passes the route id to the document lookup", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    await mountAt("/doc/some-id");
    expect(documentExists).toHaveBeenCalledWith("some-id");
  });

  it("shows 'Document not found' with a link back to / for a missing id", async () => {
    vi.mocked(documentExists).mockReturnValue(false);
    const wrapper = await mountAt("/doc/missing");
    const notFound = wrapper.get('[data-testid="not-found"]');
    expect(notFound.text()).toContain("Document not found");
    expect(notFound.get("a").attributes("href")).toBe("/");
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      "No document",
    );
  });
});
