import { describe, it, expect, vi, beforeEach } from "vitest";
import { mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import DocumentView from "./DocumentView.vue";
import { documentExists } from "../lib/documents";
import { getDocument } from "../lib/documentStore";
import { DEFAULT_TITLE } from "../lib/documentModel";
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

  it("titles the app bar 'Untitled document' for a fresh id (AC-4, AC-5)", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    const wrapper = await mountAt("/doc/some-id");
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      DEFAULT_TITLE,
    );
    expect(wrapper.find('[data-testid="document-page"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="not-found"]').exists()).toBe(false);
  });

  it("passes the route id to the document lookup", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    await mountAt("/doc/some-id");
    expect(documentExists).toHaveBeenCalledWith("some-id");
  });

  it("renders a page sized and margined from defaultStyles (AC-3)", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    const wrapper = await mountAt("/doc/page-shape-id");
    const doc = getDocument("page-shape-id");
    const page = wrapper.get('[data-testid="document-page"]');
    expect(page.attributes("style")).toContain(
      `width: ${doc!.defaultStyles.pageSize.widthCm}cm`,
    );
    expect(page.attributes("style")).toContain(
      `height: ${doc!.defaultStyles.pageSize.heightCm}cm`,
    );
    const margin = wrapper.get('[data-testid="page-margin"]');
    expect(margin.attributes("style")).toContain(
      `top: ${doc!.defaultStyles.margins.topCm}cm`,
    );
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

  it("editing the title in the app bar updates the document model (AC-6, AC-7)", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    const wrapper = await mountAt("/doc/edit-title-id");
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Project Plan");
    await input.trigger("keydown.enter");

    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      "Project Plan",
    );
    expect(getDocument("edit-title-id")?.title).toBe("Project Plan");
  });

  it("offers a Download action in the app bar for an open document (SPEC-006 AC-1)", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    const wrapper = await mountAt("/doc/download-id");
    const appBar = wrapper.get('[data-testid="app-bar"]');
    const download = appBar.get('[data-testid="download-menu"]');
    expect(download.text()).toBe("Download");
    await download.trigger("click");
    expect(wrapper.get('[data-testid="download-docx"]').text()).toBe("Word (.docx)");
  });

  it("a whitespace-only title reverts to the default (AC-9)", async () => {
    vi.mocked(documentExists).mockReturnValue(true);
    const wrapper = await mountAt("/doc/whitespace-title-id");
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("   ");
    await input.trigger("keydown.enter");

    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      DEFAULT_TITLE,
    );
    expect(getDocument("whitespace-title-id")?.title).toBe(DEFAULT_TITLE);
  });
});
