import { describe, it, expect, vi, beforeEach } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import OpenFromUrlView from "./OpenFromUrlView.vue";
import { openDocumentFromUrl, OpenFromUrlError } from "../lib/openDocumentFromUrl";
import { getDocument } from "../lib/documentStore";
import { createDocument } from "../lib/documentModel";
import { routes } from "../router";

// Only the fetch itself is mocked — OpenFromUrlError stays the real class,
// so thrown instances keep a real .message the view can render.
vi.mock("../lib/openDocumentFromUrl", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../lib/openDocumentFromUrl")>();
  return { ...actual, openDocumentFromUrl: vi.fn() };
});

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "lib", "docx", "fixtures");

async function mountAt(path: string) {
  const router: Router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  const wrapper = mount(OpenFromUrlView, { global: { plugins: [router] } });
  return { wrapper, router };
}

/** Mounts and clicks through the confirmation panel — the fetch waits for a
 * gesture, so every test past that gate goes through here. */
async function mountAndConfirm(path: string) {
  const mounted = await mountAt(path);
  await mounted.wrapper.get('[data-testid="open-from-url-confirm-button"]').trigger("click");
  return mounted;
}

async function untilPath(router: Router, path: string) {
  for (let i = 0; i < 200; i += 1) {
    if (router.currentRoute.value.path === path) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

async function untilNavigatedTo(router: Router, matcher: RegExp) {
  for (let i = 0; i < 200; i += 1) {
    if (matcher.test(router.currentRoute.value.path)) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

describe("OpenFromUrlView", () => {
  beforeEach(() => {
    vi.mocked(openDocumentFromUrl).mockReset();
  });

  // The confirmation gate (AC-9). `url` is whatever a link put in the query
  // string, so an auto-fetch on mount made any link a way to send a request
  // from the reader's browser and file the response in their library, with
  // nothing shown and nothing to decline.
  it("shows the address and fetches nothing until the reader confirms (AC-9)", async () => {
    const { wrapper } = await mountAt("/open?url=https://example.com/remote/report.docx");
    await flushPromises();

    expect(wrapper.get('[data-testid="open-from-url-host"]').text()).toBe("example.com");
    expect(wrapper.get('[data-testid="open-from-url-target"]').text()).toBe(
      "https://example.com/remote/report.docx",
    );
    expect(wrapper.find('[data-testid="open-from-url-loading"]').exists()).toBe(false);
    expect(openDocumentFromUrl).not.toHaveBeenCalled();
  });

  it("shows the full URL, not only the host, so a padded path cannot hide the target (AC-9)", async () => {
    const url = "https://trusted-looking.example.com/a/very/long/path/that/keeps/going/report.docx";
    const { wrapper } = await mountAt("/open?url=" + url);
    await flushPromises();

    expect(wrapper.get('[data-testid="open-from-url-target"]').text()).toBe(url);
  });

  it("offers a local file as an alternative to confirming (AC-9)", async () => {
    const { wrapper } = await mountAt("/open?url=https://example.com/report.docx");
    await flushPromises();

    expect(wrapper.get('[data-testid="open-from-url-pick-file-instead"]').text()).toBe(
      "Choose a file instead",
    );
  });

  it("shows a loading state while the fetch is in flight (AC-10)", async () => {
    let resolveFetch: (opened: { kind: "text"; model: ReturnType<typeof createDocument> }) => void = () => {};
    vi.mocked(openDocumentFromUrl).mockReturnValue(
      new Promise((resolve) => {
        resolveFetch = resolve;
      }),
    );
    const { wrapper } = await mountAndConfirm("/open?url=https://example.com/report.docx");

    expect(wrapper.find('[data-testid="open-from-url-loading"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="open-from-url-error"]').exists()).toBe(false);

    resolveFetch({ kind: "text", model: createDocument("resolved-id") });
    await flushPromises();
  });

  it("fetches the URL and opens the document, storing it and navigating to /doc/:id (AC-9, AC-15)", async () => {
    const doc = { ...createDocument("from-url-id"), format: "docx" as const };
    vi.mocked(openDocumentFromUrl).mockResolvedValue({ kind: "text", model: doc });
    const { wrapper, router } = await mountAndConfirm("/open?url=https://example.com/report.docx");
    await flushPromises();
    await untilPath(router, "/doc/from-url-id");

    expect(openDocumentFromUrl).toHaveBeenCalledWith("https://example.com/report.docx");
    expect(router.currentRoute.value.path).toBe("/doc/from-url-id");
    expect(getDocument("from-url-id")).toEqual(doc);
    expect(wrapper).toBeTruthy();
  });

  it("shows the failure message with the underlying reason and a fallback button (AC-11, AC-12, AC-13)", async () => {
    vi.mocked(openDocumentFromUrl).mockRejectedValue(new OpenFromUrlError("The link returned text/html, not a .docx or .odt document"));
    const { wrapper } = await mountAndConfirm("/open?url=https://example.com/page");
    await flushPromises();

    expect(wrapper.find('[data-testid="open-from-url-loading"]').exists()).toBe(false);
    const errorPanel = wrapper.get('[data-testid="open-from-url-error"]');
    expect(errorPanel.text()).toContain("Couldn't open that link");
    expect(wrapper.get('[data-testid="open-from-url-reason"]').text()).toBe(
      "The link returned text/html, not a .docx or .odt document",
    );
    expect(wrapper.get('[data-testid="open-from-url-pick-file"]').text()).toBe(
      "Choose a file instead",
    );
  });

  it("shows an error when no url query parameter is given", async () => {
    const { wrapper } = await mountAt("/open");
    await flushPromises();

    expect(wrapper.find('[data-testid="open-from-url-error"]').exists()).toBe(true);
    expect(openDocumentFromUrl).not.toHaveBeenCalled();
  });

  it("the fallback button opens a local file exactly like the picker, once chosen (AC-16 fallback path)", async () => {
    vi.mocked(openDocumentFromUrl).mockRejectedValue(new OpenFromUrlError("Failed to fetch"));
    const { wrapper, router } = await mountAndConfirm("/open?url=https://blocked.example.com/report.docx");
    await flushPromises();

    const bytes = readFileSync(join(fixturesDir, "word-authored.docx"));
    const file = new File([bytes], "word-authored.docx");

    const input = wrapper.get<HTMLInputElement>('[data-testid="open-from-url-input"]');
    Object.defineProperty(input.element, "files", { value: [file] });
    await input.trigger("change");
    await flushPromises();
    await untilNavigatedTo(router, /^\/doc\/.+/);

    expect(router.currentRoute.value.path).toMatch(/^\/doc\/.+/);
  });
});
