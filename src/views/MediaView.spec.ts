// The media viewer (SPEC-107 AC-11..13): image documents open in the
// fit/zoom viewer with a details panel; PDFs render through the browser's
// own embed with the parse-derived details beside it; a wrong-kind id hands
// off to its own view; and none of the editor's surfaces (autosave, version
// history, comments, File menu) exist here.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";

// The checked-in spec107-pages.pdf (the app's own buildPdf output).
const PAGES_PDF = new Uint8Array(readFileSync(join(__dirname, "../../e2e/fixtures/spec107-pages.pdf")));

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  // jsdom has no URL.createObjectURL.
  Object.assign(URL, {
    createObjectURL: vi.fn(() => "blob:test"),
    revokeObjectURL: vi.fn(),
  });
});

afterEach(() => {
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
});

async function loadModules() {
  const store = await import("../lib/documentStore");
  const { assembleMediaModel } = await import("../lib/media/open");
  const fixtures = await import("../lib/media/fixtures/build");
  const { routes } = await import("../router");
  const { default: MediaView } = await import("./MediaView.vue");
  async function mountMedia(id: string) {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push(`/media/${id}`);
    const wrapper = mount(MediaView, { global: { plugins: [router] } });
    await flushPromises();
    return { wrapper, router };
  }
  return { store, assembleMediaModel, fixtures, mountMedia };
}

describe("MediaView", () => {
  it("renders an image in the zoom viewer with its details (AC-11)", async () => {
    const { store, assembleMediaModel, fixtures, mountMedia } = await loadModules();
    const model = await assembleMediaModel("shot", "png", await fixtures.buildPngFixtureBytes());
    store.storeMedia(model);
    const { wrapper } = await mountMedia(model.id);

    const image = wrapper.get('[data-testid="media-image"]');
    expect(image.attributes("alt")).toBe("shot.png");
    expect(image.attributes("src")).toBe("blob:test");
    // Zoom vocabulary present: fit, 100%, the control and the level.
    expect(wrapper.find('[data-testid="media-zoom-fit"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="media-zoom-actual"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="media-zoom-level"]').exists()).toBe(true);
    // Details: dimensions, format, byte size.
    expect(wrapper.get('[data-testid="detail-dimensions"]').text()).toBe("64 × 40 px");
    expect(wrapper.get('[data-testid="detail-format"]').text()).toContain("PNG");
    expect(wrapper.get('[data-testid="detail-size"]').text()).toMatch(/B$/);
    expect(wrapper.find('[data-testid="detail-stored"]').exists()).toBe(true);
    // No editor surfaces (AC-35): no save state, no menu bar and so no
    // editor-only File menu entries, no version history, no comments.
    for (const testId of [
      "save-state",
      "menu-bar",
      "file-menu-export",
      "file-menu-save-as-template",
      "version-history-pane",
      "comments-pane",
    ]) {
      expect(wrapper.find(`[data-testid="${testId}"]`).exists()).toBe(false);
    }
    expect(wrapper.find('[data-testid="close-document"]').exists()).toBe(true);
  });

  it("creates one blob: URL per open and revokes it on unmount (AC-33)", async () => {
    const { store, assembleMediaModel, mountMedia } = await loadModules();
    const model = await assembleMediaModel("pages", "pdf", PAGES_PDF);
    store.storeMedia(model);
    const { wrapper } = await mountMedia(model.id);
    // The mount loads once and re-syncs after initLibrary resolves; the
    // re-sync of the same document must not mint (and revoke) a second URL.
    expect(URL.createObjectURL).toHaveBeenCalledTimes(1);
    expect(URL.revokeObjectURL).not.toHaveBeenCalled();
    expect(wrapper.get("object").attributes("data")).toBe("blob:test");
    wrapper.unmount();
    expect(URL.revokeObjectURL).toHaveBeenCalledWith("blob:test");
  });

  it("renders a PDF through the browser embed with parse-derived details (AC-12)", async () => {
    const { store, assembleMediaModel, mountMedia } = await loadModules();
    const model = await assembleMediaModel("pages", "pdf", PAGES_PDF);
    store.storeMedia(model);
    const { wrapper } = await mountMedia(model.id);

    expect(wrapper.find('[data-testid="media-pdf"]').exists()).toBe(true);
    expect(wrapper.find("object").attributes("type")).toBe("application/pdf");
    expect(wrapper.find("embed").exists()).toBe(true);
    expect(wrapper.find('[data-testid="media-pdf-fallback"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="detail-pages"]').text()).toBe("3 pages");
    expect(wrapper.get('[data-testid="detail-page-size"]').text()).toContain("mixed");
    expect(wrapper.get('[data-testid="detail-version"]').text()).toBe("1.7");
    expect(wrapper.get('[data-testid="detail-producer"]').text()).toBe("docu-edit");
    expect(wrapper.get('[data-testid="detail-info-title"]').text()).toBe("SPEC-107 pages");
    expect(wrapper.get('[data-testid="detail-info-author"]').text()).toBe("docu-edit factory");
    expect(wrapper.get('[data-testid="detail-created"]').text()).toBe("2026-08-30 12:00");
    expect(wrapper.get('[data-testid="detail-linearised"]').text()).toBe("No");
    expect(wrapper.get('[data-testid="detail-reconstructed"]').text()).toBe("No");
    // The fallback's actions are there for a browser with no PDF viewer.
    expect(wrapper.get('[data-testid="media-pdf-fallback"]').text()).toContain("Your browser cannot display PDFs here");
    expect(wrapper.find('[data-testid="media-pdf-download"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="media-pdf-open-tab"]').attributes("target")).toBe("_blank");
  });

  it("reports a reconstructed PDF in the details panel", async () => {
    const { store, assembleMediaModel, mountMedia } = await loadModules();
    // A nonsense startxref: only the reconstruction path can read the file.
    const text = String.fromCharCode(...PAGES_PDF);
    const broken = Uint8Array.from(text.replace(/startxref\s+\d+/, "startxref\n2"), (char) => char.charCodeAt(0));
    const model = await assembleMediaModel("broken", "pdf", broken);
    store.storeMedia(model);
    const { wrapper } = await mountMedia(model.id);
    expect(wrapper.get('[data-testid="detail-reconstructed"]').text()).toMatch(/^Yes/);
    expect(wrapper.get('[data-testid="detail-pages"]').text()).toBe("3 pages");
  });

  it("hands a presentation id off to /slides/:id", async () => {
    const { store, mountMedia } = await loadModules();
    const { createBlankPresentation } = await import("../lib/presentation/slides");
    const deck = createBlankPresentation();
    store.storePresentation(deck);
    const { router } = await mountMedia(deck.id);
    await flushPromises();
    expect(router.currentRoute.value.path).toBe(`/slides/${deck.id}`);
  });

  it("hands a text id off to /doc/:id (AC-11 kind guard)", async () => {
    const { store, mountMedia } = await loadModules();
    store.getOrCreateDocument("text-1");
    const { router } = await mountMedia("text-1");
    await flushPromises();
    expect(router.currentRoute.value.path).toBe("/doc/text-1");
  });

  it("shows a not-found state for an unknown id", async () => {
    const { mountMedia } = await loadModules();
    const { wrapper } = await mountMedia("nope");
    expect(wrapper.find('[data-testid="media-not-found"]').exists()).toBe(true);
  });
});
