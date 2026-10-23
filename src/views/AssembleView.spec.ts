// SPEC-109: the /assemble workspace, mounted. Its four regions (tray,
// preview with the ARIA list of placed rectangles, inspector over SPEC-108's
// SheetSetupPanel, output bar); tray reordering by keyboard — Alt+↑/Alt+↓ on
// a focused row, focus following the row, the move spoken in the role=status
// live region — and by drag; the refusal paths (the canvas ceiling with its
// one-click dpi fix, JPEG and transparency, more than MAX_ASSEMBLY_ITEMS);
// the deleted-item notice; Start over; and no source mutation — the sources
// are deep-frozen and a save and a download leave them exactly as they were.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import type { DecodedImage } from "../lib/media/renderEdits";

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  // jsdom has no canvas: the preview draws nothing (composeToCanvas no-ops).
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});

afterEach(() => {
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const inner of Object.values(value)) {
      deepFreeze(inner);
    }
  }
  return value;
}

function fakeSteps(overrides: Record<string, unknown> = {}) {
  const ctx = { save() {}, restore() {}, setTransform() {}, clearRect() {}, fillRect() {}, drawImage: vi.fn() };
  return {
    createCanvas: vi.fn((width: number, height: number) => ({ width, height, getContext: () => ctx }) as unknown as HTMLCanvasElement),
    decode: vi.fn(async (): Promise<DecodedImage> => ({ image: {} as CanvasImageSource, widthPx: 64, heightPx: 40, release: () => {} })),
    encode: vi.fn(async () => new Uint8Array()),
    canvasCeiling: vi.fn(() => 268_435_456),
    between: vi.fn(async () => {}),
    download: vi.fn(),
    release: vi.fn(),
    ...overrides,
  };
}

async function setup(stepOverrides: Record<string, unknown> = {}) {
  const store = await import("../lib/documentStore");
  const { assembleMediaModel } = await import("../lib/media/open");
  const fixtures = await import("../lib/media/fixtures/build");
  const { routes } = await import("../router");
  const { startAssembly } = await import("./assembly/useAssemblyPersistence");
  const { loadPreferences } = await import("../lib/preferences");
  const { default: AssembleView } = await import("./AssembleView.vue");
  const png = await fixtures.buildPngFixtureBytes();
  const steps = fakeSteps({ encode: vi.fn(async () => png), ...stepOverrides });

  async function seed(title: string) {
    const model = await assembleMediaModel(title, "png", png);
    store.storeMedia(model);
    return model;
  }

  async function mountView(): Promise<VueWrapper> {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push("/assemble");
    const wrapper = mount(AssembleView, { props: { steps }, global: { plugins: [router] }, attachTo: document.body });
    await flushPromises();
    return wrapper;
  }

  return { store, seed, startAssembly, loadPreferences, mountView, steps, png };
}

const trayTitles = (wrapper: VueWrapper) =>
  wrapper.findAll('[data-testid="assembly-item"] [data-testid="item-title"]').map((title) => title.text());

describe("the workspace", () => {
  it("is a tray, a preview with its placed rectangles, an inspector over the sheet setup, and an output bar", async () => {
    const { seed, startAssembly, mountView } = await setup();
    const a = await seed("shot");
    const b = await seed("logo");
    startAssembly([a.id, b.id]);
    const wrapper = await mountView();

    expect(trayTitles(wrapper)).toEqual(["shot", "logo"]);
    expect(wrapper.findAll('[data-testid="item-size"]').map((size) => size.text())).toEqual(["64 × 40 px", "64 × 40 px"]);
    expect(wrapper.findAll('[data-testid="assembly-item"] [data-testid="row-thumbnail"]')).toHaveLength(2);

    const canvas = wrapper.get('canvas[data-testid="assembly-preview"]');
    expect(canvas.attributes("aria-describedby")).toBe("assembly-placed-list");
    const list = wrapper.get("#assembly-placed-list");
    expect(list.attributes("aria-label")).toBe("Placed images");
    expect(list.findAll('[data-testid="assembly-placed-item"]')).toHaveLength(2);

    const inspector = wrapper.get('[data-testid="assembly-inspector"]');
    const layout = inspector.get('[data-testid="assembly-layout"]').element;
    const sheet = inspector.get('[data-testid="sheet-setup"]').element;
    // The layout controls come before the sheet setup panel.
    expect(layout.compareDocumentPosition(sheet) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();

    const output = wrapper.get('[data-testid="assembly-output"]');
    expect(output.get('[data-testid="assembly-format"]').element).toHaveProperty("value", "png");
    expect(output.find('[data-testid="assembly-quality"]').exists()).toBe(false);
    expect(output.get('[data-testid="assembly-estimate"]').text()).toBe("1240 × 1754 px at 150 dpi · PNG up to about 8.3 MB");
    expect(output.get('[data-testid="assembly-title"]').element).toHaveProperty("value", "Combined image");
    expect(output.find('[data-testid="assembly-save"]').exists()).toBe(true);
    expect(output.find('[data-testid="assembly-download"]').exists()).toBe(true);
  });

  it("opens an empty tray for New › Combined image", async () => {
    const { seed, startAssembly, mountView } = await setup();
    await seed("shot");
    startAssembly([]);
    const wrapper = await mountView();
    expect(wrapper.findAll('[data-testid="assembly-item"]')).toHaveLength(0);
    expect(wrapper.find('[data-testid="assembly-tray-empty"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="assembly-save"]').attributes("disabled")).toBeDefined();
  });

  it("adds images from the library picker, which lists only images", async () => {
    const { seed, store, mountView } = await setup();
    await seed("shot");
    store.getOrCreateDocument("doc-1").title = "Report";
    const wrapper = await mountView();
    await wrapper.get('[data-testid="assembly-add"]').trigger("click");
    const picker = wrapper.get('[data-testid="assembly-picker"]');
    expect(picker.findAll('[data-testid="picker-item"]').map((item) => item.text())).toEqual(["shotPNG"]);
    await picker.get('[data-testid="picker-check"]').setValue(true);
    await picker.get('[data-testid="assembly-picker-add"]').trigger("click");
    expect(wrapper.find('[data-testid="assembly-picker"]').exists()).toBe(false);
    expect(trayTitles(wrapper)).toEqual(["shot"]);
  });
});

describe("reordering the tray", () => {
  it("moves a focused row with Alt+↓ / Alt+↑, keeps focus on it and announces it", async () => {
    const { seed, startAssembly, mountView } = await setup();
    const ids = [(await seed("a")).id, (await seed("b")).id, (await seed("c")).id];
    startAssembly(ids);
    const wrapper = await mountView();
    const status = wrapper.get('[data-testid="assembly-tray-status"]');
    expect(status.attributes("role")).toBe("status");

    const first = wrapper.findAll('[data-testid="assembly-item"]')[0];
    (first.element as HTMLElement).focus();
    await first.trigger("keydown", { key: "ArrowDown", altKey: true });
    await flushPromises();
    expect(trayTitles(wrapper)).toEqual(["b", "a", "c"]);
    expect(status.text()).toBe("“a” moved to position 2 of 3.");
    const moved = wrapper.findAll('[data-testid="assembly-item"]')[1];
    expect(document.activeElement).toBe(moved.element);

    await moved.trigger("keydown", { key: "ArrowUp", altKey: true });
    await moved.trigger("keydown", { key: "ArrowUp", altKey: true });
    await flushPromises();
    expect(trayTitles(wrapper)).toEqual(["a", "b", "c"]);
    expect(status.text()).toBe("“a” is already first.");
    // Without Alt the arrows do not reorder.
    await wrapper.findAll('[data-testid="assembly-item"]')[0].trigger("keydown", { key: "ArrowDown" });
    expect(trayTitles(wrapper)).toEqual(["a", "b", "c"]);
    // The placed rectangles follow the tray.
    const labels = wrapper.findAll('[data-testid="assembly-placed-item"]').map((item) => item.text());
    expect(labels.map((label) => label.slice(0, 4))).toEqual(["1. a", "2. b", "3. c"]);
  });

  it("moves a dragged row to where it is dropped", async () => {
    const { seed, startAssembly, mountView } = await setup();
    startAssembly([(await seed("a")).id, (await seed("b")).id, (await seed("c")).id]);
    const wrapper = await mountView();
    const rows = wrapper.findAll('[data-testid="assembly-item"]');
    await rows[2].trigger("dragstart");
    await rows[0].trigger("dragover");
    await rows[0].trigger("drop");
    await rows[2].trigger("dragend");
    expect(trayTitles(wrapper)).toEqual(["c", "a", "b"]);
    expect(wrapper.get('[data-testid="assembly-tray-status"]').text()).toBe("“c” moved to position 1 of 3.");
  });

  it("removes and duplicates rows", async () => {
    const { seed, startAssembly, mountView } = await setup();
    startAssembly([(await seed("a")).id, (await seed("b")).id]);
    const wrapper = await mountView();
    await wrapper.findAll('[data-testid="item-duplicate"]')[0].trigger("click");
    expect(trayTitles(wrapper)).toEqual(["a", "a", "b"]);
    await wrapper.findAll('[data-testid="item-remove"]')[2].trigger("click");
    expect(trayTitles(wrapper)).toEqual(["a", "a"]);
  });
});

describe("the refusal paths", () => {
  it("refuses an output over the canvas ceiling before any work, with the dpi that fits one click away", async () => {
    const { seed, startAssembly, mountView, steps } = await setup({ canvasCeiling: vi.fn(() => 1_000_000) });
    startAssembly([(await seed("a")).id]);
    const wrapper = await mountView();
    const refusal = wrapper.get('[data-testid="assembly-refusal"]');
    expect(refusal.attributes("role")).toBe("alert");
    expect(refusal.text()).toBe(
      "At 150 dpi the image would be 1240 × 1754 px (2.2 megapixels), more than this browser's canvas limit of 1.0 megapixels. At 72 dpi it would fit.",
    );
    expect(wrapper.get('[data-testid="assembly-save"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="assembly-download"]').attributes("disabled")).toBeDefined();
    expect(steps.createCanvas).not.toHaveBeenCalled();
    await wrapper.get('[data-testid="assembly-use-dpi"]').trigger("click");
    expect(wrapper.find('[data-testid="assembly-refusal"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="sheet-dpi"]').element).toHaveProperty("value", "72");
    expect(wrapper.get('[data-testid="assembly-save"]').attributes("disabled")).toBeUndefined();
  });

  it("shows quality only for JPEG, and says before committing that its transparent sheet is filled", async () => {
    const { seed, startAssembly, mountView, steps } = await setup();
    startAssembly([(await seed("a")).id]);
    const wrapper = await mountView();
    expect(wrapper.find('[data-testid="assembly-background-notice"]').exists()).toBe(false);
    await wrapper.get('[data-testid="assembly-format"]').setValue("jpeg");
    expect(wrapper.get('[data-testid="assembly-quality"]').element).toHaveProperty("value", "90");
    const notice = wrapper.get('[data-testid="assembly-background-notice"]');
    expect(notice.text()).toBe("JPEG has no transparency, so the transparent background will be filled with the colour chosen here.");
    expect(wrapper.get('[data-testid="assembly-matte"]').element).toHaveProperty("value", "#ffffff");
    await wrapper.get('[data-testid="assembly-quality"]').setValue("150");
    expect(wrapper.get('[data-testid="assembly-quality"]').element).toHaveProperty("value", "90");
    expect(wrapper.get('[data-testid="assembly-error"]').text()).toBe("JPEG quality must be a whole number from 1 to 100.");
    await wrapper.get('[data-testid="assembly-download"]').trigger("click");
    await flushPromises();
    expect(steps.encode).toHaveBeenCalledWith(expect.anything(), "image/jpeg", 0.9);
    // A white sheet needs no notice.
    await wrapper.get('[data-testid="sheet-background"]').setValue("#ffffff");
    expect(wrapper.find('[data-testid="assembly-background-notice"]').exists()).toBe(false);
  });

  it("refuses more than 100 images, saying how many were left out", async () => {
    const { seed, startAssembly, mountView } = await setup();
    const a = await seed("a");
    startAssembly(Array.from({ length: 101 }, () => a.id));
    const wrapper = await mountView();
    expect(wrapper.findAll('[data-testid="assembly-item"]')).toHaveLength(100);
    expect(wrapper.get('[data-testid="assembly-count"]').text()).toBe("100 of 100");
    expect(wrapper.get('[data-testid="assembly-notice"]').text()).toBe("A tray holds at most 100 images; the first 100 of 101 were added.");
    expect(wrapper.get('[data-testid="assembly-add"]').attributes("disabled")).toBeDefined();
    expect(wrapper.findAll('[data-testid="item-duplicate"]')[0].attributes("disabled")).toBeDefined();
  });
});

describe("restoring and starting over", () => {
  it("drops a deleted item from the restored tray with a notice naming it", async () => {
    const { seed, startAssembly, mountView, store } = await setup();
    const a = await seed("kept");
    const b = await seed("deleted one");
    startAssembly([a.id, b.id]);
    store.deleteDocument(b.id);
    const wrapper = await mountView();
    expect(trayTitles(wrapper)).toEqual(["kept"]);
    const notice = wrapper.get('[data-testid="assembly-notice"]');
    expect(notice.text()).toBe("“deleted one” is no longer in the library and was removed from the tray.");
    expect(notice.attributes("role")).toBe("status");
  });

  it("clears the persisted plan on Start over", async () => {
    const { seed, startAssembly, mountView, loadPreferences } = await setup();
    startAssembly([(await seed("a")).id]);
    const wrapper = await mountView();
    expect(loadPreferences().assemblyPlan?.items).toHaveLength(1);
    await wrapper.get('[data-testid="assembly-reset"]').trigger("click");
    expect(loadPreferences().assemblyPlan).toBeNull();
    expect(wrapper.findAll('[data-testid="assembly-item"]')).toHaveLength(0);
  });
});

describe("no source mutation", () => {
  it("saves and downloads without touching a deep-frozen source", async () => {
    const { seed, startAssembly, mountView, store, steps } = await setup();
    const a = await seed("a");
    const b = await seed("b");
    startAssembly([a.id, b.id]);
    deepFreeze(a);
    deepFreeze(b);
    const before = [a.id, b.id].map((id) => JSON.stringify(store.getStoredRecord(id)));
    const wrapper = await mountView();

    await wrapper.get('[data-testid="assembly-save"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="assembly-saved"]').text()).toBe("Saved “Combined image” to the library.");
    await wrapper.get('[data-testid="assembly-download"]').trigger("click");
    await flushPromises();
    expect(steps.download).toHaveBeenCalledWith(expect.any(Uint8Array), "Combined image.png", "image/png");
    expect(wrapper.find('[data-testid="assembly-failure"]').exists()).toBe(false);

    // The sources: the same records, byte for byte, still in the library.
    expect([a.id, b.id].map((id) => JSON.stringify(store.getStoredRecord(id)))).toEqual(before);
    expect(store.getMedia(a.id)).toBe(a);
    const titles = store.listDocuments().map((entry) => entry.title);
    expect(titles.toSorted()).toEqual(["Combined image", "a", "b"]);
  });

  it("reports progress as 'Item n of N' and cancels at the next boundary, storing nothing", async () => {
    let wrapper: VueWrapper | null = null;
    const seen: string[] = [];
    const { seed, startAssembly, mountView, store } = await setup({
      between: vi.fn(async () => {
        const progress = wrapper?.find('[data-testid="assembly-progress"]');
        if (progress?.exists()) {
          seen.push(progress.text());
          expect(progress.attributes("role")).toBe("status");
          await wrapper?.get('[data-testid="assembly-cancel"]').trigger("click");
        }
      }),
    });
    startAssembly([(await seed("a")).id, (await seed("b")).id]);
    wrapper = await mountView();
    const before = store.listDocuments().length;
    await wrapper.get('[data-testid="assembly-save"]').trigger("click");
    await flushPromises();
    expect(seen[0]).toBe("Item 0 of 2");
    expect(store.listDocuments()).toHaveLength(before);
    expect(wrapper.find('[data-testid="assembly-progress"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="assembly-saved"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="assembly-failure"]').exists()).toBe(false);
  });
});
