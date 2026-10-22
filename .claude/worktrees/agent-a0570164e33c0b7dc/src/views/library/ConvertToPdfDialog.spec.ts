// SPEC-108's Convert to PDF dialog: the BaseDialog shell holding the shared
// panel and preview, the title field, control enablement, the transparent-
// background refusals (PDF: chosen, but white — said before committing), the
// SVG note, and the save/failure paths. The canvas-bound conversion steps
// are stubbed through the dialog's `steps` seam.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import ConvertToPdfDialog from "./ConvertToPdfDialog.vue";
import { resetSheetSetupSessions } from "../../composables/useSheetSetup";
import { resetPreferences } from "../../lib/preferences";
import { assembleMediaModel } from "../../lib/media/open";
import { buildPngFixtureBytes, buildSvgFixtureBytes } from "../../lib/media/fixtures/build";
import * as store from "../../lib/documentStore";
import type { ConvertSteps } from "./useConvertToPdf";

vi.mock("../../lib/idb", () => ({
  loadAllRecords: vi.fn().mockResolvedValue([]),
  putRecord: vi.fn().mockResolvedValue(undefined),
  deleteRecord: vi.fn().mockResolvedValue(undefined),
  getRecord: vi.fn().mockResolvedValue(undefined),
}));

function steps(encodeOk = true): ConvertSteps {
  return {
    encode: vi.fn(async () =>
      encodeOk
        ? { widthPx: 1, heightPx: 1, filter: "FlateDecode" as const, colorSpace: "DeviceRGB" as const, data: new Uint8Array([0]) }
        : null,
    ),
    crop: vi.fn(async () => "data:image/png;base64,AA=="),
    rasteriseSvg: vi.fn(async (_bytes: Uint8Array, size: { widthPx: number; heightPx: number }) => ({
      data: new Uint8Array([0]),
      ...size,
    })),
  };
}

async function mountDialog(format: "png" | "svg" = "png", encodeOk = true) {
  const bytes = format === "png" ? await buildPngFixtureBytes() : buildSvgFixtureBytes();
  const source = await assembleMediaModel("Shot", format, bytes);
  store.storeMedia(source);
  const wrapper = mount(ConvertToPdfDialog, {
    props: { sourceId: source.id, steps: steps(encodeOk) },
    attachTo: document.body,
  });
  return { wrapper, source };
}

const byId = (id: string) => `[data-testid="${id}"]`;

beforeEach(() => {
  localStorage.clear();
  resetPreferences();
  resetSheetSetupSessions();
});

afterEach(() => {
  for (const entry of store.listDocuments()) {
    store.deleteDocument(entry.id);
  }
  document.body.innerHTML = "";
});

describe("ConvertToPdfDialog", () => {
  it("is a dialog holding the setup panel, the preview, a title field and three actions", async () => {
    const { wrapper } = await mountDialog();
    expect(wrapper.get(byId("convert-pdf-dialog")).find('[role="dialog"]').exists()).toBe(true);
    expect(wrapper.find(byId("sheet-setup")).exists()).toBe(true);
    expect(wrapper.find(byId("sheet-preview")).exists()).toBe(true);
    expect((wrapper.get(byId("convert-pdf-title")).element as HTMLInputElement).value).toBe("Shot");
    expect(wrapper.get(byId("convert-pdf-save")).text()).toBe("Save to library");
    expect(wrapper.find(byId("convert-pdf-download")).exists()).toBe(true);
    expect(wrapper.find(byId("convert-pdf-cancel")).exists()).toBe(true);
    wrapper.unmount();
  });

  it("enables Save and Download only while there is a title", async () => {
    const { wrapper } = await mountDialog();
    const save = wrapper.get(byId("convert-pdf-save"));
    const download = wrapper.get(byId("convert-pdf-download"));
    expect(save.attributes("disabled")).toBeUndefined();
    expect(download.attributes("disabled")).toBeUndefined();
    await wrapper.get(byId("convert-pdf-title")).setValue("   ");
    expect(save.attributes("disabled")).toBeDefined();
    expect(download.attributes("disabled")).toBeDefined();
    expect(wrapper.get(byId("convert-pdf-blocked")).text()).toBe("Give the PDF a title.");
    await wrapper.get(byId("convert-pdf-title")).setValue("Scan");
    expect(save.attributes("disabled")).toBeUndefined();
    wrapper.unmount();
  });

  it("enables the panel's controls by what applies to the chosen sheet", async () => {
    const { wrapper } = await mountDialog();
    expect(wrapper.get(byId("sheet-width")).attributes("disabled")).toBeDefined();
    await wrapper.get(byId("sheet-paper")).setValue("Custom");
    expect(wrapper.get(byId("sheet-width")).attributes("disabled")).toBeUndefined();
    await wrapper.get(byId("sheet-paper")).setValue("Fit to content");
    expect(wrapper.get(byId("sheet-landscape")).attributes("disabled")).toBeDefined();
    await wrapper.get(byId("sheet-fit")).setValue("stretch");
    expect(wrapper.get(byId("sheet-align")).attributes("disabled")).toBeDefined();
    wrapper.unmount();
  });

  it("refuses transparency for a PDF: says so before committing, and previews white", async () => {
    const { wrapper } = await mountDialog();
    expect(wrapper.find(byId("sheet-background-notice")).exists()).toBe(false);
    await wrapper.get(byId("sheet-background")).setValue("transparent");
    expect(wrapper.get(byId("sheet-background-notice")).text()).toBe(
      "PDF pages have no transparency, so the background will be white.",
    );
    expect(wrapper.get(byId("sheet-preview-paper")).attributes("fill")).toBe("#ffffff");
    // The refusal is a fallback, not a block: the PDF can still be saved.
    expect(wrapper.get(byId("convert-pdf-save")).attributes("disabled")).toBeUndefined();
    wrapper.unmount();
  });

  it("says an SVG is rasterised", async () => {
    const png = await mountDialog("png");
    expect(png.wrapper.find(byId("convert-pdf-svg-notice")).exists()).toBe(false);
    png.wrapper.unmount();
    const svg = await mountDialog("svg");
    expect(svg.wrapper.get(byId("convert-pdf-svg-notice")).text()).toContain("rasterised");
    svg.wrapper.unmount();
  });

  it("reports the preview's placed size, and saves a new PDF to the library", async () => {
    const { wrapper, source } = await mountDialog();
    await wrapper.get(byId("sheet-margin-preset")).setValue("Narrow");
    expect(wrapper.get(byId("sheet-preview-placed-size")).text()).toBe("Placed 18.46 × 11.54 cm");
    await wrapper.get(byId("convert-pdf-save")).trigger("click");
    await vi.waitFor(() => expect(wrapper.emitted("saved")).toBeTruthy());
    const [savedId] = wrapper.emitted("saved")![0] as [string];
    const saved = store.getMedia(savedId)!;
    expect(saved).toMatchObject({ title: "Shot.pdf", format: "pdf" });
    expect(saved.pdf?.pageCount).toBe(1);
    expect(store.getMedia(source.id)?.title).toBe("Shot");
    wrapper.unmount();
  });

  it("reports a failure and stores nothing", async () => {
    const { wrapper } = await mountDialog("png", false);
    await wrapper.get(byId("convert-pdf-save")).trigger("click");
    await vi.waitFor(() => expect(wrapper.find(byId("convert-pdf-failure")).exists()).toBe(true));
    await flushPromises();
    expect(wrapper.get(byId("convert-pdf-failure")).text()).toMatch(/^Could not save the PDF/);
    expect(wrapper.emitted("saved")).toBeUndefined();
    expect(store.listDocuments()).toHaveLength(1);
    wrapper.unmount();
  });

  it("Cancel and Escape close without converting", async () => {
    const { wrapper } = await mountDialog();
    await wrapper.get(byId("convert-pdf-cancel")).trigger("click");
    await wrapper.get('[role="dialog"]').trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("close")).toHaveLength(2);
    expect(store.listDocuments()).toHaveLength(1);
    wrapper.unmount();
  });
});
