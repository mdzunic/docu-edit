// SPEC-112: the image editor inside /media/:id — the Edit action, tool
// switching, keyboard operation of the crop rectangle, the history list, the
// rotate shortcuts — and every control in every tool keyboard-reachable and
// named for a screen reader. The browser-bound steps are fakes (jsdom has no
// canvas); the session and the components are the real ones.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import type { ImageEditSteps } from "../lib/media/renderEdits";

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  vi.doMock("../lib/media/renderEdits", async (importOriginal) => {
    const actual = await importOriginal<typeof import("../lib/media/renderEdits")>();
    const steps: ImageEditSteps = {
      ...actual.BROWSER_EDIT_STEPS,
      decode: async () => ({ image: {} as CanvasImageSource, widthPx: 160, heightPx: 120, release: () => {} }),
      preview: (decoded) => ({ image: decoded.image, naturalPx: { widthPx: decoded.widthPx, heightPx: decoded.heightPx } }),
      render: async (_source, _edits, target) => ({ width: target.widthPx, height: target.heightPx }) as HTMLCanvasElement,
      encode: async () => new Uint8Array([1, 2, 3]),
      canvasCeiling: () => 4096 * 4096,
      release: () => {},
    };
    return { ...actual, BROWSER_EDIT_STEPS: steps };
  });
  Object.assign(URL, { createObjectURL: vi.fn(() => "blob:test"), revokeObjectURL: vi.fn() });
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});

let wrapper: VueWrapper | null = null;
afterEach(() => {
  wrapper?.unmount();
  wrapper = null;
  vi.doUnmock("../lib/idb");
  vi.doUnmock("../lib/media/renderEdits");
  vi.restoreAllMocks();
});

async function openPhoto(format: "jpeg" | "pdf" = "jpeg") {
  const store = await import("../lib/documentStore");
  const { assembleMediaModel } = await import("../lib/media/open");
  const fixtures = await import("../lib/media/fixtures/build");
  const { routes } = await import("../router");
  const { default: MediaView } = await import("./MediaView.vue");
  const { readFileSync } = await import("node:fs");
  const { join } = await import("node:path");
  const bytes =
    format === "pdf"
      ? new Uint8Array(readFileSync(join(__dirname, "../../e2e/fixtures/spec107-pages.pdf")))
      : fixtures.buildJpegFixtureBytes();
  const model = await assembleMediaModel("photo", format, bytes);
  store.storeMedia(model);
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/media/${model.id}`);
  wrapper = mount(MediaView, { global: { plugins: [router] }, attachTo: document.body });
  await flushPromises();
  return { wrapper, store, model };
}

async function enterEditor(view: VueWrapper) {
  await view.get('[data-testid="media-edit-open"]').trigger("click");
  await flushPromises();
  expect(view.find('[data-testid="media-edit"]').exists()).toBe(true);
}

const byTestId = (view: VueWrapper, id: string) => view.find(`[data-testid="${id}"]`);

/** A control's accessible name: aria-label(ledby), a wrapping or `for`
 * label, a button's text, or its title. */
function accessibleName(element: Element): string {
  const aria = element.getAttribute("aria-label");
  if (aria) return aria;
  const labelledBy = element.getAttribute("aria-labelledby");
  if (labelledBy) return document.getElementById(labelledBy)?.textContent?.trim() ?? "";
  const wrapping = element.closest("label");
  if (wrapping) return wrapping.textContent?.trim() ?? "";
  const id = element.getAttribute("id");
  if (id) {
    const label = document.querySelector(`label[for="${id}"]`);
    if (label) return label.textContent?.trim() ?? "";
  }
  if (element.tagName === "BUTTON") return element.textContent?.trim() || element.getAttribute("title") || "";
  return "";
}

describe("MediaView — the image editor", () => {
  it("offers Edit for an image and not for a PDF", async () => {
    const pdf = await openPhoto("pdf");
    expect(byTestId(pdf.wrapper, "media-edit-open").exists()).toBe(false);
    pdf.wrapper.unmount();
    wrapper = null;
    const { wrapper: view } = await openPhoto();
    expect(byTestId(view, "media-edit-open").text()).toBe("Edit");
    await enterEditor(view);
    // The viewer gives way to the editor; the Edit action hides; Done leaves.
    expect(byTestId(view, "media-image").exists()).toBe(false);
    expect(byTestId(view, "media-edit-open").exists()).toBe(false);
    await byTestId(view, "edit-done").trigger("click");
    expect(byTestId(view, "media-edit").exists()).toBe(false);
    expect(byTestId(view, "media-image").exists()).toBe(true);
  });

  it("switches tools from the rail, each with its own options bar", async () => {
    const { wrapper: view } = await openPhoto();
    await enterEditor(view);
    const tools = ["tool-crop", "tool-rotate", "tool-adjust", "tool-redact", "tool-annotate"];
    for (const id of tools) {
      expect(byTestId(view, id).exists()).toBe(true);
    }
    expect(byTestId(view, "tool-crop").attributes("aria-pressed")).toBe("true");
    expect(byTestId(view, "crop-aspect").exists()).toBe(true);

    await byTestId(view, "tool-rotate").trigger("click");
    expect(byTestId(view, "tool-rotate").attributes("aria-pressed")).toBe("true");
    expect(byTestId(view, "tool-crop").attributes("aria-pressed")).toBe("false");
    expect(byTestId(view, "edit-rotate-right").exists()).toBe(true);
    expect(byTestId(view, "crop-aspect").exists()).toBe(false);

    await byTestId(view, "tool-adjust").trigger("click");
    // Native sliders with number inputs beside them.
    for (const channel of ["brightness", "contrast", "saturation", "exposure"]) {
      expect(byTestId(view, `adjust-${channel}`).attributes("type")).toBe("range");
      expect(byTestId(view, `adjust-${channel}-number`).attributes("type")).toBe("number");
    }
    expect(byTestId(view, "adjust-grayscale").attributes("type")).toBe("checkbox");

    await byTestId(view, "tool-redact").trigger("click");
    expect(byTestId(view, "redact-color").exists()).toBe(true);
    expect(byTestId(view, "edit-draw-surface").exists()).toBe(true);

    await byTestId(view, "tool-annotate").trigger("click");
    expect(byTestId(view, "annotate-shape").findAll("option").map((option) => option.attributes("value"))).toEqual([
      "arrow",
      "line",
      "rectangle",
      "ellipse",
      "freehand",
      "highlight",
      "text",
    ]);

    // Arrow keys move along the rail.
    const rail = byTestId(view, "tool-crop").element as HTMLButtonElement;
    rail.focus();
    await byTestId(view, "tool-crop").trigger("keydown", { key: "ArrowDown" });
    expect(document.activeElement).toBe(byTestId(view, "tool-rotate").element);
  });

  it("operates the crop rectangle from the keyboard and announces it", async () => {
    const { wrapper: view } = await openPhoto();
    await enterEditor(view);
    await byTestId(view, "crop-aspect").setValue("16:9");
    const announce = byTestId(view, "edit-crop-announce");
    expect(announce.text()).toBe("0, 15, 160 × 90");
    expect(announce.attributes("role")).toBe("status");
    await byTestId(view, "crop-aspect").setValue("free");

    const handles = view.findAll(".crop-handle");
    expect(handles).toHaveLength(9);
    for (const handle of handles) {
      expect(handle.element.tagName).toBe("BUTTON");
      expect(handle.attributes("aria-label")).toMatch(/crop .*— 0, 15, 160 × 90$|^Crop area — 0, 15, 160 × 90$/);
    }
    const left = byTestId(view, "crop-handle-w");
    await left.trigger("keydown", { key: "ArrowRight" });
    expect(announce.text()).toBe("1, 15, 159 × 90");
    await left.trigger("keydown", { key: "ArrowRight", shiftKey: true });
    expect(announce.text()).toBe("11, 15, 149 × 90");
    await byTestId(view, "crop-handle-move").trigger("keydown", { key: "ArrowDown" });
    expect(announce.text()).toBe("11, 16, 149 × 90");
    expect(byTestId(view, "edit-undo").attributes("title")).toBe("Undo Crop");
    // The details panel reads the live stack.
    expect(byTestId(view, "detail-dimensions").text()).toBe("149 × 90 px");
  });

  it("lists the edits with a remove per row, and rotates on Cmd/Ctrl+[ and ]", async () => {
    const { wrapper: view } = await openPhoto();
    await enterEditor(view);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "]", ctrlKey: true }));
    await byTestId(view, "crop-aspect").setValue("1:1");
    await byTestId(view, "tool-adjust").trigger("click");
    await byTestId(view, "adjust-invert").setValue(true);
    await flushPromises();
    const history = byTestId(view, "media-edit-history");
    expect(history.findAll('[data-testid="edit-history-row"]').map((row) => row.find(".label").text())).toEqual([
      "Rotate right",
      "Crop",
      "Adjust invert",
    ]);
    expect(byTestId(view, "detail-dimensions").text()).toBe("120 × 120 px");

    const removes = history.findAll('[data-testid="edit-history-remove"]');
    expect(removes[1].attributes("aria-label")).toBe("Remove Crop");
    await removes[1].trigger("click");
    expect(history.findAll('[data-testid="edit-history-row"]')).toHaveLength(2);
    expect(byTestId(view, "detail-dimensions").text()).toBe("120 × 160 px");

    window.dispatchEvent(new KeyboardEvent("keydown", { key: "[", ctrlKey: true }));
    await flushPromises();
    expect(byTestId(view, "detail-dimensions").text()).toBe("160 × 120 px");
    expect(history.findAll('[data-testid="edit-history-row"]').map((row) => row.find(".label").text())).toEqual(["Adjust invert"]);
  });

  it("names every control in every tool for a screen reader", async () => {
    const { wrapper: view } = await openPhoto();
    await enterEditor(view);
    await byTestId(view, "tool-redact").trigger("click");
    await byTestId(view, "redact-add").trigger("click");
    const unnamed: string[] = [];
    for (const tool of ["tool-crop", "tool-rotate", "tool-adjust", "tool-redact", "tool-annotate"]) {
      await byTestId(view, tool).trigger("click");
      if (tool === "tool-annotate") {
        await byTestId(view, "annotate-shape").setValue("text");
      }
      const editor = byTestId(view, "media-edit").element;
      for (const control of editor.querySelectorAll("button, input, select, [tabindex='0']")) {
        if (!accessibleName(control)) {
          unnamed.push(`${tool}: ${control.outerHTML.slice(0, 90)}`);
        }
      }
    }
    expect(unnamed).toEqual([]);
    // The redaction's object row is focusable, nudged by arrows, removed by Delete.
    await byTestId(view, "tool-redact").trigger("click");
    const object = byTestId(view, "edit-object");
    expect(object.attributes("tabindex")).toBe("0");
    const before = object.attributes("aria-label");
    await object.trigger("keydown", { key: "ArrowRight", shiftKey: true });
    expect(byTestId(view, "edit-object").attributes("aria-label")).not.toBe(before);
    await byTestId(view, "edit-object").trigger("keydown", { key: "Delete" });
    expect(byTestId(view, "edit-object").exists()).toBe(false);
  });

  it("shows the redaction notice while a redaction is on the stack, and asks before flattening", async () => {
    const { wrapper: view } = await openPhoto();
    await enterEditor(view);
    expect(byTestId(view, "edit-redaction-notice").exists()).toBe(false);
    await byTestId(view, "tool-redact").trigger("click");
    await byTestId(view, "redact-add").trigger("click");
    expect(byTestId(view, "edit-redaction-notice").text()).toBe(
      "The original is still stored. Choose Flatten and discard original to remove the hidden pixels.",
    );
    await byTestId(view, "edit-flatten").trigger("click");
    const dialog = byTestId(view, "edit-flatten-dialog");
    expect(dialog.exists()).toBe(true);
    expect(byTestId(view, "edit-flatten-message").text()).toContain("permanent and cannot be undone");
    await byTestId(view, "edit-flatten-cancel").trigger("click");
    expect(byTestId(view, "edit-flatten-dialog").exists()).toBe(false);
    expect(byTestId(view, "edit-redaction-notice").exists()).toBe(true);
  });
});
