import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import StyleEditorDialog from "./StyleEditorDialog.vue";

// SPEC-074 AC-18 — a paragraph style carries borders and shading, set through
// the very dialog a paragraph's are set through. The regression this file
// exists for is what the payload is made of, not just what it says: the
// dialog holds its draft in Vue reactive state, and a reactive *proxy* stored
// in the document model cannot be structured-cloned. That is how the model
// reaches IndexedDB (documentStore) and the live word-count worker
// (workers/workerRpc), so a proxy leaking through here surfaced as a
// DataCloneError console error the moment a style with borders was saved.

function open(props: Record<string, unknown> = {}) {
  return mount(StyleEditorDialog, {
    props: { mode: "modify", initialName: "Normal", initialType: "paragraph", ...props },
  });
}

function submitted(wrapper: ReturnType<typeof open>) {
  return wrapper.emitted("save")![0][0] as {
    paragraphFormatting?: { borders?: Record<string, unknown>; shadingFill?: string };
  };
}

async function boxIt(wrapper: ReturnType<typeof open>) {
  await wrapper.get('[data-testid="style-borders-button"]').trigger("click");
  await wrapper.get('[data-testid="borders-preset-box"]').trigger("click");
  await wrapper.get('[data-testid="borders-shading-on"]').setValue(true);
  await wrapper.get('[data-testid="borders-shading-color"]').setValue("#fff2cc");
  await wrapper.get('[data-testid="borders-apply"]').trigger("click");
}

describe("StyleEditorDialog borders and shading (SPEC-074 AC-18)", () => {
  it("sets a style's borders and shading through the shared dialog", async () => {
    const wrapper = open();
    await boxIt(wrapper);
    await wrapper.get('[data-testid="style-editor-save"]').trigger("click");
    const paragraphFormatting = submitted(wrapper).paragraphFormatting;
    expect(paragraphFormatting?.borders).toEqual({
      top: { style: "single", widthPt: 1, color: "#000000" },
      bottom: { style: "single", widthPt: 1, color: "#000000" },
      left: { style: "single", widthPt: 1, color: "#000000" },
      right: { style: "single", widthPt: 1, color: "#000000" },
    });
    expect(paragraphFormatting?.shadingFill).toBe("#fff2cc");
  });

  it("summarises what the style currently carries", async () => {
    const wrapper = open();
    expect(wrapper.get('[data-testid="style-borders-button"]').text()).toBe("None…");
    await boxIt(wrapper);
    expect(wrapper.get('[data-testid="style-borders-button"]').text()).toBe("Box, shaded #fff2cc…");
  });

  it("submits borders the document model can be structured-cloned with", async () => {
    const wrapper = open();
    await boxIt(wrapper);
    await wrapper.get('[data-testid="style-editor-save"]').trigger("click");
    // Exactly what documentStore and the word-count worker do to the model
    // the catalogue entry ends up inside. A reactive proxy throws here.
    expect(() => structuredClone(submitted(wrapper).paragraphFormatting)).not.toThrow();
  });

  it("keeps borders it was seeded with clonable, even when never opened", async () => {
    const wrapper = open({
      initialBorders: { bottom: { style: "dashed", widthPt: 2, color: "#c00000" } },
      initialShadingFill: "#eeeeee",
    });
    await wrapper.get('[data-testid="style-editor-save"]').trigger("click");
    const paragraphFormatting = submitted(wrapper).paragraphFormatting;
    expect(paragraphFormatting?.borders).toEqual({
      bottom: { style: "dashed", widthPt: 2, color: "#c00000" },
    });
    expect(() => structuredClone(paragraphFormatting)).not.toThrow();
  });
});
