// SPEC-108's sheet preview draws core/sheetLayout's own numbers: the paper,
// the margin guides, the placed box and — for a cropping fit — the cut-away
// region, with the page and placed sizes beneath in the active unit.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SheetPreview from "./SheetPreview.vue";
import { placeOnSheet, resolveSheet, SHEET_DEFAULTS, type SheetFit } from "../../lib/core/sheetLayout";
import { MARGIN_PRESETS } from "../../lib/preferences";

const ITEM = { widthPx: 64, heightPx: 40 };

function mountFor(fit: SheetFit, unit = "cm", background = "#ffffff") {
  const geometry = resolveSheet({ ...SHEET_DEFAULTS, margins: { ...MARGIN_PRESETS.Narrow } });
  const placement = placeOnSheet(ITEM, geometry.content, fit, "center");
  return mount(SheetPreview, { props: { geometry, placement, unit, background, imageUrl: "data:image/png;base64,AA==" } });
}

function box(wrapper: ReturnType<typeof mountFor>, id: string) {
  const el = wrapper.get(`[data-testid="${id}"]`);
  return ["x", "y", "width", "height"].map((name) => Number(el.attributes(name)));
}

describe("SheetPreview", () => {
  it("contain: paper, margin guides and the placed box, nothing cropped", () => {
    const wrapper = mountFor("contain");
    expect(box(wrapper, "sheet-preview-paper")).toEqual([0, 0, 21, 29.7]);
    expect(box(wrapper, "sheet-preview-margins")).toEqual([1.27, 1.27, 18.46, 27.16]);
    expect(box(wrapper, "sheet-preview-placed")).toEqual([1.27, 9.081, 18.46, 11.538]);
    expect(wrapper.find('[data-testid="sheet-preview-cropped"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="sheet-preview-page-size"]').text()).toBe("Page 21 × 29.7 cm");
    expect(wrapper.get('[data-testid="sheet-preview-placed-size"]').text()).toBe("Placed 18.46 × 11.54 cm");
  });

  it("cover: the placed box is the content box and the overflow is shaded", () => {
    const wrapper = mountFor("cover");
    expect(box(wrapper, "sheet-preview-placed")).toEqual([1.27, 1.27, 18.46, 27.16]);
    // The item is 43.456 × 27.16 cm, centred: 12.498 cm overflow each side.
    const cropped = wrapper.get('[data-testid="sheet-preview-cropped"]');
    expect(cropped.attributes("d")).toBe("M-11.228 1.27h43.456v27.16h-43.456ZM1.27 1.27h18.46v27.16h-18.46Z");
    expect(cropped.attributes("fill-rule")).toBe("evenodd");
    // The view box widens to show the overflow beside the paper.
    const [left] = wrapper.get("svg").attributes("viewBox")!.split(" ").map(Number);
    expect(left).toBeLessThan(-11.228);
  });

  it("states sizes in the active unit, and draws a transparent sheet as a checkerboard", () => {
    const wrapper = mountFor("contain", "in", "transparent");
    expect(wrapper.get('[data-testid="sheet-preview-page-size"]').text()).toBe("Page 8.27 × 11.69 in");
    expect(wrapper.get('[data-testid="sheet-preview-paper"]').attributes("fill")).toMatch(/^url\(#.+-checker\)$/);
  });
});
