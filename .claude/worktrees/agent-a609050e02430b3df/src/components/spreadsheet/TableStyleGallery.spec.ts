// SPEC-133 "Table styles": the gallery offers None and Excel's 60 built-ins,
// marks the current one, emits `pick` with the name (undefined for None), and
// previews each style through tableCellLook against the workbook theme.
import { afterEach, describe, expect, it } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import TableStyleGallery from "./TableStyleGallery.vue";
import { BUILT_IN_THEMES, DEFAULT_THEME, hexToRgb } from "../../lib/core/theme";

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

function open(props: Record<string, unknown> = {}) {
  const wrapper = mount(TableStyleGallery, { props: { theme: DEFAULT_THEME, ...props } });
  mounted.push(wrapper);
  return wrapper;
}

/** A hex colour as jsdom reports an inline style's colour. */
function rgb(hex: string): string {
  return `rgb(${hexToRgb(hex).map((channel) => Math.round(channel * 255)).join(", ")})`;
}

/** The background colours of TableStyleMedium2's header preview cells. */
function header(wrapper: VueWrapper): string[] {
  return wrapper.findAll('[data-testid="sheet-table-style-TableStyleMedium2"] .preview-cell[data-row="0"]').map((cell) => (cell.element as HTMLElement).style.backgroundColor);
}

describe("TableStyleGallery (SPEC-133)", () => {
  it("offers None and the 60 built-ins as options, the current one selected", () => {
    const wrapper = open({ current: "TableStyleMedium2" });
    const gallery = wrapper.get('[data-testid="sheet-table-style-gallery"]');
    expect(gallery.attributes("role")).toBe("listbox");
    const options = gallery.findAll('[role="option"]');
    expect(options).toHaveLength(61);
    expect(options[0]!.attributes("data-testid")).toBe("sheet-table-style-none");
    expect(options[1]!.attributes("aria-label")).toBe("TableStyleLight1");
    expect(options.at(-1)!.attributes("data-testid")).toBe("sheet-table-style-TableStyleDark11");
    const selected = options.filter((option) => option.attributes("aria-selected") === "true");
    expect(selected.map((option) => option.attributes("aria-label"))).toEqual(["TableStyleMedium2"]);
    // No current style: None is selected.
    expect(open().get('[data-testid="sheet-table-style-none"]').attributes("aria-selected")).toBe("true");
  });

  it("emits pick with the style's name, and undefined for None", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="sheet-table-style-TableStyleMedium9"]').trigger("click");
    await wrapper.get('[data-testid="sheet-table-style-none"]').trigger("click");
    expect(wrapper.emitted("pick")).toEqual([["TableStyleMedium9"], [undefined]]);
  });

  it("previews TableStyleMedium2's header in the theme's accent1, banded rows beneath", () => {
    expect(header(open())).toEqual(Array(4).fill(rgb(DEFAULT_THEME.colors.accent1)));
    const facet = BUILT_IN_THEMES.find((theme) => theme.name === "Facet")!;
    expect(header(open({ theme: facet }))).toEqual(Array(4).fill(rgb(facet.colors.accent1)));
    // The second data row is unbanded; None paints nothing; the lines are drawn.
    const wrapper = open();
    const secondRow = wrapper.get('[data-testid="sheet-table-style-TableStyleMedium2"] .preview-cell[data-row="2"]');
    expect((secondRow.element as HTMLElement).style.backgroundColor).toBe("");
    expect(wrapper.findAll('[data-testid="sheet-table-style-none"] .preview-line')).toHaveLength(0);
    expect(wrapper.findAll('[data-testid="sheet-table-style-TableStyleLight1"] .preview-line').length).toBeGreaterThan(0);
  });
});
