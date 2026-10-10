// SPEC-121 "Tab colour": the core TextColorPicker opened at a point, its
// theme row from the workbook theme, a "No colour" row inside its open menu,
// and the three outcomes — a swatch (hex or ThemeColorRef), clear, close.
import { afterEach, describe, expect, it } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import SheetTabColorPicker from "./SheetTabColorPicker.vue";
import { DEFAULT_THEME } from "../../lib/core/theme";

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

async function open() {
  const wrapper = mount(SheetTabColorPicker, {
    props: { x: 40, y: 500, theme: DEFAULT_THEME, recent: ["#123456"] },
    attachTo: document.body,
  });
  mounted.push(wrapper);
  await flushPromises();
  return wrapper;
}

describe("SheetTabColorPicker (SPEC-121)", () => {
  it("opens the picker menu at the point, with No colour inside it and focused", async () => {
    const wrapper = await open();
    const root = wrapper.get('[data-testid="sheet-tab-color-picker"]');
    expect(root.attributes("style")).toContain("left: 40px");
    expect(root.attributes("style")).toContain("top: 500px");
    const menu = wrapper.get('[data-testid="text-color-menu"]');
    expect(menu.isVisible()).toBe(true);
    const none = menu.get('[data-testid="sheet-tab-color-none"]');
    expect(none.text()).toBe("No colour");
    expect(document.activeElement).toBe(none.element);
    expect(menu.findAll('[data-testid="text-color-theme-swatch"]')).toHaveLength(60);
    expect(menu.findAll('[data-testid="text-color-recent-swatch"]')).toHaveLength(1);
  });

  it("emits a theme swatch as its reference, a palette swatch as hex, and clear for No colour", async () => {
    const wrapper = await open();
    const swatch = wrapper.get('[data-testid="text-color-theme-swatch"]');
    await swatch.trigger("click");
    expect(wrapper.emitted("select")?.[0]?.[0]).toEqual({ theme: swatch.attributes("data-theme-slot") });
    const other = await open();
    await other.get('[data-testid="text-color-swatch"]').trigger("click");
    expect(other.emitted("select")).toEqual([["#000000"]]);
    expect(other.emitted("close")).toBeUndefined();
    const third = await open();
    await third.get('[data-testid="sheet-tab-color-none"]').trigger("click");
    expect(third.emitted("clear")).toHaveLength(1);
    expect(third.emitted("select")).toBeUndefined();
  });

  it("emits close when dismissed without a choice (Escape, outside click)", async () => {
    const wrapper = await open();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(wrapper.emitted("close")).toHaveLength(1);
    const other = await open();
    document.body.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(other.emitted("close")).toHaveLength(1);
    expect(other.emitted("select")).toBeUndefined();
  });
});
