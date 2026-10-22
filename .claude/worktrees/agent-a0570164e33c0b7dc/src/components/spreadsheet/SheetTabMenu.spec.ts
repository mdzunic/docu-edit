// SPEC-121 — the sheet tab context menu: shared MenuItems (disabled with a
// title reason, never hidden), focus on the first enabled item, the arrow
// ring over enabled items, Escape/Tab/backdrop closing, and viewport clamping.
import { afterEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import SheetTabMenu from "./SheetTabMenu.vue";
import { clampToViewport } from "./sheetTabBar";

const ITEMS = [
  { action: "insert", label: "Insert sheet", testid: "sheet-tab-menu-insert", disabled: true, reason: "No room" },
  { action: "rename", label: "Rename", testid: "sheet-tab-menu-rename" },
  { action: "delete", label: "Delete", testid: "sheet-tab-menu-delete", disabled: true, reason: "A workbook must keep at least one visible sheet" },
  { action: "hide", label: "Hide", testid: "sheet-tab-menu-hide" },
];

const mounted: VueWrapper[] = [];
const focusedId = () => document.activeElement?.getAttribute("data-testid");
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

async function openMenu() {
  const wrapper = mount(SheetTabMenu, { props: { items: ITEMS, x: 20, y: 30 }, attachTo: document.body });
  mounted.push(wrapper);
  await flushPromises();
  return wrapper;
}

describe("SheetTabMenu (SPEC-121)", () => {
  it("renders every item as a menuitem, disabled ones with their reason as the title", async () => {
    const wrapper = await openMenu();
    const menu = wrapper.get('[data-testid="sheet-tab-menu"]');
    expect(menu.attributes("role")).toBe("menu");
    expect(menu.attributes("aria-label")).toBe("Sheet options");
    expect(menu.findAll('[role="menuitem"]')).toHaveLength(4);
    const remove = wrapper.get('[data-testid="sheet-tab-menu-delete"]');
    expect(remove.attributes("disabled")).toBeDefined();
    expect(remove.attributes("title")).toBe("A workbook must keep at least one visible sheet");
  });

  it("focuses the first enabled item and wraps the arrow keys over enabled items only", async () => {
    const wrapper = await openMenu();
    expect(focusedId()).toBe("sheet-tab-menu-rename");
    const menu = wrapper.get('[data-testid="sheet-tab-menu"]');
    await menu.trigger("keydown", { key: "ArrowDown" });
    expect(focusedId()).toBe("sheet-tab-menu-hide");
    await menu.trigger("keydown", { key: "ArrowDown" });
    expect(focusedId()).toBe("sheet-tab-menu-rename");
    await menu.trigger("keydown", { key: "ArrowUp" });
    expect(focusedId()).toBe("sheet-tab-menu-hide");
    await menu.trigger("keydown", { key: "Home" });
    expect(focusedId()).toBe("sheet-tab-menu-rename");
    await menu.trigger("keydown", { key: "End" });
    expect(focusedId()).toBe("sheet-tab-menu-hide");
  });

  it("runs the activated item and closes on Escape, Tab and the backdrop", async () => {
    const wrapper = await openMenu();
    await wrapper.get('[data-testid="sheet-tab-menu-hide"]').trigger("click");
    await wrapper.get('[data-testid="sheet-tab-menu-delete"]').trigger("click");
    expect(wrapper.emitted("run")).toEqual([["hide"]]);
    const menu = wrapper.get('[data-testid="sheet-tab-menu"]');
    await menu.trigger("keydown", { key: "Escape" });
    await menu.trigger("keydown", { key: "Tab" });
    await wrapper.get('[data-testid="sheet-tab-menu-backdrop"]').trigger("click");
    await wrapper.get('[data-testid="sheet-tab-menu-backdrop"]').trigger("contextmenu");
    expect(wrapper.emitted("close")).toHaveLength(4);
  });

  it("stays on screen: shifted left and flipped above the point at the viewport's edges", async () => {
    expect(clampToViewport({ x: 20, y: 30 }, { width: 200, height: 300 }, { width: 1000, height: 800 })).toEqual({ x: 20, y: 30 });
    expect(clampToViewport({ x: 950, y: 780 }, { width: 200, height: 300 }, { width: 1000, height: 800 })).toEqual({ x: 792, y: 480 });
    vi.spyOn(Element.prototype, "getBoundingClientRect").mockReturnValue({ width: 200, height: 300 } as DOMRect);
    const wrapper = mount(SheetTabMenu, { props: { items: ITEMS, x: window.innerWidth - 10, y: window.innerHeight - 5 }, attachTo: document.body });
    mounted.push(wrapper);
    await flushPromises();
    const style = wrapper.get('[data-testid="sheet-tab-menu"]').attributes("style");
    expect(style).toContain(`left: ${window.innerWidth - 208}px`);
    expect(style).toContain(`top: ${window.innerHeight - 308}px`);
    vi.restoreAllMocks();
  });
});
