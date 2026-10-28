// SPEC-121 — the managed sheet tab bar: Add sheet, the tab menu's three
// openers, rename in place (input, error, keys kept inside it), activation
// keys, pointer drag-reorder behind a drop indicator (Escape / outside /
// cancel), tab colours and the live region. SPEC-116's tablist behaviour is
// pinned through the real view in SpreadsheetView.spec.ts.
import { afterEach, describe, expect, it, vi } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { nextTick } from "vue";
import SheetTabs from "./SheetTabs.vue";
import { dropGapAt, gapX, moveForDrop } from "./sheetTabBar";

const TABS = [
  { index: 0, id: "a", name: "One", color: "" },
  { index: 2, id: "b", name: "Two", color: "#ff0000" },
  { index: 3, id: "c", name: "Three", color: "#00ff00" },
];

const mounted: VueWrapper[] = [];
afterEach(() => {
  mounted.splice(0).forEach((wrapper) => wrapper.unmount());
  vi.restoreAllMocks();
});

function mountTabs(props: Record<string, unknown> = {}) {
  const wrapper = mount(SheetTabs, { props: { tabs: TABS, activeIndex: 0, ...props }, attachTo: document.body });
  mounted.push(wrapper);
  return wrapper;
}

const tabs = (wrapper: VueWrapper) => wrapper.findAll('[data-testid="sheet-tab"]');

function box(left: number, right: number): DOMRect {
  return { left, right, top: 0, bottom: 30, x: left, y: 0, width: right - left, height: 30, toJSON: () => ({}) };
}

/** Tabs 100 px apart (90 wide) in a bar spanning 0..1000 × 0..30. */
function layOut() {
  vi.spyOn(Element.prototype, "getBoundingClientRect").mockImplementation(function (this: Element) {
    const position = this.getAttribute("data-tab-position");
    return position === null ? box(0, 1000) : box(Number(position) * 100, Number(position) * 100 + 90);
  });
}

function pointer(type: string, clientX: number, clientY = 10, button = 0) {
  return new MouseEvent(type, { clientX, clientY, button, bubbles: true });
}

describe("SheetTabs management (SPEC-121)", () => {
  it("emits add from the + button after the tablist", async () => {
    const wrapper = mountTabs();
    const add = wrapper.get('[data-testid="sheet-add"]');
    expect(add.text()).toBe("+");
    expect(add.attributes("aria-label")).toBe("Add sheet");
    expect(add.attributes("title")).toBe("Add sheet");
    await add.trigger("click");
    expect(wrapper.emitted("add")).toHaveLength(1);
  });

  it("disables Add sheet while the workbook is read-only (SPEC-138)", () => {
    expect(mountTabs({ readOnly: true }).get('[data-testid="sheet-add"]').attributes("disabled")).toBeDefined();
  });

  it("asks for the tab menu on right-click (pointer point) and on Shift+F10 / ContextMenu (tab's left/bottom)", async () => {
    const wrapper = mountTabs();
    await tabs(wrapper)[1].trigger("contextmenu", { clientX: 140, clientY: 700 });
    expect(wrapper.emitted("menu")?.[0]).toEqual([{ index: 2, x: 140, y: 700 }]);
    layOut();
    await tabs(wrapper)[2].trigger("keydown", { key: "F10", shiftKey: true });
    expect(wrapper.emitted("menu")?.[1]).toEqual([{ index: 3, x: 200, y: 30 }]);
    // The browser's own contextmenu echo of the key is swallowed.
    await tabs(wrapper)[2].trigger("contextmenu", { clientX: 1, clientY: 1 });
    expect(wrapper.emitted("menu")).toHaveLength(2);
  });

  it("starts a rename on double-click and on F2; Enter and Space activate", async () => {
    const wrapper = mountTabs();
    await tabs(wrapper)[1].trigger("dblclick");
    await tabs(wrapper)[0].trigger("keydown", { key: "F2" });
    expect(wrapper.emitted("rename-start")).toEqual([[2], [0]]);
    await tabs(wrapper)[2].trigger("keydown", { key: "Enter" });
    await tabs(wrapper)[1].trigger("keydown", { key: " " });
    expect(wrapper.emitted("select")).toEqual([[3], [2]]);
  });

  it("renames in place: a focused, selected input with the draft; Enter, Escape, typing and blur emit; keys stay inside it", async () => {
    const wrapper = mountTabs({ renamingIndex: 2, renameDraft: "Two" });
    await nextTick();
    const input = wrapper.get('[data-testid="sheet-tab-rename-input"]');
    const element = input.element as HTMLInputElement;
    expect(element.value).toBe("Two");
    expect(input.attributes("aria-label")).toBe("Sheet name");
    expect(document.activeElement).toBe(element);
    expect([element.selectionStart, element.selectionEnd]).toEqual([0, 3]);
    expect(tabs(wrapper).map((tab) => tab.text())).toEqual(["One", "Three"]);

    await input.setValue("Budget");
    expect(wrapper.emitted("rename-input")?.at(-1)).toEqual(["Budget"]);
    for (const key of ["ArrowRight", "F2", "Home"]) {
      await input.trigger("keydown", { key });
    }
    expect(wrapper.emitted("select")).toBeUndefined();
    expect(wrapper.emitted("rename-start")).toBeUndefined();
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("rename-commit")).toHaveLength(1);
    const escape = new KeyboardEvent("keydown", { key: "Escape", bubbles: true });
    const outer = vi.fn();
    document.body.addEventListener("keydown", outer);
    element.dispatchEvent(escape);
    document.body.removeEventListener("keydown", outer);
    expect(wrapper.emitted("rename-cancel")).toHaveLength(1);
    expect(outer).not.toHaveBeenCalled();
    await input.trigger("blur");
    expect(wrapper.emitted("rename-blur")).toHaveLength(1);
  });

  it("shows the name error as an alert and marks the input invalid", async () => {
    const wrapper = mountTabs({ renamingIndex: 0, renameDraft: "", renameError: "A sheet name can't be blank" });
    const error = wrapper.get('[data-testid="sheet-name-error"]');
    expect(error.attributes("role")).toBe("alert");
    expect(error.text()).toBe("A sheet name can't be blank");
    expect(wrapper.get('[data-testid="sheet-tab-rename-input"]').attributes("aria-invalid")).toBe("true");
    await wrapper.setProps({ renameError: "" });
    expect(wrapper.find('[data-testid="sheet-name-error"]').exists()).toBe(false);
  });

  it("hands focus back to the tab when the rename ends", async () => {
    const wrapper = mountTabs({ renamingIndex: 0, renameDraft: "One" });
    await nextTick();
    await wrapper.setProps({ renamingIndex: -1 });
    await nextTick();
    expect(document.activeElement).toBe(tabs(wrapper)[0].element);
  });

  it("drags a tab past 4 px behind a drop indicator and emits move with model indices; the drag's click selects nothing", async () => {
    const wrapper = mountTabs();
    layOut();
    await tabs(wrapper)[0].trigger("pointerdown", { clientX: 40, clientY: 10, button: 0 });
    window.dispatchEvent(pointer("pointermove", 43));
    await nextTick();
    expect(wrapper.find('[data-testid="sheet-tab-drop-indicator"]').exists()).toBe(false);
    window.dispatchEvent(pointer("pointermove", 260));
    await nextTick();
    const indicator = wrapper.get('[data-testid="sheet-tab-drop-indicator"]');
    expect(indicator.attributes("style")).toContain("left: 290px"); // after "Three"
    window.dispatchEvent(pointer("pointerup", 260));
    await tabs(wrapper)[0].trigger("click");
    await nextTick();
    expect(wrapper.emitted("move")).toEqual([[{ index: 0, before: -1 }]]);
    expect(wrapper.emitted("select")).toBeUndefined();
    expect(wrapper.find('[data-testid="sheet-tab-drop-indicator"]').exists()).toBe(false);
  });

  it("cancels a drag on Escape, on a drop outside the bar, on pointercancel, and on a drop in place", async () => {
    const wrapper = mountTabs();
    layOut();
    const drag = async (end: () => void) => {
      await tabs(wrapper)[2].trigger("pointerdown", { clientX: 240, clientY: 10, button: 0 });
      window.dispatchEvent(pointer("pointermove", 20));
      await nextTick();
      expect(wrapper.find('[data-testid="sheet-tab-drop-indicator"]').exists()).toBe(true);
      end();
      await nextTick();
    };
    await drag(() => {
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
      window.dispatchEvent(pointer("pointerup", 20));
    });
    await drag(() => window.dispatchEvent(pointer("pointerup", 20, 400)));
    await drag(() => window.dispatchEvent(pointer("pointercancel", 20)));
    await drag(() => window.dispatchEvent(pointer("pointerup", 250)));
    expect(wrapper.emitted("move")).toBeUndefined();
    expect(wrapper.find('[data-testid="sheet-tab-drop-indicator"]').exists()).toBe(false);
  });

  it("a press that does not travel is a click; a right press never drags", async () => {
    const wrapper = mountTabs();
    layOut();
    await tabs(wrapper)[1].trigger("pointerdown", { clientX: 140, clientY: 10, button: 2 });
    window.dispatchEvent(pointer("pointermove", 20));
    await nextTick();
    expect(wrapper.find('[data-testid="sheet-tab-drop-indicator"]').exists()).toBe(false);
    await tabs(wrapper)[1].trigger("pointerdown", { clientX: 140, clientY: 10, button: 0 });
    window.dispatchEvent(pointer("pointerup", 142));
    await tabs(wrapper)[1].trigger("click");
    expect(wrapper.emitted("select")).toEqual([[2]]);
  });

  it("colours tabs: a 3 px bar on the active coloured tab, a tint on the others, data-tab-color on each", async () => {
    const wrapper = mountTabs({ activeIndex: 2 });
    const [one, two, three] = tabs(wrapper);
    expect(one.attributes("data-tab-color")).toBeUndefined();
    expect(two.attributes("data-tab-color")).toBe("#ff0000");
    expect(two.find(".sheet-tab-color-bar").exists()).toBe(true);
    expect(two.attributes("style") ?? "").not.toContain("color-mix");
    expect(three.attributes("data-tab-color")).toBe("#00ff00");
    expect(three.find(".sheet-tab-color-bar").exists()).toBe(false);
    expect(three.attributes("style")).toMatch(/color-mix\(in srgb, (#00ff00|rgb\(0, 255, 0\)) 25%, transparent\)/);
    expect(two.text()).toBe("Two");
  });

  it("reads the announcement out through a polite live region", () => {
    const wrapper = mountTabs({ announcement: 'Sheet "Two" moved to position 1' });
    const region = wrapper.get('[data-testid="sheet-tabs-announcement"]');
    expect(region.attributes("aria-live")).toBe("polite");
    expect(region.text()).toBe('Sheet "Two" moved to position 1');
  });
});

describe("sheetTabBar geometry", () => {
  const spans = [
    { left: 0, right: 90 },
    { left: 100, right: 190 },
  ];
  it("finds the nearest gap and its x", () => {
    expect([10, 50, 120, 160, 400].map((x) => dropGapAt(x, spans))).toEqual([0, 1, 1, 2, 2]);
    expect([0, 1, 2].map((gap) => gapX(gap, spans))).toEqual([0, 95, 190]);
  });
  it("maps a drop onto model indices, null when nothing moves", () => {
    expect(moveForDrop(TABS, 0, 0)).toBeNull();
    expect(moveForDrop(TABS, 0, 1)).toBeNull();
    expect(moveForDrop(TABS, 0, 2)).toEqual({ index: 0, before: 3 });
    expect(moveForDrop(TABS, 2, 0)).toEqual({ index: 3, before: 0 });
    expect(moveForDrop(TABS, 1, 3)).toEqual({ index: 2, before: -1 });
  });
});
