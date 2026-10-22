// SPEC-130: the in-cell dropdown — the button's ARIA and its place against
// the cell's right edge, the listbox (at least the cell's width, eight rows
// then scrolling) of options with the selected one marked and named by
// aria-activedescendant, focus moving into the list as it opens, the
// keyboard paths (Up/Down, Home/End, PageUp/PageDown, a typed character,
// Enter/Tab picking, Escape closing, nothing selected at first), a click
// picking, keys and presses never reaching the grid, and the unavailable
// source's one disabled option.
import { flushPromises, mount } from "@vue/test-utils";
import { afterEach, describe, expect, it } from "vitest";
import ValidationDropdown from "./ValidationDropdown.vue";

const mounted: Array<() => void> = [];
afterEach(() => {
  for (const unmount of mounted.splice(0)) {
    unmount();
  }
});

const CELL = { left: 100, top: 40, width: 64, height: 20 };
const BOUNDS = { width: 800, height: 600 };
const ITEMS = ["Low", "Medium", "High", "", "Later", "Never", "Now", "Next", "Lowest", "Max"];

async function dropdownOf(props: Record<string, unknown> = {}) {
  const host = document.createElement("div");
  document.body.append(host);
  const wrapper = mount(ValidationDropdown, { props: { cellRect: CELL, bounds: BOUNDS, ...props }, attachTo: host });
  mounted.push(() => {
    wrapper.unmount();
    host.remove();
  });
  await flushPromises();
  await pastMountMillisecond();
  return wrapper;
}

/**
 * Lets the clock leave the millisecond the dropdown mounted in, as it has by
 * the time a user presses anything. Vue stamps an event with Date.now() at the
 * first handler it reaches and skips every later handler attached in that
 * millisecond or after (its guard against listeners a patch adds mid-event),
 * so a press dispatched straight after the mount ran the button's
 * `@mousedown.prevent` but skipped the root's `@mousedown.stop`, and reached
 * the host in about half the runs.
 */
async function pastMountMillisecond() {
  const mountedAt = Date.now();
  while (Date.now() <= mountedAt) {
    await new Promise((resolve) => setTimeout(resolve, 1));
  }
}

/** A key on the open list: the emitted events it caused. */
async function key(wrapper: Awaited<ReturnType<typeof dropdownOf>>, init: KeyboardEventInit) {
  const before = { select: wrapper.emitted("select")?.length ?? 0, pick: wrapper.emitted("pick")?.length ?? 0, close: wrapper.emitted("close")?.length ?? 0 };
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  wrapper.get('[data-testid="validation-list"]').element.dispatchEvent(event);
  await flushPromises();
  return {
    prevented: event.defaultPrevented,
    select: (wrapper.emitted("select") ?? []).slice(before.select).map((args) => args[0]),
    pick: (wrapper.emitted("pick") ?? []).slice(before.pick).map((args) => args[0]),
    close: (wrapper.emitted("close") ?? []).length - before.close,
  };
}

describe("ValidationDropdown — the button", () => {
  it("is \"Choose from list\", a listbox popup, collapsed, against the cell's right edge", async () => {
    const wrapper = await dropdownOf();
    const button = wrapper.get('[data-testid="validation-dropdown"]');
    expect(button.attributes("aria-label")).toBe("Choose from list");
    expect(button.attributes("aria-haspopup")).toBe("listbox");
    expect(button.attributes("aria-expanded")).toBe("false");
    expect(button.attributes("style")).toContain("left: 164px");
    expect(button.attributes("style")).toContain("top: 40px");
    expect(button.attributes("style")).toContain("height: 20px");
    expect(wrapper.find('[data-testid="validation-list"]').exists()).toBe(false);
  });

  it("a click toggles; its press never reaches the grid and keeps focus where it is", async () => {
    const wrapper = await dropdownOf();
    let reached = 0;
    const parent = wrapper.element.parentElement!;
    parent.addEventListener("pointerdown", () => (reached += 1));
    parent.addEventListener("mousedown", () => (reached += 1));
    parent.addEventListener("dblclick", () => (reached += 1));
    const button = wrapper.get('[data-testid="validation-dropdown"]');
    const down = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    button.element.dispatchEvent(down);
    button.element.dispatchEvent(new Event("pointerdown", { bubbles: true }));
    button.element.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(down.defaultPrevented).toBe(true);
    expect(reached).toBe(0);
    await button.trigger("click");
    expect(wrapper.emitted("toggle")).toHaveLength(1);
  });
});

describe("ValidationDropdown — the list", () => {
  it("lists every item in order (a blank as an empty row), the selected one marked, at least the cell's width, eight rows tall", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: 1 });
    expect(wrapper.get('[data-testid="validation-dropdown"]').attributes("aria-expanded")).toBe("true");
    const list = wrapper.get('[data-testid="validation-list"]');
    expect(list.attributes("role")).toBe("listbox");
    const options = wrapper.findAll('[data-testid="validation-option"]');
    expect(options.map((option) => option.text())).toEqual(ITEMS);
    expect(options.map((option) => option.attributes("role"))).toEqual(ITEMS.map(() => "option"));
    expect(options.map((option) => option.attributes("aria-selected"))).toEqual(ITEMS.map((_, i) => (i === 1 ? "true" : "false")));
    expect(list.attributes("aria-activedescendant")).toBe(options[1]!.attributes("id"));
    const style = list.attributes("style") ?? "";
    expect(style).toContain("left: 100px");
    expect(style).toContain("top: 60px"); // under the cell
    expect(style).toMatch(/min-width: (\d+)px/);
    expect(Number(/min-width: (\d+)px/.exec(style)![1])).toBeGreaterThanOrEqual(CELL.width);
    expect(style).toContain("max-height: 162px"); // 8 × 20 px rows + the border
  });

  it("opens above the cell when there is no room below", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, cellRect: { ...CELL, top: 560 } });
    expect(wrapper.get('[data-testid="validation-list"]').attributes("style")).toContain("top: 398px");
  });

  it("takes focus as it opens", async () => {
    const wrapper = await dropdownOf({ items: ITEMS });
    await wrapper.setProps({ expanded: true });
    await flushPromises();
    expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-list"]').element);
  });

  it("with nothing selected, the first ArrowDown selects the first item; Up/Down move by one within the ends", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: -1 });
    expect(wrapper.get('[data-testid="validation-list"]').attributes("aria-activedescendant")).toBeUndefined();
    expect(wrapper.findAll('[aria-selected="true"]')).toHaveLength(0);
    expect((await key(wrapper, { key: "ArrowDown" })).select).toEqual([0]);
    await wrapper.setProps({ selected: 0 });
    expect((await key(wrapper, { key: "ArrowDown" })).select).toEqual([1]);
    expect((await key(wrapper, { key: "ArrowUp" })).select).toEqual([0]);
    await wrapper.setProps({ selected: 9 });
    expect((await key(wrapper, { key: "ArrowDown" })).select).toEqual([9]);
  });

  it("Home/End go to the ends, PageUp/PageDown move by eight", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: 1 });
    expect((await key(wrapper, { key: "End" })).select).toEqual([9]);
    expect((await key(wrapper, { key: "Home" })).select).toEqual([0]);
    expect((await key(wrapper, { key: "PageDown" })).select).toEqual([9]);
    await wrapper.setProps({ selected: 9 });
    expect((await key(wrapper, { key: "PageUp" })).select).toEqual([1]);
    await wrapper.setProps({ selected: -1 });
    expect((await key(wrapper, { key: "PageDown" })).select).toEqual([7]);
  });

  it("a typed character jumps to the next item starting with it, case-insensitively, wrapping", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: 0 });
    expect((await key(wrapper, { key: "l" })).select).toEqual([4]); // Later, after Low
    await wrapper.setProps({ selected: 4 });
    expect((await key(wrapper, { key: "L" })).select).toEqual([8]); // Lowest
    await wrapper.setProps({ selected: 8 });
    expect((await key(wrapper, { key: "l" })).select).toEqual([0]); // wraps to Low
    expect((await key(wrapper, { key: "q" })).select).toEqual([]); // none
  });

  it("Enter and Tab pick the selected item; with none selected they close", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: 2 });
    const enter = await key(wrapper, { key: "Enter" });
    expect([enter.pick, enter.prevented]).toEqual([[2], true]);
    expect((await key(wrapper, { key: "Tab" })).pick).toEqual([2]);
    await wrapper.setProps({ selected: -1 });
    const none = await key(wrapper, { key: "Enter" });
    expect([none.pick, none.close]).toEqual([[], 1]);
  });

  it("Escape closes; no key reaches the grid", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: 0 });
    let reached = 0;
    wrapper.element.parentElement!.addEventListener("keydown", () => (reached += 1));
    expect((await key(wrapper, { key: "Escape" })).close).toBe(1);
    await key(wrapper, { key: "ArrowDown" });
    await key(wrapper, { key: "x" });
    expect(reached).toBe(0);
  });

  it("a click on an option picks it", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS, selected: -1 });
    await wrapper.findAll('[data-testid="validation-option"]')[2]!.trigger("click");
    expect(wrapper.emitted("pick")).toEqual([[2]]);
  });

  it("focus leaving the list dismisses it", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: ITEMS });
    const outside = document.createElement("button");
    document.body.append(outside);
    outside.focus();
    expect(wrapper.emitted("dismiss")).toHaveLength(1);
    outside.remove();
  });

  it("an unavailable source shows one disabled option saying so", async () => {
    const wrapper = await dropdownOf({ expanded: true, items: [], unavailable: true });
    const option = wrapper.get('[data-testid="validation-list-unavailable"]');
    expect(option.text()).toBe("The list source isn't available");
    expect(option.attributes("role")).toBe("option");
    expect(option.attributes("aria-disabled")).toBe("true");
    expect(wrapper.findAll('[data-testid="validation-option"]')).toHaveLength(0);
    expect((await key(wrapper, { key: "ArrowDown" })).select).toEqual([]);
    expect((await key(wrapper, { key: "Enter" })).close).toBe(1);
  });
});
