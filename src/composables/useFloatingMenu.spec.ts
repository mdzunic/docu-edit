import { describe, it, expect } from "vitest";
import { nextTick, ref, type Ref } from "vue";
import { useFloatingMenu } from "./useFloatingMenu";

interface FakeRect {
  left: number;
  right: number;
  bottom: number;
  width: number;
  top?: number;
  height?: number;
}

// jsdom's getBoundingClientRect is all zeros, so the tests stub it the way
// a browser would measure the real anchor/menu.
function stubRect(el: Element, rect: FakeRect) {
  Object.defineProperty(el, "getBoundingClientRect", {
    configurable: true,
    value: () => rect,
  });
}

function setup(options: { placement?: "below" | "above" | "right"; offsetY?: number } = {}) {
  const open = ref(false);
  const anchorEl = ref<HTMLElement | null>(null) as Ref<HTMLElement | null>;
  const menuEl = ref<HTMLElement | null>(null);
  const { menuStyle } = useFloatingMenu({
    open: () => open.value,
    anchor: () => anchorEl.value,
    menu: () => menuEl.value,
    ...options,
  });
  return { open, anchorEl, menuEl, menuStyle };
}

describe("useFloatingMenu", () => {
  it("anchors the menu to a fixed position just below the trigger button", async () => {
    const { open, anchorEl, menuStyle } = setup();
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 40, right: 60, bottom: 60, width: 20 });

    open.value = true;
    await nextTick();

    expect(menuStyle.value).toEqual({
      position: "fixed",
      left: "40px",
      top: "64px",
    });
  });

  it("keeps the menu inside the viewport when the anchor hugs the right edge", async () => {
    const { open, anchorEl, menuEl, menuStyle } = setup();
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 1000, right: 1020, bottom: 60, width: 20 });

    open.value = true;
    await nextTick();
    menuEl.value = document.createElement("div");
    stubRect(menuEl.value, { left: 0, right: 200, bottom: 100, width: 200 });

    // Repositioning (as a toolbar scroll/resize would trigger) reads the
    // now-measurable menu width and clamps the left edge.
    window.dispatchEvent(new Event("resize"));

    expect(menuStyle.value.left).toBe(`${window.innerWidth - 200 - 8}px`);
    expect(menuStyle.value.position).toBe("fixed");
  });

  it("places a right-placement flyout beside the trigger, top-aligned (SPEC-059 submenu)", async () => {
    const { open, anchorEl, menuEl, menuStyle } = setup({ placement: "right", offsetY: 0 });
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 100, right: 340, top: 150, bottom: 178, width: 240 });
    menuEl.value = document.createElement("div");
    stubRect(menuEl.value, { left: 0, right: 224, bottom: 118, width: 224, top: 0, height: 118 });

    open.value = true;
    await nextTick();

    expect(menuStyle.value).toEqual({
      position: "fixed",
      left: "340px",
      top: "150px",
    });
  });

  it("flips a right-placement flyout to the trigger's left at the viewport edge", async () => {
    const { open, anchorEl, menuEl, menuStyle } = setup({ placement: "right", offsetY: 0 });
    anchorEl.value = document.createElement("button");
    const right = window.innerWidth - 40;
    stubRect(anchorEl.value, { left: right - 240, right, top: 60, bottom: 88, width: 240 });
    menuEl.value = document.createElement("div");
    stubRect(menuEl.value, { left: 0, right: 224, bottom: 118, width: 224, top: 0, height: 118 });

    open.value = true;
    await nextTick();

    expect(menuStyle.value.left).toBe(`${right - 240 - 224}px`);
    expect(menuStyle.value.top).toBe("60px");
  });

  it("opens an above-placement popup over its trigger when it fits there", async () => {
    const { open, anchorEl, menuEl, menuStyle } = setup({ placement: "above" });
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 40, right: 60, top: 500, bottom: 520, width: 20 });
    menuEl.value = document.createElement("div");
    stubRect(menuEl.value, { left: 0, right: 160, bottom: 200, width: 160, top: 0, height: 200 });

    open.value = true;
    await nextTick();

    expect(menuStyle.value.top).toBe(`${500 - 4 - 200}px`);
  });

  it("flips an above-placement popup below a trigger at the top of the viewport (SPEC-098)", async () => {
    const { open, anchorEl, menuEl, menuStyle } = setup({ placement: "above" });
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 40, right: 60, top: 10, bottom: 30, width: 20 });
    menuEl.value = document.createElement("div");
    stubRect(menuEl.value, { left: 0, right: 160, bottom: 200, width: 160, top: 0, height: 200 });

    open.value = true;
    await nextTick();

    expect(menuStyle.value.top).toBe("34px");
  });

  it("stays unpositioned while the menu is closed", async () => {
    const { open, anchorEl, menuStyle } = setup();
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 40, right: 60, bottom: 60, width: 20 });

    window.dispatchEvent(new Event("resize"));
    // No left/top before the menu opens — but `position: fixed` is there
    // from the start, so the popup's very first render is already out of
    // the toolbar's flow and cannot displace the anchor it measures.
    expect(menuStyle.value).toEqual({ position: "fixed" });

    open.value = true;
    await nextTick();
    expect(menuStyle.value.top).toBe("64px");
  });
});
