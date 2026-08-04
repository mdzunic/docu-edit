import { describe, it, expect } from "vitest";
import { nextTick, ref, type Ref } from "vue";
import { useFloatingMenu } from "./useFloatingMenu";

interface FakeRect {
  left: number;
  right: number;
  bottom: number;
  width: number;
}

// jsdom's getBoundingClientRect is all zeros, so the tests stub it the way
// a browser would measure the real anchor/menu.
function stubRect(el: Element, rect: FakeRect) {
  Object.defineProperty(el, "getBoundingClientRect", {
    configurable: true,
    value: () => rect,
  });
}

function setup() {
  const open = ref(false);
  const anchorEl = ref<HTMLElement | null>(null) as Ref<HTMLElement | null>;
  const menuEl = ref<HTMLElement | null>(null);
  const { menuStyle } = useFloatingMenu({
    open: () => open.value,
    anchor: () => anchorEl.value,
    menu: () => menuEl.value,
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

  it("stays unpositioned while the menu is closed", async () => {
    const { open, anchorEl, menuStyle } = setup();
    anchorEl.value = document.createElement("button");
    stubRect(anchorEl.value, { left: 40, right: 60, bottom: 60, width: 20 });

    window.dispatchEvent(new Event("resize"));
    expect(menuStyle.value).toEqual({});

    open.value = true;
    await nextTick();
    expect(menuStyle.value).not.toEqual({});
  });
});
