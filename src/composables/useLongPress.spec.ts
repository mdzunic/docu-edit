import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { defineComponent, h } from "vue";
import { mount } from "@vue/test-utils";
import { LONG_PRESS_MS, useLongPress } from "./useLongPress";

// A tiny host component so useLongPress's onMounted/onBeforeUnmount run.
function mountWithLongPress(onLongPress: (info: unknown) => void, opts: Record<string, unknown> = {}) {
  const Host = defineComponent({
    setup() {
      let el: HTMLElement | null = null;
      const api = useLongPress({ target: () => el, onLongPress, ...opts });
      return () => h("div", { ref: (r) => { el = r as HTMLElement; }, "data-testid": "surface" }, [api ? "" : ""]);
    },
  });
  return mount(Host, { attachTo: document.body });
}

function touch(type: string, x: number, y: number) {
  const event = new Event(type, { bubbles: true }) as PointerEvent & { pointerType: string };
  Object.defineProperty(event, "pointerType", { value: "touch" });
  Object.defineProperty(event, "clientX", { value: x });
  Object.defineProperty(event, "clientY", { value: y });
  return event;
}

describe("useLongPress (SPEC-068 AC-22/23)", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it("fires after the threshold for a still touch", () => {
    const onLongPress = vi.fn();
    const wrapper = mountWithLongPress(onLongPress);
    const surface = wrapper.get('[data-testid="surface"]').element;

    surface.dispatchEvent(touch("pointerdown", 100, 200));
    expect(onLongPress).not.toHaveBeenCalled();
    vi.advanceTimersByTime(LONG_PRESS_MS);
    expect(onLongPress).toHaveBeenCalledTimes(1);
    expect(onLongPress).toHaveBeenCalledWith(expect.objectContaining({ clientX: 100, clientY: 200 }));
    wrapper.unmount();
  });

  it("does NOT fire when the finger moves past the tolerance — a selection drag (AC-23)", () => {
    const onLongPress = vi.fn();
    const wrapper = mountWithLongPress(onLongPress);
    const surface = wrapper.get('[data-testid="surface"]').element;

    surface.dispatchEvent(touch("pointerdown", 100, 200));
    surface.dispatchEvent(touch("pointermove", 140, 200)); // moved 40px
    vi.advanceTimersByTime(LONG_PRESS_MS);
    expect(onLongPress).not.toHaveBeenCalled();
    wrapper.unmount();
  });

  it("does not fire when the touch lifts before the threshold", () => {
    const onLongPress = vi.fn();
    const wrapper = mountWithLongPress(onLongPress);
    const surface = wrapper.get('[data-testid="surface"]').element;

    surface.dispatchEvent(touch("pointerdown", 100, 200));
    vi.advanceTimersByTime(LONG_PRESS_MS / 2);
    surface.dispatchEvent(touch("pointerup", 100, 200));
    vi.advanceTimersByTime(LONG_PRESS_MS);
    expect(onLongPress).not.toHaveBeenCalled();
    wrapper.unmount();
  });

  it("swallows the synthetic click that follows a fired long-press so it cannot dismiss the menu", () => {
    const onLongPress = vi.fn();
    const wrapper = mountWithLongPress(onLongPress);
    const surface = wrapper.get('[data-testid="surface"]').element;
    const outsideClick = vi.fn();
    window.addEventListener("click", outsideClick);

    surface.dispatchEvent(touch("pointerdown", 100, 200));
    vi.advanceTimersByTime(LONG_PRESS_MS);
    surface.dispatchEvent(touch("pointerup", 100, 200));
    // The browser's synthetic post-touch click.
    surface.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    expect(outsideClick).not.toHaveBeenCalled();

    // Only the one click is eaten; a subsequent tap clicks through normally.
    surface.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    expect(outsideClick).toHaveBeenCalledTimes(1);

    window.removeEventListener("click", outsideClick);
    wrapper.unmount();
  });

  it("does not swallow a click when no long-press fired (a plain tap)", () => {
    const onLongPress = vi.fn();
    const wrapper = mountWithLongPress(onLongPress);
    const surface = wrapper.get('[data-testid="surface"]').element;
    const outsideClick = vi.fn();
    window.addEventListener("click", outsideClick);

    surface.dispatchEvent(touch("pointerdown", 100, 200));
    vi.advanceTimersByTime(LONG_PRESS_MS / 2);
    surface.dispatchEvent(touch("pointerup", 100, 200));
    surface.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    expect(outsideClick).toHaveBeenCalledTimes(1);

    window.removeEventListener("click", outsideClick);
    wrapper.unmount();
  });

  it("ignores a mouse pointer (right click already works)", () => {
    const onLongPress = vi.fn();
    const wrapper = mountWithLongPress(onLongPress);
    const surface = wrapper.get('[data-testid="surface"]').element;

    const mouseDown = new Event("pointerdown", { bubbles: true }) as PointerEvent & { pointerType: string };
    Object.defineProperty(mouseDown, "pointerType", { value: "mouse" });
    Object.defineProperty(mouseDown, "clientX", { value: 10 });
    Object.defineProperty(mouseDown, "clientY", { value: 10 });
    surface.dispatchEvent(mouseDown);
    vi.advanceTimersByTime(LONG_PRESS_MS);
    expect(onLongPress).not.toHaveBeenCalled();
    wrapper.unmount();
  });
});
