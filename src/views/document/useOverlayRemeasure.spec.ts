import { effectScope, nextTick, ref, type Ref } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useOverlayRemeasure } from "./useOverlayRemeasure";

// SPEC-035 regression (displaced squiggles with a side pane open): a resize
// of the pages container — a pane toggle narrows the workspace and the page
// re-centers — must re-fire the scroll hooks and drawing-handle measurers
// exactly like a zoom change does. jsdom has no ResizeObserver, so these
// tests install a fake that captures observe/unobserve/disconnect and lets
// the test fire the callback for "the container resized".

type ResizeCallback = () => void;

class FakeResizeObserver {
  static instances: FakeResizeObserver[] = [];
  callback: ResizeCallback;
  observed: Element[] = [];
  unobserved: Element[] = [];
  disconnected = false;
  constructor(callback: ResizeCallback) {
    this.callback = callback;
    FakeResizeObserver.instances.push(this);
  }
  observe(el: Element): void {
    this.observed.push(el);
  }
  unobserve(el: Element): void {
    this.unobserved.push(el);
  }
  disconnect(): void {
    this.disconnected = true;
  }
}

function mount(pageRef: Ref<Element | null>, zoom = ref(1)) {
  const scrollHook = vi.fn();
  const scrollHooks = new Set<() => void>([scrollHook]);
  const remeasureShapeHandles = vi.fn();
  const remeasureImageHandles = vi.fn();
  const remeasureChartHandles = vi.fn();
  const remeasureConnectors = vi.fn();
  const clearTableHover = vi.fn();
  const unmountCallbacks: Array<() => void> = [];
  const scope = effectScope();
  scope.run(() => {
    useOverlayRemeasure({
      zoom,
      pageRef,
      onUnmount: (cb) => unmountCallbacks.push(cb),
      scrollHooks,
      remeasureShapeHandles,
      remeasureImageHandles,
      remeasureChartHandles,
      remeasureConnectors,
      clearTableHover,
    });
  });
  return {
    zoom,
    scrollHook,
    remeasureShapeHandles,
    remeasureImageHandles,
    remeasureChartHandles,
    clearTableHover,
    unmount: () => {
      unmountCallbacks.forEach((cb) => cb());
      scope.stop();
    },
  };
}

describe("useOverlayRemeasure — pages-container resize (SPEC-035 pane displacement)", () => {
  afterEach(() => {
    FakeResizeObserver.instances = [];
    vi.unstubAllGlobals();
  });

  it("observes the pages container and re-fires scroll hooks and handle measurers on resize", async () => {
    vi.stubGlobal("ResizeObserver", FakeResizeObserver);
    const pageEl = document.createElement("div");
    const view = mount(ref<Element | null>(pageEl));

    const observer = FakeResizeObserver.instances[0];
    expect(observer).toBeDefined();
    expect(observer.observed).toContain(pageEl);

    observer.callback(); // the pane opened: the container's layout box changed
    await nextTick();
    expect(view.scrollHook).toHaveBeenCalledTimes(1);
    expect(view.remeasureShapeHandles).toHaveBeenCalledTimes(1);
    expect(view.remeasureImageHandles).toHaveBeenCalledTimes(1);
    expect(view.remeasureChartHandles).toHaveBeenCalledTimes(1);
    // The hover-driven table handles can't be re-measured — the layout moved
    // under the pointer — so a relayout drops them (SPEC-022 stale latch).
    expect(view.clearTableHover).toHaveBeenCalledTimes(1);
    view.unmount();
  });

  it("coalesces a same-tick burst of resize callbacks into one re-measure", async () => {
    vi.stubGlobal("ResizeObserver", FakeResizeObserver);
    const view = mount(ref<Element | null>(document.createElement("div")));
    const observer = FakeResizeObserver.instances[0];

    observer.callback();
    observer.callback();
    observer.callback();
    await nextTick();
    expect(view.scrollHook).toHaveBeenCalledTimes(1);

    // A later resize fires again — the coalescing window is one tick, not
    // once-forever.
    observer.callback();
    await nextTick();
    expect(view.scrollHook).toHaveBeenCalledTimes(2);
    view.unmount();
  });

  it("re-observes when the pages element is replaced and disconnects on unmount", async () => {
    vi.stubGlobal("ResizeObserver", FakeResizeObserver);
    const first = document.createElement("div");
    const second = document.createElement("div");
    const pageRef = ref<Element | null>(first);
    const view = mount(pageRef);
    const observer = FakeResizeObserver.instances[0];

    pageRef.value = second;
    await nextTick();
    expect(observer.unobserved).toContain(first);
    expect(observer.observed).toContain(second);

    view.unmount();
    expect(observer.disconnected).toBe(true);
  });

  it("still works (zoom watch only) where ResizeObserver does not exist", async () => {
    // No stub installed: jsdom genuinely lacks ResizeObserver.
    expect(typeof ResizeObserver).toBe("undefined");
    const view = mount(ref<Element | null>(document.createElement("div")));
    view.zoom.value = 1.5;
    await nextTick(); // zoom watch fires…
    await nextTick(); // …and defers its re-measure one tick further
    expect(view.scrollHook).toHaveBeenCalledTimes(1);
    expect(view.clearTableHover).toHaveBeenCalledTimes(1);
    view.unmount();
  });
});
