import { effectScope, nextTick, ref, type Ref } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { useZoomView, zoomStorageKey } from "./useZoomView";

// A minimal HTMLElement stand-in whose geometry the composable reads. jsdom
// reports 0 for offsetWidth/getBoundingClientRect, so the page and container
// sizes are defined explicitly.
function makePages(pageWidth: number, pageHeight: number, rectLeft = 100, rectTop = 40, pageCount = 1): HTMLElement {
  const pages = document.createElement("div");
  // Match the real DocPage roots: class `page`, test id `document-page` — the
  // id firstPageEl() selects on, and the count recomputeGrid caps columns to
  // (SPEC-063 AC-36/AC-54).
  for (let i = 0; i < Math.max(1, pageCount); i++) {
    const page = document.createElement("div");
    page.className = "page";
    page.setAttribute("data-testid", "document-page");
    Object.defineProperty(page, "offsetWidth", { value: pageWidth, configurable: true });
    Object.defineProperty(page, "offsetHeight", { value: pageHeight, configurable: true });
    pages.appendChild(page);
  }
  Object.defineProperty(pages, "offsetWidth", { value: pageWidth, configurable: true });
  Object.defineProperty(pages, "offsetHeight", { value: pageHeight, configurable: true });
  pages.getBoundingClientRect = () =>
    ({ left: rectLeft, top: rectTop, right: rectLeft, bottom: rectTop, width: 0, height: 0, x: rectLeft, y: rectTop, toJSON() {} }) as DOMRect;
  return pages;
}

function makeContainer(clientWidth: number, clientHeight: number): HTMLElement {
  const el = document.createElement("main");
  Object.defineProperty(el, "clientWidth", { value: clientWidth, configurable: true });
  Object.defineProperty(el, "clientHeight", { value: clientHeight, configurable: true });
  return el;
}

interface Harness {
  view: ReturnType<typeof useZoomView>;
  documentId: Ref<string>;
  pageRef: Ref<Element | null>;
  unmount: () => void;
}

function mountZoom(
  opts: {
    id?: string;
    pages?: HTMLElement | null;
    container?: HTMLElement | null;
    isGridEligible?: () => boolean;
  } = {},
): Harness {
  const documentId = ref(opts.id ?? "doc-1");
  const pageRef = ref<Element | null>(opts.pages ?? null);
  const unmountHooks: Array<() => void> = [];
  const scope = effectScope();
  let view!: ReturnType<typeof useZoomView>;
  scope.run(() => {
    view = useZoomView({
      documentId,
      pageRef,
      getScrollContainer: () => opts.container ?? null,
      onUnmount: (cb) => unmountHooks.push(cb),
      isGridEligible: opts.isGridEligible,
    });
  });
  return {
    view,
    documentId,
    pageRef,
    unmount: () => {
      unmountHooks.forEach((cb) => cb());
      scope.stop();
    },
  };
}

describe("useZoomView", () => {
  beforeEach(() => {
    localStorage.clear();
  });
  afterEach(() => {
    localStorage.clear();
    vi.restoreAllMocks();
  });

  it("opens at 100% with no stored value (AC-31)", () => {
    const h = mountZoom();
    expect(h.view.zoomPercent.value).toBe(100);
    expect(h.view.zoom.value).toBe(1);
    h.unmount();
  });

  it("restores the stored zoom on open (AC-30)", () => {
    localStorage.setItem(zoomStorageKey("doc-9"), "175");
    const h = mountZoom({ id: "doc-9" });
    expect(h.view.zoomPercent.value).toBe(175);
    expect(h.view.zoom.value).toBe(1.75);
    h.unmount();
  });

  it("persists under docu-edit:zoom:<id> and clamps (AC-29, AC-9)", () => {
    const h = mountZoom({ id: "doc-2" });
    h.view.setZoomPercent(175);
    expect(localStorage.getItem("docu-edit:zoom:doc-2")).toBe("175");
    h.view.setZoomPercent(1000);
    expect(h.view.zoomPercent.value).toBe(400);
    h.view.setZoomPercent(1);
    expect(h.view.zoomPercent.value).toBe(25);
    h.unmount();
  });

  it("steps the preset ladder and clamps at the ends (AC-6, AC-7)", () => {
    const h = mountZoom();
    h.view.zoomIn();
    expect(h.view.zoomPercent.value).toBe(125);
    h.view.zoomOut();
    expect(h.view.zoomPercent.value).toBe(100);
    h.view.setZoomPercent(110);
    h.view.zoomIn();
    expect(h.view.zoomPercent.value).toBe(125);
    h.view.setZoomPercent(110);
    h.view.zoomOut();
    expect(h.view.zoomPercent.value).toBe(100);
    for (let i = 0; i < 10; i += 1) h.view.zoomIn();
    expect(h.view.zoomPercent.value).toBe(200);
    for (let i = 0; i < 10; i += 1) h.view.zoomOut();
    expect(h.view.zoomPercent.value).toBe(50);
    h.unmount();
  });

  it("applies field input and reverts on invalid (AC-8, AC-10)", () => {
    const h = mountZoom();
    h.view.applyZoomInput("  250  ");
    expect(h.view.zoomPercent.value).toBe(250);
    h.view.applyZoomInput("abc");
    expect(h.view.zoomPercent.value).toBe(250);
    h.view.applyZoomInput("");
    expect(h.view.zoomPercent.value).toBe(250);
    h.view.applyZoomInput("88%");
    expect(h.view.zoomPercent.value).toBe(88);
    h.unmount();
  });

  it("resets to 100% (AC-11)", () => {
    const h = mountZoom();
    h.view.setZoomPercent(150);
    h.view.zoomReset();
    expect(h.view.zoomPercent.value).toBe(100);
    h.unmount();
  });

  it("ctrl+wheel steps zoom and prevents default; plain wheel is a no-op (AC-13, AC-14)", () => {
    const h = mountZoom();
    const prevent = vi.fn();
    h.view.onWheel({ ctrlKey: true, metaKey: false, deltaY: -1, preventDefault: prevent } as unknown as WheelEvent);
    expect(prevent).toHaveBeenCalledTimes(1);
    expect(h.view.zoomPercent.value).toBe(125);
    h.view.onWheel({ ctrlKey: true, metaKey: false, deltaY: 1, preventDefault: prevent } as unknown as WheelEvent);
    expect(h.view.zoomPercent.value).toBe(100);
    const noPrevent = vi.fn();
    h.view.onWheel({ ctrlKey: false, metaKey: false, deltaY: -1, preventDefault: noPrevent } as unknown as WheelEvent);
    expect(noPrevent).not.toHaveBeenCalled();
    expect(h.view.zoomPercent.value).toBe(100);
    h.unmount();
  });

  it("round-trips client<->content coordinates at 0.5/1/2 (AC-22, AC-23)", () => {
    const pages = makePages(800, 1000, 120, 60);
    const h = mountZoom({ pages });
    for (const percent of [50, 100, 200]) {
      h.view.setZoomPercent(percent);
      for (const [cx, cy] of [
        [120, 60],
        [520, 560],
      ]) {
        const content = h.view.clientToContentPx(cx, cy);
        const back = h.view.contentPxToClient(content.x, content.y);
        expect(back.x).toBeCloseTo(cx, 6);
        expect(back.y).toBeCloseTo(cy, 6);
      }
    }
    // a point 100px right of the scaled left edge is 200 model-px at 50%
    h.view.setZoomPercent(50);
    expect(h.view.clientToContentPx(220, 60).x).toBe(200);
    h.unmount();
  });

  it("fit-width solves pageWidth*zoom == availableWidth - 48 (AC-26)", async () => {
    const pages = makePages(800, 1100);
    const container = makeContainer(1400, 900);
    const h = mountZoom({ pages, container });
    await nextTick();
    h.view.setFitWidth();
    // (1400 - 48) / 800 = 1.69 → 169%
    expect(h.view.zoomPercent.value).toBeCloseTo(169, 5);
    h.unmount();
  });

  it("fit-page solves pageHeight*zoom == visibleHeight (AC-28)", async () => {
    const pages = makePages(800, 1000);
    const container = makeContainer(1400, 600);
    const h = mountZoom({ pages, container });
    await nextTick();
    h.view.setFitPage();
    // 600 / 1000 = 0.60 → 60%
    expect(h.view.zoomPercent.value).toBeCloseTo(60, 5);
    h.unmount();
  });

  it("recomputes an active fit mode against new geometry (AC-27)", async () => {
    const pages = makePages(800, 1100);
    const container = makeContainer(1400, 900);
    const h = mountZoom({ pages, container });
    await nextTick();
    h.view.setFitWidth();
    expect(h.view.zoomPercent.value).toBeCloseTo(169, 5);
    Object.defineProperty(container, "clientWidth", { value: 1000, configurable: true });
    h.view.recomputeFit();
    // (1000 - 48) / 800 = 1.19 → 119%
    expect(h.view.zoomPercent.value).toBeCloseTo(119, 5);
    h.unmount();
  });

  it("an explicit zoom clears the fit mode so resize no longer recomputes", async () => {
    const pages = makePages(800, 1100);
    const container = makeContainer(1400, 900);
    const h = mountZoom({ pages, container });
    await nextTick();
    h.view.setFitWidth();
    h.view.setZoomPercent(100);
    Object.defineProperty(container, "clientWidth", { value: 800, configurable: true });
    h.view.recomputeFit();
    expect(h.view.zoomPercent.value).toBe(100);
    h.unmount();
  });

  it("sizes the ruler wrapper to the scaled footprint, empty at 100% (AC-53, AC-54)", async () => {
    const pages = makePages(800, 1100);
    const h = mountZoom({ pages });
    await nextTick();
    // Measure the natural page size (ResizeObserver is absent in jsdom, so the
    // window-resize path stands in for it).
    window.dispatchEvent(new Event("resize"));
    await nextTick();
    // At 100% the ruler centres itself normally — no sizing wrapper.
    expect(h.view.rulerSizerStyle.value).toEqual({});
    h.view.setZoomPercent(50);
    await nextTick();
    // The wrapper takes the scaled footprint (800 * 0.5) so it shares `.pages`'
    // centring/overflow and the ruler tracks the page at every zoom.
    expect(h.view.rulerSizerStyle.value).toEqual({ width: "400px" });
    h.view.setZoomPercent(200);
    await nextTick();
    expect(h.view.rulerSizerStyle.value).toEqual({ width: "1600px" });
    h.unmount();
  });

  describe("side-by-side page grid (AC-36)", () => {
    it("tiles multiple columns when the zoom leaves room", async () => {
      const pages = makePages(800, 1100, 100, 40, 4);
      const container = makeContainer(1400, 900);
      const h = mountZoom({ pages, container });
      await nextTick();
      // 1400 / (800*0.5 + 24) = 3.30 → 3 columns at 50% (4 pages available).
      h.view.setZoomPercent(50);
      h.view.recomputeGrid();
      expect(h.view.pageColumns.value).toBe(3);
      expect(h.view.pageGrid.value).toBe(true);
      expect(h.view.gridStyle.value).toEqual({ "--page-grid-cols": "3" });
      h.unmount();
    });

    it("never reserves more columns than there are pages (AC-54)", async () => {
      // The zoom leaves room for 3 columns, but a single-page document must
      // stay in one centred column so the page keeps `.pages`' centre —
      // otherwise the ruler and fixed overlays drift off it.
      const pages = makePages(800, 1100, 100, 40, 1);
      const container = makeContainer(1400, 900);
      const h = mountZoom({ pages, container });
      await nextTick();
      h.view.setZoomPercent(50);
      h.view.recomputeGrid();
      expect(h.view.pageColumns.value).toBe(1);
      expect(h.view.pageGrid.value).toBe(false);
      expect(h.view.gridStyle.value).toEqual({});
      // Two pages cap the 3-column fit to two.
      h.pageRef.value = makePages(800, 1100, 100, 40, 2);
      h.view.recomputeGrid();
      expect(h.view.pageColumns.value).toBe(2);
      h.unmount();
    });

    it("stays a single column at 100% and exposes no grid style", async () => {
      const pages = makePages(800, 1100);
      const container = makeContainer(1400, 900);
      const h = mountZoom({ pages, container });
      await nextTick();
      h.view.recomputeGrid();
      expect(h.view.pageColumns.value).toBe(1);
      expect(h.view.pageGrid.value).toBe(false);
      expect(h.view.gridStyle.value).toEqual({});
      h.unmount();
    });

    it("never tiles when the layout is ineligible (Web layout)", async () => {
      const pages = makePages(800, 1100);
      const container = makeContainer(1400, 900);
      const h = mountZoom({ pages, container, isGridEligible: () => false });
      await nextTick();
      h.view.setZoomPercent(50);
      h.view.recomputeGrid();
      expect(h.view.pageColumns.value).toBe(1);
      expect(h.view.pageGrid.value).toBe(false);
      h.unmount();
    });
  });
});
