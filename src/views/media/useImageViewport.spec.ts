// The image viewer's fit/zoom/pan state (SPEC-107): fit tracks the window,
// the buttons walk SPEC-063's preset ladder, Cmd/Ctrl +/-/0 is the keyboard
// vocabulary, and double-click toggles fit ⇄ 100%.
import { describe, expect, it } from "vitest";
import { useImageViewport, FIT_MARGIN_PX, FIT_MIN_PERCENT } from "./useImageViewport";

function sized(): ReturnType<typeof useImageViewport> {
  const viewport = useImageViewport();
  viewport.setContainerSize(800 + FIT_MARGIN_PX * 2, 600 + FIT_MARGIN_PX * 2);
  viewport.setImageSize(1600, 600);
  return viewport;
}

describe("useImageViewport", () => {
  it("opens fit-to-window: scale derived from the container and image", () => {
    const viewport = sized();
    expect(viewport.mode.value).toBe("fit");
    // Width is the binding axis: 800 / 1600 = 50%.
    expect(viewport.percent.value).toBe(50);
    expect(viewport.scale.value).toBeCloseTo(0.5);
  });

  it("fit recomputes live when the window resizes", () => {
    const viewport = sized();
    viewport.setContainerSize(400 + FIT_MARGIN_PX * 2, 600 + FIT_MARGIN_PX * 2);
    expect(viewport.percent.value).toBe(25);
  });

  it("fits a large photo below the 25% manual floor, so it really fits", () => {
    const viewport = useImageViewport();
    viewport.setContainerSize(1000 + FIT_MARGIN_PX * 2, 640 + FIT_MARGIN_PX * 2);
    viewport.setImageSize(4032, 3024);
    // min(1000 / 4032, 640 / 3024) = 21.16% — floored, never rounded up.
    expect(viewport.percent.value).toBe(21);
    expect(4032 * viewport.scale.value).toBeLessThanOrEqual(1000);
    expect(3024 * viewport.scale.value).toBeLessThanOrEqual(640);
    expect(FIT_MIN_PERCENT).toBeLessThan(25);
  });

  it("never steps the wrong way from a fit outside the preset ladder", () => {
    const large = useImageViewport();
    large.setContainerSize(500 + FIT_MARGIN_PX * 2, 500 + FIT_MARGIN_PX * 2);
    large.setImageSize(5000, 5000);
    expect(large.percent.value).toBe(10);
    large.zoomOut(); // no preset below 10%: stays put, still fitting
    expect(large.mode.value).toBe("fit");
    expect(large.percent.value).toBe(10);
    large.zoomIn();
    expect(large.percent.value).toBe(50);

    const icon = useImageViewport();
    icon.setContainerSize(800 + FIT_MARGIN_PX * 2, 800 + FIT_MARGIN_PX * 2);
    icon.setImageSize(16, 16);
    expect(icon.percent.value).toBe(400); // the input band's ceiling
    icon.zoomIn(); // nothing above: stays put
    expect(icon.mode.value).toBe("fit");
    icon.zoomOut();
    expect(icon.percent.value).toBe(200);
  });

  it("defaults to 100% before sizes are known", () => {
    const viewport = useImageViewport();
    expect(viewport.percent.value).toBe(100);
  });

  it("zoom in/out walk the preset ladder from the effective percentage", () => {
    const viewport = sized();
    viewport.zoomIn(); // from 50 → next preset up (75)
    expect(viewport.mode.value).toBe("manual");
    expect(viewport.percent.value).toBe(75);
    viewport.zoomIn();
    expect(viewport.percent.value).toBe(90);
    viewport.zoomOut();
    expect(viewport.percent.value).toBe(75);
  });

  it("100% and Fit are direct targets; double-click toggles between them", () => {
    const viewport = sized();
    viewport.toggleFitActual();
    expect(viewport.mode.value).toBe("manual");
    expect(viewport.percent.value).toBe(100);
    viewport.toggleFitActual();
    expect(viewport.mode.value).toBe("fit");
    expect(viewport.percent.value).toBe(50);
  });

  it("handles the Cmd/Ctrl +, - and 0 keys and ignores everything else", () => {
    const viewport = sized();
    expect(viewport.handleZoomKey(new KeyboardEvent("keydown", { key: "=", ctrlKey: true }))).toBe(true);
    expect(viewport.percent.value).toBe(75);
    expect(viewport.handleZoomKey(new KeyboardEvent("keydown", { key: "-", metaKey: true }))).toBe(true);
    expect(viewport.percent.value).toBe(50);
    expect(viewport.handleZoomKey(new KeyboardEvent("keydown", { key: "0", ctrlKey: true }))).toBe(true);
    expect(viewport.percent.value).toBe(100);
    expect(viewport.handleZoomKey(new KeyboardEvent("keydown", { key: "0" }))).toBe(false);
    expect(viewport.handleZoomKey(new KeyboardEvent("keydown", { key: "a", ctrlKey: true }))).toBe(false);
  });

  it("pans by deltas and resets the pan on Fit", () => {
    const viewport = sized();
    viewport.zoomActual();
    viewport.panBy(10, -5);
    viewport.panBy(2, 3);
    expect(viewport.pan.value).toEqual({ x: 12, y: -2 });
    viewport.zoomFit();
    expect(viewport.pan.value).toEqual({ x: 0, y: 0 });
  });
});
