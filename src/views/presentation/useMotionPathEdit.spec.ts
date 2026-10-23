// SPEC-105 motion-path editor: the overlay geometry (slide fractions from the
// frame's centre, in surface px), the endpoint drag — pointer capture,
// clientToCm, a live draft refreshed while dragging, one record + commit on
// pointer-up — the arrow nudge, and a custom path's refusal.
import { describe, expect, it, vi } from "vitest";
import { computed, ref } from "vue";
import type { MotionPath, PresentationModel, Slide } from "../../lib/presentation/model";
import { cmToPx } from "../../lib/core/units";
import { useMotionPathEdit } from "./useMotionPathEdit";

/** A 20cm × 10cm deck with a 4cm × 2cm shape centred at (5cm, 5cm). */
function harness(path: MotionPath) {
  const slide: Slide = {
    id: "s",
    frames: [{ id: "box", kind: "shape", geometry: "rect", xCm: 3, yCm: 4, widthCm: 4, heightCm: 2 }],
    animations: [{ frameId: "box", effect: "motionPath", trigger: "click", path }],
  };
  const model = ref<PresentationModel | null>({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format: "pptx",
    slides: [slide],
    slideWidthCm: 20,
    slideHeightCm: 10,
  });
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const record = vi.fn();
  const commit = vi.fn();
  const edit = useMotionPathEdit({ model, currentSlide, selectedIndex: ref(0), record, commit });
  return { edit, model, record, commit };
}

/** An overlay element whose on-screen rect is the slide at 10px per cm. */
function overlayEl(): Element {
  return { getBoundingClientRect: () => ({ left: 0, top: 0, width: 200, height: 100 }) } as unknown as Element;
}
function pointer(clientX: number, clientY: number, target: unknown = null): PointerEvent {
  return { clientX, clientY, pointerId: 7, target, preventDefault: () => {} } as unknown as PointerEvent;
}

const px = (cm: number): number => cmToPx(cm);
function key(k: string): KeyboardEvent {
  return { key: k, preventDefault: () => {} } as unknown as KeyboardEvent;
}

describe("useMotionPathEdit", () => {
  it("draws nothing until armed, then the resolved path, endpoint and ghost in surface px", () => {
    const { edit } = harness({ preset: "line" });
    expect(edit.overlay.value).toBeNull();
    edit.toggle();
    const o = edit.overlay.value!;
    expect([o.widthPx, o.heightPx]).toEqual([px(20), px(10)]);
    // 25% of the 20cm width = 5cm to the right of the frame centre (5cm, 5cm).
    expect(o.handle.x).toBeCloseTo(px(10));
    expect(o.handle.y).toBeCloseTo(px(5));
    expect(o.ghost.x).toBeCloseTo(px(3 + 5));
    expect(o.d.startsWith(`M ${px(5).toFixed(2)} ${px(5).toFixed(2)} L`)).toBe(true);
    expect(o.readOnly).toBe(false);
  });

  it("drags the endpoint with pointer capture: a live draft, then ONE record + commit on pointer-up", () => {
    const { edit, model, record, commit } = harness({ preset: "line" });
    edit.toggle();
    const capture = { setPointerCapture: vi.fn(), releasePointerCapture: vi.fn() };
    const target = Object.assign(Object.create(Element.prototype), capture);
    expect(edit.onEndpointPointerDown(pointer(100, 50, target), overlayEl())).toBe(true);
    expect(capture.setPointerCapture).toHaveBeenCalledWith(7);
    // Down 3cm from the centre (5cm,5cm) = 30% of the 10cm height.
    edit.onEndpointPointerMove(pointer(50, 80));
    expect(edit.draft.value).toEqual({ preset: "line", direction: "down", distanceFraction: 0.3 });
    expect(edit.overlay.value!.handle.y).toBeCloseTo(cmToPx(8));
    expect(record).not.toHaveBeenCalled();
    edit.onEndpointPointerMove(pointer(50, 90));
    edit.onEndpointPointerUp(pointer(50, 90, target));
    expect(capture.releasePointerCapture).toHaveBeenCalledWith(7);
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Edit motion path");
    expect(commit).toHaveBeenCalledTimes(1);
    expect(model.value!.slides[0].animations![0].path).toEqual({ preset: "line", direction: "down", distanceFraction: 0.4 });
    expect(model.value!.slides[0].dirty).toBe(true);
    expect(edit.draft.value).toBeNull();
  });

  it("a drag back to where it started records nothing", () => {
    const { edit, record } = harness({ preset: "line" });
    edit.toggle();
    edit.onEndpointPointerDown(pointer(100, 50), overlayEl());
    edit.onEndpointPointerUp(pointer(100, 50));
    expect(record).not.toHaveBeenCalled();
  });

  it("arrow keys nudge the endpoint one percent — one entry per press", () => {
    const { edit, model, record } = harness({ preset: "line" });
    edit.toggle();
    edit.onEndpointKeydown(key("ArrowRight"));
    edit.onEndpointKeydown(key("ArrowRight"));
    expect(record).toHaveBeenCalledTimes(2);
    expect(model.value!.slides[0].animations![0].path).toEqual({ preset: "line", distanceFraction: 0.27 });
  });

  it("an imported custom path renders read-only: dragging and nudging are refused", () => {
    const { edit, model, record } = harness({ preset: "custom", d: "M 0 0 L 0.1 0.1" });
    edit.toggle();
    expect(edit.overlay.value!.readOnly).toBe(true);
    expect(edit.onEndpointPointerDown(pointer(100, 50), overlayEl())).toBe(false);
    edit.onEndpointPointerMove(pointer(10, 10));
    edit.onEndpointPointerUp(pointer(10, 10));
    edit.onEndpointKeydown({ key: "ArrowUp", preventDefault: () => {} } as unknown as KeyboardEvent);
    expect(record).not.toHaveBeenCalled();
    expect(model.value!.slides[0].animations![0].path).toEqual({ preset: "custom", d: "M 0 0 L 0.1 0.1" });
  });
});
