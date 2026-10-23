// SPEC-105 motion-path editor: the overlay geometry (slide fractions from the
// frame's centre, in surface px), the endpoint drag — pointer capture, focus,
// clientToCm, a live draft refreshed while dragging, one record + commit on
// pointer-up — the arrow nudge (kept from the editor's window keys), the
// label following committed edits, and a custom path's refusal.
import { describe, expect, it, vi } from "vitest";
import { computed, ref, shallowRef, triggerRef } from "vue";
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
function key(k: string): KeyboardEvent & { stopPropagation: ReturnType<typeof vi.fn> } {
  return { key: k, preventDefault: () => {}, stopPropagation: vi.fn() } as unknown as KeyboardEvent & {
    stopPropagation: ReturnType<typeof vi.fn>;
  };
}

/** The app's shape: the model behind a shallowRef, rows edited in place, and a
 * commit that re-identifies the current slide and its frames (useSlideRender)
 * — never the animation rows themselves. */
function appHarness(path: MotionPath) {
  const slide: Slide = {
    id: "s",
    frames: [{ id: "box", kind: "shape", geometry: "rect", xCm: 3, yCm: 4, widthCm: 4, heightCm: 2 }],
    animations: [{ frameId: "box", effect: "motionPath", trigger: "click", path }],
  };
  const model = shallowRef<PresentationModel | null>({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format: "pptx",
    slides: [slide],
    slideWidthCm: 20,
    slideHeightCm: 10,
  });
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const commit = vi.fn(() => {
    const m = model.value!;
    m.slides[0] = { ...m.slides[0], frames: m.slides[0].frames.map((f) => ({ ...f })) };
    m.slides = m.slides.slice();
    triggerRef(model);
  });
  const edit = useMotionPathEdit({ model, currentSlide, selectedIndex: ref(0), record: vi.fn(), commit });
  return { edit, model };
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

  it("the endpoint follows the pointer while the path and ghost show the snapped draft", () => {
    const { edit } = harness({ preset: "line" });
    edit.toggle();
    edit.onEndpointPointerDown(pointer(100, 50), overlayEl());
    // (7cm, 8cm): 2cm right and 3cm down of the centre — snaps to down, 30%.
    edit.onEndpointPointerMove(pointer(70, 80));
    expect(edit.draft.value).toEqual({ preset: "line", direction: "down", distanceFraction: 0.3 });
    expect(edit.overlay.value!.handle.x).toBeCloseTo(px(7));
    expect(edit.overlay.value!.handle.y).toBeCloseTo(px(8));
    expect(edit.overlay.value!.ghost.x).toBeCloseTo(px(3));
    expect(edit.overlay.value!.ghost.y).toBeCloseTo(px(4 + 3));
    expect(edit.endpointLabel.value).toBe("Path end: down, 30%");
    // Released: the endpoint sits on the snapped path's end.
    edit.onEndpointPointerUp(pointer(70, 80));
    expect(edit.overlay.value!.handle.x).toBeCloseTo(px(5));
    expect(edit.overlay.value!.handle.y).toBeCloseTo(px(8));
  });

  it("a pressed endpoint takes focus (its default is prevented)", () => {
    const { edit } = harness({ preset: "line" });
    edit.toggle();
    const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("tabindex", "0");
    document.body.append(circle);
    const preventDefault = vi.fn();
    const down = { clientX: 100, clientY: 50, pointerId: 7, target: circle, preventDefault } as unknown as PointerEvent;
    expect(edit.onEndpointPointerDown(down, overlayEl())).toBe(true);
    expect(preventDefault).toHaveBeenCalled();
    expect(document.activeElement).toBe(circle);
    circle.remove();
  });

  it("the label and overlay follow committed edits though the row is edited in place", () => {
    const { edit, model } = appHarness({ preset: "line" });
    edit.toggle();
    expect(edit.endpointLabel.value).toBe("Path end: right, 25%");
    edit.onEndpointPointerDown(pointer(100, 50), overlayEl());
    edit.onEndpointPointerMove(pointer(50, 90));
    edit.onEndpointPointerUp(pointer(50, 90));
    expect(model.value!.slides[0].animations![0].path).toEqual({ preset: "line", direction: "down", distanceFraction: 0.4 });
    expect(edit.endpointLabel.value).toBe("Path end: down, 40%");
    expect(edit.overlay.value!.handle.y).toBeCloseTo(px(9));
    edit.onEndpointKeydown(key("ArrowDown"));
    expect(edit.endpointLabel.value).toBe("Path end: down, 41%");
    // Another effect: no longer a path row, so the overlay goes.
    model.value!.slides[0].animations![0] = { frameId: "box", effect: "spin", trigger: "click" };
    model.value!.slides[0] = { ...model.value!.slides[0] };
    triggerRef(model);
    expect(edit.isPathRow.value).toBe(false);
    expect(edit.overlay.value).toBeNull();
  });

  it("a drag back to where it started records nothing", () => {
    const { edit, record } = harness({ preset: "line" });
    edit.toggle();
    edit.onEndpointPointerDown(pointer(100, 50), overlayEl());
    edit.onEndpointPointerUp(pointer(100, 50));
    expect(record).not.toHaveBeenCalled();
  });

  it("arrow keys nudge the endpoint one percent — one entry per press, never the frame's", () => {
    const { edit, model, record } = harness({ preset: "line" });
    edit.toggle();
    const right = key("ArrowRight");
    edit.onEndpointKeydown(right);
    edit.onEndpointKeydown(key("ArrowRight"));
    expect(record).toHaveBeenCalledTimes(2);
    expect(model.value!.slides[0].animations![0].path).toEqual({ preset: "line", distanceFraction: 0.27 });
    // Stopped short of the editor's window keys (which nudge the frame).
    expect(right.stopPropagation).toHaveBeenCalled();
    const tab = key("Tab");
    edit.onEndpointKeydown(tab);
    expect(tab.stopPropagation).not.toHaveBeenCalled();
  });

  it("an imported custom path renders read-only: dragging and nudging are refused", () => {
    const { edit, model, record } = harness({ preset: "custom", d: "M 0 0 L 0.1 0.1" });
    edit.toggle();
    expect(edit.overlay.value!.readOnly).toBe(true);
    expect(edit.onEndpointPointerDown(pointer(100, 50), overlayEl())).toBe(false);
    edit.onEndpointPointerMove(pointer(10, 10));
    edit.onEndpointPointerUp(pointer(10, 10));
    const up = key("ArrowUp");
    edit.onEndpointKeydown(up);
    expect(up.stopPropagation).toHaveBeenCalled();
    expect(record).not.toHaveBeenCalled();
    expect(model.value!.slides[0].animations![0].path).toEqual({ preset: "custom", d: "M 0 0 L 0.1 0.1" });
  });
});
