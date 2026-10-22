// SPEC-096: the arrange composable — align/distribute over the selection,
// group creation (animations dropped with a notice), the confirm-to-release
// flow for a rotated verbatim group, the marquee gesture, and the Position &
// Size dialog command — each as ONE undo entry (record before the first
// mutation, commit once).
import { computed, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { insertShapeFrame } from "../../lib/presentation/frames";
import type { GroupFrame, PresentationModel, ShapeFrame } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { useFrameArrange } from "./useFrameArrange";

function setup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  model.value!.slides[0].frames = [];
  const currentIndex = shallowRef(0);
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const selection = useFrameSelection({ currentSlide });
  const record = vi.fn();
  const commit = vi.fn();
  // A surface whose px box equals the slide's cm box: 1 client unit = 1cm.
  const surface = {
    getBoundingClientRect: () => ({
      left: 0,
      top: 0,
      width: model.value!.slideWidthCm,
      height: model.value!.slideHeightCm,
    }),
  } as unknown as HTMLElement;
  const arrange = useFrameArrange({
    model,
    currentSlide,
    currentIndex,
    selection,
    surface: () => surface,
    record,
    refresh: vi.fn(),
    commit,
  });
  const frameAt = (id: string) => model.value!.slides[0].frames.find((f) => f.id === id)!;
  return { model, selection, arrange, record, commit, frameAt };
}

function shape(model: PresentationModel, x: number, y: number, w = 2, h = 2): string {
  return insertShapeFrame(model, 0, "rect", { xCm: x, yCm: y, widthCm: w, heightCm: h })!;
}

function ptr(type: string, clientX: number, clientY: number, init: MouseEventInit = {}): PointerEvent {
  return new MouseEvent(type, { clientX, clientY, ...init }) as unknown as PointerEvent;
}

describe("useFrameArrange align/distribute (SPEC-096)", () => {
  it("aligns every selected frame's left edge to the selection box in one entry", () => {
    const { model, selection, arrange, record, commit, frameAt } = setup();
    const a = shape(model.value!, 2, 1);
    const b = shape(model.value!, 6, 4);
    const c = shape(model.value!, 10, 8);
    selection.selectMany([a, b, c]);
    arrange.align("left");
    expect([frameAt(a).xCm, frameAt(b).xCm, frameAt(c).xCm]).toEqual([2, 2, 2]);
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Align left");
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("aligns a single frame against the slide (the only reference at 1)", () => {
    const { model, selection, arrange, frameAt } = setup();
    const a = shape(model.value!, 5, 5);
    selection.select(a);
    arrange.align("centerX");
    expect(frameAt(a).xCm + 1).toBeCloseTo(model.value!.slideWidthCm / 2, 3);
  });

  it("aligns against the first selected (anchor) when the reference is 'first'", () => {
    const { model, selection, arrange, frameAt } = setup();
    const a = shape(model.value!, 4, 1);
    const b = shape(model.value!, 9, 4);
    selection.selectMany([a, b]);
    arrange.alignReference.value = "first";
    arrange.align("left");
    expect(frameAt(b).xCm).toBe(4);
    expect(frameAt(a).xCm).toBe(4);
  });

  it("distribute needs ≥3 and equalises gaps holding the outermost two", () => {
    const { model, selection, arrange, frameAt, record } = setup();
    const a = shape(model.value!, 1, 1);
    const b = shape(model.value!, 4, 1);
    selection.selectMany([a, b]);
    expect(arrange.canDistribute.value).toBe(false);
    expect(arrange.distributeReason.value).not.toBe("");
    arrange.distribute("x");
    expect(record).not.toHaveBeenCalled();

    const c = shape(model.value!, 12, 1);
    selection.selectMany([a, b, c]);
    expect(arrange.canDistribute.value).toBe(true);
    arrange.distribute("x");
    const gap1 = frameAt(b).xCm - (frameAt(a).xCm + 2);
    const gap2 = frameAt(c).xCm - (frameAt(b).xCm + 2);
    expect(gap1).toBeCloseTo(gap2, 3);
    expect(frameAt(a).xCm).toBe(1); // outermost held
    expect(frameAt(c).xCm).toBe(12);
    expect(record).toHaveBeenCalledWith("Distribute horizontally");
  });
});

describe("useFrameArrange group / release (SPEC-096)", () => {
  it("groups the selection, selects the group, and notices dropped animations", () => {
    const { model, selection, arrange, record } = setup();
    const a = shape(model.value!, 2, 2);
    const b = shape(model.value!, 6, 2);
    model.value!.slides[0].animations = [{ frameId: a, effect: "appear", trigger: "click" }];
    selection.selectMany([a, b]);
    arrange.group();
    const frames = model.value!.slides[0].frames;
    expect(frames).toHaveLength(1);
    expect(frames[0].kind).toBe("group");
    expect(selection.selectedFrameIds.value).toEqual([frames[0].id]);
    expect(model.value!.slides[0].animations).toEqual([]);
    expect(arrange.notice.value).toContain("animations");
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Group");
  });

  it("releases an unrotated verbatim group silently, all groups in one entry", () => {
    const { model, selection, arrange, record, commit } = setup();
    const group: GroupFrame = {
      id: "g1",
      kind: "group",
      xCm: 1,
      yCm: 1,
      widthCm: 4,
      heightCm: 4,
      xml: "<p:grpSp/>",
      children: [{ id: "c1", kind: "shape", geometry: "rect", xCm: 1, yCm: 1, widthCm: 2, heightCm: 2 }],
    };
    model.value!.slides[0].frames.push(group);
    selection.select("g1");
    arrange.ungroup();
    expect(arrange.pendingRelease.value).toBeNull();
    expect(model.value!.slides[0].frames.map((f) => f.id)).toEqual(["c1"]);
    expect(selection.selectedFrameIds.value).toEqual(["c1"]);
    expect(record).toHaveBeenCalledWith("Ungroup");
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("parks a ROTATED verbatim group in pendingRelease; cancel changes nothing; confirm releases", () => {
    const { model, selection, arrange, record } = setup();
    const group: GroupFrame = {
      id: "g1",
      kind: "group",
      xCm: 1,
      yCm: 1,
      widthCm: 4,
      heightCm: 4,
      rotationDeg: 30,
      xml: "<p:grpSp/>",
      children: [{ id: "c1", kind: "shape", geometry: "rect", xCm: 1, yCm: 1, widthCm: 2, heightCm: 2 }],
    };
    model.value!.slides[0].frames.push(group);
    selection.select("g1");
    arrange.ungroup();
    expect(arrange.pendingRelease.value).toEqual(["g1"]);
    expect(record).not.toHaveBeenCalled();

    arrange.cancelRelease();
    expect(arrange.pendingRelease.value).toBeNull();
    expect(model.value!.slides[0].frames[0].id).toBe("g1"); // nothing changed

    arrange.ungroup();
    arrange.confirmRelease();
    expect(model.value!.slides[0].frames.map((f) => f.id)).toEqual(["c1"]);
  });

  it("an AUTHORED rotated selection path never confirms (no xml → not lossy)", () => {
    const { model, selection, arrange } = setup();
    const a = shape(model.value!, 2, 2);
    const b = shape(model.value!, 6, 2);
    selection.selectMany([a, b]);
    arrange.group();
    arrange.ungroup();
    expect(arrange.pendingRelease.value).toBeNull();
    expect(model.value!.slides[0].frames).toHaveLength(2);
  });
});

describe("useFrameArrange marquee (SPEC-096 AC-6..10)", () => {
  it("selects every frame whose rotated box intersects the marquee", () => {
    const { model, selection, arrange } = setup();
    const a = shape(model.value!, 1, 1);
    const b = shape(model.value!, 5, 1);
    shape(model.value!, 15, 15);
    arrange.onSurfacePointerDown(ptr("pointerdown", 0.5, 0.5));
    window.dispatchEvent(ptr("pointermove", 8, 4));
    expect(arrange.marquee.value).not.toBeNull();
    window.dispatchEvent(ptr("pointerup", 8, 4));
    expect(arrange.marquee.value).toBeNull();
    expect(selection.selectedFrameIds.value).toEqual([a, b]);
  });

  it("Shift adds to the existing selection instead of replacing it", () => {
    const { model, selection, arrange } = setup();
    const a = shape(model.value!, 1, 1);
    const far = shape(model.value!, 15, 15);
    selection.select(far);
    arrange.onSurfacePointerDown(ptr("pointerdown", 0.5, 0.5, { shiftKey: true }));
    window.dispatchEvent(ptr("pointermove", 4, 4));
    window.dispatchEvent(ptr("pointerup", 4, 4));
    expect(selection.selectedFrameIds.value).toEqual([a, far]);
  });

  it("a release under the threshold is the plain deselecting click", () => {
    const { model, selection, arrange } = setup();
    const a = shape(model.value!, 1, 1);
    selection.select(a);
    arrange.onSurfacePointerDown(ptr("pointerdown", 10, 10));
    window.dispatchEvent(ptr("pointerup", 10.01, 10.01));
    expect(selection.selectedFrameIds.value).toEqual([]);
  });

  it("Escape mid-drag abandons the marquee and changes nothing", () => {
    const { model, selection, arrange } = setup();
    const a = shape(model.value!, 1, 1);
    selection.select(a);
    arrange.onSurfacePointerDown(ptr("pointerdown", 8, 8));
    window.dispatchEvent(ptr("pointermove", 0.5, 0.5));
    expect(arrange.marquee.value).not.toBeNull();
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(arrange.marquee.value).toBeNull();
    window.dispatchEvent(ptr("pointerup", 0.5, 0.5)); // listener already gone
    expect(selection.selectedFrameIds.value).toEqual([a]);
  });
});

describe("useFrameArrange Position & Size (SPEC-096 AC-62..65)", () => {
  it("applies one patch as one undo entry, clamped to the minimum size", () => {
    const { model, selection, arrange, record, commit, frameAt } = setup();
    const a = shape(model.value!, 2, 2, 4, 3);
    selection.select(a);
    arrange.openPosition();
    expect(arrange.positionOpen.value).toBe(true);
    arrange.applyPosition({ xCm: 5, yCm: 6, widthCm: 0.1, heightCm: 2, rotationDeg: 90 });
    const f = frameAt(a) as ShapeFrame;
    expect(f.xCm).toBe(5);
    expect(f.yCm).toBe(6);
    expect(f.widthCm).toBe(0.5); // MIN_FRAME_SIZE_CM clamp
    expect(f.rotationDeg).toBe(90);
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Position and size");
    expect(commit).toHaveBeenCalledTimes(1);
    arrange.closePosition();
    expect(arrange.positionOpen.value).toBe(false);
  });
});
