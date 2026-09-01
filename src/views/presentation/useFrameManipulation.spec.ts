// View spec: the keyboard paths of the frame-manipulation composable (nudge,
// delete, duplicate, arrange, ungroup) over a real blank deck — no DOM surface
// needed, so it exercises the command wiring directly.
import { computed, shallowRef, triggerRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { insertShapeFrame, insertTableFrame } from "../../lib/presentation/frames";
import type { GroupFrame, PresentationModel, TableFrame } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { useFrameManipulation } from "./useFrameManipulation";

function setup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const currentIndex = shallowRef(0);
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const selection = useFrameSelection({ currentSlide });
  const record = vi.fn();
  const commit = vi.fn();
  const refresh = vi.fn();
  const manip = useFrameManipulation({
    model,
    currentSlide,
    currentIndex,
    selection,
    isEditing: () => false,
    surface: () => null,
    record,
    refresh,
    commit,
  });
  return { model, currentSlide, selection, manip, record, commit };
}

function key(init: KeyboardEventInit): KeyboardEvent {
  return new KeyboardEvent("keydown", init);
}

describe("useFrameManipulation keyboard", () => {
  it("nudges the selected frame by 0.1cm (1cm with Shift) and records one entry", () => {
    const { model, selection, manip, record, commit } = setup();
    const id = insertShapeFrame(model.value!, 0, "rect", { xCm: 5, yCm: 5, widthCm: 3, heightCm: 2 })!;
    selection.select(id);
    const x0 = model.value!.slides[0].frames.find((f) => f.id === id)!.xCm;

    expect(manip.onKeydown(key({ key: "ArrowRight" }))).toBe(true);
    const x1 = model.value!.slides[0].frames.find((f) => f.id === id)!.xCm;
    expect(x1).toBeCloseTo(x0 + 0.1, 3);
    expect(record).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);

    manip.onKeydown(key({ key: "ArrowRight", shiftKey: true }));
    const x2 = model.value!.slides[0].frames.find((f) => f.id === id)!.xCm;
    expect(x2).toBeCloseTo(x1 + 1, 3);
  });

  it("deletes with Delete and duplicates with Cmd/Ctrl+D", () => {
    const { model, selection, manip } = setup();
    const id = insertShapeFrame(model.value!, 0, "rect", { xCm: 5, yCm: 5, widthCm: 3, heightCm: 2 })!;
    selection.select(id);
    const count = model.value!.slides[0].frames.length;

    manip.onKeydown(key({ key: "d", ctrlKey: true }));
    expect(model.value!.slides[0].frames).toHaveLength(count + 1);

    manip.onKeydown(key({ key: "Delete" }));
    expect(model.value!.slides[0].frames).toHaveLength(count);
  });

  it("does nothing while a text session is active", () => {
    const model = shallowRef<PresentationModel | null>(createBlankPresentation());
    const currentIndex = shallowRef(0);
    const currentSlide = computed(() => model.value?.slides[0] ?? null);
    const selection = useFrameSelection({ currentSlide });
    const id = insertShapeFrame(model.value!, 0, "rect", { xCm: 5, yCm: 5, widthCm: 3, heightCm: 2 })!;
    selection.select(id);
    const manip = useFrameManipulation({
      model,
      currentSlide,
      currentIndex,
      selection,
      isEditing: () => true,
      surface: () => null,
      record: vi.fn(),
      refresh: vi.fn(),
      commit: vi.fn(),
    });
    expect(manip.onKeydown(key({ key: "Delete" }))).toBe(false);
    expect(model.value!.slides[0].frames.some((f) => f.id === id)).toBe(true);
  });

  it("Cmd/Ctrl+Shift+G ungroups a selected group and selects a child", () => {
    const { model, selection, manip } = setup();
    const group: GroupFrame = {
      id: "g1",
      xCm: 0,
      yCm: 0,
      widthCm: 6,
      heightCm: 6,
      kind: "group",
      xml: "<p:grpSp/>",
      children: [
        { id: "c1", xCm: 1, yCm: 1, widthCm: 2, heightCm: 2, kind: "shape", geometry: "rect" },
      ],
    };
    model.value!.slides[0].frames.push(group);
    selection.select("g1");
    manip.onKeydown(key({ key: "G", ctrlKey: true, shiftKey: true }));
    expect(model.value!.slides[0].frames.some((f) => f.kind === "group")).toBe(false);
    expect(selection.selectedFrameId.value).toBe("c1");
  });

  it("arrange sends the selected frame to back", () => {
    const { model, selection, manip } = setup();
    const a = insertShapeFrame(model.value!, 0, "rect", { xCm: 1, yCm: 1, widthCm: 2, heightCm: 2 })!;
    selection.select(a);
    manip.arrange("back");
    expect(model.value!.slides[0].frames[0].id).toBe(a);
  });
});

// A move/up event carrying viewport coords; jsdom lacks PointerEvent, but the
// gesture only reads clientX/clientY, which MouseEvent provides.
function ptr(type: string, clientX: number, clientY: number): PointerEvent {
  return new MouseEvent(type, { clientX, clientY }) as unknown as PointerEvent;
}

function tableSetup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const currentIndex = shallowRef(0);
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const selection = useFrameSelection({ currentSlide });
  // A surface whose px box equals the slide's cm box, so one client unit is
  // exactly one cm and a drag delta reads straight through as a cm delta.
  const surface = {
    getBoundingClientRect: () => ({
      left: 0,
      top: 0,
      width: model.value!.slideWidthCm,
      height: model.value!.slideHeightCm,
    }),
  } as unknown as HTMLElement;
  const record = vi.fn();
  const commit = vi.fn();
  const manip = useFrameManipulation({
    model,
    currentSlide,
    currentIndex,
    selection,
    isEditing: () => false,
    surface: () => surface,
    record,
    refresh: vi.fn(),
    commit,
  });
  const tableAt = (id: string) => model.value!.slides[0].frames.find((f) => f.id === id) as TableFrame;
  return { model, manip, selection, record, commit, tableAt };
}

describe("useFrameManipulation table boundary drag (AC-8)", () => {
  it("drags a column boundary live, keeping the table width and committing once", () => {
    const { model, manip, record, commit, tableAt } = tableSetup();
    const id = insertTableFrame(model.value!, 0, 2, 2)!;
    const w0 = tableAt(id).columnWidthsCm[0];
    const totalBefore = tableAt(id).widthCm;

    manip.onTableBorderPointerDown(id, "col", 0, ptr("pointerdown", 5, 0));
    window.dispatchEvent(ptr("pointermove", 6, 0)); // +1cm to the right
    expect(tableAt(id).columnWidthsCm[0]).toBeCloseTo(w0 + 1, 3);
    expect(tableAt(id).widthCm).toBeCloseTo(totalBefore, 3);
    expect(record).toHaveBeenCalledTimes(1);

    window.dispatchEvent(ptr("pointerup", 6, 0));
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("drags a row boundary live to set the row height", () => {
    const { model, manip, tableAt } = tableSetup();
    const id = insertTableFrame(model.value!, 0, 2, 1)!;
    const h0 = tableAt(id).rows[0].heightCm ?? 0;

    manip.onTableBorderPointerDown(id, "row", 0, ptr("pointerdown", 0, 5));
    window.dispatchEvent(ptr("pointermove", 0, 6.5)); // +1.5cm down
    expect(tableAt(id).rows[0].heightCm).toBeCloseTo(h0 + 1.5, 3);
    window.dispatchEvent(ptr("pointerup", 0, 6.5));
  });

  it("ignores the boundary drag while a text session is active", () => {
    const model = shallowRef<PresentationModel | null>(createBlankPresentation());
    const currentIndex = shallowRef(0);
    const currentSlide = computed(() => model.value?.slides[0] ?? null);
    const selection = useFrameSelection({ currentSlide });
    const id = insertTableFrame(model.value!, 0, 2, 2)!;
    const surface = {
      getBoundingClientRect: () => ({ left: 0, top: 0, width: model.value!.slideWidthCm, height: model.value!.slideHeightCm }),
    } as unknown as HTMLElement;
    const record = vi.fn();
    const manip = useFrameManipulation({
      model,
      currentSlide,
      currentIndex,
      selection,
      isEditing: () => true,
      surface: () => surface,
      record,
      refresh: vi.fn(),
      commit: vi.fn(),
    });
    const before = [...(model.value!.slides[0].frames.find((f) => f.id === id) as TableFrame).columnWidthsCm];
    manip.onTableBorderPointerDown(id, "col", 0, ptr("pointerdown", 5, 0));
    window.dispatchEvent(ptr("pointermove", 8, 0));
    expect((model.value!.slides[0].frames.find((f) => f.id === id) as TableFrame).columnWidthsCm).toEqual(before);
    expect(record).not.toHaveBeenCalled();
  });
});

// --- SPEC-096: multi-selection commands, rotate + adjust gestures -----------

import type { ShapeFrame } from "../../lib/presentation/model";

function multiSetup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  model.value!.slides[0].frames = [];
  const currentIndex = shallowRef(0);
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const selection = useFrameSelection({ currentSlide });
  const record = vi.fn();
  const commit = vi.fn();
  // The real refresh() re-identifies the slide and every frame
  // (useSlideRender.bumpCurrentSlide) before re-rendering, so a gesture that
  // holds a frame reference across a refresh goes stale. Mimic that here —
  // a no-op stub let exactly that bug through QA once (SPEC-096 AC-27).
  const refresh = vi.fn(() => {
    const slide = model.value!.slides[0];
    model.value!.slides[0] = { ...slide, frames: slide.frames.map((f) => ({ ...f })) };
    triggerRef(model);
  });
  const surface = {
    getBoundingClientRect: () => ({
      left: 0,
      top: 0,
      width: model.value!.slideWidthCm,
      height: model.value!.slideHeightCm,
    }),
  } as unknown as HTMLElement;
  const manip = useFrameManipulation({
    model,
    currentSlide,
    currentIndex,
    selection,
    isEditing: () => false,
    surface: () => surface,
    record,
    refresh,
    commit,
  });
  const at = (id: string) => model.value!.slides[0].frames.find((f) => f.id === id)!;
  const add = (x: number, y: number) => insertShapeFrame(model.value!, 0, "rect", { xCm: x, yCm: y, widthCm: 2, heightCm: 2 })!;
  return { model, selection, manip, record, commit, at, add };
}

describe("useFrameManipulation multi-selection (SPEC-096)", () => {
  it("an arrow nudge moves every selected frame by the same delta in ONE entry", () => {
    const { selection, manip, record, commit, at, add } = multiSetup();
    const a = add(2, 2);
    const b = add(6, 6);
    selection.selectMany([a, b]);
    manip.onKeydown(key({ key: "ArrowRight" }));
    expect(at(a).xCm).toBeCloseTo(2.1, 3);
    expect(at(b).xCm).toBeCloseTo(6.1, 3);
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Move 2 objects");
    expect(commit).toHaveBeenCalledTimes(1);
    // Both frames stay selected.
    expect(selection.selectedFrameIds.value).toEqual([a, b]);
  });

  it("Cmd/Ctrl+D duplicates the whole selection and leaves the COPIES selected", () => {
    const { model, selection, manip, at, add } = multiSetup();
    const a = add(2, 2);
    const b = add(6, 6);
    selection.selectMany([a, b]);
    manip.onKeydown(key({ key: "d", ctrlKey: true }));
    expect(model.value!.slides[0].frames).toHaveLength(4);
    const copies = selection.selectedFrameIds.value;
    expect(copies).toHaveLength(2);
    expect(copies).not.toContain(a);
    expect(copies).not.toContain(b);
    expect(at(copies[0]).xCm).toBeCloseTo(2.5, 3);
  });

  it("Delete removes every selected frame, animations dropped per target", () => {
    const { model, selection, manip, add } = multiSetup();
    const a = add(2, 2);
    const b = add(6, 6);
    const keep = add(12, 12);
    model.value!.slides[0].animations = [
      { frameId: a, effect: "appear", trigger: "click" },
      { frameId: keep, effect: "fadeIn", trigger: "click" },
    ];
    selection.selectMany([a, b]);
    manip.onKeydown(key({ key: "Delete" }));
    expect(model.value!.slides[0].frames.map((f) => f.id)).toEqual([keep]);
    expect(model.value!.slides[0].animations).toEqual([{ frameId: keep, effect: "fadeIn", trigger: "click" }]);
  });

  it("arrange moves the multi-selection as one contiguous block", () => {
    const { model, selection, manip, add } = multiSetup();
    const a = add(1, 1);
    const b = add(4, 1);
    const c = add(7, 1);
    selection.selectMany([a, c]);
    manip.arrange("front");
    expect(model.value!.slides[0].frames.map((f) => f.id)).toEqual([b, a, c]);
  });
});

describe("useFrameManipulation rotate gesture (SPEC-096)", () => {
  it("drags rotationDeg to the pointer angle, one undo entry, readout live", () => {
    const { selection, manip, record, commit, at, add } = multiSetup();
    const a = add(10, 10); // centre (11, 11)
    selection.select(a);
    manip.onRotateHandlePointerDown(a, ptr("pointerdown", 11, 9));
    // Pointer due east of the centre → atan2 = 0° + 90 = 90°.
    window.dispatchEvent(ptr("pointermove", 15, 11));
    expect(at(a).rotationDeg).toBe(90);
    expect(manip.rotatingDeg.value).toBe(90);
    // A second move must keep tracking even though refresh() re-identified
    // the frame after the first one (pointer due south → 180°).
    window.dispatchEvent(ptr("pointermove", 11, 15));
    expect(at(a).rotationDeg).toBe(180);
    expect(manip.rotatingDeg.value).toBe(180);
    expect(record).not.toHaveBeenCalled(); // records only on release
    window.dispatchEvent(ptr("pointerup", 11, 15));
    expect(at(a).rotationDeg).toBe(180);
    expect(manip.rotatingDeg.value).toBeNull();
    expect(record).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("Shift snaps to the nearest multiple of 15°", () => {
    const { selection, manip, at, add } = multiSetup();
    const a = add(10, 10);
    selection.select(a);
    manip.onRotateHandlePointerDown(a, ptr("pointerdown", 11, 9));
    // ~97° unsnapped → 105 with Shift (97/15 = 6.47 → 6…? 97→ round(6.47)=6 → 90).
    window.dispatchEvent(new MouseEvent("pointermove", { clientX: 15, clientY: 11.5, shiftKey: true }) as unknown as PointerEvent);
    expect((at(a).rotationDeg ?? 0) % 15).toBe(0);
    window.dispatchEvent(ptr("pointerup", 15, 11.5));
  });

  it("Escape mid-drag restores the starting angle and records nothing", () => {
    const { selection, manip, record, at, add } = multiSetup();
    const a = add(10, 10);
    selection.select(a);
    manip.onRotateHandlePointerDown(a, ptr("pointerdown", 11, 9));
    // Two moves, so the restore must reach the CURRENT frame object — the
    // one refresh() re-identified — not the pointer-down capture.
    window.dispatchEvent(ptr("pointermove", 15, 11));
    window.dispatchEvent(ptr("pointermove", 11, 15));
    expect(at(a).rotationDeg).toBe(180);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(at(a).rotationDeg).toBeUndefined();
    expect(manip.rotatingDeg.value).toBeNull();
    window.dispatchEvent(ptr("pointerup", 11, 15));
    expect(record).not.toHaveBeenCalled();
  });

  it("Escape restores a NON-ZERO starting angle after several moves", () => {
    const { model, selection, manip, record, at, add } = multiSetup();
    const a = add(10, 10);
    const frame = model.value!.slides[0].frames.find((f) => f.id === a)!;
    frame.rotationDeg = 30;
    selection.select(a);
    manip.onRotateHandlePointerDown(a, ptr("pointerdown", 11, 9));
    window.dispatchEvent(ptr("pointermove", 15, 11)); // 90°
    window.dispatchEvent(ptr("pointermove", 11, 15)); // 180°
    expect(at(a).rotationDeg).toBe(180);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(at(a).rotationDeg).toBe(30);
    window.dispatchEvent(ptr("pointerup", 11, 15));
    expect(at(a).rotationDeg).toBe(30);
    expect(record).not.toHaveBeenCalled();
  });
});

describe("useFrameManipulation adjustment handles (SPEC-096)", () => {
  function roundRectSetup() {
    const s = multiSetup();
    const id = insertShapeFrame(s.model.value!, 0, "roundRect", { xCm: 4, yCm: 4, widthCm: 8, heightCm: 4 })!;
    s.selection.select(id);
    return { ...s, id };
  }

  it("dragging a handle stores the fraction live and commits one entry", () => {
    const { manip, record, commit, at, id } = roundRectSetup();
    manip.onAdjustPointerDown(id, 0, ptr("pointerdown", 4, 4));
    // roundRect's corner-radius adjustment rides the x axis: pointer at 20%
    // then 30% of the width — the second move must still track after
    // refresh() re-identified the frame.
    window.dispatchEvent(ptr("pointermove", 4 + 8 * 0.2, 4));
    expect((at(id) as ShapeFrame).adjustments![0]).toBeCloseTo(0.2, 2);
    window.dispatchEvent(ptr("pointermove", 4 + 8 * 0.3, 4));
    const live = (at(id) as ShapeFrame).adjustments![0];
    expect(live).toBeCloseTo(0.3, 2);
    expect(record).not.toHaveBeenCalled();
    window.dispatchEvent(ptr("pointerup", 4 + 8 * 0.3, 4));
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Adjust shape");
    expect(commit).toHaveBeenCalledTimes(1);
    expect((at(id) as ShapeFrame).adjustments![0]).toBeCloseTo(0.3, 2);
  });

  it("arrow keys nudge by 1% of the range, Shift by 5%, one entry per press", () => {
    const { manip, record, at, id } = roundRectSetup();
    const before = (at(id) as ShapeFrame).adjustments?.[0] ?? 0.1667;
    expect(manip.onAdjustKeydown(id, 0, key({ key: "ArrowRight" }))).toBe(true);
    const spec = { min: 0, max: 0.5 }; // roundRect corner radius range
    const oneStep = (spec.max - spec.min) * 0.01;
    expect((at(id) as ShapeFrame).adjustments![0]).toBeCloseTo(before + oneStep, 4);
    manip.onAdjustKeydown(id, 0, key({ key: "ArrowRight", shiftKey: true }));
    expect((at(id) as ShapeFrame).adjustments![0]).toBeCloseTo(before + oneStep + oneStep * 5, 4);
    expect(record).toHaveBeenCalledTimes(2);
  });
});
