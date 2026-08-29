// View spec: the keyboard paths of the frame-manipulation composable (nudge,
// delete, duplicate, arrange, ungroup) over a real blank deck — no DOM surface
// needed, so it exercises the command wiring directly.
import { computed, shallowRef } from "vue";
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
