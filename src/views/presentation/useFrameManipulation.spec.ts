// View spec: the keyboard paths of the frame-manipulation composable (nudge,
// delete, duplicate, arrange, ungroup) over a real blank deck — no DOM surface
// needed, so it exercises the command wiring directly.
import { computed, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { insertShapeFrame } from "../../lib/presentation/frames";
import type { GroupFrame, PresentationModel } from "../../lib/presentation/model";
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
