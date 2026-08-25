// SPEC-080 — the presentation undo stack: record-before-command, undo/redo
// swaps, the 100-entry cap, and reset.
import { describe, expect, it, vi } from "vitest";
import { shallowRef } from "vue";
import { MAX_UNDO_STEPS, usePresentationUndo } from "./usePresentationUndo";
import { addSlide, createBlankPresentation, deleteSlide } from "../../lib/presentation/slides";
import type { PresentationModel } from "../../lib/presentation/model";

function setup() {
  // A shallowRef mirrors PresentationView: the model object stays a plain
  // (non-proxied) object, so structuredClone can snapshot its slides.
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const onApply = vi.fn();
  const undo = usePresentationUndo({ model, onApply });
  return { model, onApply, undo };
}

describe("usePresentationUndo", () => {
  it("undoes and redoes a structural command (AC-40/41)", () => {
    const { model, undo } = setup();
    expect(undo.canUndo.value).toBe(false);

    undo.record("Add slide");
    addSlide(model.value!, 0);
    expect(model.value!.slides).toHaveLength(2);
    expect(undo.canUndo.value).toBe(true);
    expect(undo.canRedo.value).toBe(false);

    undo.undo();
    expect(model.value!.slides).toHaveLength(1);
    expect(undo.canRedo.value).toBe(true);

    undo.redo();
    expect(model.value!.slides).toHaveLength(2);
  });

  it("restores a deleted slide's frames and structure flag on undo (AC-44)", () => {
    const { model, undo } = setup();
    addSlide(model.value!, 0); // 2 slides, structureDirty set
    const secondFrames = model.value!.slides[1].frames.map((f) => f.id);

    undo.record("Delete slide");
    deleteSlide(model.value!, 1);
    expect(model.value!.slides).toHaveLength(1);

    undo.undo();
    expect(model.value!.slides).toHaveLength(2);
    expect(model.value!.slides[1].frames.map((f) => f.id)).toEqual(secondFrames);
    expect(model.value!.structureDirty).toBe(true);
  });

  it("persists through onApply on undo and redo (AC-46)", () => {
    const { model, onApply, undo } = setup();
    undo.record();
    addSlide(model.value!, 0);
    undo.undo();
    undo.redo();
    expect(onApply).toHaveBeenCalledTimes(2);
  });

  it("caps the stack at 100 entries, dropping the oldest (AC-43)", () => {
    const { model, undo } = setup();
    for (let i = 0; i < MAX_UNDO_STEPS + 20; i++) {
      undo.record();
      addSlide(model.value!, 0);
    }
    let depth = 0;
    while (undo.canUndo.value) {
      undo.undo();
      depth++;
    }
    expect(depth).toBe(MAX_UNDO_STEPS);
  });

  it("reset clears both stacks (per-open-view, not persisted — AC-45)", () => {
    const { model, undo } = setup();
    undo.record();
    addSlide(model.value!, 0);
    undo.reset();
    expect(undo.canUndo.value).toBe(false);
    expect(undo.canRedo.value).toBe(false);
  });
});
