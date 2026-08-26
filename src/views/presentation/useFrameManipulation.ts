// Direct manipulation of the selected slide frame (SPEC-082): pointer-drag
// move (with pointer capture, cursor feedback and centre/edge alignment
// guides), eight-handle resize (aspect rules, 0.5cm min), arrow-key nudge,
// delete, duplicate, z-order arrange and group ungroup. Every gesture records
// ONE undo entry with a human label, keeps the frame selected, and persists
// write-through on release. The pure geometry lives in stageMath.ts; this
// composable owns the DOM listeners and the command wiring.
import { ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import {
  GUIDE_SNAP_CM,
  NUDGE_LARGE_STEP_CM,
  NUDGE_STEP_CM,
  deleteFrame,
  duplicateFrame,
  moveFrame,
  reorderFrame,
  resizeFrame,
  setFramePosition,
  ungroupFrame,
  type FrameRect,
  type ReorderTarget,
} from "../../lib/presentation/frames";
import { shapeDefinition } from "../../lib/core/shapeCatalogue";
import type { FrameSelection } from "./useFrameSelection";
import { applyGuides, clientToCm, resizeRect, type ResizeHandle, type StageGuides } from "./stageMath";

export interface FrameManipulationDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  selection: FrameSelection;
  /** Whether a text-editing session is active (blocks delete/nudge, AC-27). */
  isEditing: () => boolean;
  /** The stage surface element, for client→cm conversion. */
  surface: () => HTMLElement | null;
  record: (label?: string) => void;
  /** Re-render the stage without persisting (live drag/resize preview). */
  refresh: () => void;
  /** Persist write-through + re-render (gesture end). */
  commit: () => void;
}

/** A human noun for a frame kind, for undo labels ("Move shape") and the
 * accessible name. */
export function frameNoun(frame: SlideFrame): string {
  switch (frame.kind) {
    case "image":
      return "image";
    case "shape":
      return "shape";
    case "group":
      return "group";
    case "foreign":
      return frame.label || "object";
    default:
      return "text box";
  }
}

/** The accessible name a frame is exposed under (AC-52). */
export function frameAccessibleName(frame: SlideFrame): string {
  switch (frame.kind) {
    case "image":
      return "Image";
    case "group":
      return "Group";
    case "foreign":
      return frame.label || "Object";
    case "shape":
      return shapeDefinition(frame.geometry)?.label ?? "Shape";
    default:
      return frame.role ? "Placeholder" : "Text box";
  }
}

export function useFrameManipulation(deps: FrameManipulationDeps) {
  const guides = ref<StageGuides>({ vertical: [], horizontal: [] });
  const dragging = ref(false);

  function frameById(id: string | null): SlideFrame | null {
    const slide = deps.currentSlide.value;
    if (!slide || id === null) return null;
    return slide.frames.find((f) => f.id === id) ?? null;
  }

  function surfaceRect(): DOMRect | null {
    return deps.surface()?.getBoundingClientRect() ?? null;
  }

  function pointerCm(event: PointerEvent): { xCm: number; yCm: number } | null {
    const rect = surfaceRect();
    const model = deps.model.value;
    if (!rect || !model) return null;
    return clientToCm(event.clientX, event.clientY, rect, model.slideWidthCm, model.slideHeightCm);
  }

  // --- move (pointer drag) --------------------------------------------------

  // A pointer travels at least this far (cm) before a press counts as a drag
  // rather than a click — so a click/double-click that selects or opens a text
  // session is never swallowed by pointer capture (the reason we do NOT capture
  // the pointer up front).
  const DRAG_ENGAGE_CM = (3 * 2.54) / 96;

  function onFramePointerDown(frameId: string, event: PointerEvent): void {
    if (deps.isEditing()) return;
    const frame = frameById(frameId);
    const model = deps.model.value;
    const start = pointerCm(event);
    if (!frame || !model || !start) return;
    deps.selection.select(frameId);

    const origin: FrameRect = { xCm: frame.xCm, yCm: frame.yCm, widthCm: frame.widthCm, heightCm: frame.heightCm };
    let recorded = false;
    let engaged = false;

    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now) return;
      if (!engaged) {
        if (Math.abs(now.xCm - start.xCm) < DRAG_ENGAGE_CM && Math.abs(now.yCm - start.yCm) < DRAG_ENGAGE_CM) {
          return;
        }
        engaged = true;
        dragging.value = true;
      }
      if (!recorded) {
        deps.record(`Move ${frameNoun(frame)}`);
        recorded = true;
      }
      const rect: FrameRect = { ...origin, xCm: origin.xCm + (now.xCm - start.xCm), yCm: origin.yCm + (now.yCm - start.yCm) };
      const snapped = applyGuides(rect, model.slideWidthCm, model.slideHeightCm, GUIDE_SNAP_CM);
      guides.value = snapped.guides;
      setFramePosition(model, deps.currentIndex.value, frameId, snapped.rect.xCm, snapped.rect.yCm);
      deps.refresh();
    };
    const onUp = (): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      dragging.value = false;
      guides.value = { vertical: [], horizontal: [] };
      if (recorded) deps.commit();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  // --- resize (handle drag) -------------------------------------------------

  function onHandlePointerDown(frameId: string, handle: ResizeHandle, event: PointerEvent): void {
    const frame = frameById(frameId);
    const model = deps.model.value;
    if (!frame || !model) return;
    event.stopPropagation();
    deps.selection.select(frameId);

    const origin: FrameRect = { xCm: frame.xCm, yCm: frame.yCm, widthCm: frame.widthCm, heightCm: frame.heightCm };
    // Images keep aspect by default (Shift frees); other frames keep aspect
    // only while Shift is held (AC-22/23). Tracked via window listeners (no
    // pointer capture, matching the move gesture).
    let recorded = false;
    dragging.value = true;

    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now) return;
      if (!recorded) {
        deps.record(`Resize ${frameNoun(frame)}`);
        recorded = true;
      }
      const aspect = frame.kind === "image" ? !move.shiftKey : move.shiftKey;
      const rect = resizeRect(handle, origin, now, aspect);
      resizeFrame(model, deps.currentIndex.value, frameId, rect);
      deps.refresh();
    };
    const onUp = (): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      dragging.value = false;
      if (recorded) deps.commit();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  // --- keyboard (nudge / delete / duplicate / arrange / ungroup) ------------

  function run(label: string, apply: (model: PresentationModel, index: number) => string | string[] | null): void {
    const model = deps.model.value;
    if (!model) return;
    deps.record(label);
    const selected = apply(model, deps.currentIndex.value);
    if (Array.isArray(selected)) {
      deps.selection.select(selected[0] ?? null);
    } else {
      deps.selection.select(selected);
    }
    deps.commit();
  }

  function nudge(dxCm: number, dyCm: number): void {
    const id = deps.selection.selectedFrameId.value;
    const frame = frameById(id);
    if (!id || !frame) return;
    run(`Move ${frameNoun(frame)}`, (model, index) => moveFrame(model, index, id, dxCm, dyCm));
  }

  function remove(): void {
    const id = deps.selection.selectedFrameId.value;
    const frame = frameById(id);
    if (!id || !frame) return;
    run(`Delete ${frameNoun(frame)}`, (model, index) => deleteFrame(model, index, id));
  }

  function duplicate(): void {
    const id = deps.selection.selectedFrameId.value;
    const frame = frameById(id);
    if (!id || !frame) return;
    run(`Duplicate ${frameNoun(frame)}`, (model, index) => duplicateFrame(model, index, id));
  }

  function arrange(target: ReorderTarget): void {
    const id = deps.selection.selectedFrameId.value;
    const frame = frameById(id);
    if (!id || !frame) return;
    const labels: Record<ReorderTarget, string> = {
      front: "Bring to front",
      forward: "Bring forward",
      backward: "Send backward",
      back: "Send to back",
    };
    run(labels[target], (model, index) => reorderFrame(model, index, id, target));
  }

  function ungroup(): void {
    const id = deps.selection.selectedFrameId.value;
    const frame = frameById(id);
    if (!id || frame?.kind !== "group") return;
    run("Ungroup", (model, index) => ungroupFrame(model, index, id));
  }

  /** Returns true when it handled the key (the caller then stops). */
  function onKeydown(event: KeyboardEvent): boolean {
    if (deps.isEditing()) return false;
    const id = deps.selection.selectedFrameId.value;
    if (id === null) return false;
    const mod = event.metaKey || event.ctrlKey;

    if (mod && event.shiftKey && (event.key === "g" || event.key === "G")) {
      event.preventDefault();
      ungroup();
      return true;
    }
    if (mod && (event.key === "d" || event.key === "D")) {
      event.preventDefault();
      duplicate();
      return true;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      remove();
      return true;
    }
    const step = event.shiftKey ? NUDGE_LARGE_STEP_CM : NUDGE_STEP_CM;
    if (event.key === "ArrowLeft") return nudgeKey(event, -step, 0);
    if (event.key === "ArrowRight") return nudgeKey(event, step, 0);
    if (event.key === "ArrowUp") return nudgeKey(event, 0, -step);
    if (event.key === "ArrowDown") return nudgeKey(event, 0, step);
    return false;
  }

  function nudgeKey(event: KeyboardEvent, dxCm: number, dyCm: number): boolean {
    event.preventDefault();
    nudge(dxCm, dyCm);
    return true;
  }

  return {
    guides,
    dragging,
    onFramePointerDown,
    onHandlePointerDown,
    onKeydown,
    remove,
    duplicate,
    arrange,
    ungroup,
  };
}
