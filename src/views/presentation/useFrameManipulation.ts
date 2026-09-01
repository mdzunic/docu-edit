// Direct manipulation of the selected slide frames (SPEC-082, multi-selection
// + rotation + adjustments SPEC-096): pointer-drag move (with pointer capture,
// cursor feedback and centre/edge alignment guides — now including other
// frames' edges/centres), eight-handle resize (aspect rules, 0.5cm min,
// composed with rotation in the frame's unrotated local space), a rotate
// handle (Shift snaps 15°, Escape restores), diamond adjustment handles
// (dragged along their declared axis, arrow-nudged 1%/5%), arrow-key nudge,
// delete, duplicate, z-order arrange and ungroup — each applying to EVERY
// selected frame in ONE undo entry. The pure geometry lives in stageMath.ts /
// core/arrange.ts; this composable owns the DOM listeners and the command
// wiring.
import { ref, type ComputedRef, type Ref } from "vue";
import { DEFAULT_TABLE_ROW_HEIGHT_CM, type PresentationModel, type Slide, type SlideFrame } from "../../lib/presentation/model";
import {
  GUIDE_SNAP_CM,
  NUDGE_LARGE_STEP_CM,
  NUDGE_STEP_CM,
  deleteFrame,
  duplicateFrame,
  frameRotatedBox,
  moveFrame,
  moveFrames,
  reorderFrame,
  reorderFrames,
  resizeFrame,
  resizeTableColumnBoundary,
  resizeTableRowBoundary,
  setFrameAdjustments,
  setFramePosition,
  setFrameRotation,
  ungroupFrame,
  type FrameRect,
  type ReorderTarget,
} from "../../lib/presentation/frames";
import { shapeDefinition } from "../../lib/core/shapeCatalogue";
import { normalizeRotationDeg } from "../../lib/core/shapes";
import { unrotatePoint } from "../../lib/core/shapeGeometry";
import { adjustmentValueFromPoint, nudgedAdjustments } from "../../lib/shapeAdjustHandles";
import { resolvedShapeAdjustments } from "../../lib/core/shapeCatalogue";
import type { FrameSelection } from "./useFrameSelection";
import { applyGuides, clientToCm, DRAG_ENGAGE_CM, resizeRotatedRect, type ResizeHandle, type StageGuides } from "./stageMath";

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
    case "table":
      return "table";
    case "chart":
      return "chart";
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
    case "table":
      return "Table";
    case "chart":
      return "Chart";
    default:
      return frame.role ? "Placeholder" : "Text box";
  }
}

export function useFrameManipulation(deps: FrameManipulationDeps) {
  const guides = ref<StageGuides>({ vertical: [], horizontal: [] });
  const dragging = ref(false);
  /** The live rotation angle while a rotate drag runs (the aria-live
   * rotation-readout), null otherwise. */
  const rotatingDeg = ref<number | null>(null);

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

  /** Every OTHER top-level frame's rotated box — the smart-guide candidates
   * (SPEC-096: guide activation uses rotatedBounds). */
  function otherBoxes(excluded: readonly string[]): FrameRect[] {
    const slide = deps.currentSlide.value;
    if (!slide) return [];
    return slide.frames.filter((f) => !excluded.includes(f.id)).map(frameRotatedBox);
  }

  // --- move (pointer drag) --------------------------------------------------

  // A pointer travels at least DRAG_ENGAGE_CM (stageMath) before a press
  // counts as a drag rather than a click — so a click/double-click that
  // selects or opens a text session is never swallowed by pointer capture
  // (the reason we do NOT capture the pointer up front).
  function onFramePointerDown(frameId: string, event: PointerEvent): void {
    if (deps.isEditing()) return;
    const frame = frameById(frameId);
    const model = deps.model.value;
    const start = pointerCm(event);
    if (!frame || !model || !start) return;
    // Shift+click toggles membership (SPEC-096 AC-2) and starts no drag; a
    // plain press on an already-selected frame keeps the multi-selection so
    // the drag moves the whole set.
    if (event.shiftKey) {
      deps.selection.toggle(frameId);
      return;
    }
    if (!deps.selection.isSelected(frameId)) {
      deps.selection.select(frameId);
    }
    const ids = [...deps.selection.selectedFrameIds.value];
    if (ids.length > 1) {
      multiMoveDrag(ids, start);
      return;
    }

    const origin: FrameRect = { xCm: frame.xCm, yCm: frame.yCm, widthCm: frame.widthCm, heightCm: frame.heightCm };
    const box = frameRotatedBox(frame);
    const pad = { x: origin.xCm - box.xCm, y: origin.yCm - box.yCm };
    const others = otherBoxes([frameId]);
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
      // Snap the frame's ROTATED axis-aligned box, then translate back.
      const rect: FrameRect = { ...box, xCm: box.xCm + (now.xCm - start.xCm), yCm: box.yCm + (now.yCm - start.yCm) };
      const snapped = applyGuides(rect, model.slideWidthCm, model.slideHeightCm, GUIDE_SNAP_CM, others);
      guides.value = snapped.guides;
      setFramePosition(model, deps.currentIndex.value, frameId, snapped.rect.xCm + pad.x, snapped.rect.yCm + pad.y);
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

  /** A drag over ≥2 selected frames: every frame moves by the same delta,
   * clamped on the union box (moveFrames), in ONE undo entry. */
  function multiMoveDrag(ids: string[], start: { xCm: number; yCm: number }): void {
    const model = deps.model.value;
    if (!model) return;
    let recorded = false;
    let engaged = false;
    const applied = { x: 0, y: 0 };

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
        deps.record(`Move ${ids.length} objects`);
        recorded = true;
      }
      const total = { x: now.xCm - start.xCm, y: now.yCm - start.yCm };
      moveFrames(model, deps.currentIndex.value, ids, total.x - applied.x, total.y - applied.y);
      applied.x = total.x;
      applied.y = total.y;
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

  // --- resize (handle drag) -------------------------------------------------

  function onHandlePointerDown(frameId: string, handle: ResizeHandle, event: PointerEvent): void {
    const frame = frameById(frameId);
    const model = deps.model.value;
    if (!frame || !model) return;
    event.stopPropagation();
    deps.selection.select(frameId);

    const origin: FrameRect = { xCm: frame.xCm, yCm: frame.yCm, widthCm: frame.widthCm, heightCm: frame.heightCm };
    const rotation = normalizeRotationDeg(frame.rotationDeg ?? 0);
    const others = otherBoxes([frameId]);
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
      const rect = resizeRotatedRect(handle, origin, rotation, now, aspect);
      resizeFrame(model, deps.currentIndex.value, frameId, rect);
      // Guides fire during a resize too (SPEC-096 AC) — display-only lines on
      // the result's rotated box; the resize itself never snaps (the anchored
      // corner must stay put).
      const resized = frameById(frameId);
      if (resized) {
        guides.value = applyGuides(frameRotatedBox(resized), model.slideWidthCm, model.slideHeightCm, GUIDE_SNAP_CM, others).guides;
      }
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

  // --- rotate (rotate-handle drag, SPEC-096) --------------------------------

  function onRotateHandlePointerDown(frameId: string, event: PointerEvent): void {
    if (deps.isEditing()) return;
    const frame = frameById(frameId);
    const model = deps.model.value;
    if (!frame || !model) return;
    event.stopPropagation();
    deps.selection.select(frameId);

    const startDeg = normalizeRotationDeg(frame.rotationDeg ?? 0);
    const centre = { x: frame.xCm + frame.widthCm / 2, y: frame.yCm + frame.heightCm / 2 };
    let moved = false;
    dragging.value = true;
    rotatingDeg.value = startDeg;

    const restore = (): void => {
      if (startDeg === 0) {
        delete frame.rotationDeg;
      } else {
        frame.rotationDeg = startDeg;
      }
    };
    const cleanup = (): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("keydown", onKey, true);
      dragging.value = false;
      rotatingDeg.value = null;
    };
    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now) return;
      moved = true;
      // deg = atan2 about the centre + 90 (the handle sits above the frame).
      let deg = normalizeRotationDeg((Math.atan2(now.yCm - centre.y, now.xCm - centre.x) * 180) / Math.PI + 90);
      if (move.shiftKey) {
        deg = normalizeRotationDeg(Math.round(deg / 15) * 15);
      }
      // Live preview mutates the frame directly (no dirty flag until the
      // gesture commits — an Escape must leave a clean model).
      if (deg === 0) {
        delete frame.rotationDeg;
      } else {
        frame.rotationDeg = deg;
      }
      rotatingDeg.value = deg;
      deps.refresh();
    };
    const onKey = (key: KeyboardEvent): void => {
      if (key.key !== "Escape") return;
      key.stopPropagation();
      // Escape mid-drag restores the starting angle and records nothing.
      restore();
      cleanup();
      deps.refresh();
    };
    const onUp = (): void => {
      const finalDeg = normalizeRotationDeg(frame.rotationDeg ?? 0);
      cleanup();
      if (!moved || finalDeg === startDeg) {
        restore();
        deps.refresh();
        return;
      }
      // One undo entry per gesture: restore the pre-gesture state, record it,
      // then apply the final angle as the command.
      restore();
      deps.record(`Rotate ${frameNoun(frame)}`);
      setFrameRotation(model, deps.currentIndex.value, frameId, finalDeg);
      deps.commit();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("keydown", onKey, true);
  }

  // --- adjustment handles (SPEC-096) ----------------------------------------

  function onAdjustPointerDown(frameId: string, index: number, event: PointerEvent): void {
    if (deps.isEditing()) return;
    const frame = frameById(frameId);
    const model = deps.model.value;
    if (!frame || frame.kind !== "shape" || !model) return;
    event.stopPropagation();

    const startStored = frame.adjustments ? [...frame.adjustments] : undefined;
    let moved = false;
    dragging.value = true;

    const restore = (): void => {
      if (startStored) {
        frame.adjustments = startStored;
      } else {
        delete frame.adjustments;
      }
    };
    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now || frame.widthCm <= 0 || frame.heightCm <= 0) return;
      // Map the pointer into the shape's own unrotated unit square.
      const local = unrotatePoint(
        { x: now.xCm - frame.xCm, y: now.yCm - frame.yCm },
        frame.widthCm,
        frame.heightCm,
        frame.rotationDeg ?? 0,
      );
      const value = adjustmentValueFromPoint(frame.geometry, index, local.x / frame.widthCm, local.y / frame.heightCm);
      if (value === null) return;
      moved = true;
      const next = resolvedShapeAdjustments(frame.geometry, frame.adjustments);
      next[index] = value;
      frame.adjustments = next; // live, no dirty flag until commit
      deps.refresh();
    };
    const onUp = (): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      dragging.value = false;
      if (!moved) return;
      const final = frame.adjustments ? [...frame.adjustments] : [];
      restore();
      deps.record("Adjust shape");
      setFrameAdjustments(model, deps.currentIndex.value, frameId, final);
      deps.commit();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  /** Arrow keys on a focused adjust-handle nudge the value by 1% of its range
   * (5% with Shift), one undo entry per press. Returns true when handled. */
  function onAdjustKeydown(frameId: string, index: number, event: KeyboardEvent): boolean {
    const frame = frameById(frameId);
    const model = deps.model.value;
    if (!frame || frame.kind !== "shape" || !model) return false;
    const steps =
      event.key === "ArrowRight" || event.key === "ArrowDown" ? 1 : event.key === "ArrowLeft" || event.key === "ArrowUp" ? -1 : 0;
    if (steps === 0) return false;
    event.preventDefault();
    const next = nudgedAdjustments(frame.geometry, frame.adjustments, index, steps, event.shiftKey);
    deps.record("Adjust shape");
    setFrameAdjustments(model, deps.currentIndex.value, frameId, next);
    deps.commit();
    return true;
  }

  // --- table grid resize (boundary drag) ------------------------------------

  // Dragging an internal column/row boundary of the selected table live-resizes
  // it (AC-8a/b): one undo entry, refresh per move, commit on release — the same
  // record/refresh/commit shape as the frame move/resize gestures. The origin
  // (widths / row height) is snapshotted at pointer-down so each move recomputes
  // from a stable base rather than accumulating rounding.
  function onTableBorderPointerDown(frameId: string, axis: "col" | "row", index: number, event: PointerEvent): void {
    if (deps.isEditing()) return;
    const frame = frameById(frameId);
    const model = deps.model.value;
    const start = pointerCm(event);
    if (!frame || frame.kind !== "table" || !model || !start) return;
    deps.selection.select(frameId);

    const originWidths = [...frame.columnWidthsCm];
    const originRowHeight = frame.rows[index]?.heightCm ?? DEFAULT_TABLE_ROW_HEIGHT_CM;
    let recorded = false;
    dragging.value = true;

    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now) return;
      if (!recorded) {
        deps.record(axis === "col" ? "Resize column" : "Resize row");
        recorded = true;
      }
      if (axis === "col") {
        resizeTableColumnBoundary(model, deps.currentIndex.value, frameId, index, originWidths, now.xCm - start.xCm);
      } else {
        resizeTableRowBoundary(model, deps.currentIndex.value, frameId, index, originRowHeight, now.yCm - start.yCm);
      }
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
      deps.selection.selectMany(selected);
    } else {
      deps.selection.select(selected);
    }
    deps.commit();
  }

  /** The whole selection, z-order stable. */
  function selectedIds(): string[] {
    return [...deps.selection.selectedFrameIds.value];
  }

  function nudge(dxCm: number, dyCm: number): void {
    const ids = selectedIds();
    if (ids.length === 0) return;
    if (ids.length > 1) {
      run(`Move ${ids.length} objects`, (model, index) => moveFrames(model, index, ids, dxCm, dyCm));
      return;
    }
    const frame = frameById(ids[0]);
    if (!frame) return;
    run(`Move ${frameNoun(frame)}`, (model, index) => moveFrame(model, index, ids[0], dxCm, dyCm));
  }

  function remove(): void {
    const ids = selectedIds();
    if (ids.length === 0) return;
    if (ids.length > 1) {
      // One undo entry for the whole set; each target's animations are
      // dropped by deleteFrame through removeAnimationsForFrame (AC-15).
      run(`Delete ${ids.length} objects`, (model, index) => {
        let neighbour: string | null = null;
        for (const id of ids) {
          neighbour = deleteFrame(model, index, id);
        }
        return neighbour;
      });
      return;
    }
    const frame = frameById(ids[0]);
    if (!frame) return;
    run(`Delete ${frameNoun(frame)}`, (model, index) => deleteFrame(model, index, ids[0]));
  }

  function duplicate(): void {
    const ids = selectedIds();
    if (ids.length === 0) return;
    if (ids.length > 1) {
      // Duplicates all selected frames by DUPLICATE_OFFSET_CM and leaves the
      // COPIES selected (AC-14), one undo entry.
      run(`Duplicate ${ids.length} objects`, (model, index) => {
        const copies: string[] = [];
        for (const id of ids) {
          const copy = duplicateFrame(model, index, id);
          if (copy) copies.push(copy);
        }
        return copies;
      });
      return;
    }
    const frame = frameById(ids[0]);
    if (!frame) return;
    run(`Duplicate ${frameNoun(frame)}`, (model, index) => duplicateFrame(model, index, ids[0]));
  }

  function arrange(target: ReorderTarget): void {
    const ids = selectedIds();
    if (ids.length === 0) return;
    const labels: Record<ReorderTarget, string> = {
      front: "Bring to front",
      forward: "Bring forward",
      backward: "Send backward",
      back: "Send to back",
    };
    if (ids.length > 1) {
      run(labels[target], (model, index) => reorderFrames(model, index, ids, target));
      return;
    }
    run(labels[target], (model, index) => reorderFrame(model, index, ids[0], target));
  }

  /** Releases every selected group in one undo entry (AC-20). The
   * confirm-to-release flow for a rotated VERBATIM group lives in
   * useFrameArrange, which calls this once confirmed. */
  function ungroup(): void {
    const groups = selectedIds().filter((id) => frameById(id)?.kind === "group");
    if (groups.length === 0) return;
    run("Ungroup", (model, index) => {
      const promoted: string[] = [];
      for (const id of groups) {
        promoted.push(...(ungroupFrame(model, index, id) ?? []));
      }
      return promoted;
    });
  }

  /** Returns true when it handled the key (the caller then stops). */
  function onKeydown(event: KeyboardEvent): boolean {
    if (deps.isEditing()) return false;
    if (deps.selection.selectedFrameIds.value.length === 0) return false;
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
    rotatingDeg,
    onFramePointerDown,
    onHandlePointerDown,
    onRotateHandlePointerDown,
    onAdjustPointerDown,
    onAdjustKeydown,
    onTableBorderPointerDown,
    onKeydown,
    remove,
    duplicate,
    arrange,
    ungroup,
  };
}
