// Arranging slide objects (SPEC-096): the marquee selection gesture, the
// Align/Distribute menu commands (over core/arrange's pure rect algebra),
// authored Group creation, confirm-to-release of a rotated verbatim group,
// and the "Position and size…" dialog's command. Constructed inside
// usePresentationEditing and re-exported as `stageEdit.arrange`, so
// PresentationView gains no new use*() call — only template hookups.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { GroupFrame, PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import {
  FRAME_MIN_ON_SLIDE_CM,
  MIN_FRAME_SIZE_CM,
  frameRotatedBox,
  groupFrames,
  resizeFrame,
  setFramePosition,
  setFrameRotation,
  ungroupFrame,
  type FrameRect,
} from "../../lib/presentation/frames";
import { alignRects, distributeRects, type AlignEdge, type AlignReference, type Rect } from "../../lib/core/arrange";
import { normalizeRotationDeg } from "../../lib/core/shapes";
import { asMeasurementUnit } from "../../lib/core/units";
import { usePreferences } from "../../composables/usePreferences";
import type { FrameSelection } from "./useFrameSelection";
import { clientToCm, DRAG_ENGAGE_CM } from "./stageMath";

export interface FrameArrangeDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  selection: FrameSelection;
  /** The stage surface element, for client→cm conversion. */
  surface: () => HTMLElement | null;
  record: (label?: string) => void;
  /** Re-render the stage without persisting. */
  refresh: () => void;
  /** Persist write-through + re-render (command end). */
  commit: () => void;
}

const ALIGN_LABELS: Record<AlignEdge, string> = {
  left: "Align left",
  centerX: "Align centre",
  right: "Align right",
  top: "Align top",
  middleY: "Align middle",
  bottom: "Align bottom",
};

/** Whether two cm boxes overlap (marquee intersection, on rotated boxes). */
function intersects(a: Rect, b: Rect): boolean {
  return a.xCm < b.xCm + b.widthCm && b.xCm < a.xCm + a.widthCm && a.yCm < b.yCm + b.heightCm && b.yCm < a.yCm + a.heightCm;
}

/** What the Position & Size dialog applies (already parsed to cm/degrees). */
export interface PositionPatch {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
  rotationDeg: number;
}

export function useFrameArrange(deps: FrameArrangeDeps) {
  /** The live marquee box (slide cm) while a drag runs, null otherwise. */
  const marquee = ref<FrameRect | null>(null);
  /** Group ids parked awaiting the lossy-release confirm (AC-42). */
  const pendingRelease = ref<string[] | null>(null);
  /** One-line notice (e.g. animations dropped on grouping), aria-live. */
  const notice = ref<string | null>(null);
  /** The Align menu's reference select. */
  const alignReference = ref<AlignReference>("selection");
  /** Position & Size dialog visibility. */
  const positionOpen = ref(false);
  /** The measurement unit the dialog's fields read/write in. */
  const { preferences } = usePreferences();
  const unit = computed(() => asMeasurementUnit(preferences.measurementUnit));

  const selectedCount = computed(() => deps.selection.selectedFrameIds.value.length);
  const canDistribute = computed(() => selectedCount.value >= 3);
  const distributeReason = computed(() =>
    canDistribute.value ? "" : "Select at least three objects to distribute",
  );

  function slideFrames(): SlideFrame[] {
    return deps.currentSlide.value?.frames ?? [];
  }

  function selectedFrames(): SlideFrame[] {
    const ids = deps.selection.selectedFrameIds.value;
    return slideFrames().filter((frame) => ids.includes(frame.id));
  }

  function pointerCm(event: PointerEvent): { xCm: number; yCm: number } | null {
    const rect = deps.surface()?.getBoundingClientRect();
    const model = deps.model.value;
    if (!rect || !model) return null;
    return clientToCm(event.clientX, event.clientY, rect, model.slideWidthCm, model.slideHeightCm);
  }

  // --- marquee selection (AC-6..10) -----------------------------------------

  /** A pointer-down on empty stage: past DRAG_ENGAGE_CM it becomes a marquee
   * (release selects every top-level frame whose rotatedBounds box intersects
   * it; Shift adds to the existing selection); under the threshold the
   * release is the plain deselecting click (SPEC-082 unchanged); Escape
   * mid-drag abandons it and changes nothing. */
  function onSurfacePointerDown(event: PointerEvent): void {
    const start = pointerCm(event);
    if (!start) {
      return;
    }
    const additive = event.shiftKey;
    let engaged = false;

    const cleanup = (): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("keydown", onKey, true);
      marquee.value = null;
    };
    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now) return;
      if (!engaged) {
        if (Math.abs(now.xCm - start.xCm) < DRAG_ENGAGE_CM && Math.abs(now.yCm - start.yCm) < DRAG_ENGAGE_CM) {
          return;
        }
        engaged = true;
      }
      marquee.value = {
        xCm: Math.min(start.xCm, now.xCm),
        yCm: Math.min(start.yCm, now.yCm),
        widthCm: Math.abs(now.xCm - start.xCm),
        heightCm: Math.abs(now.yCm - start.yCm),
      };
    };
    const onKey = (key: KeyboardEvent): void => {
      if (key.key !== "Escape") return;
      key.stopPropagation();
      cleanup(); // abandoned — nothing changes
    };
    const onUp = (): void => {
      const box = marquee.value;
      cleanup();
      if (!engaged || !box) {
        deps.selection.clear(); // the plain deselecting click
        return;
      }
      const hit = slideFrames()
        .filter((frame) => intersects(frameRotatedBox(frame), box))
        .map((frame) => frame.id);
      deps.selection.selectMany(hit, additive);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("keydown", onKey, true);
  }

  // --- align / distribute (AC-53..56) ---------------------------------------

  /** The effective reference: with exactly one frame selected, Slide is the
   * only sensible (and only offered) reference. */
  function effectiveReference(): AlignReference {
    return selectedCount.value <= 1 ? "container" : alignReference.value;
  }

  function applyRects(label: string, frames: SlideFrame[], rects: Rect[]): void {
    const model = deps.model.value;
    if (!model) return;
    notice.value = null;
    deps.record(label);
    frames.forEach((frame, i) => {
      // The algebra ran on rotated boxes — translate back to frame origins.
      const box = frameRotatedBox(frame);
      const pad = { x: frame.xCm - box.xCm, y: frame.yCm - box.yCm };
      setFramePosition(model, deps.currentIndex.value, frame.id, rects[i].xCm + pad.x, rects[i].yCm + pad.y);
    });
    deps.commit();
  }

  function align(edge: AlignEdge): void {
    const model = deps.model.value;
    const frames = selectedFrames();
    if (!model || frames.length === 0) return;
    const container: Rect = { xCm: 0, yCm: 0, widthCm: model.slideWidthCm, heightCm: model.slideHeightCm };
    const rects = frames.map(frameRotatedBox);
    applyRects(ALIGN_LABELS[edge], frames, alignRects(rects, edge, effectiveReference(), container));
  }

  function distribute(axis: "x" | "y"): void {
    const frames = selectedFrames();
    if (frames.length < 3) return;
    const rects = frames.map(frameRotatedBox);
    const label = axis === "x" ? "Distribute horizontally" : "Distribute vertically";
    applyRects(label, frames, distributeRects(rects, axis));
  }

  // --- group / release (AC-40..44) ------------------------------------------

  function group(): void {
    const model = deps.model.value;
    const slide = deps.currentSlide.value;
    const ids = deps.selection.selectedFrameIds.value;
    if (!model || !slide || ids.length < 2) return;
    const animated = (slide.animations ?? []).some((a) => ids.includes(a.frameId));
    notice.value = null;
    deps.record("Group");
    const groupId = groupFrames(model, deps.currentIndex.value, ids);
    deps.commit();
    if (groupId) {
      deps.selection.select(groupId);
    }
    if (animated) {
      // Out of scope: a group and its descendants stay unanimatable, so
      // grouping animated frames drops those animations — the one-line notice.
      notice.value = "Grouping removed this selection's animations.";
    }
  }

  function selectedGroups(): GroupFrame[] {
    return selectedFrames().filter((frame): frame is GroupFrame => frame.kind === "group");
  }

  /** Releases every selected group in one undo entry (AC-20). A VERBATIM
   * group carrying a non-zero rotation is lossy to release (its children
   * were flattened without the rotation), so it parks in the pendingRelease
   * confirm; a group with rotation 0 releases silently. */
  function ungroup(): void {
    const groups = selectedGroups();
    if (groups.length === 0) return;
    const lossy = groups.some(
      (g) => g.xml !== undefined && normalizeRotationDeg(g.rotationDeg ?? 0) !== 0,
    );
    if (lossy) {
      pendingRelease.value = groups.map((g) => g.id);
      return;
    }
    release(groups.map((g) => g.id));
  }

  function release(ids: string[]): void {
    const model = deps.model.value;
    if (!model) return;
    notice.value = null;
    deps.record("Ungroup");
    const promoted: string[] = [];
    for (const id of ids) {
      promoted.push(...(ungroupFrame(model, deps.currentIndex.value, id) ?? []));
    }
    deps.commit();
    deps.selection.selectMany(promoted);
  }

  function confirmRelease(): void {
    const ids = pendingRelease.value;
    pendingRelease.value = null;
    if (ids) release(ids);
  }

  function cancelRelease(): void {
    pendingRelease.value = null; // changes nothing
  }

  // --- Position & Size dialog (AC-62..66) -----------------------------------

  function openPosition(): void {
    if (deps.selection.anchorFrame.value) {
      positionOpen.value = true;
    }
  }

  function closePosition(): void {
    positionOpen.value = false;
  }

  /** One field commit from the dialog = one undo entry. The commands clamp:
   * sizes to MIN_FRAME_SIZE_CM, position to FRAME_MIN_ON_SLIDE_CM on-slide. */
  function applyPosition(patch: PositionPatch): void {
    const model = deps.model.value;
    const frame = deps.selection.anchorFrame.value;
    if (!model || !frame) return;
    deps.record("Position and size");
    const index = deps.currentIndex.value;
    const widthCm = Math.max(patch.widthCm, MIN_FRAME_SIZE_CM);
    const heightCm = Math.max(patch.heightCm, MIN_FRAME_SIZE_CM);
    if (widthCm !== frame.widthCm || heightCm !== frame.heightCm) {
      resizeFrame(model, index, frame.id, { xCm: patch.xCm, yCm: patch.yCm, widthCm, heightCm });
    }
    setFramePosition(model, index, frame.id, patch.xCm, patch.yCm);
    if (normalizeRotationDeg(patch.rotationDeg) !== normalizeRotationDeg(frame.rotationDeg ?? 0)) {
      setFrameRotation(model, index, frame.id, patch.rotationDeg);
    }
    deps.commit();
  }

  return {
    marquee,
    notice,
    pendingRelease,
    alignReference,
    positionOpen,
    unit,
    selectedCount,
    canDistribute,
    distributeReason,
    onSurfacePointerDown,
    align,
    distribute,
    group,
    ungroup,
    confirmRelease,
    cancelRelease,
    openPosition,
    closePosition,
    applyPosition,
    clampMin: { sizeCm: MIN_FRAME_SIZE_CM, onSlideCm: FRAME_MIN_ON_SLIDE_CM },
  };
}
