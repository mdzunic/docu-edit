// The motion-path editor (SPEC-105): the stage overlay the Animations panel's
// "Edit path" arms for the selected motion-path row — its resolved path drawn
// dashed from the frame's centre, a ghost of the frame where the path ends,
// and a draggable endpoint. The useAnimations controller creates this itself
// (PresentationView gains no wiring); AnimationStageLayer renders `overlay`
// and forwards the endpoint's pointer/keyboard events here.
//
// The drag owns the pointer (capture), converts each move through
// stageMath's `clientToCm` against the overlay's on-screen rect, and refreshes
// a live DRAFT path the overlay draws — the model is untouched until
// pointer-up, which records one undo entry ("Edit motion path"), writes the
// path and commits. Path coordinates are slide fractions from the frame's
// centre; the drag snaps to what the model can express (four directions, a
// distance of 5–100%) — the endpoint follows the pointer while the dashed path
// and the ghost show the snapped result. A pressed endpoint takes focus (its
// press's default is prevented), and the arrows it handles stop there, so
// they never also nudge the selected frame through the editor's window keys.
// An imported custom path is drawn read-only: its endpoint refuses the drag
// and the arrows, and only choosing another preset in the panel replaces it.
import { computed, ref, shallowRef, type ComputedRef, type Ref } from "vue";
import type { MotionPath, PresentationModel, Slide, SlideAnimation } from "../../lib/presentation/model";
import {
  motionPathDirection,
  motionPathDistance,
  motionPathFromHandle,
  motionPathHandle,
  pathDInPixels,
  pathEndpoints,
  resolveMotionPath,
  type PathPoint,
} from "../../lib/presentation/motionPaths";
import { updateAnimation } from "../../lib/presentation/slideAnimations";
import { cmToPx } from "../../lib/core/units";
import { clientToCm } from "./stageMath";

export interface MotionPathEditDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  /** The Animations panel's selected row. */
  selectedIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
}

/** What the overlay draws, in unscaled surface px (the stage's viewBox). */
export interface MotionPathOverlay {
  widthPx: number;
  heightPx: number;
  /** The dashed path, from the frame's centre. */
  d: string;
  /** The draggable endpoint. */
  handle: { x: number; y: number };
  /** The frame's box where the path leaves it. */
  ghost: { x: number; y: number; width: number; height: number };
  readOnly: boolean;
}

/** The arrow keys' nudge: one percent of the slide dimension. */
const NUDGE = 0.01;

export function useMotionPathEdit(deps: MotionPathEditDeps) {
  const armed = ref(false);
  /** The path the live drag would write — drawn instead of the stored one. */
  const draft = shallowRef<MotionPath | null>(null);
  /** Where the live drag's pointer is (slide fractions from the frame's
   * centre) — the endpoint follows it; the path shows the snapped draft. */
  const pointer = shallowRef<PathPoint | null>(null);
  const dragging = ref(false);
  // Bumped by refresh() so the overlay re-derives mid-drag.
  const revision = ref(0);
  let surface: DOMRect | null = null;

  // A fresh copy per read: the model's rows are edited in place, so handing
  // out the row itself would leave every computed below on its old fields.
  const selected = computed<SlideAnimation | null>(() => {
    void revision.value;
    const row = deps.currentSlide.value?.animations?.[deps.selectedIndex.value];
    return row ? { ...row } : null;
  });
  const isPathRow = computed(() => selected.value?.effect === "motionPath" && !!selected.value.path);
  const readOnly = computed(() => selected.value?.path?.preset === "custom");
  const frame = computed(() => {
    const id = selected.value?.frameId;
    return id ? (deps.currentSlide.value?.frames.find((f) => f.id === id) ?? null) : null;
  });

  function deckSize(): { widthCm: number; heightCm: number } {
    const model = deps.model.value;
    return { widthCm: model?.slideWidthCm ?? 33.867, heightCm: model?.slideHeightCm ?? 19.05 };
  }

  /** The overlay geometry, or null when nothing is armed on a path row. */
  const overlay = computed<MotionPathOverlay | null>(() => {
    void revision.value;
    const path = draft.value ?? selected.value?.path;
    const target = frame.value;
    const commands = path ? resolveMotionPath(path) : undefined;
    if (!armed.value || !isPathRow.value || !path || !target || !commands) {
      return null;
    }
    const { widthCm, heightCm } = deckSize();
    const widthPx = cmToPx(widthCm);
    const heightPx = cmToPx(heightCm);
    const origin = { x: cmToPx(target.xCm + target.widthCm / 2), y: cmToPx(target.yCm + target.heightCm / 2) };
    const end = pathEndpoints(commands).end;
    const handle = pointer.value ?? (path.preset === "custom" ? end : motionPathHandle(path));
    return {
      widthPx,
      heightPx,
      d: pathDInPixels(commands, origin, widthPx, heightPx),
      handle: { x: origin.x + handle.x * widthPx, y: origin.y + handle.y * heightPx },
      ghost: {
        x: cmToPx(target.xCm) + end.x * widthPx,
        y: cmToPx(target.yCm) + end.y * heightPx,
        width: cmToPx(target.widthCm),
        height: cmToPx(target.heightCm),
      },
      readOnly: path.preset === "custom",
    };
  });

  /** Re-derives the overlay (the live drag calls it on every move). */
  function refresh(): void {
    revision.value += 1;
  }

  function toggle(): void {
    armed.value = !armed.value;
    draft.value = null;
    pointer.value = null;
  }
  function disarm(): void {
    armed.value = false;
    draft.value = null;
    pointer.value = null;
    dragging.value = false;
  }

  /** Writes `path` onto the selected row as one undo entry. */
  function writePath(path: MotionPath): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    const current = selected.value?.path;
    if (!slide || !model || !current || JSON.stringify(path) === JSON.stringify(current)) {
      return;
    }
    deps.record("Edit motion path");
    updateAnimation(slide, deps.selectedIndex.value, { path }, model.format);
    deps.commit();
  }

  /** Starts the endpoint drag: captures the pointer against the overlay's
   * on-screen rect. Refused (false) for a read-only custom path. */
  function onEndpointPointerDown(event: PointerEvent, overlayEl: Element | null): boolean {
    if (!isPathRow.value || readOnly.value || !overlayEl) {
      return false;
    }
    event.preventDefault();
    const target = event.target;
    if (target instanceof Element && typeof target.setPointerCapture === "function") {
      try {
        target.setPointerCapture(event.pointerId);
      } catch {
        // A synthetic pointer has nothing to capture.
      }
    }
    // The prevented default took the browser's focus with it.
    if (target instanceof SVGElement || target instanceof HTMLElement) {
      target.focus({ preventScroll: true });
    }
    surface = overlayEl.getBoundingClientRect();
    dragging.value = true;
    draft.value = selected.value?.path ?? null;
    return true;
  }

  /** The live drag: the pointer, in slide fractions from the frame's centre,
   * becomes the draft path. */
  function onEndpointPointerMove(event: PointerEvent): void {
    const path = selected.value?.path;
    const target = frame.value;
    if (!dragging.value || !surface || !path || !target) {
      return;
    }
    const { widthCm, heightCm } = deckSize();
    const { xCm, yCm } = clientToCm(event.clientX, event.clientY, surface, widthCm, heightCm);
    const point = {
      x: (xCm - (target.xCm + target.widthCm / 2)) / widthCm,
      y: (yCm - (target.yCm + target.heightCm / 2)) / heightCm,
    };
    pointer.value = point;
    draft.value = motionPathFromHandle(path, point);
    refresh();
  }

  /** Ends the drag: one record + commit of the dragged path. */
  function onEndpointPointerUp(event: PointerEvent): void {
    if (!dragging.value) {
      return;
    }
    dragging.value = false;
    const target = event.target;
    if (target instanceof Element && typeof target.releasePointerCapture === "function") {
      try {
        target.releasePointerCapture(event.pointerId);
      } catch {
        // Nothing captured.
      }
    }
    const path = draft.value;
    draft.value = null;
    pointer.value = null;
    surface = null;
    if (path) {
      writePath(path);
    }
    refresh();
  }

  /** Keyboard: an arrow nudges the endpoint one percent that way — one undo
   * entry per key press. */
  function onEndpointKeydown(event: KeyboardEvent): void {
    const path = selected.value?.path;
    if (!path || !event.key.startsWith("Arrow")) {
      return;
    }
    // The endpoint's arrows, even a read-only one's, are never the frame's.
    event.preventDefault();
    event.stopPropagation();
    if (readOnly.value) {
      return;
    }
    const handle = motionPathHandle(path);
    const dx = event.key === "ArrowLeft" ? -NUDGE : event.key === "ArrowRight" ? NUDGE : 0;
    const dy = event.key === "ArrowUp" ? -NUDGE : event.key === "ArrowDown" ? NUDGE : 0;
    writePath(motionPathFromHandle(path, { x: handle.x + dx, y: handle.y + dy }));
    refresh();
  }

  /** The endpoint's accessible description (the live draft's while dragging). */
  const endpointLabel = computed(() => {
    const path = draft.value ?? selected.value?.path;
    if (!path) return "";
    if (path.preset === "custom") return "Custom path end (read-only)";
    return `Path end: ${motionPathDirection(path)}, ${Math.round(motionPathDistance(path) * 100)}%`;
  });

  return {
    armed,
    dragging,
    draft,
    readOnly,
    isPathRow,
    overlay,
    endpointLabel,
    toggle,
    disarm,
    refresh,
    onEndpointPointerDown,
    onEndpointPointerMove,
    onEndpointPointerUp,
    onEndpointKeydown,
  };
}
