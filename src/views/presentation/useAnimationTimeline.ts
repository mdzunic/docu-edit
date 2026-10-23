// The Animations panel's timeline view (SPEC-105), split out of
// useAnimations (which creates it and spreads it into its controller): the
// persisted open state (`presentationAnimationTimelineOpen`), the current
// slide's timeline (geometry in components/presentation/animationTimeline.ts)
// with a live edge drag applied, a bar selecting its row, and the edge edits —
// a left-edge (delay) or right-edge (duration) pointer drag recorded once on
// pointer-up, or a Left/Right arrow step of 100ms, one undo entry per key. A
// pressed edge takes focus, so the arrows then step the edge just dragged, and
// an arrow it handles stops there (the editor's window keys would otherwise
// also nudge the selected frame).
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { Slide, SlideAnimation } from "../../lib/presentation/model";
import { effectiveAnimationDelayMs, effectiveAnimationDurationMs } from "../../lib/presentation/animations";
import type { AnimationPatch } from "../../lib/presentation/slideAnimations";
import {
  draggedDelayMs,
  draggedDurationMs,
  msFromDx,
  timelineModel,
  TIMELINE_STEP_MS,
  type TimelineBar,
} from "../../components/presentation/animationTimeline";
import { usePreferences } from "../../composables/usePreferences";

/** An edge of a timeline bar: its left (the delay) or right (the duration). */
export type TimelineEdge = "delay" | "duration";

/** The delay/duration an edge edit of `deltaMs` leaves (clamped, 100ms steps). */
function edgeValue(bar: TimelineBar, edge: TimelineEdge, deltaMs: number): number {
  return edge === "delay" ? draggedDelayMs(bar.delayMs, deltaMs) : draggedDurationMs(bar.singleMs, deltaMs, bar.passes);
}

export interface AnimationTimelineDeps {
  currentSlide: ComputedRef<Slide | null>;
  animations: ComputedRef<SlideAnimation[]>;
  selectedIndex: Ref<number>;
  selectRow: (index: number) => void;
  /** One edit of the selected row: record(label) → update → commit. */
  patch: (label: string, change: AnimationPatch) => void;
}

export function useAnimationTimeline(deps: AnimationTimelineDeps) {
  const { preferences, setPreference } = usePreferences();
  const { animations, selectedIndex, selectRow, patch } = deps;
  function at(index: number): SlideAnimation | undefined {
    return animations.value[index];
  }

  const timelineOpen = computed(() => preferences.presentationAnimationTimelineOpen);
  function toggleTimeline(): void {
    setPreference("presentationAnimationTimelineOpen", !timelineOpen.value);
  }
  /** A live edge drag: the row and the value it would write on pointer-up. */
  const timelineDraft = ref<{ animationIndex: number; edge: TimelineEdge; valueMs: number } | null>(null);
  let edgeDrag: { bar: TimelineBar; edge: TimelineEdge; startX: number } | null = null;
  /** The timeline of the current slide, a live drag's value applied. */
  const timeline = computed(() => {
    const slide = deps.currentSlide.value;
    if (!slide) return null;
    const draftValue = timelineDraft.value;
    const shown = draftValue
      ? {
          ...slide,
          animations: animations.value.map((a, i) =>
            i === draftValue.animationIndex
              ? { ...a, [draftValue.edge === "delay" ? "delayMs" : "durationMs"]: draftValue.valueMs }
              : a,
          ),
        }
      : { ...slide, animations: animations.value };
    return timelineModel(shown);
  });
  function writeEdge(animationIndex: number, edge: TimelineEdge, valueMs: number): void {
    const a = at(animationIndex);
    if (!a) return;
    const current = edge === "delay" ? effectiveAnimationDelayMs(a) : effectiveAnimationDurationMs(a);
    if (valueMs === current) return;
    selectedIndex.value = animationIndex;
    patch(
      edge === "delay" ? "Change animation delay" : "Change animation duration",
      edge === "delay" ? { delayMs: valueMs } : { durationMs: valueMs },
    );
  }
  /** A bar drives the same selection as its row. */
  function selectBar(bar: TimelineBar): void {
    selectRow(bar.animationIndex);
  }
  function startEdgeDrag(bar: TimelineBar, edge: TimelineEdge, event: PointerEvent): void {
    selectBar(bar);
    const target = event.target;
    if (target instanceof Element && typeof target.setPointerCapture === "function") {
      try {
        target.setPointerCapture(event.pointerId);
      } catch {
        // A synthetic pointer has nothing to capture.
      }
    }
    // The press's default (and with it the browser's focus) is prevented.
    if (target instanceof HTMLElement) {
      target.focus({ preventScroll: true });
    }
    edgeDrag = { bar, edge, startX: event.clientX };
    timelineDraft.value = { animationIndex: bar.animationIndex, edge, valueMs: edgeValue(bar, edge, 0) };
  }
  function moveEdgeDrag(event: PointerEvent): void {
    if (!edgeDrag) return;
    const { bar, edge, startX } = edgeDrag;
    timelineDraft.value = { animationIndex: bar.animationIndex, edge, valueMs: edgeValue(bar, edge, msFromDx(event.clientX - startX)) };
  }
  /** Pointer-up ends the drag: one undo entry for the whole gesture (none
   * when the edge ends where it started). */
  function endEdgeDrag(): void {
    const draftValue = timelineDraft.value;
    edgeDrag = null;
    timelineDraft.value = null;
    if (draftValue) writeEdge(draftValue.animationIndex, draftValue.edge, draftValue.valueMs);
  }
  /** A cancelled pointer (the browser took it) discards the drag. */
  function cancelEdgeDrag(): void {
    edgeDrag = null;
    timelineDraft.value = null;
  }
  /** Left/Right on an edge: ∓100ms, one undo entry per key press. */
  function edgeKeydown(bar: TimelineBar, edge: TimelineEdge, event: KeyboardEvent): void {
    const delta = event.key === "ArrowLeft" ? -TIMELINE_STEP_MS : event.key === "ArrowRight" ? TIMELINE_STEP_MS : 0;
    if (delta === 0) return;
    event.preventDefault();
    event.stopPropagation();
    writeEdge(bar.animationIndex, edge, edgeValue(bar, edge, edge === "duration" ? delta * bar.passes : delta));
  }

  return {
    timelineOpen,
    toggleTimeline,
    timeline,
    timelineDraft,
    selectBar,
    startEdgeDrag,
    moveEdgeDrag,
    endEdgeDrag,
    cancelEdgeDrag,
    edgeKeydown,
  };
}
