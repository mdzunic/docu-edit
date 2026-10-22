// Composes SPEC-081's slide editing for PresentationView so the view stays a
// thin wiring hub (PresentationView.size.spec.ts): frame selection
// (useFrameSelection) + the text-editing session (useSlideTextEditing) + the
// stage's global keyboard (undo/redo, slide navigation, Tab/Shift+Tab frame
// cycling, Enter-to-edit, Escape) + the toolbar's recent-colour memory.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFill } from "../../lib/presentation/model";
import { findShortcut, isOutsideDocumentInput } from "../../lib/shortcuts";
import { loadRecentColors, pushRecentColor } from "../../lib/core/colors";
import type { ModelColor } from "../../lib/core/theme";
import { setFrameFill, setFrameStroke } from "../../lib/presentation/frames";
import { advanceKeys, backKeys } from "../../lib/presentation/playbackKeys";
import type { FrameStroke } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { useSlideTextEditing, type CaretPoint } from "./useSlideTextEditing";
import { useFrameInsertion } from "./useFrameInsertion";
import { useFrameManipulation } from "./useFrameManipulation";
import { useFrameArrange } from "./useFrameArrange";
import type { ResizeHandle } from "./stageMath";

function isTypingTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
  );
}

export interface PresentationEditingDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  /** The current slide's index (the frame commands' target). */
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** Marks `model.media` changed so `commit` re-slices it (SPEC-095 AC-13). */
  bumpMedia: () => void;
  /** Re-render the stage without persisting — live drag/resize preview. */
  refresh: () => void;
  undo: () => void;
  redo: () => void;
  /** Move the current slide by `delta` (PageUp/PageDown). */
  navigate: (delta: number) => void;
  /** Surface a broken/unreadable image (AC-9). */
  reportError: (message: string) => void;
}

export function usePresentationEditing(deps: PresentationEditingDeps) {
  const selection = useFrameSelection({ currentSlide: deps.currentSlide });
  const editing = useSlideTextEditing({
    model: deps.model,
    currentSlide: deps.currentSlide,
    selection,
    record: deps.record,
    commit: deps.commit,
  });

  // The stage surface element, registered by SlideStage, shared by the
  // insertion + manipulation gestures for client→cm conversion.
  const surfaceEl = ref<HTMLElement | null>(null);
  const surface = (): HTMLElement | null => surfaceEl.value;

  const insertion = useFrameInsertion({
    model: deps.model,
    currentSlide: deps.currentSlide,
    currentIndex: deps.currentIndex,
    selection,
    surface,
    record: deps.record,
    commit: deps.commit,
    bumpMedia: deps.bumpMedia,
    startEditing: (id) => editing.startEditing(id),
    reportError: deps.reportError,
  });

  const manipulation = useFrameManipulation({
    model: deps.model,
    currentSlide: deps.currentSlide,
    currentIndex: deps.currentIndex,
    selection,
    isEditing: () => editing.editingFrameId.value !== null,
    surface,
    record: deps.record,
    refresh: deps.refresh,
    commit: deps.commit,
  });

  // Marquee + align/distribute + group/release + Position & Size (SPEC-096),
  // re-exported as stageEdit.arrange so the view gains only template hookups.
  const arrange = useFrameArrange({
    model: deps.model,
    currentSlide: deps.currentSlide,
    currentIndex: deps.currentIndex,
    selection,
    surface,
    record: deps.record,
    refresh: deps.refresh,
    commit: deps.commit,
  });

  function registerSurface(el: HTMLElement | null): void {
    surfaceEl.value = el;
  }

  // --- style popover commands (one undo entry per commit, AC-42) ------------
  // SPEC-096: fill/stroke apply to EVERY selected frame that supports them
  // (the commands skip unsupported kinds by returning null), still one entry.
  function withSelected(label: string, apply: (model: PresentationModel, index: number, id: string) => void): void {
    const model = deps.model.value;
    const ids = selection.selectedFrameIds.value;
    if (!model || ids.length === 0) return;
    deps.record(label);
    for (const id of ids) {
      apply(model, deps.currentIndex.value, id);
    }
    deps.commit();
  }
  function setFill(fill: SlideFill): void {
    withSelected("Fill", (model, index, id) => setFrameFill(model, index, id, fill));
  }
  function setStroke(stroke: FrameStroke | null): void {
    withSelected("Stroke", (model, index, id) => setFrameStroke(model, index, id, stroke));
  }

  function onSurfaceDraw(event: PointerEvent): void {
    insertion.onSurfacePointerDown(event);
  }
  function onFramePointerDown(id: string, event: PointerEvent): void {
    manipulation.onFramePointerDown(id, event);
  }
  function onHandlePointerDown(id: string, handle: ResizeHandle, event: PointerEvent): void {
    manipulation.onHandlePointerDown(id, handle, event);
  }
  function onTableBorderPointerDown(id: string, axis: "col" | "row", index: number, event: PointerEvent): void {
    manipulation.onTableBorderPointerDown(id, axis, index, event);
  }

  const recentColors = ref<string[]>(loadRecentColors());

  // Double-clicking a chart frame opens its data sheet instead of a text
  // session (SPEC-092 AC-14). useSlideCharts registers the handler after the
  // view builds it (it needs this composable's selection), so onStartEdit can
  // route a chart's dblclick to it by frame kind without a construction cycle.
  let chartEditHandler: ((id: string) => void) | null = null;
  function setChartEditHandler(fn: (id: string) => void): void {
    chartEditHandler = fn;
  }

  function onFrameSelect(id: string, event?: MouseEvent): void {
    // Shift+click membership is handled by the pointerdown path (toggle) —
    // the trailing mousedown must not collapse it (SPEC-096 AC-2).
    if (event?.shiftKey) return;
    if (editing.editingFrameId.value && editing.editingFrameId.value !== id) {
      editing.stopEditing();
    }
    // A plain press on a frame already inside a multi-selection keeps the set
    // (so a drag moves all of it); anything else selects that one frame.
    if (!selection.isSelected(id)) {
      selection.select(id);
    }
  }
  function onStartEdit(id: string, point?: CaretPoint): void {
    // Enter / double-click collapses a multi-selection to the one frame
    // before starting its session (SPEC-096 AC-22).
    selection.select(id);
    const frame = deps.currentSlide.value?.frames.find((f) => f.id === id);
    if (frame?.kind === "chart") {
      chartEditHandler?.(id);
      return;
    }
    editing.startEditing(id, false, point);
  }
  function onStartEditCell(id: string, row: number, col: number, point?: CaretPoint): void {
    editing.startEditingCell(id, row, col, point);
  }
  function onCellSelect(id: string, row: number, col: number): void {
    if (editing.editingFrameId.value && editing.editingFrameId.value !== id) {
      editing.stopEditing();
    }
    editing.focusCell(id, row, col);
  }
  function onSurfaceDeselect(): void {
    // A press on empty stage stops any text session immediately; the
    // SELECTION only clears on release, through the marquee gesture (a
    // sub-threshold release is the plain deselecting click — SPEC-096).
    if (editing.editingFrameId.value) editing.stopEditing();
  }
  /** Pointer-down on empty stage while no tool is armed: the marquee. */
  function onSurfaceMarquee(event: PointerEvent): void {
    arrange.onSurfacePointerDown(event);
  }

  function setColor(color: ModelColor): void {
    editing.setColor(color);
    // A theme swatch stores the reference (SPEC-094 AC-49); only literals
    // enter the recents strip, which stays hex-only.
    if (typeof color === "string") {
      recentColors.value = pushRecentColor(color, recentColors.value);
    }
  }

  function onGlobalKeydown(event: KeyboardEvent): void {
    // Undo/redo drive the presentation stack (SPEC-080), ignoring dialogs/inputs.
    const shortcut = findShortcut(event);
    if ((shortcut?.id === "edit.undo" || shortcut?.id === "edit.redo") && !isOutsideDocumentInput(event.target)) {
      event.preventDefault();
      if (shortcut.id === "edit.undo") deps.undo();
      else deps.redo();
      return;
    }
    // The editing session owns its own keys (the host's keydown listener).
    if (editing.editingFrameId.value) return;
    if (isTypingTarget(event.target)) return;

    const mod = event.metaKey || event.ctrlKey;
    // Cmd/Ctrl+A selects every frame on the slide (SPEC-096 AC-3).
    if (mod && !event.shiftKey && (event.key === "a" || event.key === "A")) {
      event.preventDefault();
      selection.selectAll();
      return;
    }
    // Cmd/Ctrl+Shift+G routes through arrange so releasing a rotated
    // VERBATIM group parks in the pendingRelease confirm (SPEC-096 AC-42).
    if (mod && event.shiftKey && (event.key === "g" || event.key === "G") && selection.selectedFrameIds.value.length > 0) {
      event.preventDefault();
      arrange.ungroup();
      return;
    }

    // SPEC-082 manipulation keys (nudge / delete / duplicate / ungroup).
    if (manipulation.onKeydown(event)) return;

    if (event.key === "Tab" && selection.selectedFrameId.value !== null) {
      event.preventDefault();
      selection.cycle(event.shiftKey ? -1 : 1);
      return;
    }
    if (event.key === "Enter" && selection.selectedFrameId.value !== null) {
      const frame = selection.anchorFrame.value;
      if (frame && frame.kind === "text") {
        event.preventDefault();
        // Collapse a multi-selection to the one frame before its session
        // starts (SPEC-096 AC-22).
        selection.select(frame.id);
        editing.startEditing(frame.id);
      } else if (frame && frame.kind === "table") {
        event.preventDefault();
        selection.select(frame.id);
        const cell = editing.activeCell.value;
        editing.startEditingCell(frame.id, cell?.row ?? 0, cell?.col ?? 0);
      }
      return;
    }
    if (event.key === "Escape") {
      if (insertion.isArmed.value) {
        event.preventDefault();
        insertion.disarm();
        return;
      }
      if (selection.selectedFrameId.value !== null) {
        event.preventDefault();
        selection.clear();
        return;
      }
    }
    // PageUp/PageDown stay editor-side (SPEC-084 keeps the richer map in the
    // play view), but read the shared playback table so the two maps cannot
    // drift — if PageDown/PageUp ever left advanceKeys/backKeys, this breaks.
    if (event.key === "PageDown" && advanceKeys.includes(event.key)) {
      event.preventDefault();
      deps.navigate(1);
    } else if (event.key === "PageUp" && backKeys.includes(event.key)) {
      event.preventDefault();
      deps.navigate(-1);
    }
  }

  return {
    selection,
    editing,
    insertion,
    manipulation,
    arrange,
    recentColors: computed(() => recentColors.value),
    onFrameSelect,
    onStartEdit,
    onStartEditCell,
    onCellSelect,
    onSurfaceDeselect,
    onSurfaceMarquee,
    onSurfaceDraw,
    onFramePointerDown,
    onHandlePointerDown,
    onTableBorderPointerDown,
    registerSurface,
    setColor,
    setFill,
    setStroke,
    setChartEditHandler,
    onGlobalKeydown,
  };
}
