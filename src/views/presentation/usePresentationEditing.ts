// Composes SPEC-081's slide editing for PresentationView so the view stays a
// thin wiring hub (PresentationView.size.spec.ts): frame selection
// (useFrameSelection) + the text-editing session (useSlideTextEditing) + the
// stage's global keyboard (undo/redo, slide navigation, Tab/Shift+Tab frame
// cycling, Enter-to-edit, Escape) + the toolbar's recent-colour memory.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFill } from "../../lib/presentation/model";
import { findShortcut, isOutsideDocumentInput } from "../../lib/shortcuts";
import { loadRecentColors, pushRecentColor } from "../../lib/core/colors";
import { setFrameFill, setFrameStroke } from "../../lib/presentation/frames";
import { advanceKeys, backKeys } from "../../lib/presentation/playbackKeys";
import type { FrameStroke } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { useSlideTextEditing, type CaretPoint } from "./useSlideTextEditing";
import { useFrameInsertion } from "./useFrameInsertion";
import { useFrameManipulation } from "./useFrameManipulation";
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

  function registerSurface(el: HTMLElement | null): void {
    surfaceEl.value = el;
  }

  // --- style popover commands (one undo entry per commit, AC-42) ------------
  function withSelected(label: string, apply: (model: PresentationModel, index: number, id: string) => void): void {
    const model = deps.model.value;
    const id = selection.selectedFrameId.value;
    if (!model || !id) return;
    deps.record(label);
    apply(model, deps.currentIndex.value, id);
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

  const recentColors = ref<string[]>(loadRecentColors());

  function onFrameSelect(id: string): void {
    if (editing.editingFrameId.value && editing.editingFrameId.value !== id) {
      editing.stopEditing();
    }
    selection.select(id);
  }
  function onStartEdit(id: string, point?: CaretPoint): void {
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
    if (editing.editingFrameId.value) editing.stopEditing();
    selection.clear();
  }

  function setColor(color: string): void {
    editing.setColor(color);
    recentColors.value = pushRecentColor(color, recentColors.value);
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

    // SPEC-082 manipulation keys (nudge / delete / duplicate / ungroup).
    if (manipulation.onKeydown(event)) return;

    if (event.key === "Tab" && selection.selectedFrameId.value !== null) {
      event.preventDefault();
      selection.cycle(event.shiftKey ? -1 : 1);
      return;
    }
    if (event.key === "Enter" && selection.selectedFrameId.value !== null) {
      const frame = selection.selectedFrame.value;
      if (frame && frame.kind === "text") {
        event.preventDefault();
        editing.startEditing(frame.id);
      } else if (frame && frame.kind === "table") {
        event.preventDefault();
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
    recentColors: computed(() => recentColors.value),
    onFrameSelect,
    onStartEdit,
    onStartEditCell,
    onCellSelect,
    onSurfaceDeselect,
    onSurfaceDraw,
    onFramePointerDown,
    onHandlePointerDown,
    registerSurface,
    setColor,
    setFill,
    setStroke,
    onGlobalKeydown,
  };
}
