// Composes SPEC-081's slide editing for PresentationView so the view stays a
// thin wiring hub (PresentationView.size.spec.ts): frame selection
// (useFrameSelection) + the text-editing session (useSlideTextEditing) + the
// stage's global keyboard (undo/redo, slide navigation, Tab/Shift+Tab frame
// cycling, Enter-to-edit, Escape) + the toolbar's recent-colour memory.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { findShortcut, isOutsideDocumentInput } from "../../lib/shortcuts";
import { loadRecentColors, pushRecentColor } from "../../lib/core/colors";
import { useFrameSelection } from "./useFrameSelection";
import { useSlideTextEditing, type CaretPoint } from "./useSlideTextEditing";

function isTypingTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
  );
}

export interface PresentationEditingDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  record: (label?: string) => void;
  commit: () => void;
  undo: () => void;
  redo: () => void;
  /** Move the current slide by `delta` (PageUp/PageDown). */
  navigate: (delta: number) => void;
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
      }
      return;
    }
    if (event.key === "Escape" && selection.selectedFrameId.value !== null) {
      event.preventDefault();
      selection.clear();
      return;
    }
    if (event.key === "PageUp" || event.key === "PageDown") {
      event.preventDefault();
      deps.navigate(event.key === "PageDown" ? 1 : -1);
    }
  }

  return {
    selection,
    editing,
    recentColors: computed(() => recentColors.value),
    onFrameSelect,
    onStartEdit,
    onSurfaceDeselect,
    setColor,
    onGlobalKeydown,
  };
}
