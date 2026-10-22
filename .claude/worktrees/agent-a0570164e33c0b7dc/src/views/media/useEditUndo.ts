// The image editor's undo stack (SPEC-112) in SPEC-015's vocabulary and
// SPEC-080's order: the caller `record(label)`s BEFORE mutating the history
// and `commit()`s after; undo/redo swap the history list back. One step per
// completed gesture — a drag records once on its first move and commits on
// release, so it is one entry, not sixty. Per open editor (a fresh stack per
// session), never persisted, capped at IMAGE_UNDO_LIMIT with the oldest
// dropped. Snapshots are the history arrays themselves: the history is
// immutable (core/imageEditHistory.ts returns new arrays), so a step holds a
// reference, never a copy.
import { ref, type Ref } from "vue";
import type { ImageEdit } from "../../lib/core/imageEdits";

/** The most steps a session keeps. */
export const IMAGE_UNDO_LIMIT = 100;

interface Step {
  label: string;
  history: readonly ImageEdit[];
}

export interface EditUndo {
  /** BEFORE mutating: remembers the history as it stands. */
  record: (label: string) => void;
  /** AFTER mutating: one step, unless the mutation changed nothing. */
  commit: () => void;
  /** Abandons a recorded step, putting its history back. */
  cancel: () => void;
  undo: () => void;
  redo: () => void;
  /** Forgets every step (a new session, or after Flatten — which no undo
   * can take back). */
  reset: () => void;
  canUndo: Ref<boolean>;
  canRedo: Ref<boolean>;
  undoLabel: Ref<string>;
  redoLabel: Ref<string>;
}

/** Whether two histories hold the same edits. */
export function sameHistory(a: readonly ImageEdit[], b: readonly ImageEdit[]): boolean {
  return a === b || JSON.stringify(a) === JSON.stringify(b);
}

export function useEditUndo(history: Ref<ImageEdit[]>): EditUndo {
  const undoStack: Step[] = [];
  const redoStack: Step[] = [];
  const canUndo = ref(false);
  const canRedo = ref(false);
  const undoLabel = ref("");
  const redoLabel = ref("");
  let pending: Step | null = null;

  function refresh(): void {
    canUndo.value = undoStack.length > 0;
    canRedo.value = redoStack.length > 0;
    undoLabel.value = undoStack.at(-1)?.label ?? "";
    redoLabel.value = redoStack.at(-1)?.label ?? "";
  }

  function record(label: string): void {
    pending = { label, history: history.value };
  }

  function commit(): void {
    const step = pending;
    pending = null;
    if (!step || sameHistory(step.history, history.value)) {
      return;
    }
    undoStack.push(step);
    if (undoStack.length > IMAGE_UNDO_LIMIT) {
      undoStack.shift();
    }
    redoStack.length = 0;
    refresh();
  }

  function cancel(): void {
    if (pending) {
      history.value = [...pending.history];
      pending = null;
    }
  }

  function swap(from: Step[], to: Step[]): void {
    const step = from.pop();
    if (!step) {
      return;
    }
    to.push({ label: step.label, history: history.value });
    history.value = [...step.history];
    refresh();
  }

  function reset(): void {
    undoStack.length = 0;
    redoStack.length = 0;
    pending = null;
    refresh();
  }

  return {
    record,
    commit,
    cancel,
    undo: () => swap(undoStack, redoStack),
    redo: () => swap(redoStack, undoStack),
    reset,
    canUndo,
    canRedo,
    undoLabel,
    redoLabel,
  };
}
