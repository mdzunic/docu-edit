// The image editor's gesture API (SPEC-112): every tool mutates the history
// list through these four calls, so every mutation is undoable the same way
// — `begin(label)` records BEFORE the first change, `update(next)` replaces
// the history live (the preview follows; no undo entry), `end()` commits one
// step. A drag is begin → many updates → end; a click, a keypress or a typed
// value is `apply(label, next)`, the same three in one. What depends on the
// geometry follows it (rebaseDependents): a resize keeps its scale, and
// redactions and annotations keep their place on the picture, when a gesture
// crops, turns or resizes under them — always measured from the gesture's
// START, so a long drag never compounds: the entries an earlier update
// derived are put back to the start's before the next derivation.
import { computed, type ComputedRef, type Ref } from "vue";
import { rebaseDependents } from "../../lib/core/imageEditHistory";
import {
  normaliseStack,
  resolveEditGeometry,
  type EditGeometry,
  type ImageEdit,
  type PixelSize,
} from "../../lib/core/imageEdits";
import type { EditUndo } from "./useEditUndo";

export interface EditSession {
  /** The session's edits in the order the user made them (intent). */
  history: Ref<ImageEdit[]>;
  /** The canonical program the history normalises to. */
  stack: ComputedRef<ImageEdit[]>;
  /** The stack base's natural size (the edits' coordinate space). */
  naturalPx: Ref<PixelSize | null>;
  geometry: ComputedRef<EditGeometry | null>;
  begin: (label: string) => void;
  update: (next: ImageEdit[]) => void;
  end: () => void;
  /** Abandons the open gesture, restoring the history it started from. */
  cancel: () => void;
  apply: (label: string, next: ImageEdit[]) => void;
  /** Whether a gesture is open. */
  active: () => boolean;
}

export function createEditSession(
  history: Ref<ImageEdit[]>,
  naturalPx: Ref<PixelSize | null>,
  undo: EditUndo,
): EditSession {
  const stack = computed(() => normaliseStack(history.value));
  const geometry = computed(() => (naturalPx.value ? resolveEditGeometry(stack.value, naturalPx.value) : null));
  let start: ImageEdit[] | null = null;
  // Entries this gesture derived (a rebased resize, a carried object) →
  // the start's entry each was derived from.
  let origin = new WeakMap<ImageEdit, ImageEdit>();

  function end(): void {
    if (start !== null) {
      start = null;
      undo.commit();
    }
  }

  function begin(label: string): void {
    end();
    start = history.value;
    origin = new WeakMap();
    undo.record(label);
  }

  function update(next: ImageEdit[]): void {
    const from = start ?? history.value;
    const base = next.map((edit) => origin.get(edit) ?? edit);
    const derived = naturalPx.value ? rebaseDependents(from, base, naturalPx.value) : base;
    derived.forEach((edit, index) => {
      if (edit !== base[index]) {
        origin.set(edit, base[index]);
      }
    });
    history.value = derived;
  }

  function cancel(): void {
    if (start !== null) {
      start = null;
      undo.cancel();
    }
  }

  function apply(label: string, next: ImageEdit[]): void {
    begin(label);
    update(next);
    end();
  }

  return { history, stack, naturalPx, geometry, begin, update, end, cancel, apply, active: () => start !== null };
}
