// The presentation undo stack (SPEC-080) — the first module of the
// presentation-side composable directory, mirroring src/views/document/.
// This spec is the normative source for it. A structural command snapshots
// `{ slides, structureDirty }` via `structuredClone` BEFORE it runs (opaque
// and media are excluded — the commands never touch them); undo swaps those
// fields back onto the live model and re-renders + persists through the
// caller's `onApply`. Per-open-view (a fresh stack per mounted view), never
// persisted; capped at 100 entries, oldest dropped. SPEC-081+ push their
// edits through the same `record`.
import { ref, type Ref } from "vue";
import type { DocumentProperties } from "../../lib/core/documentProperties";
import type { PresentationModel, Slide } from "../../lib/presentation/model";

/** The most undo steps a view keeps; the oldest is dropped past this (AC-43). */
export const MAX_UNDO_STEPS = 100;

/** The deck-level state a snapshot captures beside the slides (SPEC-090), so an
 * undo restores size, first-slide-number, properties and `deckDirty` together
 * with the slide geometry a rescale changed. */
interface DeckSnapshot {
  slideWidthCm: number;
  slideHeightCm: number;
  firstSlideNumber?: number;
  deckDirty?: true;
  properties?: DocumentProperties;
}

interface Snapshot {
  slides: Slide[];
  structureDirty?: true;
  deck: DeckSnapshot;
}

export interface PresentationUndoDeps {
  /** The live model whose `slides`/`structureDirty` a snapshot swaps. */
  model: Ref<PresentationModel | null>;
  /** Called after undo/redo mutates the model — the view re-renders the
   * shallowRef and writes the deck through (`touchDocument`). */
  onApply: () => void;
}

export interface PresentationUndo {
  record: (label?: string) => void;
  undo: () => void;
  redo: () => void;
  canUndo: Ref<boolean>;
  canRedo: Ref<boolean>;
  reset: () => void;
}

export function usePresentationUndo(deps: PresentationUndoDeps): PresentationUndo {
  const undoStack: Snapshot[] = [];
  const redoStack: Snapshot[] = [];
  const canUndo = ref(false);
  const canRedo = ref(false);

  function refresh(): void {
    canUndo.value = undoStack.length > 0;
    canRedo.value = redoStack.length > 0;
  }

  function snapshot(model: PresentationModel): Snapshot {
    const deck: DeckSnapshot = {
      slideWidthCm: model.slideWidthCm,
      slideHeightCm: model.slideHeightCm,
    };
    if (model.firstSlideNumber !== undefined) {
      deck.firstSlideNumber = model.firstSlideNumber;
    }
    if (model.deckDirty) {
      deck.deckDirty = true;
    }
    if (model.properties) {
      deck.properties = structuredClone(model.properties);
    }
    const snap: Snapshot = { slides: structuredClone(model.slides), deck };
    if (model.structureDirty) {
      snap.structureDirty = true;
    }
    return snap;
  }

  function applySnapshot(model: PresentationModel, snap: Snapshot): void {
    model.slides = snap.slides;
    if (snap.structureDirty) {
      model.structureDirty = true;
    } else {
      delete model.structureDirty;
    }
    // Deck-level state (SPEC-090): size, first-slide-number, properties and
    // deckDirty restore together; absent keys are deleted, not left stale.
    model.slideWidthCm = snap.deck.slideWidthCm;
    model.slideHeightCm = snap.deck.slideHeightCm;
    if (snap.deck.firstSlideNumber !== undefined) {
      model.firstSlideNumber = snap.deck.firstSlideNumber;
    } else {
      delete model.firstSlideNumber;
    }
    if (snap.deck.deckDirty) {
      model.deckDirty = true;
    } else {
      delete model.deckDirty;
    }
    if (snap.deck.properties) {
      model.properties = snap.deck.properties;
    } else {
      delete model.properties;
    }
    deps.onApply();
  }

  /** Captures the pre-command state. Call BEFORE running a structural command
   * (label is for parity with the text side / future tooltips). */
  function record(label = ""): void {
    void label;
    const model = deps.model.value;
    if (!model) {
      return;
    }
    undoStack.push(snapshot(model));
    if (undoStack.length > MAX_UNDO_STEPS) {
      undoStack.shift();
    }
    redoStack.length = 0;
    refresh();
  }

  function undo(): void {
    const model = deps.model.value;
    if (!model || undoStack.length === 0) {
      return;
    }
    redoStack.push(snapshot(model));
    applySnapshot(model, undoStack.pop()!);
    refresh();
  }

  function redo(): void {
    const model = deps.model.value;
    if (!model || redoStack.length === 0) {
      return;
    }
    undoStack.push(snapshot(model));
    applySnapshot(model, redoStack.pop()!);
    refresh();
  }

  function reset(): void {
    undoStack.length = 0;
    redoStack.length = 0;
    refresh();
  }

  return { record, undo, redo, canUndo, canRedo, reset };
}
