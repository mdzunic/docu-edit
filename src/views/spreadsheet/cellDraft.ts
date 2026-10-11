// The cell editor's draft (SPEC-122 "Editor DOM"): the text and caret the
// in-cell editor and the editable formula bar share, the error shown under
// it, the hosts that render it (each useDraftSurface registers one), and the
// draft's own text undo — Cmd/Ctrl+Z and +Y while editing step through the
// draft's edits, never the workbook's. `revision` is bumped whenever the
// draft changes other than by a host's own input, so every host re-renders.
import { ref } from "vue";

export type SurfaceKind = "cell" | "bar";

/** What useDraftSurface registers for each editable host. */
export interface DraftSurface {
  /** The live selection in the host as offsets, or null. */
  selection(): { start: number; end: number } | null;
  /** The offset under a viewport point, or null. */
  caretAt(x: number, y: number): number | null;
  focus(): void;
  focused(): boolean;
}

interface DraftState {
  text: string;
  caret: number;
}

export function createDraft() {
  const text = ref("");
  const caret = ref(0);
  const error = ref("");
  const revision = ref(0);
  let history: DraftState[] = [];
  let future: DraftState[] = [];
  const surfaces = new Map<SurfaceKind, DraftSurface>();
  let focusRequested = false;

  /** A new draft state (undoable by `step`). */
  function set(next: string, at: number): void {
    history.push({ text: text.value, caret: caret.value });
    future = [];
    text.value = next;
    caret.value = Math.max(0, Math.min(at, next.length));
    error.value = "";
    revision.value += 1;
  }

  /** A fresh draft (the editor opens or closes): no history. */
  function reset(next = "", at = next.length): void {
    history = [];
    future = [];
    text.value = next;
    caret.value = Math.max(0, Math.min(at, next.length));
    error.value = "";
    revision.value += 1;
  }

  /** The draft's undo (`back`) or redo; false when there is none. */
  function step(back: boolean): boolean {
    const state = (back ? history : future).pop();
    if (!state) {
      return false;
    }
    (back ? future : history).push({ text: text.value, caret: caret.value });
    text.value = state.text;
    caret.value = state.caret;
    revision.value += 1;
    return true;
  }

  /** Shows a commit's error with the caret where it lies. */
  function fail(message: string, at: number): void {
    error.value = message;
    caret.value = at;
    revision.value += 1;
  }

  function focusedSurface(): DraftSurface | undefined {
    return [...surfaces.values()].find((surface) => surface.focused());
  }

  /** Reads the caret from the host that has focus. */
  function syncCaret(): void {
    const sel = focusedSurface()?.selection();
    if (sel) {
      caret.value = sel.end;
    }
  }

  /** The focused host's selection (else the caret) replaced by `insert`. */
  function replaceSelection(insert: string, maxLength: number): boolean {
    const sel = focusedSurface()?.selection() ?? { start: caret.value, end: caret.value };
    const next = text.value.slice(0, sel.start) + insert + text.value.slice(sel.end);
    if (next.length > maxLength) {
      return false;
    }
    set(next, sel.start + insert.length);
    return true;
  }

  /** text[from, to) replaced by `insert`, the caret at `caretAfter` (one
   * undoable step) — SPEC-127's function assists insert through this. */
  function replaceRange(from: number, to: number, insert: string, caretAfter: number): void {
    set(text.value.slice(0, from) + insert + text.value.slice(to), caretAfter);
  }

  /** Registers a host; returns its unregistration. A cell host mounted
   * while focus was requested (the grid re-mounted on a sheet switch) takes
   * it. */
  function register(kind: SurfaceKind, surface: DraftSurface): () => void {
    surfaces.set(kind, surface);
    if (kind === "cell" && focusRequested) {
      focusRequested = false;
      surface.focus();
    }
    return () => {
      if (surfaces.get(kind) === surface) {
        surfaces.delete(kind);
      }
    };
  }

  return {
    text,
    caret,
    error,
    revision,
    set,
    reset,
    step,
    fail,
    syncCaret,
    replaceSelection,
    replaceRange,
    register,
    /** The in-cell editor's offset under a viewport point. */
    caretAt: (x: number, y: number): number | null => surfaces.get("cell")?.caretAt(x, y) ?? null,
    focusCell: (): void => surfaces.get("cell")?.focus(),
    /** Focus goes to the next cell host to mount (a sheet switch). */
    requestFocus: (): void => {
      focusRequested = true;
    },
  };
}

export type CellDraft = ReturnType<typeof createDraft>;
