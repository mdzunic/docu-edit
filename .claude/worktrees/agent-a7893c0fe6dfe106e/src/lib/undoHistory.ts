// Per-document undo/redo history (SPEC-015). Every command in `commands.ts`
// is a pure function that clones-then-edits the model (see that file's
// header comment) and never mutates its input — so undo needs no
// per-command inverse. An entry simply remembers the whole model+selection
// pair from before and after an edit, and undo/redo swap between them. That
// is also what the shared test in `commands.undo.spec.ts` (AC-20, AC-21)
// exploits: undoing any registered command is nothing more than restoring
// the untouched `before` object, so a command is undoable for free as long
// as it upholds the immutability contract every command already promises.
//
// Framework-free by design (see CLAUDE.md) so it is unit-testable without
// mounting DocumentView.vue, which owns only the DOM-facing wiring: keyboard
// shortcuts, toolbar buttons and the caret-move that breaks a typing group.
import type { DocumentModel } from "./documentModel";
import type { ModelSelection } from "./commands";

/** A full editor state an undo/redo step can restore: the model plus the
 * selection that was active alongside it. */
export interface HistorySnapshot {
  model: DocumentModel;
  selection: ModelSelection;
}

/** Typed `null` seed for history-snapshot refs in `.vue` scripts, which
 * espree bars from writing the `HistorySnapshot | null` annotation
 * themselves (see commands.ts's `noSelection` for the same constraint). */
export function noHistorySnapshot(): HistorySnapshot | null {
  return null;
}

/** One undo step: the state before and after an edit, its human-readable
 * action name (AC-7, AC-8's "Undo <action>" / "Redo <action>" tooltip), and
 * whether it may absorb later same-label edits into itself (AC-9..12's
 * typing grouping) instead of always starting a new step. */
export interface HistoryEntry {
  before: HistorySnapshot;
  after: HistorySnapshot;
  label: string;
  groupable: boolean;
}

/** The history holds at least this many steps per document (AC-16) and is
 * bounded to it (AC-17): pushing past the cap evicts the oldest step. Model
 * snapshots are cheap to hold — commands only clone the containers they
 * touch (see commands.ts), so an untouched paragraph is shared by reference
 * across every step that didn't edit it — which is what makes a 200-step
 * cap enough to bound memory without needing to discard content eagerly. */
export const MAX_HISTORY_STEPS = 200;

/** A typing group is broken automatically after this long without a further
 * groupable edit (AC-11), mirroring the autosave debounce's shape
 * (autosaveScheduler.ts) but as its own constant — the two cadences are
 * unrelated and must be free to diverge. */
export const TYPING_GROUP_PAUSE_MS = 2000;

interface DocumentHistory {
  undoStack: HistoryEntry[];
  redoStack: HistoryEntry[];
  /** True while the most recent step is still open to absorbing further
   * same-label groupable edits (a typing run in progress). */
  groupOpen: boolean;
  pauseTimer: ReturnType<typeof setTimeout> | null;
}

const histories = new Map<string, DocumentHistory>();

function historyFor(docId: string): DocumentHistory {
  let history = histories.get(docId);
  if (!history) {
    history = { undoStack: [], redoStack: [], groupOpen: false, pauseTimer: null };
    histories.set(docId, history);
  }
  return history;
}

function clearPauseTimer(history: DocumentHistory): void {
  if (history.pauseTimer !== null) {
    clearTimeout(history.pauseTimer);
    history.pauseTimer = null;
  }
}

/** Records a completed edit. A groupable edit merges into the still-open
 * previous step when its label matches (AC-9: consecutive typed characters
 * collapse into one step); anything else — a non-groupable edit (AC-12,
 * AC-13), or a groupable one whose group was already broken — starts a new
 * step. Any recorded edit clears the redo stack (AC-15): the new edit
 * invalidates whatever the redo stack would have replayed. A groupable edit
 * (re)arms the pause timer that breaks the group after 2 s of inactivity
 * (AC-11); a non-groupable one leaves no group open to break. */
export function recordEdit(docId: string, entry: HistoryEntry): void {
  const history = historyFor(docId);
  history.redoStack = [];
  clearPauseTimer(history);
  const top = history.undoStack[history.undoStack.length - 1];
  if (entry.groupable && history.groupOpen && top && top.groupable && top.label === entry.label) {
    top.after = entry.after;
  } else {
    history.undoStack.push(entry);
    if (history.undoStack.length > MAX_HISTORY_STEPS) {
      history.undoStack.shift();
    }
    history.groupOpen = entry.groupable;
  }
  if (entry.groupable) {
    history.pauseTimer = setTimeout(() => {
      history.pauseTimer = null;
      history.groupOpen = false;
    }, TYPING_GROUP_PAUSE_MS);
  }
}

/** Breaks the current typing group (AC-10: a caret move) without recording
 * an edit, so the next groupable edit starts a fresh step instead of
 * absorbing into the one before the break. A no-op for a document with no
 * history yet. */
export function breakGroup(docId: string): void {
  const history = histories.get(docId);
  if (history) {
    history.groupOpen = false;
    clearPauseTimer(history);
  }
}

/** Undoes the most recent step, moving it to the redo stack and returning
 * the state to restore — the step's `before` snapshot, which is exactly the
 * selection that existed before the undone operation (AC-14) — or null when
 * there is nothing to undo. */
export function undo(docId: string): HistorySnapshot | null {
  const history = histories.get(docId);
  const entry = history?.undoStack.pop();
  if (!history || !entry) {
    return null;
  }
  clearPauseTimer(history);
  history.groupOpen = false;
  history.redoStack.push(entry);
  return entry.before;
}

/** Redoes the most recently undone step, moving it back to the undo stack
 * and returning its `after` snapshot, or null when there is nothing to
 * redo. */
export function redo(docId: string): HistorySnapshot | null {
  const history = histories.get(docId);
  const entry = history?.redoStack.pop();
  if (!history || !entry) {
    return null;
  }
  clearPauseTimer(history);
  history.groupOpen = false;
  history.undoStack.push(entry);
  return entry.after;
}

export function canUndo(docId: string): boolean {
  return (histories.get(docId)?.undoStack.length ?? 0) > 0;
}

export function canRedo(docId: string): boolean {
  return (histories.get(docId)?.redoStack.length ?? 0) > 0;
}

/** The action name the next undo would report, for the toolbar's
 * "Undo <action>" tooltip (AC-7); null when the undo stack is empty. */
export function undoLabel(docId: string): string | null {
  const stack = histories.get(docId)?.undoStack;
  return stack && stack.length > 0 ? stack[stack.length - 1].label : null;
}

/** The action name the next redo would report (AC-8); null when the redo
 * stack is empty. */
export function redoLabel(docId: string): string | null {
  const stack = histories.get(docId)?.redoStack;
  return stack && stack.length > 0 ? stack[stack.length - 1].label : null;
}

/** Discards a document's history entirely (AC-18: per document, since every
 * function here is keyed by docId; AC-19: cleared when the document is
 * closed — call this from the close path). */
export function clearHistory(docId: string): void {
  const history = histories.get(docId);
  if (history) {
    clearPauseTimer(history);
    histories.delete(docId);
  }
}
