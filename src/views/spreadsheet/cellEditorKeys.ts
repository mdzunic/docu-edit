// The cell editor's keys (SPEC-122 "Editor keys") — the table, applied by
// useCellEditing's onKeydown — and the moves Enter and Tab make, with the
// Tab-run return: Tab (in a single-cell selection) remembers the column a
// run of Tabs started in, and Enter after it lands one row down in that
// column; any other navigation — a navigation key, a click — forgets it,
// while typing the next entry does not. Inside a multi-cell
// selection Enter and Tab cycle within it (SPEC-117's cycleWithin) without
// collapsing it. These are conventions of the editor, not registry
// commands, so Ctrl/Cmd+Enter never meets the text product's
// insert.pageBreak.
import type { Ref } from "vue";
import { findShortcut } from "../../lib/shortcuts";
import { cycleWithin, moveActive, navigateByKey, selectCell, type NavContext, type SheetSelection } from "../../lib/spreadsheet/selection";
import type { PointMode } from "./usePointMode";
import { primaryModifier } from "./useSheetSelection";

export type EditMode = "ready" | "enter" | "edit";

/** How the active cell moves on: Enter/Tab (`back` = Shift), or a key
 * SPEC-117's map applies (an arrow in Enter mode). */
export type Move = { kind: "enter" | "tab"; back: boolean } | { kind: "key"; event: KeyboardEvent };

/** SPEC-117's navigation keys: in Ready they move the selection, which
 * forgets a Tab-run's return column (typing the next entry does not). */
const NAVIGATION_KEYS = new Set(["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Home", "End", "PageUp", "PageDown"]);

function singleCell(sel: SheetSelection): boolean {
  const range = sel.ranges[0];
  return sel.ranges.length === 1 && range !== undefined && range.top === range.bottom && range.left === range.right;
}

/** The moves, and the Tab-run they keep. */
export function createCellMoves(deps: { sheetId: () => string; selection: () => SheetSelection; set: (next: SheetSelection) => void; nav: () => NavContext }) {
  let tabRun: { sheetId: string; col: number } | null = null;

  /** The selection a move gives (null: the key moves nothing). */
  function next(move: Move): SheetSelection | null {
    const nav = deps.nav();
    const sel = deps.selection();
    const sheetId = deps.sheetId();
    if (move.kind === "key") {
      tabRun = null;
      const event = move.event;
      return navigateByKey(sel, { key: event.key, shift: event.shiftKey, mod: primaryModifier(event), alt: event.altKey, ctrl: event.ctrlKey }, nav);
    }
    if (move.kind === "tab") {
      if (singleCell(sel) && tabRun?.sheetId !== sheetId) {
        tabRun = { sheetId, col: sel.active.col };
      }
      return cycleWithin(sel, "row", move.back, nav);
    }
    const run = tabRun;
    tabRun = null;
    if (run && run.sheetId === sheetId && !move.back && singleCell(sel)) {
      return moveActive(selectCell({ row: sel.active.row, col: run.col }, nav), "down", nav);
    }
    return cycleWithin(sel, "column", move.back, nav);
  }

  /** Applies a move; true when the active cell moved. */
  function apply(move: Move): boolean {
    const sel = next(move);
    if (sel) {
      deps.set(sel);
    }
    return sel !== null;
  }

  return {
    apply,
    forget: (): void => {
      tabRun = null;
    },
  };
}

/** What the key table drives (useCellEditing's actions). */
export interface EditorKeyActions {
  mode: Ref<EditMode>;
  point: Pick<PointMode, "arrow" | "cycle" | "end">;
  moves: ReturnType<typeof createCellMoves>;
  /** Moved the active cell: the grid reveals it. */
  revealed(): void;
  openEdit(): boolean;
  focusCell(): void;
  clearSelection(): void;
  cancel(): void;
  commit(move: Move | null): boolean;
  fill(): void;
  insertText(text: string): void;
  historyStep(back: boolean): void;
}

/** A key in Ready state; true when the editor took it (otherwise SPEC-117's
 * navigation runs). */
export function readyKey(event: KeyboardEvent, actions: EditorKeyActions): boolean {
  const plain = !event.ctrlKey && !event.metaKey && !event.altKey;
  if (event.key === "F2" && plain && !event.shiftKey) {
    event.preventDefault();
    if (actions.openEdit()) {
      actions.focusCell();
    }
    return true;
  }
  if ((event.key === "Delete" || event.key === "Backspace") && plain) {
    event.preventDefault();
    actions.clearSelection();
    return true;
  }
  if (event.key === "Tab" || event.key === "Enter") {
    event.preventDefault();
    if (plain && actions.moves.apply({ kind: event.key === "Tab" ? "tab" : "enter", back: event.shiftKey })) {
      actions.revealed();
    }
    return true;
  }
  if (NAVIGATION_KEYS.has(event.key)) {
    actions.moves.forget();
  }
  return false;
}

/** A key in Enter or Edit mode. Every key is the editor's: what it does not
 * handle here is native text editing (typing, Delete, Home, the caret). */
export function editKey(event: KeyboardEvent, actions: EditorKeyActions): boolean {
  const id = findShortcut(event)?.id;
  if (id === "edit.undo" || id === "edit.redo") {
    event.preventDefault();
    actions.historyStep(id === "edit.undo");
    return true;
  }
  const ctrlOrCmd = event.ctrlKey || event.metaKey;
  switch (event.key) {
    case "Escape":
      event.preventDefault();
      actions.cancel();
      return true;
    case "Enter":
      event.preventDefault();
      if (event.altKey && !ctrlOrCmd) {
        actions.insertText("\n");
      } else if (ctrlOrCmd) {
        actions.fill();
      } else {
        actions.commit({ kind: "enter", back: event.shiftKey });
      }
      return true;
    case "Tab":
      event.preventDefault();
      if (!ctrlOrCmd && !event.altKey) {
        actions.commit({ kind: "tab", back: event.shiftKey });
      }
      return true;
    case "F2":
      event.preventDefault();
      actions.mode.value = actions.mode.value === "enter" ? "edit" : "enter";
      actions.point.end();
      return true;
    case "F4":
      event.preventDefault();
      actions.point.cycle();
      return true;
    case "ArrowUp":
    case "ArrowDown":
    case "ArrowLeft":
    case "ArrowRight":
      if (actions.mode.value === "enter") {
        event.preventDefault();
        if (!actions.point.arrow(event)) {
          actions.commit({ kind: "key", event });
        }
      }
      return true;
    default:
      return true;
  }
}
