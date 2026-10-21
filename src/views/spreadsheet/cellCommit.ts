// The cell editor's commit path (SPEC-122 "Cell commands", SPEC-130 "Entry
// gate"): what a commit, a Fill (Ctrl/Cmd+Enter) and a dropdown pick write,
// and when — the editing state itself is useCellEditing's, which builds this
// over its draft and hands it the moves back to Ready.
//
// - Prepare: the edited cell's write context and its entry (cellEntry.ts's
//   entryFromText over the draft — a formula error keeps the editor open
//   with the message and the caret at the error; a pick brings its entry);
//   a Fill is refused first by protection or a partial array (a toast).
// - The entry gate (`entryGate`, useDataValidation): after protection and
//   parsing, the entry is checked against the data-validation rules — a
//   commit at its cell, a Fill at each target a rule covers (the first
//   failure deciding), a pick at its cell. A valid entry is written at once
//   (one undo step, rows refitted), exactly as without the gate; an invalid
//   one opens the alert and the editor stays open — the commit reports false,
//   so a click-away is "handled" and the selection stays — until it is
//   answered: "commit" writes and moves on, "retry" goes back into Edit mode
//   with the whole entered text selected in the focused cell editor,
//   "cancel" closes the editor leaving the cell as it was. Nothing commits
//   while the alert asks.
// - An unchanged text records nothing and validates nothing.
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import type { CellEntry } from "../../lib/spreadsheet/cells";
import type { CellValue, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";
import type { CellDraft } from "./cellDraft";
import type { Move } from "./cellEditorKeys";
import { entryFromText, refusalFor, writeEntry, writeFill, type CellWriteContext } from "./cellEntry";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";
import type { EntryGate, EntryVerdict } from "./validationGate";

export interface CommitPathDeps {
  model: () => SpreadsheetModel | null;
  /** The cell being edited and the selection a Fill fills (null: Ready). */
  target: () => { sheetId: string; at: CellRef; ranges: GridRange[] } | null;
  draft: CellDraft;
  /** The edit text the editor opened on. */
  original: () => string;
  editing: () => boolean;
  locale: () => SpreadsheetLocale;
  format: () => FormatContext;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  measurer: () => CellTextMeasurer;
  toast: (message: string) => void;
  /** The view state into the live model (useCellEditing's). */
  syncViewState: () => void;
  /** Opens Enter mode on the active cell holding `text` (a pick); false
   * when refused (a toast). */
  open: (text: string) => boolean;
  /** Back to Ready (no move). */
  finish: () => void;
  /** Back to Ready after a write, applying a commit's move. */
  done: (move: Move | null) => void;
  /** Back into Edit mode (the alert's "retry"; the text is then selected). */
  edit: () => void;
  entryGate?: EntryGate;
}

interface Prepared {
  ctx: CellWriteContext;
  at: CellRef;
  entry: CellEntry;
}

export function createCommitPath(deps: CommitPathDeps) {
  const draft = deps.draft;
  /** The entry the gate's alert is asking about: nothing commits until it
   * is answered; `forget` (useCellEditing's finish) drops it. */
  let gated: object | null = null;

  /** The edited cell's write context and the draft (or `given`) as its
   * entry — or null after showing why it cannot be written. */
  function prepared(forFill: boolean, given?: CellEntry): Prepared | null {
    const m = deps.model();
    const target = deps.target();
    const sheetIndex = m?.sheets.findIndex((sheet) => sheet.id === target?.sheetId) ?? -1;
    const sheet = m?.sheets[sheetIndex];
    if (!m || !target || !sheet) {
      return null;
    }
    const refusal = forFill ? refusalFor(m, sheetIndex, target.ranges) : null;
    if (refusal) {
      deps.toast(refusal);
      return null;
    }
    const style = cellStyle(m, effectiveStyleId(sheet, target.at.row, target.at.col));
    const outcome = given ? { entry: given } : entryFromText(draft.text.value, { locale: deps.locale(), dateSystem: m.dateSystem, style });
    if ("error" in outcome) {
      draft.fail(outcome.error, outcome.caret);
      return null;
    }
    deps.syncViewState();
    return { ctx: { model: m, sheetIndex, undo: deps.undo, measurer: deps.measurer(), format: deps.format() }, at: target.at, entry: outcome.entry };
  }

  /** Asks the entry gate about `ready`'s entry at `targets`: false when it
   * may be written now, true when the alert opened (see the header). */
  function held(ready: Prepared, targets: Iterable<CellRef>, write: () => void, after: () => void): boolean {
    const gate = deps.entryGate;
    const sheet = ready.ctx.model.sheets[ready.ctx.sheetIndex];
    const failure = gate && sheet ? gate.firstFailure(sheet, targets, { value: ready.entry.value, formula: ready.entry.formula }, ready.at) : null;
    if (!gate || !failure) {
      return false;
    }
    const token = {};
    gated = token;
    void gate.alert(failure).then((verdict) => {
      if (gated === token) {
        gated = null;
        answered(verdict, write, after);
      }
    });
    return true;
  }

  function answered(verdict: EntryVerdict, write: () => void, after: () => void): void {
    if (verdict === "commit") {
      deps.syncViewState();
      write();
      after();
    } else if (verdict === "retry") {
      deps.edit();
      draft.focusCell();
      draft.selectAll();
      return;
    } else {
      deps.finish();
    }
    draft.focusCell();
  }

  /** Writes the draft into the edited cell and moves on; false when it
   * cannot be written, or the gate's alert holds it (kept open). */
  function commit(move: Move | null): boolean {
    if (gated) {
      return false;
    }
    if (deps.target() && draft.text.value !== deps.original()) {
      const ready = prepared(false);
      if (!ready) {
        return false;
      }
      const text = draft.text.value;
      const write = (): void => writeEntry(ready.ctx, ready.at, ready.entry, text);
      if (held(ready, [ready.at], write, () => deps.done(move))) {
        return false;
      }
      write();
    }
    deps.done(move);
    return true;
  }

  /** Ctrl/Cmd+Enter: the entry into every cell of the selection ("Fill"). */
  function fill(): void {
    const ranges = deps.target()?.ranges ?? [];
    const ready = gated ? null : prepared(true);
    if (!ready) {
      return;
    }
    const write = (): void => writeFill(ready.ctx, ranges, ready.at, ready.entry);
    const sheet = ready.ctx.model.sheets[ready.ctx.sheetIndex];
    if (!held(ready, deps.entryGate && sheet ? deps.entryGate.fillTargets(sheet, ranges, ready.at) : [], write, deps.finish)) {
      write();
      deps.finish();
    }
  }

  /** A dropdown pick: `value` entered into the active cell (an edit in
   * progress commits first, then protection, the gate, one undo step
   * `Typing "<text>" in B2`), without moving. False when nothing was
   * written (yet). */
  function enterValue(value: CellValue, text: string): boolean {
    if (gated || (deps.editing() && !commit(null)) || !deps.open(text)) {
      return false;
    }
    const ready = prepared(false, { value });
    if (!ready) {
      deps.finish();
      return false;
    }
    const write = (): void => writeEntry(ready.ctx, ready.at, ready.entry, text);
    if (held(ready, [ready.at], write, () => deps.done(null))) {
      return false;
    }
    write();
    deps.done(null);
    return true;
  }

  return {
    commit,
    fill,
    enterValue,
    forget: (): void => {
      gated = null;
    },
  };
}
