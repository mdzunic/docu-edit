// Point mode (SPEC-122 "Point mode", "Editor keys"): while a formula is
// being entered, the arrow keys, clicks and drags write references into it.
// THIS composable owns the point reference, the reference colouring and
// highlights, and which sheet a point lands on; useCellEditing owns the
// draft it writes into.
//
// - A point starts when lib's canPointAt(text, caret) allows it — the text
//   starts with "=" and the caret follows "=", an operator, "(" or "," — and
//   inserts a reference at the caret; while it lasts, each further arrow,
//   click or drag REPLACES that reference (placePointReference). Arrows move
//   it from the cell it names (from the edited cell, or the shown sheet's
//   active cell, at first) through SPEC-117's keyboard map, so Shift+arrows
//   extend it to a range and Ctrl/Cmd+arrows jump to the data edge
//   (edgeInDirection); a click names the cell (its merge's top-left), Shift+
//   click and drags extend it. Typing anything else ends the point
//   (`end`, called by useCellEditing on input).
// - On another sheet (a tab clicked while pointing) references are
//   sheet-qualified and quoted as SPEC-119 prints them (`'My Sheet'!B2`).
// - Every reference in the text is coloured by referenceSpans' colour index
//   (first appearance, modulo 8, the `--sheet-ref-1..8` tokens); each one on
//   the sheet on show is a highlight the grid outlines in the same colour,
//   the one being pointed at dashed.
// - F4 cycles the reference at the caret A1 → $A$1 → A$1 → $A1
//   (cycleAbsolute).
import { computed, shallowRef, type ComputedRef } from "vue";
import { formatA1, formatRange, type GridRange } from "../../lib/spreadsheet/address";
import { quoteSheetName } from "../../lib/spreadsheet/formula/print";
import type { ReferenceNode } from "../../lib/spreadsheet/formula/references";
import { mergeAt } from "../../lib/spreadsheet/merges";
import { MAX_COLS, MAX_ROWS, type Sheet } from "../../lib/spreadsheet/model";
import { canPointAt, cycleAbsolute, placePointReference, referenceSpans } from "../../lib/spreadsheet/pointMode";
import { extendTo, navigateByKey, selectCell, type NavContext, type SheetSelection } from "../../lib/spreadsheet/selection";
import type { DraftSpan } from "./cellEditorDom";
import { primaryModifier } from "./useSheetSelection";

/** A reference outlined on the grid. */
export interface RefHighlight {
  range: GridRange;
  colorIndex: number;
  /** The reference being pointed at. */
  dashed: boolean;
}

export interface PointModeDeps {
  /** The draft (useCellEditing's) and how a point writes it. */
  text: () => string;
  caret: () => number;
  setDraft: (text: string, caret: number) => void;
  /** Whether a cell is being edited (Enter or Edit mode). */
  editing: () => boolean;
  /** The edited cell's sheet. */
  editedSheet: () => Sheet | null;
  /** The sheet on show, its active cell and its NavContext. */
  shownSheet: () => Sheet | null;
  shownSelection: () => SheetSelection;
  nav: () => NavContext;
}

interface Point {
  sheetId: string;
  span: { start: number; end: number };
  sel: SheetSelection;
}

/** A reference's range on its sheet (a whole row or column spans the grid). */
export function referenceRange(ref: ReferenceNode): GridRange {
  const end = ref.end ?? ref.start;
  const rows = [ref.start.row, end.row];
  const cols = [ref.start.col, end.col];
  const top = Math.min(rows[0] ?? 0, rows[1] ?? 0);
  const bottom = rows[0] === undefined || rows[1] === undefined ? MAX_ROWS - 1 : Math.max(rows[0], rows[1]);
  const left = Math.min(cols[0] ?? 0, cols[1] ?? 0);
  const right = cols[0] === undefined || cols[1] === undefined ? MAX_COLS - 1 : Math.max(cols[0], cols[1]);
  return { top, left, bottom, right };
}

export function usePointMode(deps: PointModeDeps) {
  const point = shallowRef<Point | null>(null);
  let dragging = false;

  /** Whether the draft is a formula at all. */
  const isFormula = (): boolean => deps.editing() && deps.text().startsWith("=");

  /** Whether an arrow, click or drag points now: the current point is still
   * at the caret, or the caret is where a reference may start. */
  function allowed(): boolean {
    if (!isFormula()) {
      return false;
    }
    const current = point.value;
    if (current && current.span.end === deps.caret()) {
      return true;
    }
    return canPointAt(deps.text(), deps.caret());
  }

  /** The reference text of a point selection on `sheet`. */
  function refText(sel: SheetSelection, sheet: Sheet, nav: NavContext): string {
    const range = sel.ranges[sel.activeRangeIndex] ?? sel.ranges[0];
    const merge = mergeAt(nav.merges, sel.active.row, sel.active.col);
    const single =
      !range ||
      (range.top === range.bottom && range.left === range.right) ||
      (merge !== undefined && merge.top === range.top && merge.left === range.left && merge.bottom === range.bottom && merge.right === range.right);
    const text = single ? formatA1(sel.active) : formatRange(range);
    const edited = deps.editedSheet();
    return edited && edited.id !== sheet.id ? `${quoteSheetName(sheet.name)}!${text}` : text;
  }

  /** Writes the point `sel` into the draft: inserted at the caret, or
   * replacing the current point's reference. */
  function place(sel: SheetSelection): void {
    const sheet = deps.shownSheet();
    if (!sheet) {
      return;
    }
    const current = point.value;
    const caret = deps.caret();
    const span = current && current.span.end === caret ? current.span : { start: caret, end: caret };
    const placed = placePointReference(deps.text(), span, refText(sel, sheet, deps.nav()));
    deps.setDraft(placed.text, placed.span.end);
    point.value = { sheetId: sheet.id, span: placed.span, sel };
  }

  /** Where the next arrow starts from: the point on show, else the shown
   * sheet's active cell (the edited cell on its own sheet). */
  function base(nav: NavContext): SheetSelection {
    const current = point.value;
    const sheet = deps.shownSheet();
    if (current && sheet && current.sheetId === sheet.id) {
      return current.sel;
    }
    return selectCell(deps.shownSelection().active, nav);
  }

  /** An arrow key while pointing is allowed: moves, extends or jumps the
   * point. True when it was taken. */
  function arrow(event: KeyboardEvent): boolean {
    if (!event.key.startsWith("Arrow") || !allowed()) {
      return false;
    }
    const nav = deps.nav();
    const next = navigateByKey(
      base(nav),
      { key: event.key, shift: event.shiftKey, mod: primaryModifier(event), alt: event.altKey, ctrl: event.ctrlKey },
      nav,
    );
    if (!next) {
      return false;
    }
    place(next);
    return true;
  }

  /** A press on a cell while editing: points when allowed (Shift extends
   * the point). True when it was taken. */
  function press(cell: { row: number; col: number }, shift: boolean): boolean {
    if (!allowed()) {
      return false;
    }
    const nav = deps.nav();
    const current = point.value;
    const sheet = deps.shownSheet();
    const sameSheet = current !== null && sheet !== null && current.sheetId === sheet.id;
    place(shift && sameSheet ? extendTo(current.sel, cell, nav) : selectCell(cell, nav));
    dragging = true;
    return true;
  }

  /** The pointer moved to `cell` during a pointing drag. */
  function drag(cell: { row: number; col: number }): void {
    const current = point.value;
    if (dragging && current && current.sheetId === deps.shownSheet()?.id) {
      place(extendTo(current.sel, cell, deps.nav()));
    }
  }

  function release(): void {
    dragging = false;
  }

  /** Ends the current point (typing, commit, cancel). */
  function end(): void {
    point.value = null;
    dragging = false;
  }

  /** F4: cycles the `$` markers of the reference at the caret. */
  function cycle(): void {
    if (!isFormula()) {
      return;
    }
    const next = cycleAbsolute(deps.text(), deps.caret());
    end();
    deps.setDraft(next.text, next.caret);
  }

  /** The coloured references of the draft. */
  const references = computed(() => (isFormula() ? referenceSpans(deps.text()) : []));
  const spans: ComputedRef<DraftSpan[]> = computed(() =>
    references.value.map((span) => ({ start: span.start, end: span.end, colorIndex: span.colorIndex })),
  );

  /** The references on the sheet on show, outlined in their colours. */
  const highlights = computed((): RefHighlight[] => {
    const shown = deps.shownSheet();
    const edited = deps.editedSheet();
    if (!shown || !edited) {
      return [];
    }
    const current = point.value;
    return references.value.flatMap((span) => {
      const name = span.ref.sheet ?? edited.name;
      if (span.ref.sheetEnd !== undefined || name.toLowerCase() !== shown.name.toLowerCase()) {
        return [];
      }
      const dashed = current !== null && current.span.start === span.start && current.span.end === span.end;
      return [{ range: referenceRange(span.ref), colorIndex: span.colorIndex, dashed }];
    });
  });

  return {
    point,
    pointing: computed(() => point.value !== null),
    allowed,
    arrow,
    press,
    drag,
    release,
    end,
    cycle,
    spans,
    highlights,
  };
}

export type PointMode = ReturnType<typeof usePointMode>;
