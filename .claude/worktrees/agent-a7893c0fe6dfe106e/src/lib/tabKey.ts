// SPEC-054's Tab precedence table, as a decision separated from its effects.
//
// Tab is the most overloaded key in the editor: inside a table cell it moves
// between cells and grows the table (SPEC-022 AC-10..12); at the start of a
// list item it changes list level (SPEC-018, narrowed by SPEC-054 rule 2);
// at the start of any other paragraph it indents (rule 3); anywhere else it
// types a literal tab (rule 4). Which one applies is derived entirely from
// the model and the selection, so it belongs here rather than in the
// composable — four ordered rules are worth unit tests, and they cannot have
// any while the branch that picks them is welded to `applyEdit`.
//
// The caller does the applying. `resolveTabAction` only names the outcome.
import { INDENT_STEP_CM, isCollapsed, touchedBlocks, type ModelSelection } from "./commands";
import { caretTableLocation, nextCellCaret, prevCellCaret, type TableLocation } from "./tables";
import { paragraphAt, type CaretPosition } from "./textEditing";
import type { DocumentModel } from "./documentModel";

export type TabAction =
  /** Rule 1: move the caret to the neighbouring cell. */
  | { kind: "moveCaret"; caret: CaretPosition }
  /** Rule 1, at the last cell: grow the table by a row. */
  | { kind: "appendRow"; location: TableLocation }
  /** Rule 2: promote or demote the list item. */
  | { kind: "listLevel"; delta: number }
  /** Rule 3: indent or outdent the paragraph. */
  | { kind: "indent"; deltaCm: number }
  /** Rule 4: a literal tab character. */
  | { kind: "insertTab" }
  /** Tab is swallowed: prevented, but there is nothing to do. */
  | { kind: "none" };

/** The action Tab (or Shift+Tab) means for `selection` in `model`, resolved
 * in SPEC-054's precedence order. */
export function resolveTabAction(
  model: DocumentModel,
  selection: ModelSelection,
  shiftKey: boolean,
): TabAction {
  const collapsed = isCollapsed(selection);

  // Rule 1 — inside a table cell, Tab is cell navigation and nothing else.
  if (collapsed && selection.focus.cell) {
    if (shiftKey) {
      const previous = prevCellCaret(model, selection.focus);
      return previous ? { kind: "moveCaret", caret: previous } : { kind: "none" };
    }
    const next = nextCellCaret(model, selection.focus);
    if (next) {
      return { kind: "moveCaret", caret: next };
    }
    // Past the last cell: Tab appends a row rather than leaving the table.
    const location = caretTableLocation(model, selection.focus);
    return location ? { kind: "appendRow", location } : { kind: "none" };
  }

  // Rules 2 and 3 are both gated on a caret collapsed at offset 0 — SPEC-054
  // narrowed SPEC-018's "any offset", so a Tab mid-paragraph types rather
  // than silently restructuring the list around it.
  const atStart = collapsed && selection.focus.offset === 0;
  if (atStart) {
    const inList = touchedBlocks(model, selection).some(
      (address) => !!paragraphAt(model, address)?.list,
    );
    const direction = shiftKey ? -1 : 1;
    return inList
      ? { kind: "listLevel", delta: direction }
      : { kind: "indent", deltaCm: direction * INDENT_STEP_CM };
  }

  // Rule 4 — a literal tab. Shift+Tab has nothing to insert.
  return shiftKey ? { kind: "none" } : { kind: "insertTab" };
}
