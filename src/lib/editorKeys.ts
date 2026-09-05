// The contextual key meanings the canvas keydown map reads (SPEC-021's
// useKeyboardMap), as plain predicates over a KeyboardEvent.
//
// These are deliberately NOT in shortcuts.ts. That file is the registry of
// named commands a user could look up in the shortcuts reference; its header
// excludes exactly this kind of thing — the contextual conventions (Escape
// dismisses, Delete removes whatever is selected, an arrow nudges) that no
// word processor lists as a command. They still need somewhere testable to
// live, because each one encodes a real decision about which modifiers must
// NOT be held, and getting that wrong silently steals a native browser key.
import type { NudgeDirection } from "./core/shapes";

/** Whether a keydown would type or delete content — a plain printable
 * character, or Enter/Backspace/Delete without a modifier. This is the set
 * the selected-TOC and selected-table-of-figures guards intercept (SPEC-033
 * AC-15, SPEC-055). Excludes navigation, Tab and every modifier combo, so
 * Ctrl/Cmd+C and the arrow keys still work around a selected atom. */
export function isContentKey(event: KeyboardEvent): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return false;
  }
  if (event.key === "Enter" || event.key === "Backspace" || event.key === "Delete") {
    return true;
  }
  return event.key.length === 1;
}

/** Whether the key removes the current selection. Both keys delete a
 * selected shape/image/chart/TOC outright — there is no backwards-vs-forwards
 * distinction to make when the target is a whole atom rather than a caret. */
export function isDeleteKey(event: KeyboardEvent): boolean {
  return event.key === "Delete" || event.key === "Backspace";
}

/** Ctrl/Cmd+D, the duplicate combo for a selected shape or chart (SPEC-051
 * AC-41, SPEC-052). Shift must not be held: Ctrl+Shift+D is free for other
 * bindings, and matching it here would swallow one. */
export function isDuplicateCombo(event: KeyboardEvent): boolean {
  return (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "d";
}

/** A plain Ctrl/Cmd+V. Observed rather than bound (SPEC-016): the browser
 * owns this key and fires its own paste event, so the map only watches it to
 * clear a stale plain-paste flag. Ctrl/Cmd+Shift+V is a real registry
 * command and must not match here. */
export function isPlainPasteCombo(event: KeyboardEvent): boolean {
  return (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "v";
}

/** The nudge direction an arrow key means for a selected shape (SPEC-051
 * AC-43), or null for any other key. */
export function nudgeDirectionForKey(key: string): NudgeDirection | null {
  switch (key) {
    case "ArrowUp":
      return "up";
    case "ArrowDown":
      return "down";
    case "ArrowLeft":
      return "left";
    case "ArrowRight":
      return "right";
    default:
      return null;
  }
}

/** An arrow key with no modifier at all — the only shape that gets the
 * cell-boundary and cross-column caret handling (SPEC-022 AC-13/14,
 * SPEC-030 AC-20/21). Shift would be extending a selection and Ctrl/Alt are
 * word/paragraph jumps, all of which the browser handles natively. */
export function isUnmodifiedArrowKey(event: KeyboardEvent): boolean {
  if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) {
    return false;
  }
  return nudgeDirectionForKey(event.key) !== null;
}
