// The find bar's keys (SPEC-134 "Keys"): which action a keydown asks of
// useSheetFind, matched through the text registry's findShortcut
// (find.open, find.openReplace). Cmd/Ctrl+F opens the bar, Cmd/Ctrl+H opens it
// with the replace row; on macOS, while the bar is open, Cmd+G steps to the
// next match and Cmd+Shift+G to the previous (Go To is Control+G there, as
// in Excel for Mac); Escape on the bar's buttons, toggles and selects or in
// Find all's results closes it (its text fields close it themselves). Nothing
// is asked from inside a dialog. The caller prevents the default of every
// action it is given — and, while the cell editor is open, does nothing else.
import { findShortcut } from "../../lib/shortcuts";

export type FindKeyAction = "open" | "openReplace" | "next" | "prev" | "close";

const inDialog = (target: EventTarget | null): boolean => target instanceof Element && target.closest('[role="dialog"]') !== null;
const inFindSurface = (target: EventTarget | null): boolean =>
  target instanceof Element && target.closest('[data-testid="find-bar"], [data-testid="sheet-find-results"]') !== null;

/** The action `event` asks for, or null (see the header). */
export function findKeyAction(event: KeyboardEvent, barOpen: boolean, isMac: boolean): FindKeyAction | null {
  if (event.defaultPrevented || inDialog(event.target)) {
    return null;
  }
  if (event.key === "Escape") {
    return barOpen && inFindSurface(event.target) && !(event.target instanceof HTMLInputElement) ? "close" : null;
  }
  if (barOpen && isMac && event.metaKey && !event.ctrlKey && !event.altKey && event.key.toLowerCase() === "g") {
    return event.shiftKey ? "prev" : "next";
  }
  const id = findShortcut(event, isMac)?.id;
  return id === "find.open" ? "open" : id === "find.openReplace" ? "openReplace" : null;
}
