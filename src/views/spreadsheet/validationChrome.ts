// Where SPEC-130's grid chrome sits and how its list answers keys — pure, so
// the .vue components (which cannot carry TS syntax) bind the results and the
// rules are tested without a DOM.
//
// Geometry is the grid's: unscaled px relative to the paper's origin (the
// top-left of the cell area, past the headers), inside the zoomed content box
// — so everything placed from the active cell's rect follows scroll and zoom.
// `bounds` is the cell area's size (the viewport minus the headers).
//
// - The input message sits below-right of the cell (its top-left at the
//   cell's bottom-right corner, a small gap down), flipping to the cell's
//   left when it would cross the right edge and above the cell when it would
//   cross the bottom.
// - The dropdown button sits against the cell's right edge, just outside it
//   (inside it when the cell touches the area's right edge), the cell's
//   height tall.
// - The list opens under the cell (above it when there is no room below and
//   there is above), at least the cell's width, eight rows tall at most —
//   and always inside the area: with room for it on neither side it takes
//   the roomier side and scrolls in what that side holds (or, when neither
//   holds a row, sits over the cell at the area's bottom); it shifts left to
//   keep its width inside the right edge, a long item cut at that edge.
// - List keys: Up/Down move by one, Home/End to the ends, PageUp/PageDown by
//   eight; with nothing selected the first move selects the first item (a
//   PageDown the eighth). A printable character selects the next item whose
//   text starts with it (case-insensitively, wrapping past the end). Enter
//   or Tab picks the selected item (closes when none); Escape and Alt+Up
//   close.

export interface ChromeRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ChromeBounds {
  width: number;
  height: number;
}

/** The input message's element id (the active cell's aria-describedby). */
export const VALIDATION_INPUT_MESSAGE_ID = "validation-input-message";
/** The dropdown list's element id (the button's aria-controls). */
export const VALIDATION_LIST_ID = "validation-list";
/** The unavailable source's one disabled option. */
export const LIST_UNAVAILABLE_TEXT = "The list source isn't available";
export const INPUT_MESSAGE_GAP_PX = 4;
export const DROPDOWN_BUTTON_PX = 16;
export const LIST_ROW_PX = 20;
export const LIST_VISIBLE_ROWS = 8;
/** The list's border, top and bottom together. */
const LIST_BORDER_PX = 2;
/** The least the list shows: one row. */
const LIST_MIN_PX = LIST_ROW_PX + LIST_BORDER_PX;

/** The input message's top-left for its measured size (see the header). */
export function placeInputMessage(cell: ChromeRect, size: { width: number; height: number }, bounds: ChromeBounds): { left: number; top: number } {
  let left = cell.left + cell.width;
  let top = cell.top + cell.height + INPUT_MESSAGE_GAP_PX;
  if (left + size.width > bounds.width) {
    left = cell.left - size.width;
  }
  if (top + size.height > bounds.height) {
    top = cell.top - size.height - INPUT_MESSAGE_GAP_PX;
  }
  return { left: Math.max(0, left), top: Math.max(0, top) };
}

/** The dropdown button's box (see the header). */
export function placeDropdownButton(cell: ChromeRect, bounds: ChromeBounds): ChromeRect {
  const outside = cell.left + cell.width;
  const left = outside + DROPDOWN_BUTTON_PX > bounds.width ? Math.max(cell.left, outside - DROPDOWN_BUTTON_PX) : outside;
  return { left, top: cell.top, width: DROPDOWN_BUTTON_PX, height: cell.height };
}

/** The list's position, width limits and height for `rows` items (see the
 * header). */
export function placeValidationList(
  cell: ChromeRect,
  rows: number,
  bounds: ChromeBounds,
): { left: number; top: number; minWidth: number; maxWidth: number; maxHeight: number } {
  const wanted = Math.min(LIST_VISIBLE_ROWS, Math.max(1, rows)) * LIST_ROW_PX + LIST_BORDER_PX;
  const below = cell.top + cell.height;
  const roomBelow = bounds.height - below;
  const roomAbove = cell.top;
  let top: number;
  let maxHeight: number;
  if (wanted <= roomBelow) {
    [top, maxHeight] = [below, wanted];
  } else if (wanted <= roomAbove) {
    [top, maxHeight] = [cell.top - wanted, wanted];
  } else if (Math.max(roomBelow, roomAbove) >= LIST_MIN_PX) {
    [top, maxHeight] = roomBelow >= roomAbove ? [below, roomBelow] : [0, roomAbove];
  } else {
    maxHeight = Math.max(LIST_MIN_PX, Math.min(wanted, bounds.height));
    top = Math.max(0, bounds.height - maxHeight);
  }
  const minWidth = cell.width + DROPDOWN_BUTTON_PX;
  const left = Math.max(0, Math.min(cell.left, bounds.width - minWidth));
  return { left, top, minWidth, maxWidth: Math.max(minWidth, bounds.width - left), maxHeight };
}

/** What a key in the list does. */
export type ListKeyAction = { kind: "select"; index: number } | { kind: "pick"; index: number } | { kind: "close" } | null;

/** The next item whose text starts with `char` after `from` (wrapping), or
 * -1. */
export function nextItemStartingWith(texts: readonly string[], from: number, char: string): number {
  const wanted = char.toLocaleLowerCase();
  for (let step = 1; step <= texts.length; step += 1) {
    const index = (((from + step) % texts.length) + texts.length) % texts.length;
    if (texts[index]!.toLocaleLowerCase().startsWith(wanted)) {
      return index;
    }
  }
  return -1;
}

/** A key on the open list (see the header); null: not the list's. */
export function listKeyAction(
  event: Pick<KeyboardEvent, "key" | "altKey" | "ctrlKey" | "metaKey">,
  texts: readonly string[],
  selected: number,
): ListKeyAction {
  const last = texts.length - 1;
  const select = (index: number): ListKeyAction => (last < 0 ? null : { kind: "select", index: Math.max(0, Math.min(last, index)) });
  if (event.key === "Escape" || (event.key === "ArrowUp" && event.altKey)) {
    return { kind: "close" };
  }
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return null;
  }
  switch (event.key) {
    case "ArrowDown":
      return select(selected < 0 ? 0 : selected + 1);
    case "ArrowUp":
      return select(selected < 0 ? 0 : selected - 1);
    case "Home":
      return select(0);
    case "End":
      return select(last);
    case "PageDown":
      return select(selected < 0 ? LIST_VISIBLE_ROWS - 1 : selected + LIST_VISIBLE_ROWS);
    case "PageUp":
      return select(selected < 0 ? 0 : selected - LIST_VISIBLE_ROWS);
    case "Enter":
    case "Tab":
      return selected >= 0 && selected <= last ? { kind: "pick", index: selected } : { kind: "close" };
    default: {
      if (Array.from(event.key).length !== 1 || event.key.trim() === "") {
        return null;
      }
      const index = nextItemStartingWith(texts, selected, event.key);
      return index >= 0 ? { kind: "select", index } : null;
    }
  }
}
