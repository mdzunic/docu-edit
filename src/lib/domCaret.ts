// Maps between the DOM selection inside the editable canvas and model-space
// caret positions (SPEC-010). Paragraph elements carry `data-section` /
// `data-block` attributes; the character offset within a paragraph is the
// length of the text between the paragraph's start and the DOM point, which
// a Range computes for us regardless of how the runs are split into spans.
// Framework-free on purpose: it needs a DOM, but no Vue.
import type { CaretPosition } from "./textEditing";

/** Selects the canvas's paragraph elements (shared with caretGeometry). */
export const BLOCK_SELECTOR = "[data-section][data-block]";

/** Typed `null` seed for element variables in `.vue` scripts, which espree
 * bars from writing the `Element | null` annotation themselves. */
export function noElement(): Element | null {
  return null;
}

/** Typed `null` seed for EventTarget parameters in `.vue` scripts (same
 * espree constraint as noElement). */
export function noEventTarget(): EventTarget | null {
  return null;
}

/** A table cell as an event target resolves it (SPEC-022): the containing
 * table's model address plus the cell's row/column. */
export interface TableCellTarget {
  section: number;
  block: number;
  row: number;
  col: number;
}

/** Typed `null` seed for cell-target variables in `.vue` scripts (espree
 * bars the annotation, same constraint as noElement). */
export function noTableCellTarget(): TableCellTarget | null {
  return null;
}

/** The table cell an event target sits in, or null outside tables. The
 * table's model address is read off the cell's first paragraph — the td
 * itself carries only its row/column (SPEC-022). */
export function tableCellFromEventTarget(
  root: Element,
  target: EventTarget | null,
): TableCellTarget | null {
  if (!(target instanceof Node)) {
    return null;
  }
  const element = target instanceof Element ? target : target.parentElement;
  const td = element?.closest("td[data-row]");
  if (!td || !root.contains(td)) {
    return null;
  }
  const paragraph = td.querySelector("[data-section][data-block]");
  if (!paragraph) {
    return null;
  }
  const section = Number(paragraph.getAttribute("data-section"));
  const block = Number(paragraph.getAttribute("data-block"));
  const row = Number(td.getAttribute("data-row"));
  const col = Number(td.getAttribute("data-col"));
  if (![section, block, row, col].every(Number.isInteger)) {
    return null;
  }
  return { section, block, row, col };
}

/** The cell path a canvas paragraph element carries (SPEC-022): cell
 * paragraphs hold their table's section/block address plus
 * data-row/data-col/data-cell-block; top-level paragraphs have no
 * data-row. Exported for caretGeometry's fallback path. */
export function cellPathOf(
  paragraph: Element,
): { row: number; col: number; block: number } | undefined {
  const row = Number(paragraph.getAttribute("data-row"));
  const col = Number(paragraph.getAttribute("data-col"));
  const block = Number(paragraph.getAttribute("data-cell-block"));
  if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(block)) {
    return undefined;
  }
  return { row, col, block };
}

/** The model caret for a DOM point (node + offset), or null when the point
 * is not inside one of the root's paragraph elements. */
export function caretFromDomPoint(
  root: Element,
  node: Node,
  offset: number,
): CaretPosition | null {
  const element = node instanceof Element ? node : node.parentElement;
  const paragraph = element?.closest(BLOCK_SELECTOR);
  if (!paragraph || !root.contains(paragraph)) {
    return null;
  }
  const section = Number(paragraph.getAttribute("data-section"));
  const block = Number(paragraph.getAttribute("data-block"));
  if (!Number.isInteger(section) || !Number.isInteger(block)) {
    return null;
  }
  const range = paragraph.ownerDocument.createRange();
  range.selectNodeContents(paragraph);
  try {
    range.setEnd(node, offset);
  } catch {
    return null;
  }
  const caret: CaretPosition = { section, block, offset: range.toString().length };
  const cell = paragraph.hasAttribute("data-row") ? cellPathOf(paragraph) : undefined;
  if (cell) {
    caret.cell = cell;
  }
  return caret;
}

/** The current DOM selection as an ordered model range, or null when either
 * end lies outside the root's paragraphs. */
export function modelRangeFromSelection(
  root: Element,
  selection: Selection | null,
): { start: CaretPosition; end: CaretPosition } | null {
  if (!selection || selection.rangeCount === 0) {
    return null;
  }
  const range = selection.getRangeAt(0);
  const start = caretFromDomPoint(root, range.startContainer, range.startOffset);
  const end = caretFromDomPoint(root, range.endContainer, range.endOffset);
  if (!start || !end) {
    return null;
  }
  return { start, end };
}

/** A DOM point (node + offset) usable as a Range/Selection boundary. */
export interface DomPoint {
  node: Node;
  offset: number;
}

/** Resolves a model caret to the DOM point inside its paragraph element:
 * the text node containing the offset, the last text node's end when the
 * offset overshoots, or the paragraph element itself when it has no text.
 * Null when the addressed paragraph is not in the DOM. */
export function domPointForCaret(
  root: Element,
  caret: CaretPosition,
): DomPoint | null {
  const address = `[data-section="${caret.section}"][data-block="${caret.block}"]`;
  let paragraph: Element | null;
  let atomOffset: number | undefined;
  if (caret.cell) {
    paragraph = root.querySelector(
      `${address}[data-row="${caret.cell.row}"][data-col="${caret.cell.col}"]` +
        `[data-cell-block="${caret.cell.block}"]`,
    );
  } else {
    // Cell paragraphs carry their table's section/block address too, so a
    // top-level lookup must exclude them. When the address itself is a
    // table (an atom caret at its edge), resolve to the first or last cell
    // paragraph instead — offset 0 is the position before the table.
    paragraph = root.querySelector(`${address}:not([data-row])`);
    if (!paragraph) {
      const cellParagraphs = root.querySelectorAll(`${address}[data-row]`);
      if (cellParagraphs.length === 0) {
        return null;
      }
      // The atom offset (0 = before, 1 = after) maps to the first cell
      // paragraph's start or the last one's end, not to a character count.
      const atEnd = caret.offset > 0;
      paragraph = atEnd ? cellParagraphs[cellParagraphs.length - 1] : cellParagraphs[0];
      atomOffset = atEnd ? Number.MAX_SAFE_INTEGER : 0;
    }
  }
  if (!paragraph) {
    return null;
  }
  const walker = paragraph.ownerDocument.createTreeWalker(
    paragraph,
    NodeFilter.SHOW_TEXT,
  );
  let remaining = atomOffset ?? caret.offset;
  let lastText: Text | null = null;
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    lastText = text;
    if (remaining <= text.data.length) {
      return { node: text, offset: remaining };
    }
    remaining -= text.data.length;
  }
  if (lastText) {
    return { node: lastText, offset: lastText.data.length };
  }
  return { node: paragraph, offset: 0 };
}

function selectionFor(root: Element): Selection | null {
  return root.ownerDocument.defaultView?.getSelection() ?? null;
}

/** Places the DOM caret (a collapsed selection) at a model caret position.
 * Silently does nothing when the addressed paragraph is not in the DOM —
 * callers re-render first, so that only happens for stale carets. */
export function setDomCaret(root: Element, caret: CaretPosition): void {
  const point = domPointForCaret(root, caret);
  const selection = selectionFor(root);
  if (!point || !selection) {
    return;
  }
  const range = root.ownerDocument.createRange();
  range.setStart(point.node, point.offset);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
}

/** Sets the DOM selection to span two model positions, preserving direction:
 * `anchor` stays fixed, `focus` is the end a subsequent Shift+movement would
 * keep extending. Does nothing when either paragraph is not in the DOM. */
export function setDomSelection(
  root: Element,
  anchor: CaretPosition,
  focus: CaretPosition,
): void {
  const anchorPoint = domPointForCaret(root, anchor);
  const focusPoint = domPointForCaret(root, focus);
  const selection = selectionFor(root);
  if (!anchorPoint || !focusPoint || !selection) {
    return;
  }
  selection.setBaseAndExtent(
    anchorPoint.node,
    anchorPoint.offset,
    focusPoint.node,
    focusPoint.offset,
  );
}

/** Moves only the focus end of the DOM selection to a model position,
 * keeping the current anchor — Shift+movement's "extend the selection".
 * With no selection to extend it degrades to placing the caret. */
export function extendDomSelection(root: Element, focus: CaretPosition): void {
  const point = domPointForCaret(root, focus);
  const selection = selectionFor(root);
  if (!point || !selection) {
    return;
  }
  if (selection.rangeCount === 0) {
    setDomCaret(root, focus);
    return;
  }
  selection.extend(point.node, point.offset);
}
