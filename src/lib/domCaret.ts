// Maps between the DOM selection inside the editable canvas and model-space
// caret positions (SPEC-010). Paragraph elements carry `data-section` /
// `data-block` attributes; the character offset within a paragraph is the
// length of the text between the paragraph's start and the DOM point, which
// a Range computes for us regardless of how the runs are split into spans.
// A paragraph split across pages (SPEC-028) renders as several elements
// sharing one address, each carrying `data-slice-start` — the character
// offset its own text begins at — so both directions of the mapping add or
// subtract that base. Framework-free on purpose: it needs a DOM, but no Vue.
import type { CaretPosition } from "./textEditing";

/** Selects the canvas's paragraph elements (shared with caretGeometry). */
export const BLOCK_SELECTOR = "[data-section][data-block]";

/** The character offset a rendered element's text starts at within its
 * block — 0 unless the paragraph is split across pages (SPEC-028). */
export function sliceStartOf(element: Element): number {
  const value = Number(element.getAttribute("data-slice-start"));
  return Number.isInteger(value) && value > 0 ? value : 0;
}

// All rendered slices of a block address in document order, with each
// slice's character window; a whole (unsplit) block is its own single
// entry.
function slicesFor(root: Element, address: string): Element[] {
  return Array.from(root.querySelectorAll(`${address}:not([data-row])`)).sort(
    (a, b) => sliceStartOf(a) - sliceStartOf(b),
  );
}

// The slice whose character window contains `offset`: the first slice the
// offset falls strictly inside, preferring the LATER slice at an exact
// boundary (the caret at a page-split boundary belongs to the line that
// starts the next page), and the last slice for an offset past the end.
function sliceForOffset(slices: Element[], offset: number): Element {
  for (let i = slices.length - 1; i >= 0; i -= 1) {
    if (offset >= sliceStartOf(slices[i])) {
      return slices[i];
    }
  }
  return slices[0];
}

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

/** A paragraph's model address under an event target (SPEC-032): the
 * section/block the element's data attributes carry, descending through
 * `cell` for a paragraph inside a table. What the context menu's
 * numbering-control entries target — the right-click point, which need not
 * be where the caret is. */
export interface ParagraphTarget {
  section: number;
  block: number;
  cell?: { row: number; col: number; block: number };
}

/** Typed `null` seed for paragraph-target variables in `.vue` scripts
 * (espree bars the annotation, same constraint as noElement). */
export function noParagraphTarget(): ParagraphTarget | null {
  return null;
}

export function paragraphFromEventTarget(
  root: Element,
  target: EventTarget | null,
): ParagraphTarget | null {
  if (!(target instanceof Node)) {
    return null;
  }
  const element = target instanceof Element ? target : target.parentElement;
  const paragraph = element?.closest("[data-section][data-block]");
  if (!paragraph || !root.contains(paragraph)) {
    return null;
  }
  const section = Number(paragraph.getAttribute("data-section"));
  const block = Number(paragraph.getAttribute("data-block"));
  if (!Number.isInteger(section) || !Number.isInteger(block)) {
    return null;
  }
  const cell = cellPathOf(paragraph);
  return cell ? { section, block, cell } : { section, block };
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
  const caret: CaretPosition = {
    section,
    block,
    offset: sliceStartOf(paragraph) + range.toString().length,
  };
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
    const slices = slicesFor(root, address);
    paragraph = slices.length > 0 ? sliceForOffset(slices, caret.offset) : null;
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
  let remaining = atomOffset ?? Math.max(0, caret.offset - sliceStartOf(paragraph));
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

/** The canvas element rendering a caret's paragraph, or null when it is
 * not in the DOM — the measurement anchor for caretGeometry's boundary-line
 * checks (SPEC-022). */
export function paragraphElementForCaret(
  root: Element,
  caret: CaretPosition,
): Element | null {
  const address = `[data-section="${caret.section}"][data-block="${caret.block}"]`;
  if (caret.cell) {
    return root.querySelector(
      `${address}[data-row="${caret.cell.row}"][data-col="${caret.cell.col}"]` +
        `[data-cell-block="${caret.cell.block}"]`,
    );
  }
  const slices = slicesFor(root, address);
  return slices.length > 0 ? sliceForOffset(slices, caret.offset) : null;
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
