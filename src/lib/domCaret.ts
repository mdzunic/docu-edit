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

/** Marks a rendered inline element as a caret atom (SPEC-031's note
 * reference marks): its rendered text may be any length (a note number like
 * "12"), but the model text it stands for is this attribute's value (the
 * one-character placeholder). Both mapping directions below — and every
 * consumer that equates rendered text with model text — count the attribute
 * value's length instead of the rendered text's. */
export const CARET_ATOM_ATTRIBUTE = "data-caret-atom";
const CARET_ATOM_SELECTOR = "[data-caret-atom]";

function atomModelLength(atom: Element): number {
  return (atom.getAttribute(CARET_ATOM_ATTRIBUTE) ?? "").length || 1;
}

/** The model text a rendered canvas subtree stands for: its text content
 * with every caret atom's rendered text replaced by the atom's model text,
 * and every `data-canvas-ignore` subtree (header/footer bands and note
 * areas — content outside the body flow) skipped. Identical to
 * `textContent` for plain body content. */
export function canvasModelText(root: Element): string {
  if (root.hasAttribute("data-canvas-ignore")) {
    return "";
  }
  if (root.hasAttribute(CARET_ATOM_ATTRIBUTE)) {
    return root.getAttribute(CARET_ATOM_ATTRIBUTE) ?? "";
  }
  let text = "";
  for (const node of Array.from(root.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      text += (node as Text).data;
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      text += canvasModelText(node as Element);
    }
  }
  return text;
}

/** The model text length a rendered element stands for (see
 * canvasModelText) — what `element.textContent.length` used to mean before
 * caret atoms existed. */
export function modelTextLengthOf(element: Element): number {
  return canvasModelText(element).length;
}

/** Maps line-start offsets measured in an element's RENDERED text (see
 * pageMetrics.lineStartOffsets) into model text space: offsets past an
 * atom's rendered label shrink by the label's extra length; an offset
 * inside a label snaps to the atom's start. */
export function renderedOffsetsToModel(element: Element, offsets: number[]): number[] {
  const atoms = Array.from(element.querySelectorAll(CARET_ATOM_SELECTOR));
  if (atoms.length === 0) {
    return offsets;
  }
  // Each atom's rendered start offset within the element, in document
  // order, with its rendered/model lengths.
  const doc = element.ownerDocument;
  const spans = atoms.map((atom) => {
    const range = doc.createRange();
    range.selectNodeContents(element);
    range.setEndBefore(atom);
    const start = range.toString().length;
    return {
      start,
      rendered: (atom.textContent ?? "").length,
      model: atomModelLength(atom),
    };
  });
  return offsets.map((offset) => {
    let mapped = offset;
    for (const span of spans) {
      if (offset >= span.start + span.rendered) {
        mapped -= span.rendered - span.model;
      } else if (offset > span.start) {
        // Inside the rendered label: snap to the atom's start.
        mapped -= offset - span.start;
        break;
      } else {
        break;
      }
    }
    return mapped;
  });
}

/** The character offset a rendered element's text starts at within its
 * block — 0 unless the paragraph is split across pages (SPEC-028). */
export function sliceStartOf(element: Element): number {
  const value = Number(element.getAttribute("data-slice-start"));
  return Number.isInteger(value) && value > 0 ? value : 0;
}

/** True when `element` sits inside a `data-canvas-ignore` subtree WITHIN
 * `root` — a note area or header/footer band (SPEC-029/031) whose
 * paragraphs reuse synthetic section-0 addressing that must never shadow
 * the root's own blocks. Relative on purpose: when `root` itself is the
 * editable surface inside such a subtree (a note body, a band), the
 * ignored ancestor lies outside it and nothing is excluded. */
export function inIgnoredSubtree(root: Element, element: Element): boolean {
  const ignored = element.closest("[data-canvas-ignore]");
  return !!ignored && ignored !== root && root.contains(ignored);
}

/** Per-pass memo of block address → rendered slices for `domPointForCaret`.
 * Resolving a caret costs one `querySelectorAll` over the whole realized
 * document; a caller that resolves MANY carets in one synchronous pass (the
 * spell-check overlay re-measures every visible issue on every scroll idle
 * slot — SPEC-047 AC-8) would otherwise pay that scan twice per range. Only
 * valid for one root and only while the DOM is unchanged: build one, use it
 * synchronously, drop it. */
export type CaretSliceCache = Map<string, Element[]>;

/** A fresh {@link CaretSliceCache} — a named constructor so `.vue` scripts,
 * which cannot write the generic annotation, can seed one too. */
export function createCaretSliceCache(): CaretSliceCache {
  return new Map();
}

// All rendered slices of a block address in document order, with each
// slice's character window; a whole (unsplit) block is its own single
// entry.
function slicesFor(root: Element, address: string, cache?: CaretSliceCache): Element[] {
  const cached = cache?.get(address);
  if (cached) {
    return cached;
  }
  const slices = Array.from(root.querySelectorAll(`${address}:not([data-row])`))
    .filter((el) => !inIgnoredSubtree(root, el))
    .sort((a, b) => sliceStartOf(a) - sliceStartOf(b));
  cache?.set(address, slices);
  return slices;
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
  if (!td || !root.contains(td) || inIgnoredSubtree(root, td)) {
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
  if (!paragraph || !root.contains(paragraph) || inIgnoredSubtree(root, paragraph)) {
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
  if (!paragraph || !root.contains(paragraph) || inIgnoredSubtree(root, paragraph)) {
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
  // Caret atoms (SPEC-031): an atom before the point contributed its whole
  // rendered label to the range's text but stands for its model text; a
  // point inside an atom snaps to just after it.
  let adjusted: number | null = null;
  let adjustment = 0;
  for (const atom of Array.from(paragraph.querySelectorAll(CARET_ATOM_SELECTOR))) {
    const renderedLength = (atom.textContent ?? "").length;
    const modelLength = atomModelLength(atom);
    const atomRange = paragraph.ownerDocument.createRange();
    atomRange.selectNodeContents(atom);
    if (atomRange.compareBoundaryPoints(Range.END_TO_END, range) <= 0) {
      adjustment += renderedLength - modelLength;
      continue;
    }
    if (atomRange.compareBoundaryPoints(Range.END_TO_START, range) < 0) {
      const before = paragraph.ownerDocument.createRange();
      before.selectNodeContents(paragraph);
      before.setEnd(atomRange.startContainer, atomRange.startOffset);
      adjusted = before.toString().length - adjustment + modelLength;
    }
    break; // Atoms come in document order: the rest start after the point.
  }
  const caret: CaretPosition = {
    section,
    block,
    offset: sliceStartOf(paragraph) + (adjusted ?? range.toString().length - adjustment),
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
 * Null when the addressed paragraph is not in the DOM. Pass a
 * {@link CaretSliceCache} when resolving many carets over one unchanged DOM
 * to share the per-block element lookup. */
export function domPointForCaret(
  root: Element,
  caret: CaretPosition,
  sliceCache?: CaretSliceCache,
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
    const slices = slicesFor(root, address, sliceCache);
    paragraph = slices.length > 0 ? sliceForOffset(slices, caret.offset) : null;
    if (!paragraph) {
      const cellParagraphs = Array.from(root.querySelectorAll(`${address}[data-row]`)).filter(
        (el) => !inIgnoredSubtree(root, el),
      );
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
  // Walk text nodes and caret atoms in order. An atom is opaque: it counts
  // as its model text's length and the point lands before or after it,
  // never inside its rendered label (SPEC-031).
  const units: Array<{ text: Text } | { atom: Element }> = [];
  const visit = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      units.push({ text: node as Text });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }
    if ((node as Element).hasAttribute(CARET_ATOM_ATTRIBUTE)) {
      units.push({ atom: node as Element });
      return;
    }
    for (const child of Array.from(node.childNodes)) {
      visit(child);
    }
  };
  for (const child of Array.from(paragraph.childNodes)) {
    visit(child);
  }
  const afterAtom = (atom: Element): DomPoint => {
    const parent = atom.parentNode as Node;
    const index = Array.prototype.indexOf.call(parent.childNodes, atom);
    return { node: parent, offset: index + 1 };
  };
  let remaining = atomOffset ?? Math.max(0, caret.offset - sliceStartOf(paragraph));
  let last: { text: Text } | { atom: Element } | null = null;
  for (const unit of units) {
    if ("text" in unit) {
      if (remaining <= unit.text.data.length) {
        return { node: unit.text, offset: remaining };
      }
      remaining -= unit.text.data.length;
    } else {
      const modelLength = atomModelLength(unit.atom);
      if (remaining < modelLength) {
        const parent = unit.atom.parentNode as Node;
        const index = Array.prototype.indexOf.call(parent.childNodes, unit.atom);
        return { node: parent, offset: index };
      }
      remaining -= modelLength;
    }
    last = unit;
  }
  if (last && "text" in last) {
    return { node: last.text, offset: last.text.data.length };
  }
  if (last) {
    return afterAtom(last.atom);
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
