// Geometry between viewport points and model caret positions (SPEC-011):
// where does a mouse click or a PageUp/PageDown jump land in the document?
// Built on the browser's caret-from-point hit testing with a paragraph-rect
// fallback, so clicks in page margins and empty space still resolve to the
// nearest text position. Framework-free: needs a DOM, but no Vue.
import {
  BLOCK_SELECTOR,
  caretFromDomPoint,
  cellPathOf,
  domPointForCaret,
  paragraphElementForCaret,
  type DomPoint,
} from "./domCaret";
import type { CaretPosition } from "./textEditing";

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), Math.max(min, max));
}

// Chromium ships caretRangeFromPoint (typed on Document by lib.dom), Firefox
// ships caretPositionFromPoint (the standard name, not in lib.dom yet);
// jsdom ships neither, which the callers' fallbacks cover.
function caretPointAt(doc: Document, x: number, y: number): DomPoint | null {
  if (typeof doc.caretRangeFromPoint === "function") {
    const range = doc.caretRangeFromPoint(x, y);
    return range ? { node: range.startContainer, offset: range.startOffset } : null;
  }
  const positional = doc as Document & {
    caretPositionFromPoint?: (
      x: number,
      y: number,
    ) => { offsetNode: Node; offset: number } | null;
  };
  if (typeof positional.caretPositionFromPoint === "function") {
    const position = positional.caretPositionFromPoint(x, y);
    return position ? { node: position.offsetNode, offset: position.offset } : null;
  }
  return null;
}

/** The model caret at an exact viewport point, or null when the point does
 * not hit test into one of the root's paragraphs (outside the viewport, or
 * over other UI). PageUp/PageDown uses this strict form so an unresolvable
 * jump can fall back to the document's start/end. */
export function caretAtClientPoint(
  root: Element,
  x: number,
  y: number,
): CaretPosition | null {
  const point = caretPointAt(root.ownerDocument, x, y);
  return point ? caretFromDomPoint(root, point.node, point.offset) : null;
}

/** The model caret NEAREST a viewport point (AC-1): the point is clamped
 * into the content area and into the vertical band of the closest
 * paragraph, then hit tested; when hit testing is unavailable or misses,
 * the closest paragraph's start or end is chosen by which side of it the
 * point lies on. Null only when the canvas has no paragraphs. */
export function caretNearClientPoint(
  root: Element,
  x: number,
  y: number,
): CaretPosition | null {
  const blocks = Array.from(root.querySelectorAll(BLOCK_SELECTOR));
  if (blocks.length === 0) {
    return null;
  }
  const rootRect = root.getBoundingClientRect();
  const cx = clamp(x, rootRect.left + 1, rootRect.right - 1);
  let best = blocks[0];
  let bestDistance = Infinity;
  for (const element of blocks) {
    const rect = element.getBoundingClientRect();
    const distance = y < rect.top ? rect.top - y : y > rect.bottom ? y - rect.bottom : 0;
    if (distance < bestDistance) {
      bestDistance = distance;
      best = element;
    }
  }
  const rect = best.getBoundingClientRect();
  const cy = clamp(y, rect.top + 1, rect.bottom - 1);
  const point = caretPointAt(best.ownerDocument, cx, cy);
  if (point) {
    const caret = caretFromDomPoint(root, point.node, point.offset);
    if (caret) {
      return caret;
    }
  }
  const section = Number(best.getAttribute("data-section"));
  const block = Number(best.getAttribute("data-block"));
  if (!Number.isInteger(section) || !Number.isInteger(block)) {
    return null;
  }
  const pastEnd =
    y > rect.bottom || (y >= rect.top && x > (rect.left + rect.right) / 2);
  const caret: CaretPosition = {
    section,
    block,
    offset: pastEnd ? (best.textContent ?? "").length : 0,
  };
  const cell = best.hasAttribute("data-row") ? cellPathOf(best) : undefined;
  if (cell) {
    caret.cell = cell;
  }
  return caret;
}

/** Whether a caret sits on its paragraph's first/last visual line — what
 * decides if ArrowUp/ArrowDown at a table-cell boundary should leave the
 * cell (SPEC-022 AC-13) or keep moving within the paragraph's wrapped
 * lines. True when the environment cannot measure (jsdom): a paragraph is
 * then treated as a single line. */
export function caretOnBoundaryLine(
  root: Element,
  caret: CaretPosition,
  edge: "first" | "last",
): boolean {
  const rect = caretClientRect(root, caret);
  const paragraph = paragraphElementForCaret(root, caret);
  if (!rect || !paragraph) {
    return true;
  }
  const paragraphRect = paragraph.getBoundingClientRect();
  if (rect.height <= 0) {
    return true;
  }
  if (edge === "last") {
    return paragraphRect.bottom - rect.bottom < rect.height * 0.5;
  }
  return rect.top - paragraphRect.top < rect.height * 0.5;
}

/** The measured geometry the table hover handles position from (SPEC-022
 * AC-20, AC-21): the hovered table's rectangle plus per-column and per-row
 * bands, all relative to the page element the overlay is absolutely
 * positioned in. */
export interface TableHandleGeometry {
  section: number;
  block: number;
  left: number;
  top: number;
  width: number;
  height: number;
  cols: Array<{ left: number; width: number }>;
  rows: Array<{ top: number; height: number }>;
}

/** Typed `null` seed for hover-geometry refs in `.vue` scripts (espree bars
 * the annotation — see domCaret.noElement). */
export function noTableHandleGeometry(): TableHandleGeometry | null {
  return null;
}

/** Measures a canvas table element for the hover handles. Null when the
 * table holds no addressable cell paragraph (never the case for a rendered
 * model table). */
export function measureTableHandles(
  pageEl: Element,
  table: Element,
): TableHandleGeometry | null {
  const paragraph = table.querySelector("[data-section][data-block]");
  if (!paragraph) {
    return null;
  }
  const section = Number(paragraph.getAttribute("data-section"));
  const block = Number(paragraph.getAttribute("data-block"));
  if (!Number.isInteger(section) || !Number.isInteger(block)) {
    return null;
  }
  const pageRect = pageEl.getBoundingClientRect();
  const tableRect = table.getBoundingClientRect();
  const firstRow = table.querySelector("tr");
  const cols = firstRow
    ? Array.from(firstRow.children).map((td) => {
        const rect = td.getBoundingClientRect();
        return { left: rect.left - pageRect.left, width: rect.width };
      })
    : [];
  const rows = Array.from(table.querySelectorAll("tr")).map((tr) => {
    const rect = tr.getBoundingClientRect();
    return { top: rect.top - pageRect.top, height: rect.height };
  });
  return {
    section,
    block,
    left: tableRect.left - pageRect.left,
    top: tableRect.top - pageRect.top,
    width: tableRect.width,
    height: tableRect.height,
    cols,
    rows,
  };
}

/** The viewport rectangle of a model caret — where the blinking caret would
 * paint. Falls back to the paragraph element's own rect when the position
 * has no measurable text (an empty paragraph). Null when the paragraph is
 * not in the DOM or the environment cannot measure (jsdom). */
export function caretClientRect(
  root: Element,
  caret: CaretPosition,
): DOMRect | null {
  const point = domPointForCaret(root, caret);
  if (!point) {
    return null;
  }
  const range = root.ownerDocument.createRange();
  try {
    range.setStart(point.node, point.offset);
  } catch {
    return null;
  }
  range.collapse(true);
  let rect: DOMRect | null = null;
  try {
    rect = range.getClientRects()[0] ?? range.getBoundingClientRect();
  } catch {
    rect = null;
  }
  if (rect && (rect.top !== 0 || rect.left !== 0 || rect.height !== 0)) {
    return rect;
  }
  const element =
    point.node instanceof Element ? point.node : point.node.parentElement;
  const elementRect = element?.getBoundingClientRect() ?? null;
  if (
    elementRect &&
    (elementRect.top !== 0 || elementRect.left !== 0 || elementRect.height !== 0)
  ) {
    return elementRect;
  }
  return null;
}
