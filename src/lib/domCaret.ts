// Maps between the DOM selection inside the editable canvas and model-space
// caret positions (SPEC-010). Paragraph elements carry `data-section` /
// `data-block` attributes; the character offset within a paragraph is the
// length of the text between the paragraph's start and the DOM point, which
// a Range computes for us regardless of how the runs are split into spans.
// Framework-free on purpose: it needs a DOM, but no Vue.
import type { CaretPosition } from "./textEditing";

const BLOCK_SELECTOR = "[data-section][data-block]";

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
  return { section, block, offset: range.toString().length };
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

/** Places the DOM caret (a collapsed selection) at a model caret position.
 * Silently does nothing when the addressed paragraph is not in the DOM —
 * callers re-render first, so that only happens for stale carets. */
export function setDomCaret(root: Element, caret: CaretPosition): void {
  const paragraph = root.querySelector(
    `[data-section="${caret.section}"][data-block="${caret.block}"]`,
  );
  if (!paragraph) {
    return;
  }
  const doc = paragraph.ownerDocument;
  const view = doc.defaultView;
  if (!view) {
    return;
  }
  const walker = doc.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let remaining = caret.offset;
  let target: Text | null = null;
  let targetOffset = 0;
  let lastText: Text | null = null;
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    lastText = text;
    if (remaining <= text.data.length) {
      target = text;
      targetOffset = remaining;
      break;
    }
    remaining -= text.data.length;
  }
  const range = doc.createRange();
  if (target) {
    range.setStart(target, targetOffset);
  } else if (lastText) {
    range.setStart(lastText, lastText.data.length);
  } else {
    range.setStart(paragraph, 0);
  }
  range.collapse(true);
  const selection = view.getSelection();
  if (!selection) {
    return;
  }
  selection.removeAllRanges();
  selection.addRange(range);
}
