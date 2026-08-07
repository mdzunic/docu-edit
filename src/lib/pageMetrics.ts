// DOM measurement for the pagination engine (SPEC-028): line boxes, block
// margins and line-start character offsets, read from the rendered canvas
// with Range hit-testing. Framework-free: needs a DOM, but no Vue (see
// domCaret.ts). Every reader degrades gracefully where the environment
// cannot measure (jsdom): a block then reads as a single line of its
// element height — which collapses pagination to one page, exactly what
// the unit-test environment should see.
import type { BlockMetrics, LineBox } from "./pagination";

function lineBoxesOf(element: Element, elementRect: DOMRect): LineBox[] {
  const range = element.ownerDocument.createRange();
  range.selectNodeContents(element);
  let rects: DOMRect[] = [];
  try {
    rects = Array.from(range.getClientRects());
  } catch {
    rects = [];
  }
  const usable = rects
    .filter((rect) => rect.height > 0)
    .sort((a, b) => a.top - b.top || a.left - b.left);
  if (usable.length === 0) {
    return [{ top: 0, height: elementRect.height }];
  }
  const lines: LineBox[] = [];
  let top = usable[0].top;
  let bottom = usable[0].bottom;
  for (const rect of usable) {
    if (rect.top < bottom - 1) {
      // Vertical overlap with the current line (sub/superscript runs,
      // taller inline fragments): one visual line.
      top = Math.min(top, rect.top);
      bottom = Math.max(bottom, rect.bottom);
    } else {
      lines.push({ top: top - elementRect.top, height: bottom - top });
      top = rect.top;
      bottom = rect.bottom;
    }
  }
  lines.push({ top: top - elementRect.top, height: bottom - top });
  return lines;
}

/** Measures one rendered block element: border-box height, collapsible
 * margins, and its visual line boxes (element-relative px). `splittable`
 * false (tables, images, breaks) skips line detection — the engine treats
 * an empty line list as an atom. */
export function measureBlockElement(
  element: Element,
  contentWidthPx: number,
  splittable: boolean,
): BlockMetrics {
  const rect = element.getBoundingClientRect();
  const view = element.ownerDocument.defaultView;
  const style = view ? view.getComputedStyle(element) : null;
  return {
    contentWidthPx,
    heightPx: rect.height,
    marginTopPx: style ? Number.parseFloat(style.marginTop) || 0 : 0,
    marginBottomPx: style ? Number.parseFloat(style.marginBottom) || 0 : 0,
    lines: splittable ? lineBoxesOf(element, rect) : [],
  };
}

/** A paragraph rendered as several page slices measures as the
 * concatenation of its slices: line boxes stack in order (slice boundaries
 * are line boundaries by construction), margins come from the outer
 * edges. */
export function combineSliceMetrics(slices: BlockMetrics[]): BlockMetrics {
  if (slices.length === 1) {
    return slices[0];
  }
  const lines: LineBox[] = [];
  let height = 0;
  for (const slice of slices) {
    for (const line of slice.lines) {
      lines.push({ top: height + line.top, height: line.height });
    }
    height += slice.heightPx;
  }
  return {
    contentWidthPx: slices[0].contentWidthPx,
    heightPx: height,
    marginTopPx: slices[0].marginTopPx,
    marginBottomPx: slices[slices.length - 1].marginBottomPx,
    lines,
  };
}

interface TextSpan {
  node: Text;
  start: number;
  end: number;
}

function textSpansOf(element: Element): TextSpan[] {
  const walker = element.ownerDocument.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  const spans: TextSpan[] = [];
  let offset = 0;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    spans.push({ node, start: offset, end: offset + node.data.length });
    offset += node.data.length;
  }
  return spans;
}

function pointFor(spans: TextSpan[], offset: number): { node: Text; offset: number } | null {
  for (const span of spans) {
    if (offset < span.end || (offset === span.end && span === spans[spans.length - 1])) {
      return { node: span.node, offset: Math.max(0, offset - span.start) };
    }
  }
  return null;
}

// The element-relative top of the character at `offset` — measured as the
// rect of the range [offset, offset+1), which unambiguously belongs to the
// line the character renders on (a collapsed caret rect at a wrap boundary
// would report the previous line instead).
function charTop(
  element: Element,
  spans: TextSpan[],
  offset: number,
  length: number,
  elementTop: number,
): number | null {
  const start = pointFor(spans, Math.min(offset, Math.max(0, length - 1)));
  const end = pointFor(spans, Math.min(offset + 1, length));
  if (!start || !end) {
    return null;
  }
  const range = element.ownerDocument.createRange();
  try {
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
  } catch {
    return null;
  }
  let rects: DOMRect[] = [];
  try {
    rects = Array.from(range.getClientRects());
  } catch {
    return null;
  }
  const rect = rects.find((entry) => entry.height > 0) ?? rects[0];
  return rect ? rect.top - elementTop : null;
}

/** The character offset each of the element's visual lines starts at —
 * offsets are element-relative (a slice's own text space). Line 0 starts
 * at 0; later entries are found by binary search over per-character line
 * tops, so the cost is O(lines · log length). Falls back to proportional
 * estimates where the environment cannot measure a character. */
export function lineStartOffsets(element: Element, lines: LineBox[]): number[] {
  const spans = textSpansOf(element);
  const length = spans.length > 0 ? spans[spans.length - 1].end : 0;
  const offsets: number[] = [0];
  if (lines.length <= 1 || length === 0) {
    return offsets;
  }
  const rect = element.getBoundingClientRect();
  for (let lineIndex = 1; lineIndex < lines.length; lineIndex += 1) {
    const threshold = lines[lineIndex].top - 1;
    let low = offsets[lineIndex - 1];
    let high = length - 1;
    let found = -1;
    while (low <= high) {
      const mid = (low + high) >> 1;
      const top = charTop(element, spans, mid, length, rect.top);
      if (top === null) {
        found = -1;
        break;
      }
      if (top >= threshold) {
        found = mid;
        high = mid - 1;
      } else {
        low = mid + 1;
      }
    }
    if (found < 0) {
      // Unmeasurable: distribute the remaining characters evenly across
      // the remaining lines rather than losing the split entirely.
      const remaining = lines.length - lineIndex + 1;
      found = offsets[lineIndex - 1] + Math.max(1, Math.floor((length - offsets[lineIndex - 1]) / remaining));
    }
    offsets.push(Math.min(Math.max(found, offsets[lineIndex - 1]), length));
  }
  return offsets;
}
