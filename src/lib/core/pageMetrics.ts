// DOM measurement for a paginated layout engine (SPEC-028): line boxes,
// block margins and line-start character offsets, read from the rendered
// canvas with Range hit-testing. Product-neutral: measures any rendered
// block element — the pagination engine that consumes these metrics stays
// in lib/pagination.ts. Framework-free: needs a DOM, but no Vue (see
// domCaret.ts). Every reader degrades gracefully where the environment
// cannot measure (jsdom): a block then reads as a single line of its
// element height — which collapses pagination to one page, exactly what
// the unit-test environment should see.

/** One rendered line of a paragraph, relative to the block's border-box
 * top, in px. */
export interface LineBox {
  top: number;
  height: number;
}

/** A block's measured (or estimated) geometry at a given content width. */
export interface BlockMetrics {
  /** The content width the block was measured at, px — stale when the
   * section's content width changes. */
  contentWidthPx: number;
  /** Border-box height, px; margins excluded. */
  heightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
  /** Line boxes for a paragraph, in document order; empty for atoms
   * (tables, images, breaks) and for paragraphs that cannot be split. */
  lines: LineBox[];
  /** True when these metrics are an estimate rather than a DOM measurement
   * — the caller re-measures such blocks and re-runs the layout. */
  estimated?: boolean;
}

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
    .toSorted((a, b) => a.top - b.top || a.left - b.left);
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

/** Marks a rendered table row as a repeat of the table's header run rather
 * than a row of its own (SPEC-076 AC-7). Every consumer that walks the
 * canvas — measurement here, the model-text mapping in domCaret — skips such
 * a row, because the repeat is a layout artefact with no model presence. */
export const REPEATED_HEADER_ATTRIBUTE = "data-repeated-header";

/** Measures a rendered table as one line box per MODEL row (SPEC-076): the
 * unit a table splits at across pages is a row, exactly as a paragraph's is
 * a line, so the pagination engine needs no table-specific arithmetic.
 *
 * Repeated header rows are skipped and the remaining rows are re-based to
 * start at 0, so a table rendered as several page slices combines through
 * `combineSliceMetrics` without double-counting the header run each
 * continuation redraws (AC-11: the header run is measured once per table,
 * never once per page).
 *
 * Degrades in jsdom exactly as `measureBlockElement` does: unmeasurable rows
 * read as zero-height boxes, which collapses the table onto one page.
 */
export function measureTableElement(element: Element, contentWidthPx: number): BlockMetrics {
  const view = element.ownerDocument.defaultView;
  const style = view ? view.getComputedStyle(element) : null;
  const rows = Array.from(element.querySelectorAll(":scope > tbody > tr")).filter(
    (row) => !row.hasAttribute(REPEATED_HEADER_ATTRIBUTE),
  );
  const lines: LineBox[] = [];
  let height = 0;
  for (const row of rows) {
    const rect = row.getBoundingClientRect();
    lines.push({ top: height, height: rect.height });
    height += rect.height;
  }
  return {
    contentWidthPx,
    heightPx: height,
    marginTopPx: style ? Number.parseFloat(style.marginTop) || 0 : 0,
    marginBottomPx: style ? Number.parseFloat(style.marginBottom) || 0 : 0,
    lines,
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
