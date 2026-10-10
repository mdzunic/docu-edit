// The DOM side of cell editing (SPEC-122 "Editor DOM"): the in-cell editor
// and the editable formula bar are contenteditable hosts edited
// normalise-after-input (SPEC-081's pattern) — after the browser applies an
// edit the host is read back to plain text, re-rendered as text plus one
// coloured `<span>` per reference, and the caret restored by character
// offset. Framework-free so it is testable without a component.
//
// - Rendering writes text nodes (line breaks as literal "\n" under
//   `white-space: pre-wrap`), one `<span class="sheet-ref">` per reference
//   coloured from `--sheet-ref-1..8` by its colour index, and — when the text
//   ends with a line break — a trailing `<br data-draft-end>` so the empty
//   last line shows and can hold the caret.
// - Reading back accepts whatever the browser made of that: `<br>` → "\n"
//   (the trailing marker excepted), a block element (`<div>`, `<p>`) starts a
//   new line, every other element contributes its text.
// - Offsets are UTF-16 offsets into the read-back text: the caret's offset
//   is the read-back length of everything before it, so reading and placing
//   always agree.

/** A coloured reference in the draft text (pointMode's referenceSpans). */
export interface DraftSpan {
  start: number;
  end: number;
  colorIndex: number;
}

/** The number of `--sheet-ref-N` colour tokens. */
export const SHEET_REF_COLORS = 8;

/** The CSS colour of a reference colour index (0-based). */
export function refColor(colorIndex: number): string {
  return `var(--sheet-ref-${(((colorIndex % SHEET_REF_COLORS) + SHEET_REF_COLORS) % SHEET_REF_COLORS) + 1})`;
}

const BLOCK_TAGS = new Set(["DIV", "P", "LI", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "PRE"]);
const END_MARKER = "data-draft-end";

/** Appends the plain text of `node`'s children to `out`. */
function collect(node: Node, out: string[], root: Node): void {
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 3) {
      out.push((child as Text).data);
    } else if (child.nodeType === 1) {
      const element = child as Element;
      if (element.tagName === "BR") {
        const isEndMarker = element.hasAttribute(END_MARKER) && !hasFollowingContent(element, root);
        if (!isEndMarker) {
          out.push("\n");
        }
      } else {
        if (BLOCK_TAGS.has(element.tagName)) {
          const sofar = out.join("");
          if (sofar !== "" && !sofar.endsWith("\n")) {
            out.push("\n");
          }
        }
        collect(element, out, root);
      }
    }
  }
}

/** Whether anything with text follows `node` inside `root`. */
function hasFollowingContent(node: Node, root: Node): boolean {
  for (let at: Node | null = node; at && at !== root; at = at.parentNode) {
    for (let next = at.nextSibling; next; next = next.nextSibling) {
      if ((next.textContent ?? "") !== "" || (next.nodeType === 1 && (next as Element).tagName === "BR")) {
        return true;
      }
    }
  }
  return false;
}

/** The host's content as plain text ("\n" for line breaks). */
export function readDraftText(host: Node): string {
  const out: string[] = [];
  collect(host, out, host);
  return out.join("").replace(/\r\n?/g, "\n");
}

/** Re-renders `host` as `text` with one coloured span per reference. */
export function renderDraft(host: HTMLElement, text: string, spans: readonly DraftSpan[]): void {
  const doc = host.ownerDocument;
  const fragment = doc.createDocumentFragment();
  let at = 0;
  for (const span of spans.toSorted((a, b) => a.start - b.start)) {
    if (span.start < at || span.end <= span.start || span.end > text.length) {
      continue;
    }
    if (span.start > at) {
      fragment.append(doc.createTextNode(text.slice(at, span.start)));
    }
    const element = doc.createElement("span");
    element.className = "sheet-ref";
    element.dataset.refIndex = String(span.colorIndex);
    element.style.color = refColor(span.colorIndex);
    element.textContent = text.slice(span.start, span.end);
    fragment.append(element);
    at = span.end;
  }
  if (at < text.length) {
    fragment.append(doc.createTextNode(text.slice(at)));
  }
  if (text.endsWith("\n")) {
    const marker = doc.createElement("br");
    marker.setAttribute(END_MARKER, "");
    fragment.append(marker);
  }
  host.replaceChildren(fragment);
}

/** The read-back offset of the DOM point (node, offset) inside `host`. */
function offsetOfPoint(host: HTMLElement, node: Node, offset: number): number {
  const range = host.ownerDocument.createRange();
  range.setStart(host, 0);
  try {
    range.setEnd(node, offset);
  } catch {
    return readDraftText(host).length;
  }
  const out: string[] = [];
  collect(range.cloneContents(), out, host);
  return out.join("").replace(/\r\n?/g, "\n").length;
}

function inside(host: HTMLElement, node: Node | null): node is Node {
  return node !== null && (node === host || host.contains(node));
}

/** The selection inside `host` as read-back offsets (start ≤ end), or null
 * when the selection is elsewhere. */
export function selectionOffsetsIn(host: HTMLElement): { start: number; end: number } | null {
  const selection = host.ownerDocument.getSelection();
  if (!selection || selection.rangeCount === 0 || !inside(host, selection.anchorNode) || !inside(host, selection.focusNode)) {
    return null;
  }
  const anchor = offsetOfPoint(host, selection.anchorNode, selection.anchorOffset);
  const focus = offsetOfPoint(host, selection.focusNode, selection.focusOffset);
  return { start: Math.min(anchor, focus), end: Math.max(anchor, focus) };
}

/** The caret (the selection's focus) inside `host`, or null. */
export function caretOffsetIn(host: HTMLElement): number | null {
  const selection = host.ownerDocument.getSelection();
  if (!selection || selection.rangeCount === 0 || !inside(host, selection.focusNode)) {
    return null;
  }
  return offsetOfPoint(host, selection.focusNode, selection.focusOffset);
}

/** The DOM point of read-back offset `offset` in a host renderDraft wrote. */
function pointAt(host: HTMLElement, offset: number): { node: Node; offset: number } {
  const walker = host.ownerDocument.createTreeWalker(host, NodeFilter.SHOW_TEXT);
  let remaining = Math.max(0, offset);
  let last: Text | null = null;
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    if (remaining <= node.data.length) {
      return { node, offset: remaining };
    }
    remaining -= node.data.length;
    last = node;
  }
  if (last) {
    return { node: last, offset: last.data.length };
  }
  return { node: host, offset: 0 };
}

/** Places the caret (or the selection start..end) at read-back offsets. */
export function placeCaretIn(host: HTMLElement, start: number, end = start): void {
  const doc = host.ownerDocument;
  const selection = doc.getSelection();
  if (!selection) {
    return;
  }
  const a = pointAt(host, start);
  const b = end === start ? a : pointAt(host, end);
  const range = doc.createRange();
  range.setStart(a.node, a.offset);
  range.setEnd(b.node, b.offset);
  selection.removeAllRanges();
  selection.addRange(range);
}

/** The read-back offset under the viewport point (x, y) when it lands in
 * `host`; null when it does not, or the browser cannot tell (jsdom). */
export function caretOffsetAtPoint(host: HTMLElement, x: number, y: number): number | null {
  const doc = host.ownerDocument as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
  };
  let node: Node | null = null;
  let offset = 0;
  if (typeof doc.caretRangeFromPoint === "function") {
    const range = doc.caretRangeFromPoint(x, y);
    node = range?.startContainer ?? null;
    offset = range?.startOffset ?? 0;
  } else if (typeof doc.caretPositionFromPoint === "function") {
    const position = doc.caretPositionFromPoint(x, y);
    node = position?.offsetNode ?? null;
    offset = position?.offset ?? 0;
  }
  return inside(host, node) ? offsetOfPoint(host, node, offset) : null;
}
