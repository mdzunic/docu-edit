// The DOM side of slide text editing (SPEC-081): render a text frame's model
// paragraphs into a contenteditable host, and read the host's DOM back into
// paragraphs after the browser applies an edit ("normalise-after-input"). The
// renderer only ever emits block `<div>`s of styled `<span>`s, so the
// normaliser reads back exactly what rendering wrote. Kept beside the stage
// renderer (a rendering concern), NOT under src/lib/presentation — it touches
// the DOM and only needs the model's plain types.
import type { SlideParagraph, SlideRun } from "../../lib/presentation/model";

const PT_TO_PX = 96 / 72;

type RunMarks = Omit<SlideRun, "text">;

/** Renders paragraphs into an editable host: one `<div class="ce-p">` per
 * paragraph, one `<span>` per run with inline styles for its DEFINED marks
 * only (so the normaliser reads `undefined` back as `undefined`). An empty
 * paragraph gets a lone `<br>` so it keeps a caret line. */
export function renderEditableParagraphs(host: HTMLElement, paragraphs: SlideParagraph[]): void {
  const doc = host.ownerDocument;
  host.textContent = "";
  const paras = paragraphs.length > 0 ? paragraphs : [{ runs: [] as SlideRun[] }];
  for (const para of paras) {
    const block = doc.createElement("div");
    block.className = "ce-p";
    block.style.textAlign = para.align ?? "left";
    if (para.runs.length === 0) {
      block.appendChild(doc.createElement("br"));
    } else {
      for (const run of para.runs) {
        block.appendChild(runSpan(doc, run));
      }
    }
    host.appendChild(block);
  }
}

function runSpan(doc: Document, run: SlideRun): HTMLSpanElement {
  const span = doc.createElement("span");
  if (run.bold) span.style.fontWeight = "700";
  if (run.italic) span.style.fontStyle = "italic";
  if (run.underline) span.style.textDecoration = "underline";
  if (run.fontFamily) span.style.fontFamily = `'${run.fontFamily}', sans-serif`;
  if (run.sizePt !== undefined) span.style.fontSize = `${run.sizePt * PT_TO_PX}px`;
  if (run.color) span.style.color = run.color;
  span.textContent = run.text;
  return span;
}

/** Reads a contenteditable host back into paragraphs: block children and
 * `<br>`s split paragraphs, styled spans carry run marks, bare text is a plain
 * run. Paragraph-level properties (align/bullet/level) are NOT recovered here —
 * they are set by toolbar commands on the model and merged back by the caller. */
export function normalizeTextFrameDom(host: HTMLElement): SlideParagraph[] {
  const lines: SlideRun[][] = [[]];
  const line = (): SlideRun[] => lines[lines.length - 1];
  const push = (text: string, marks: RunMarks): void => {
    if (text.length === 0) return;
    line().push(buildRun(text, marks));
  };
  const newLine = (): void => {
    lines.push([]);
  };

  const walk = (node: Node, marks: RunMarks): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === Node.TEXT_NODE) {
        push(child.textContent ?? "", marks);
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const el = child as HTMLElement;
        if (el.tagName === "BR") {
          newLine();
        } else if (isBlock(el)) {
          if (line().length > 0) newLine();
          walk(el, marksFrom(el, marks));
          newLine();
        } else {
          walk(el, marksFrom(el, marks));
        }
      }
    }
  };
  walk(host, {});

  // Drop a single trailing empty line left by a block boundary, but keep an
  // all-empty host as one empty paragraph.
  while (lines.length > 1 && lines[lines.length - 1].length === 0) {
    lines.pop();
  }
  return lines.map((runs) => ({ runs: coalesce(runs) }));
}

function isBlock(el: HTMLElement): boolean {
  return el.tagName === "DIV" || el.tagName === "P";
}

/** Folds an element's inline styles into the inherited marks. */
function marksFrom(el: HTMLElement, base: RunMarks): RunMarks {
  const marks: RunMarks = { ...base };
  const s = el.style;
  const weight = s.fontWeight;
  if (weight === "700" || weight === "bold" || Number(weight) >= 600) marks.bold = true;
  if (s.fontStyle === "italic" || s.fontStyle === "oblique") marks.italic = true;
  const decoration = s.textDecorationLine || s.textDecoration;
  if (decoration && decoration.includes("underline")) marks.underline = true;
  if (s.fontFamily) marks.fontFamily = stripFontFamily(s.fontFamily);
  if (s.fontSize) {
    const px = parseFloat(s.fontSize);
    if (Number.isFinite(px)) marks.sizePt = Math.round((px / PT_TO_PX) * 10) / 10;
  }
  const hex = cssColorToHex(s.color);
  if (hex) marks.color = hex;
  return marks;
}

function stripFontFamily(value: string): string {
  return value.split(",")[0].trim().replace(/^['"]|['"]$/g, "");
}

/** Parses a CSS colour (`#rrggbb`, `#rgb`, or `rgb(...)`) to lowercase
 * `#rrggbb`, or undefined when it can't. */
export function cssColorToHex(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const v = value.trim();
  if (/^#[0-9a-fA-F]{6}$/.test(v)) return v.toLowerCase();
  if (/^#[0-9a-fA-F]{3}$/.test(v)) {
    return `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`.toLowerCase();
  }
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);
  if (rgb) {
    const toHex = (n: string): string => Number(n).toString(16).padStart(2, "0");
    return `#${toHex(rgb[1])}${toHex(rgb[2])}${toHex(rgb[3])}`.toLowerCase();
  }
  return undefined;
}

function buildRun(text: string, marks: RunMarks): SlideRun {
  const run: SlideRun = { text };
  if (marks.bold) run.bold = true;
  if (marks.italic) run.italic = true;
  if (marks.underline) run.underline = true;
  if (marks.fontFamily) run.fontFamily = marks.fontFamily;
  if (marks.sizePt !== undefined) run.sizePt = marks.sizePt;
  if (marks.color) run.color = marks.color;
  return run;
}

function marksEqual(a: SlideRun, b: SlideRun): boolean {
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    a.fontFamily === b.fontFamily &&
    a.sizePt === b.sizePt &&
    a.color === b.color
  );
}

/** Merges adjacent runs with identical marks. */
function coalesce(runs: SlideRun[]): SlideRun[] {
  const out: SlideRun[] = [];
  for (const run of runs) {
    const prev = out[out.length - 1];
    if (prev && marksEqual(prev, run)) prev.text += run.text;
    else out.push({ ...run });
  }
  return out;
}

// --- caret addressing ------------------------------------------------------

/** A caret position as a paragraph index and a character offset within that
 * paragraph — the address the editor restores after a re-render. */
export interface CaretAddress {
  para: number;
  offset: number;
}

/** The current selection's focus as a paragraph/offset address inside `host`,
 * or null when the selection is outside it. */
export function caretAddressOf(host: HTMLElement): CaretAddress | null {
  const sel = host.ownerDocument.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const { focusNode, focusOffset } = sel;
  if (!focusNode || !host.contains(focusNode)) return null;
  const blocks = Array.from(host.children);
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block === focusNode || block.contains(focusNode)) {
      return { para: i, offset: offsetWithin(block, focusNode, focusOffset) };
    }
  }
  return { para: Math.max(0, blocks.length - 1), offset: 0 };
}

function offsetWithin(block: Element, focusNode: Node, focusOffset: number): number {
  if (focusNode === block) {
    // Offset is a child index; sum the text length of preceding children.
    let total = 0;
    for (let i = 0; i < focusOffset && i < block.childNodes.length; i++) {
      total += (block.childNodes[i].textContent ?? "").length;
    }
    return total;
  }
  let total = 0;
  const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    if (node === focusNode) return total + focusOffset;
    total += (node.textContent ?? "").length;
    node = walker.nextNode();
  }
  return total;
}

/** Places the caret at an address inside `host` (clamped). */
export function placeCaret(host: HTMLElement, address: CaretAddress): void {
  const doc = host.ownerDocument;
  const block = host.children[Math.min(address.para, host.children.length - 1)];
  if (!block) return;
  const range = doc.createRange();
  const walker = doc.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let remaining = address.offset;
  let node = walker.nextNode();
  let last: Text | null = null;
  while (node) {
    const len = (node.textContent ?? "").length;
    if (remaining <= len) {
      range.setStart(node, remaining);
      range.collapse(true);
      applyRange(doc, range);
      return;
    }
    remaining -= len;
    last = node as Text;
    node = walker.nextNode();
  }
  if (last) {
    range.setStart(last, (last.textContent ?? "").length);
  } else {
    range.setStart(block, 0);
  }
  range.collapse(true);
  applyRange(doc, range);
}

function applyRange(doc: Document, range: Range): void {
  const sel = doc.getSelection();
  if (!sel) return;
  sel.removeAllRanges();
  sel.addRange(range);
}

/** Selects all text in `host` (Cmd/Ctrl+A within the frame, AC-10). */
export function selectAllIn(host: HTMLElement): void {
  const doc = host.ownerDocument;
  const range = doc.createRange();
  range.selectNodeContents(host);
  applyRange(doc, range);
}

/** Locates the DOM (node, offset) for a paragraph/character address. */
function locate(host: HTMLElement, address: CaretAddress): { node: Node; offset: number } {
  const block = host.children[Math.min(address.para, host.children.length - 1)];
  if (!block) return { node: host, offset: 0 };
  const walker = host.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let remaining = address.offset;
  let node = walker.nextNode();
  let last: Node | null = null;
  while (node) {
    const len = (node.textContent ?? "").length;
    if (remaining <= len) return { node, offset: remaining };
    remaining -= len;
    last = node;
    node = walker.nextNode();
  }
  return last ? { node: last, offset: (last.textContent ?? "").length } : { node: block, offset: 0 };
}

/** Restores a selection spanning `start`→`end` after a re-render. */
export function setSelectionRange(host: HTMLElement, start: CaretAddress, end: CaretAddress): void {
  const doc = host.ownerDocument;
  const a = locate(host, start);
  const b = locate(host, end);
  const range = doc.createRange();
  range.setStart(a.node, a.offset);
  range.setEnd(b.node, b.offset);
  applyRange(doc, range);
}

/** The full selection over `host` as a paragraph range, or a collapsed caret
 * address when nothing is selected. Returns start/end paragraph+offset. */
export function selectionRangeOf(host: HTMLElement): { start: CaretAddress; end: CaretAddress } | null {
  const sel = host.ownerDocument.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const anchor = pointOf(host, sel.anchorNode, sel.anchorOffset);
  const focus = pointOf(host, sel.focusNode, sel.focusOffset);
  if (!anchor || !focus) return null;
  const ordered =
    anchor.para < focus.para || (anchor.para === focus.para && anchor.offset <= focus.offset)
      ? { start: anchor, end: focus }
      : { start: focus, end: anchor };
  return ordered;
}

function pointOf(host: HTMLElement, node: Node | null, offset: number): CaretAddress | null {
  if (!node || !host.contains(node)) return null;
  const blocks = Array.from(host.children);
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i] === node || blocks[i].contains(node)) {
      return { para: i, offset: offsetWithin(blocks[i], node, offset) };
    }
  }
  return { para: Math.max(0, blocks.length - 1), offset: 0 };
}
