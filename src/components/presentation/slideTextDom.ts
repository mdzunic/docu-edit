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
 * paragraph gets a lone `<br>` so it keeps a caret line. Paragraph-level props
 * (align/bullet/level) ride along as `data-*` attributes plus a `data-p`
 * marker, so `normalizeTextFrameDom` recovers them per-block — that keeps
 * bullet/indent glued to the paragraph's DOM element across a keystroke that
 * changes the paragraph count (Enter, Backspace-merge), rather than matching by
 * a shifting array index. */
export function renderEditableParagraphs(host: HTMLElement, paragraphs: SlideParagraph[]): void {
  const doc = host.ownerDocument;
  host.textContent = "";
  const paras = paragraphs.length > 0 ? paragraphs : [{ runs: [] as SlideRun[] }];
  for (const para of paras) {
    const block = doc.createElement("div");
    block.className = "ce-p";
    block.dataset.p = "1";
    block.style.textAlign = para.align ?? "left";
    if (para.align) block.dataset.align = para.align;
    if (para.bullet) block.dataset.bullet = para.bullet;
    if (para.level) block.dataset.level = String(para.level);
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

/** A paragraph's block-level properties, the subset `renderEditableParagraphs`
 * stamps onto each block and `normalizeTextFrameDom` reads back. */
type ParaProps = Pick<SlideParagraph, "align" | "bullet" | "level">;

interface PendingLine {
  runs: SlideRun[];
  props: ParaProps;
  /** Props came from a block WE rendered (`data-p`), so they are authoritative
   * — even an empty `props` is an explicit "no bullet/align" and must not be
   * overwritten by inheritance. A browser-created block (`data-p` absent) is
   * `false` and inherits from the preceding paragraph. */
  explicit: boolean;
}

/** Reads a contenteditable host back into paragraphs: block children and
 * `<br>`s split paragraphs, styled spans carry run marks, bare text is a plain
 * run. Paragraph-level properties (align/bullet/level) are recovered from each
 * block's `data-*` attributes; a block the browser created (Enter/paste split,
 * so it carries no `data-p` marker) inherits them from the preceding paragraph,
 * matching how an editor carries bullet/indent onto a new line. */
export function normalizeTextFrameDom(host: HTMLElement): SlideParagraph[] {
  const lines: PendingLine[] = [{ runs: [], props: {}, explicit: false }];
  const line = (): PendingLine => lines[lines.length - 1];
  const push = (text: string, marks: RunMarks): void => {
    if (text.length === 0) return;
    line().runs.push(buildRun(text, marks));
  };
  const newLine = (props: ParaProps, explicit: boolean): void => {
    lines.push({ runs: [], props, explicit });
  };

  const walk = (node: Node, marks: RunMarks, blockProps: ParaProps, blockExplicit: boolean): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === Node.TEXT_NODE) {
        push(child.textContent ?? "", marks);
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const el = child as HTMLElement;
        if (el.tagName === "BR") {
          // A soft break stays inside the enclosing block, so the new line keeps
          // that block's props.
          newLine(blockProps, blockExplicit);
        } else if (isBlock(el)) {
          const { props, explicit } = paraPropsFrom(el);
          if (line().runs.length > 0) {
            newLine(props, explicit);
          } else {
            const cur = line();
            cur.props = props;
            cur.explicit = explicit;
          }
          walk(el, marksFrom(el, marks), props, explicit);
          newLine(blockProps, blockExplicit);
        } else {
          walk(el, marksFrom(el, marks), blockProps, blockExplicit);
        }
      }
    }
  };
  walk(host, {}, {}, false);

  // Drop a single trailing empty line left by a block boundary, but keep an
  // all-empty host as one empty paragraph.
  while (lines.length > 1 && lines[lines.length - 1].runs.length === 0) {
    lines.pop();
  }

  // Inherit block props onto browser-created paragraphs from the last rendered
  // one; an explicitly-rendered paragraph keeps its own props (including an
  // explicit "no bullet" after a toggle-off), so nothing clobbers them.
  let carried: ParaProps = {};
  for (const l of lines) {
    if (l.explicit) carried = l.props;
    else l.props = carried;
  }

  return lines.map((l) => {
    const para: SlideParagraph = { runs: coalesce(l.runs) };
    if (l.props.align) para.align = l.props.align;
    if (l.props.bullet) para.bullet = l.props.bullet;
    if (l.props.level) para.level = l.props.level;
    return para;
  });
}

/** Recovers a block's paragraph-level props from its `data-*` attributes.
 * `explicit` is true when the block carries our `data-p` marker (we rendered
 * it), false for a block the browser minted during an edit. */
function paraPropsFrom(el: HTMLElement): { props: ParaProps; explicit: boolean } {
  const explicit = el.dataset.p === "1";
  const props: ParaProps = {};
  const align = el.dataset.align;
  if (align === "left" || align === "center" || align === "right" || align === "justify") {
    props.align = align;
  }
  const bullet = el.dataset.bullet;
  if (bullet === "disc" || bullet === "number") props.bullet = bullet;
  const level = Number(el.dataset.level);
  if (Number.isFinite(level) && level > 0) props.level = level;
  return { props, explicit };
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
  if (!focusNode || (focusNode !== host && !host.contains(focusNode))) return null;
  return addressAt(host, focusNode, focusOffset);
}

/** Maps a DOM (node, offset) inside `host` to a paragraph/character address.
 * Handles a point on the host itself (a whole-content selection's endpoints,
 * where the node is the host and the offset is a block index). */
function addressAt(host: HTMLElement, node: Node, offset: number): CaretAddress {
  const blocks = Array.from(host.children);
  if (node === host) {
    if (offset >= blocks.length) {
      const last = Math.max(0, blocks.length - 1);
      return { para: last, offset: (blocks[last]?.textContent ?? "").length };
    }
    return { para: offset, offset: 0 };
  }
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i] === node || blocks[i].contains(node)) {
      return { para: i, offset: offsetWithin(blocks[i], node, offset) };
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
  if (!node || (node !== host && !host.contains(node))) return null;
  return addressAt(host, node, offset);
}
