// The DOM side of slide text editing (SPEC-081): render a text frame's model
// paragraphs into a contenteditable host, and read the host's DOM back into
// paragraphs after the browser applies an edit ("normalise-after-input"). The
// renderer only ever emits block `<div>`s of styled `<span>`s, so the
// normaliser reads back exactly what rendering wrote. Kept beside the stage
// renderer (a rendering concern), NOT under src/lib/presentation — it touches
// the DOM and only needs the model's plain types.
//
// SPEC-097 widens the round trip: every new paragraph field rides on its
// block as a `data-*` attribute (list style/char/picture, start-at, indent,
// hanging, line and paragraph spacing, and the pptx/odp preservation slots),
// and a run span carries its letter spacing (inline `letter-spacing`) and its
// link (`data-link-href`/`data-link-slide`/`data-link-tooltip`, drawn
// underlined by the `ce-link` class — never an inline underline, which would
// read back as a run mark). The host renders the paragraph's indent/spacing
// so editing is WYSIWYG — at full scale: run sizes stay plain px (an autofit
// scale is never baked into a run the browser copies), and list markers are
// the stage's alone (a marker in the host would shift caret offsets).
import type { SlideParagraph, SlideRun } from "../../lib/presentation/model";
import type { ListMarkerStyle } from "../../lib/core/lists";
import { DEFAULT_THEME, resolveColor, resolveFont, type DocumentTheme } from "../../lib/core/theme";
import { isListParagraph } from "../../lib/presentation/slideLists";
import { paragraphStyle } from "./slideRender";

const PT_TO_PX = 96 / 72;

type RunMarks = Omit<SlideRun, "text">;

/** The class a linked run's span carries in the editing host (the stage's
 * scoped style underlines it). */
export const EDIT_LINK_CLASS = "ce-link";

const MARKER_STYLES: ReadonlySet<string> = new Set<ListMarkerStyle>([
  "disc",
  "circle",
  "square",
  "decimal",
  "lowerAlpha",
  "upperAlpha",
  "lowerRoman",
  "upperRoman",
]);

/** The paragraph fields stamped onto a block as `data-*` (dataset keys) —
 * the scalar half of `ParaProps` beyond align/bullet/level. */
const NUMBER_FIELDS = [
  ["startAt", "startAt"],
  ["indentCm", "indent"],
  ["hangingCm", "hanging"],
  ["lineSpacing", "lineSpacing"],
  ["lineSpacingPt", "lineSpacingPt"],
  ["spaceBeforePt", "spaceBefore"],
  ["spaceAfterPt", "spaceAfter"],
  ["pptxSourceParagraph", "pptxSourceParagraph"],
] as const;
const STRING_FIELDS = [
  ["bulletChar", "bulletChar"],
  ["bulletImageId", "bulletImage"],
  ["pptxPPr", "pptxPpr"],
  ["odpParaProps", "odpParaProps"],
  ["sourceWireId", "sourceWireId"],
] as const;

/** Renders paragraphs into an editable host: one `<div class="ce-p">` per
 * paragraph, one `<span>` per run with inline styles for its DEFINED marks
 * only (so the normaliser reads `undefined` back as `undefined`). An empty
 * paragraph gets a lone `<br>` so it keeps a caret line. Paragraph-level props
 * ride along as `data-*` attributes plus a `data-p` marker (and a per-render
 * `data-pi` index), so `normalizeTextFrameDom` recovers them per-block — that
 * keeps bullet/indent glued to the paragraph's DOM element across a keystroke
 * that changes the paragraph count (Enter, Backspace-merge), rather than
 * matching by a shifting array index. `defaultAlign` is how a paragraph with
 * no alignment of its own lays out — a shape's text centres (SPEC-079) — and,
 * like every block style, is never read back. */
export function renderEditableParagraphs(
  host: HTMLElement,
  paragraphs: SlideParagraph[],
  theme: DocumentTheme = DEFAULT_THEME,
  defaultAlign: NonNullable<SlideParagraph["align"]> = "left",
): void {
  const doc = host.ownerDocument;
  host.textContent = "";
  const paras = paragraphs.length > 0 ? paragraphs : [{ runs: [] as SlideRun[] }];
  paras.forEach((para, index) => {
    const block = doc.createElement("div");
    block.className = "ce-p";
    block.dataset.p = "1";
    block.dataset.pi = String(index);
    applyBlockStyle(block, para, defaultAlign);
    if (para.align) block.dataset.align = para.align;
    if (para.bullet) block.dataset.bullet = para.bullet;
    if (para.level) block.dataset.level = String(para.level);
    if (para.bulletStyle) block.dataset.bulletStyle = para.bulletStyle;
    if (para.lineSpacingRule) block.dataset.lineSpacingRule = para.lineSpacingRule;
    for (const [field, key] of NUMBER_FIELDS) {
      const value = para[field];
      if (value !== undefined) block.dataset[key] = String(value);
    }
    for (const [field, key] of STRING_FIELDS) {
      const value = para[field];
      if (value !== undefined) block.dataset[key] = value;
    }
    if (para.runs.length === 0) {
      block.appendChild(doc.createElement("br"));
    } else {
      for (const run of para.runs) {
        block.appendChild(runSpan(doc, run, theme));
      }
    }
    host.appendChild(block);
  });
}

/** A block's WYSIWYG paragraph style — the stage's own (`paragraphStyle`:
 * alignment, indent, line pitch, space before/after) — except that a list
 * item drops its hanging first-line indent: the host shows no marker, so its
 * first line starts at the text edge, exactly where the stage puts the text
 * after the marker box. Style only; nothing here is read back. */
function applyBlockStyle(block: HTMLElement, para: SlideParagraph, defaultAlign: NonNullable<SlideParagraph["align"]>): void {
  const style = paragraphStyle(para, defaultAlign);
  if (isListParagraph(para)) delete style.textIndent;
  for (const [key, value] of Object.entries(style)) {
    block.style.setProperty(key.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`), value);
  }
}

function runSpan(doc: Document, run: SlideRun, theme: DocumentTheme): HTMLSpanElement {
  const span = doc.createElement("span");
  if (run.bold) span.style.fontWeight = "700";
  if (run.italic) span.style.fontStyle = "italic";
  if (run.underline) span.style.textDecoration = "underline";
  // Refs resolve into the editing DOM (the DOM only carries literal CSS); the
  // read-back normaliser then sees literals, so an edited run's marks are the
  // effective values (SPEC-094 — only the edited frame is affected).
  if (run.fontFamily) span.style.fontFamily = `'${resolveFont(run.fontFamily, theme)}', sans-serif`;
  if (run.sizePt !== undefined) span.style.fontSize = `${run.sizePt * PT_TO_PX}px`;
  if (run.color) span.style.color = resolveColor(run.color, theme);
  if (run.letterSpacingPt) span.style.letterSpacing = `${run.letterSpacingPt * PT_TO_PX}px`;
  if (run.linkHref) span.dataset.linkHref = run.linkHref;
  else if (run.linkSlide) span.dataset.linkSlide = String(run.linkSlide);
  if ((run.linkHref || run.linkSlide) && run.linkTooltip) span.dataset.linkTooltip = run.linkTooltip;
  if (run.linkHref || run.linkSlide) span.classList.add(EDIT_LINK_CLASS);
  // SPEC-098: the run's language and noProof ride on the span so a session
  // never strips them (the host is spellcheck="false"; `lang` is inert there).
  if (run.lang) span.lang = run.lang;
  if (run.noProof) span.dataset.noProof = "1";
  span.textContent = run.text;
  return span;
}

/** A paragraph's block-level properties, the subset `renderEditableParagraphs`
 * stamps onto each block and `normalizeTextFrameDom` reads back. */
type ParaProps = Omit<SlideParagraph, "runs">;

interface PendingLine {
  runs: SlideRun[];
  props: ParaProps;
  /** Props came from a block WE rendered (`data-p`), so they are authoritative
   * — even an empty `props` is an explicit "no bullet/align" and must not be
   * overwritten by inheritance. A browser-created block (`data-p` absent) is
   * `false` and inherits from the preceding paragraph. */
  explicit: boolean;
}

/** The props a NEW line takes from the paragraph it split off: everything
 * except `startAt`. A start-at seeds a numbered run's FIRST item (slideLists
 * derives the rest by adjacency), so a new line continues the run rather than
 * re-seeding it — and a writer emitting a per-item restart (ODF
 * `text:start-value`) never sees a duplicated seed. The preservation slots
 * (`pptxPPr`/`odpParaProps`) DO carry over: they hold the paragraph's
 * unmodelled formatting — bullet font/colour/size, tabs, default run props —
 * which the new line visibly shares, exactly as PowerPoint and Impress copy a
 * paragraph's properties onto the line Enter creates; dropping them would
 * re-render a Wingdings `buChar` bullet in the wrong font on that line. A
 * paragraph's `sourceWireId` (its source id, which preserved animation
 * targets) is an identity, not formatting: like `startAt` it stays on the
 * first line, or two paragraphs would claim one target. Its pptx
 * `pptxSourceParagraph` DOES carry over: it says which source paragraph the
 * text came from, which both halves of a split did — exactly as the reader
 * gives every line an `a:br` split off its paragraph's index — so a
 * preserved by-paragraph build still reveals all of that text. */
function carriedProps(props: ParaProps): ParaProps {
  const { startAt: _startAt, sourceWireId: _sourceWireId, ...rest } = props;
  return rest;
}

/** Reads a contenteditable host back into paragraphs: block children and
 * `<br>`s split paragraphs, styled spans carry run marks, bare text is a plain
 * run. Paragraph-level properties are recovered from each block's `data-*`
 * attributes; a block the browser created (Enter/paste split, so it carries
 * no `data-p` marker) inherits them from the preceding paragraph (bar
 * `startAt`, see `carriedProps`), matching how an editor carries bullet/indent
 * onto a new line. A block the browser CLONED from one of ours (Chrome copies
 * a block's attributes when Enter splits it) repeats an earlier block's
 * `data-pi`; that later copy is the split-off line and drops `startAt` too. */
export function normalizeTextFrameDom(host: HTMLElement): SlideParagraph[] {
  const lines: PendingLine[] = [{ runs: [], props: {}, explicit: false }];
  const seenIndexes = new Set<string>();
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
          // that block's props (bar a start-at seed, which stays on the first).
          newLine(carriedProps(blockProps), blockExplicit);
        } else if (isBlock(el)) {
          const read = paraPropsFrom(el);
          const index = el.dataset.pi;
          const isSplitCopy = read.explicit && index !== undefined && seenIndexes.has(index);
          if (index !== undefined) seenIndexes.add(index);
          const props = isSplitCopy ? carriedProps(read.props) : read.props;
          const explicit = read.explicit;
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
    if (l.explicit) carried = carriedProps(l.props);
    else l.props = carried;
  }

  return lines.map((l) => ({ runs: coalesce(l.runs), ...definedProps(l.props) }));
}

/** The defined props in the model's canonical key order (absent stays
 * absent, so a round trip is `toEqual` to its input). */
function definedProps(props: ParaProps): ParaProps {
  const out: ParaProps = {};
  if (props.align) out.align = props.align;
  if (props.bullet) out.bullet = props.bullet;
  if (props.level) out.level = props.level;
  if (props.bulletStyle) out.bulletStyle = props.bulletStyle;
  if (props.bulletChar !== undefined) out.bulletChar = props.bulletChar;
  if (props.bulletImageId !== undefined) out.bulletImageId = props.bulletImageId;
  if (props.startAt !== undefined) out.startAt = props.startAt;
  if (props.indentCm !== undefined) out.indentCm = props.indentCm;
  if (props.hangingCm !== undefined) out.hangingCm = props.hangingCm;
  if (props.lineSpacing !== undefined) out.lineSpacing = props.lineSpacing;
  if (props.lineSpacingPt !== undefined) out.lineSpacingPt = props.lineSpacingPt;
  if (props.lineSpacingRule) out.lineSpacingRule = props.lineSpacingRule;
  if (props.spaceBeforePt !== undefined) out.spaceBeforePt = props.spaceBeforePt;
  if (props.spaceAfterPt !== undefined) out.spaceAfterPt = props.spaceAfterPt;
  if (props.pptxPPr !== undefined) out.pptxPPr = props.pptxPPr;
  if (props.odpParaProps !== undefined) out.odpParaProps = props.odpParaProps;
  if (props.sourceWireId !== undefined) out.sourceWireId = props.sourceWireId;
  if (props.pptxSourceParagraph !== undefined) out.pptxSourceParagraph = props.pptxSourceParagraph;
  return out;
}

/** Recovers a block's paragraph-level props from its `data-*` attributes.
 * `explicit` is true when the block carries our `data-p` marker (we rendered
 * it), false for a block the browser minted during an edit. */
function paraPropsFrom(el: HTMLElement): { props: ParaProps; explicit: boolean } {
  const explicit = el.dataset.p === "1";
  const props: ParaProps = {};
  const data = el.dataset;
  const align = data.align;
  if (align === "left" || align === "center" || align === "right" || align === "justify") {
    props.align = align;
  }
  const bullet = data.bullet;
  if (bullet === "disc" || bullet === "number") props.bullet = bullet;
  const level = Number(data.level);
  if (Number.isFinite(level) && level > 0) props.level = level;
  const style = data.bulletStyle;
  if (style !== undefined && MARKER_STYLES.has(style)) props.bulletStyle = style as ListMarkerStyle;
  if (data.lineSpacingRule === "exact") props.lineSpacingRule = "exact";
  for (const [field, key] of NUMBER_FIELDS) {
    const raw = data[key];
    if (raw === undefined || raw === "") continue;
    const value = Number(raw);
    if (Number.isFinite(value)) props[field] = value;
  }
  for (const [field, key] of STRING_FIELDS) {
    const raw = data[key];
    if (raw !== undefined) props[field] = raw;
  }
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
  if (s.letterSpacing) {
    const px = parseFloat(s.letterSpacing);
    if (Number.isFinite(px)) marks.letterSpacingPt = Math.round((px / PT_TO_PX) * 10) / 10;
  }
  // A link mark: at most one of href/slide (an inner span's link replaces
  // an outer one's wholesale, tooltip included).
  const href = el.dataset.linkHref;
  const slide = Number(el.dataset.linkSlide);
  if (href || (Number.isInteger(slide) && slide >= 1)) {
    delete marks.linkHref;
    delete marks.linkSlide;
    delete marks.linkTooltip;
    if (href) marks.linkHref = href;
    else marks.linkSlide = slide;
    if (el.dataset.linkTooltip) marks.linkTooltip = el.dataset.linkTooltip;
  }
  const lang = el.getAttribute("lang");
  if (lang) marks.lang = lang;
  if (el.dataset.noProof === "1") marks.noProof = true;
  return marks;
}

function stripFontFamily(value: string): string {
  return value.split(",")[0].trim().replace(/^['"]|['"]$/g, "");
}

const toHex = (n: string): string => Number(n).toString(16).padStart(2, "0");

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
  if (marks.letterSpacingPt) run.letterSpacingPt = marks.letterSpacingPt;
  if (marks.linkHref) run.linkHref = marks.linkHref;
  else if (marks.linkSlide) run.linkSlide = marks.linkSlide;
  if ((marks.linkHref || marks.linkSlide) && marks.linkTooltip) run.linkTooltip = marks.linkTooltip;
  if (marks.lang) run.lang = marks.lang;
  if (marks.noProof) run.noProof = true;
  return run;
}

/** Whether two runs carry the same marks — so coalescing merges only runs
 * that would render identically (SPEC-097 AC-10: letter spacing and the link
 * scalars compare `===` per key, like text.ts#marksEqual). */
export function marksEqual(a: SlideRun, b: SlideRun): boolean {
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    a.fontFamily === b.fontFamily &&
    a.sizePt === b.sizePt &&
    a.color === b.color &&
    (a.letterSpacingPt ?? 0) === (b.letterSpacingPt ?? 0) &&
    a.linkHref === b.linkHref &&
    a.linkSlide === b.linkSlide &&
    a.linkTooltip === b.linkTooltip &&
    a.lang === b.lang &&
    !!a.noProof === !!b.noProof
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

/** Places the caret at the viewport point (clientX, clientY) when it lands
 * inside `host`, returning true on success (AC-1: a double-click drops the
 * caret at the click point, not at the end). Uses `caretRangeFromPoint`
 * (Chromium/WebKit) or `caretPositionFromPoint` (standard); returns false when
 * neither exists (jsdom) or the point resolves outside the host, so the caller
 * can fall back to placing the caret at the end.
 *
 * The host sits inside its frame's text body, whose insets are padding and
 * whose anchor can hug the host to its content (SPEC-097), so a click in the
 * frame's margin lands beside the host rather than on it. Such a point is
 * retried at the nearest point inside the host's box: a double-click just left
 * of the first character still drops the caret before it. */
export function placeCaretAtPoint(host: HTMLElement, clientX: number, clientY: number): boolean {
  if (placeCaretInsideAt(host, clientX, clientY)) return true;
  const box = host.getBoundingClientRect();
  if (box.width <= 0 || box.height <= 0) return false;
  const x = Math.min(Math.max(clientX, box.left + 1), box.right - 1);
  const y = Math.min(Math.max(clientY, box.top + 1), box.bottom - 1);
  if (x === clientX && y === clientY) return false;
  return placeCaretInsideAt(host, x, y);
}

function placeCaretInsideAt(host: HTMLElement, x: number, y: number): boolean {
  const doc = host.ownerDocument;
  const range = caretRangeAtPoint(doc, x, y);
  if (!range) return false;
  const node = range.startContainer;
  if (node !== host && !host.contains(node)) return false;
  range.collapse(true);
  applyRange(doc, range);
  return true;
}

function caretRangeAtPoint(doc: Document, x: number, y: number): Range | null {
  if (typeof doc.caretRangeFromPoint === "function") {
    return doc.caretRangeFromPoint(x, y);
  }
  // caretPositionFromPoint is the standardised form (Firefox); not in every
  // lib.dom, so reach it through a narrow cast rather than the Document type.
  const posFromPoint = (
    doc as Document & {
      caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    }
  ).caretPositionFromPoint;
  if (typeof posFromPoint === "function") {
    const pos = posFromPoint.call(doc, x, y);
    if (!pos) return null;
    const range = doc.createRange();
    range.setStart(pos.offsetNode, pos.offset);
    return range;
  }
  return null;
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
