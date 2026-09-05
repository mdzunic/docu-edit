// External HTML → document model import for paste (SPEC-016): Word Online,
// Google Docs and plain web-page fragments. Whitelist import — only what
// the spec's mapping table lists is read; everything else (scripts, styles,
// classes, positioning, event handlers, embedded remote content, unknown
// markup) is dropped. Parsing uses DOMParser, whose detached document never
// executes scripts or fetches subresources, so hostile payloads are inert
// while being walked.
//
// Deferred structure imports at the text level until its spec lands:
// list items become plain paragraphs (SPEC-018), tables flatten to
// tab-separated rows (SPEC-022), anchors keep their runs but lose the URL
// (SPEC-025), and every <img> is dropped behind the "Images aren't
// supported yet" gate (SPEC-024).
import type {
  Paragraph,
  ParagraphAlignment,
  ParagraphFormatting,
  Run,
  RunFormatting,
} from "./documentModel";
import { normalizeRuns } from "./textEditing";
import { cleanFontFamily } from "./odf/properties";
import { HIGHLIGHT_NAMES } from "./core/colors";
import { parseColorInput } from "./core/colors";
import { cmToTwips, twipsToCm } from "./core/units";

export interface HtmlImportResult {
  blocks: Paragraph[];
  /** True when the fragment contained `<img>` content, which is dropped
   * until SPEC-024 — the caller shows "Images aren't supported yet". */
  imagesDropped: boolean;
}

// Active/embedded content removed wholesale, descendants included — the
// sanitisation contract's "drops scripts, classes, positioning and unknown
// markup" plus every "embedded remote content" carrier.
const DROP_TAGS = new Set([
  "script",
  "style",
  "link",
  "meta",
  "title",
  "head",
  "base",
  "iframe",
  "frame",
  "frameset",
  "object",
  "embed",
  "applet",
  "param",
  "video",
  "audio",
  "source",
  "track",
  "form",
  "button",
  "input",
  "select",
  "option",
  "textarea",
  "svg",
  "math",
  "template",
  "noscript",
  "canvas",
  "map",
  "area",
  "dialog",
]);

// Elements that start a new paragraph. Containers that merely group blocks
// (ul/ol, table sections, dl, body) recurse without opening one.
const BLOCK_TAGS = new Set([
  "p",
  "div",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "li",
  "dt",
  "dd",
  "blockquote",
  "pre",
  "address",
  "section",
  "article",
  "aside",
  "header",
  "footer",
  "main",
  "nav",
  "figure",
  "figcaption",
  "caption",
  "center",
  "details",
  "summary",
]);

// Blocks that represent a real (possibly empty) paragraph on their own —
// an empty <p></p> imports as an empty paragraph, an empty wrapper <div>
// does not.
const PARAGRAPH_TAGS = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "dt", "dd", "pre"]);

// A style value that could smuggle a fetch or script is never interpreted.
function unsafeStyleValue(value: string): boolean {
  const lowered = value.toLowerCase();
  return (
    lowered.includes("url(") ||
    lowered.includes("expression") ||
    lowered.includes("javascript:") ||
    lowered.includes("vbscript:") ||
    lowered.includes("data:")
  );
}

function styleDeclarations(el: Element): Array<[string, string]> {
  const raw = el.getAttribute("style");
  if (!raw) {
    return [];
  }
  const declarations: Array<[string, string]> = [];
  for (const part of raw.split(";")) {
    const colon = part.indexOf(":");
    if (colon <= 0) {
      continue;
    }
    const property = part.slice(0, colon).trim().toLowerCase();
    const value = part.slice(colon + 1).trim();
    if (value.length === 0 || unsafeStyleValue(value)) {
      continue;
    }
    declarations.push([property, value]);
  }
  return declarations;
}

const CSS_KEYWORDS = new Set([
  "transparent",
  "inherit",
  "initial",
  "unset",
  "revert",
  "currentcolor",
  "none",
  "auto",
  "normal",
]);

/** A CSS colour as the model stores it: lowercase hex from hex/rgb() forms,
 * a plain named colour kept as-is. Undefined for keywords and anything not
 * recognised — an unparseable colour is dropped, never passed through. */
function parseCssColor(value: string): string | undefined {
  const lowered = value.trim().toLowerCase();
  if (CSS_KEYWORDS.has(lowered)) {
    return undefined;
  }
  const hex = parseColorInput(lowered);
  if (hex) {
    return hex;
  }
  return /^[a-z]+$/.test(lowered) ? lowered : undefined;
}

/** Background colour → highlight: the OOXML palette name when the colour is
 * one of the 16 palette entries, else the lowercase hex. White and keyword
 * backgrounds mean "no highlight" on import. */
function parseHighlight(value: string): string | undefined {
  const color = parseCssColor(value);
  if (!color) {
    return undefined;
  }
  const hex = parseColorInput(color) ?? HIGHLIGHT_HEX_BY_NAME.get(color);
  if (!hex || hex === "#ffffff") {
    return undefined;
  }
  return HIGHLIGHT_NAMES[hex] ?? hex;
}

// CSS named colours that coincide with the OOXML highlight palette, so
// `background: yellow` round-trips as the palette name.
const HIGHLIGHT_HEX_BY_NAME = new Map<string, string>(
  Object.entries(HIGHLIGHT_NAMES)
    .filter(([, name]) => /^[a-z]+$/.test(name))
    .map(([hex, name]) => [name, hex]),
);

/** Font size in points from a CSS length: pt kept, px converted (96 px =
 * 72 pt), snapped to Word's half-point grid. Relative units are skipped —
 * without a computed-style context they cannot be resolved. */
function parseFontSizePt(value: string): number | undefined {
  const match = /^([\d.]+)(px|pt)$/.exec(value.trim().toLowerCase());
  if (!match) {
    return undefined;
  }
  const size = Number(match[1]);
  if (!Number.isFinite(size) || size <= 0) {
    return undefined;
  }
  const points = match[2] === "px" ? size * 0.75 : size;
  const snapped = Math.round(points * 2) / 2;
  return snapped > 0 ? snapped : undefined;
}

function snapCm(cm: number): number {
  return twipsToCm(cmToTwips(cm));
}

/** A CSS length in centimetres; undefined for relative/unresolvable units. */
function parseLengthCm(value: string): number | undefined {
  const match = /^(-?[\d.]+)(px|pt|cm|mm|in)?$/.exec(value.trim().toLowerCase());
  if (!match) {
    return undefined;
  }
  const size = Number(match[1]);
  if (!Number.isFinite(size)) {
    return undefined;
  }
  if (size === 0) {
    return 0;
  }
  switch (match[2]) {
    case "px":
      return snapCm((size / 96) * 2.54);
    case "pt":
      return snapCm((size / 72) * 2.54);
    case "cm":
      return snapCm(size);
    case "mm":
      return snapCm(size / 10);
    case "in":
      return snapCm(size * 2.54);
    default:
      return undefined;
  }
}

function parseLengthPt(value: string): number | undefined {
  const match = /^(-?[\d.]+)(px|pt)?$/.exec(value.trim().toLowerCase());
  if (!match) {
    return undefined;
  }
  const size = Number(match[1]);
  if (!Number.isFinite(size)) {
    return undefined;
  }
  if (size === 0) {
    return 0;
  }
  if (match[2] === "px") {
    return Math.round(size * 0.75 * 20) / 20;
  }
  if (match[2] === "pt") {
    return Math.round(size * 20) / 20;
  }
  return undefined;
}

// The legacy <font size="1".."7"> scale, in points.
const FONT_SIZE_STEPS: Record<string, number> = {
  "1": 8,
  "2": 10,
  "3": 12,
  "4": 14,
  "5": 18,
  "6": 24,
  "7": 36,
};

function withKey<K extends keyof RunFormatting>(
  formatting: RunFormatting,
  key: K,
  value: RunFormatting[K] | undefined,
): RunFormatting {
  const next = { ...formatting };
  if (value === undefined) {
    delete next[key];
  } else {
    next[key] = value;
  }
  return next;
}

/** The run formatting a tag itself implies, layered over the inherited
 * context: <b>/<strong> bold, <i>/<em> italic, <u>/<ins> underline,
 * <s>/<strike>/<del> strikethrough, <sub>/<sup> vertical alignment. */
function formattingFromTag(tag: string, base: RunFormatting): RunFormatting {
  switch (tag) {
    case "b":
    case "strong":
      return { ...base, bold: true };
    case "i":
    case "em":
      return { ...base, italic: true };
    case "u":
    case "ins":
      return { ...base, underline: true };
    case "s":
    case "strike":
    case "del":
      return { ...base, strikethrough: true };
    case "sub":
      return { ...base, vertAlign: "subscript" };
    case "sup":
      return { ...base, vertAlign: "superscript" };
    default:
      return base;
  }
}

/** Legacy <font color/face/size> attributes. */
function formattingFromFontElement(el: Element, base: RunFormatting): RunFormatting {
  let formatting = base;
  const color = el.getAttribute("color");
  if (color) {
    const parsed = parseCssColor(color);
    if (parsed) {
      formatting = { ...formatting, color: parsed };
    }
  }
  const face = el.getAttribute("face");
  if (face) {
    const family = cleanFontFamily(face);
    if (family) {
      formatting = { ...formatting, fontFamily: family };
    }
  }
  const size = el.getAttribute("size");
  if (size && FONT_SIZE_STEPS[size.trim()]) {
    formatting = { ...formatting, fontSizePt: FONT_SIZE_STEPS[size.trim()] };
  }
  return formatting;
}

/** Inline CSS → run formatting. Google Docs expresses bold/italic/underline
 * as CSS on spans rather than tags, so font-weight, font-style,
 * text-decoration and vertical-align are read alongside the mapping table's
 * color/font-family/font-size; an explicit "normal"/"none" turns the
 * inherited toggle off (Google wraps whole fragments in
 * <b style="font-weight:normal">). */
function formattingFromStyles(el: Element, base: RunFormatting): RunFormatting {
  let formatting = base;
  for (const [property, value] of styleDeclarations(el)) {
    const lowered = value.toLowerCase();
    if (property === "font-weight") {
      const numeric = Number(lowered);
      if (lowered === "bold" || lowered === "bolder" || numeric >= 600) {
        formatting = { ...formatting, bold: true };
      } else if (lowered === "normal" || lowered === "lighter" || (Number.isFinite(numeric) && numeric < 600)) {
        formatting = withKey(formatting, "bold", undefined);
      }
    } else if (property === "font-style") {
      if (lowered === "italic" || lowered === "oblique") {
        formatting = { ...formatting, italic: true };
      } else if (lowered === "normal") {
        formatting = withKey(formatting, "italic", undefined);
      }
    } else if (property === "text-decoration" || property === "text-decoration-line") {
      if (lowered.includes("underline")) {
        formatting = { ...formatting, underline: true };
      }
      if (lowered.includes("line-through")) {
        formatting = { ...formatting, strikethrough: true };
      }
      if (lowered === "none") {
        formatting = withKey(formatting, "underline", undefined);
        formatting = withKey(formatting, "strikethrough", undefined);
      }
    } else if (property === "vertical-align") {
      if (lowered === "sub") {
        formatting = { ...formatting, vertAlign: "subscript" };
      } else if (lowered === "super") {
        formatting = { ...formatting, vertAlign: "superscript" };
      } else if (lowered === "baseline") {
        formatting = withKey(formatting, "vertAlign", undefined);
      }
    } else if (property === "color") {
      const color = parseCssColor(value);
      if (color) {
        formatting = { ...formatting, color };
      }
    } else if (property === "background-color" || property === "background") {
      const highlight = parseHighlight(value);
      if (highlight) {
        formatting = { ...formatting, highlight };
      }
    } else if (property === "font-size") {
      const size = parseFontSizePt(value);
      if (size) {
        formatting = { ...formatting, fontSizePt: size };
      }
    } else if (property === "font-family") {
      const family = cleanFontFamily(value);
      if (family) {
        formatting = { ...formatting, fontFamily: family };
      }
    }
  }
  return formatting;
}

interface BlockProps {
  formatting?: ParagraphFormatting;
  headingLevel?: number;
}

function parseAlignment(value: string): ParagraphAlignment | undefined {
  switch (value.toLowerCase()) {
    case "left":
    case "start":
      return "left";
    case "center":
      return "center";
    case "right":
    case "end":
      return "right";
    case "justify":
      return "justify";
    default:
      return undefined;
  }
}

/** Block CSS → paragraph formatting, per the mapping table: text-align,
 * margin-top/bottom, line-height, text-indent, padding-left/right, plus
 * dir/direction for RTL. Inherits the enclosing block's properties so a
 * wrapper's alignment reaches the paragraphs inside it. */
function blockPropsFrom(el: Element, tag: string, inherited: BlockProps): BlockProps {
  const formatting: ParagraphFormatting = { ...inherited.formatting };
  for (const [property, value] of styleDeclarations(el)) {
    if (property === "text-align") {
      const alignment = parseAlignment(value);
      if (alignment) {
        formatting.alignment = alignment;
      }
    } else if (property === "margin-top") {
      const points = parseLengthPt(value);
      if (points !== undefined && points >= 0) {
        formatting.spaceBeforePt = points;
      }
    } else if (property === "margin-bottom") {
      const points = parseLengthPt(value);
      if (points !== undefined && points >= 0) {
        formatting.spaceAfterPt = points;
      }
    } else if (property === "line-height") {
      const lowered = value.toLowerCase();
      if (/^[\d.]+$/.test(lowered)) {
        const multiple = Math.round(Number(lowered) * 20) / 20;
        if (multiple > 0) {
          formatting.lineSpacing = multiple;
          delete formatting.lineSpacingPt;
        }
      } else if (/^[\d.]+%$/.test(lowered)) {
        const multiple = Math.round(Number(lowered.slice(0, -1)) / 5) / 20;
        if (multiple > 0) {
          formatting.lineSpacing = multiple;
          delete formatting.lineSpacingPt;
        }
      } else {
        const points = parseLengthPt(lowered);
        if (points !== undefined && points > 0) {
          formatting.lineSpacingPt = points;
          delete formatting.lineSpacing;
        }
      }
    } else if (property === "text-indent") {
      const cm = parseLengthCm(value);
      if (cm !== undefined) {
        formatting.indentFirstLineCm = cm;
      }
    } else if (property === "padding-left") {
      const cm = parseLengthCm(value);
      if (cm !== undefined && cm >= 0) {
        formatting.indentStartCm = cm;
      }
    } else if (property === "padding-right") {
      const cm = parseLengthCm(value);
      if (cm !== undefined && cm >= 0) {
        formatting.indentEndCm = cm;
      }
    } else if (property === "direction" && value.toLowerCase() === "rtl") {
      formatting.direction = "rtl";
    }
  }
  if ((el.getAttribute("dir") ?? "").toLowerCase() === "rtl") {
    formatting.direction = "rtl";
  }
  const props: BlockProps = {};
  if (Object.keys(formatting).length > 0) {
    props.formatting = formatting;
  }
  const heading = /^h([1-6])$/.exec(tag);
  if (heading) {
    props.headingLevel = Number(heading[1]);
  } else if (inherited.headingLevel !== undefined) {
    props.headingLevel = inherited.headingLevel;
  }
  return props;
}

interface RowState {
  cells: number;
  cellBlocks: number;
}

function paragraphFromProps(runs: Run[], props: BlockProps): Paragraph {
  const paragraph: Paragraph = { type: "paragraph", runs };
  if (props.formatting && Object.keys(props.formatting).length > 0) {
    paragraph.formatting = { ...props.formatting };
  }
  if (props.headingLevel !== undefined) {
    paragraph.headingLevel = props.headingLevel;
  }
  return paragraph;
}

/** Imports an external HTML fragment into model paragraphs. */
export function importHtml(html: string): HtmlImportResult {
  const parsed = new DOMParser().parseFromString(html, "text/html");
  const blocks: Paragraph[] = [];
  let current: { runs: Run[]; props: BlockProps } | null = null;
  let imagesDropped = false;

  function flush(): void {
    if (!current) {
      return;
    }
    let runs = normalizeRuns(current.runs);
    // Trailing whitespace before the paragraph break is source formatting
    // noise; one trailing <br> is invisible in HTML rendering, so a lone
    // trailing "\n" goes with it.
    if (runs.length > 0) {
      const lastRun = runs[runs.length - 1];
      const trimmed = lastRun.text.replace(/[ ]+$/, "").replace(/\n$/, "");
      runs =
        trimmed === lastRun.text
          ? runs
          : normalizeRuns([
              ...runs.slice(0, -1),
              { text: trimmed, formatting: lastRun.formatting },
            ]);
    }
    blocks.push(paragraphFromProps(runs, current.props));
    current = null;
  }

  function ensure(props: BlockProps): { runs: Run[]; props: BlockProps } {
    if (!current) {
      current = { runs: [], props };
    }
    return current;
  }

  function lastChar(): string {
    if (!current) {
      return "";
    }
    for (let i = current.runs.length - 1; i >= 0; i--) {
      const text = current.runs[i].text;
      if (text.length > 0) {
        return text[text.length - 1];
      }
    }
    return "";
  }

  function appendText(
    raw: string,
    formatting: RunFormatting,
    props: BlockProps,
    pre: boolean,
  ): void {
    let text = raw;
    if (!pre) {
      text = text.replace(/\s+/g, " ");
      if (text.length === 0) {
        return;
      }
      const open = current !== null && current.runs.length > 0;
      if (text === " " && !open) {
        return;
      }
      if (!open) {
        text = text.replace(/^ +/, "");
      } else if (text.startsWith(" ") && lastChar() === " ") {
        text = text.slice(1);
      }
      if (text.length === 0) {
        return;
      }
    } else if (text.length === 0) {
      return;
    }
    ensure(props).runs.push({ text, formatting: { ...formatting } });
  }

  function appendBreak(formatting: RunFormatting, props: BlockProps): void {
    ensure(props).runs.push({ text: "\n", formatting: { ...formatting } });
  }

  function walk(
    node: Node,
    formatting: RunFormatting,
    props: BlockProps,
    pre: boolean,
    row: RowState | null,
  ): void {
    if (node.nodeType === Node.TEXT_NODE) {
      appendText((node as CharacterData).data, formatting, props, pre);
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }
    const el = node as Element;
    const tag = el.tagName.toLowerCase();
    if (DROP_TAGS.has(tag)) {
      return;
    }
    if (tag === "img") {
      imagesDropped = true;
      return;
    }
    if (tag === "br") {
      appendBreak(formatting, props);
      return;
    }
    if (tag === "hr") {
      flush();
      return;
    }
    if (tag === "tr") {
      flush();
      const rowState: RowState = { cells: 0, cellBlocks: 0 };
      ensure(props);
      for (const child of Array.from(el.childNodes)) {
        walk(child, formatting, props, pre, rowState);
      }
      flush();
      return;
    }
    if (tag === "td" || tag === "th") {
      if (row) {
        if (row.cells > 0) {
          ensure(props).runs.push({ text: "\t", formatting: {} });
        }
        row.cells += 1;
        row.cellBlocks = 0;
      }
      const cellFormatting = formattingFromStyles(el, formatting);
      for (const child of Array.from(el.childNodes)) {
        walk(child, cellFormatting, props, pre, row);
      }
      return;
    }
    if (BLOCK_TAGS.has(tag)) {
      const childFormatting = formattingFromStyles(el, formatting);
      if (row) {
        // A block inside a table cell stays on the row's paragraph; a
        // second block in the same cell separates with a forced line break.
        if (row.cellBlocks > 0) {
          appendBreak({}, props);
        }
        row.cellBlocks += 1;
        for (const child of Array.from(el.childNodes)) {
          walk(child, childFormatting, props, pre, row);
        }
        return;
      }
      flush();
      const childProps = blockPropsFrom(el, tag, props);
      const before = blocks.length;
      for (const child of Array.from(el.childNodes)) {
        walk(child, childFormatting, childProps, pre || tag === "pre", null);
      }
      flush();
      if (blocks.length === before && PARAGRAPH_TAGS.has(tag)) {
        blocks.push(paragraphFromProps([], childProps));
      }
      return;
    }
    // Inline whitelist tags contribute formatting; unknown inline markup is
    // dropped while its characters survive (recursed transparently). <a>
    // lands here too: runs kept, href discarded — the model has no link
    // field until SPEC-025, and a javascript: URL never survives because no
    // URL does.
    let childFormatting = formattingFromTag(tag, formatting);
    if (tag === "font") {
      childFormatting = formattingFromFontElement(el, childFormatting);
    }
    childFormatting = formattingFromStyles(el, childFormatting);
    for (const child of Array.from(el.childNodes)) {
      walk(child, childFormatting, props, pre, row);
    }
  }

  const body = parsed.body;
  for (const child of Array.from(body.childNodes)) {
    walk(child, {}, {}, false, null);
  }
  flush();
  return { blocks, imagesDropped };
}
