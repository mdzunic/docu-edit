// Hand-written Markdown parser (SPEC-050): maps the "Supported subset" onto
// the document model. No Markdown dependency — this is deliberately a small,
// documented subset (see specs/050-markdown-open-edit-and-convert.md), not a
// CommonMark/GFM-conformant parser. Anything outside the subset (raw HTML
// blocks, reference-style link definitions, an image whose `src` cannot be
// resolved) is preserved verbatim as an OpaqueBlock rather than interpreted
// or dropped — the same SPEC-008 fidelity rule the .docx/.odt readers follow,
// reusing `OpaqueBlock.xml` as a plain-text (not XML) verbatim payload, which
// lib/equality.ts's canonicalizeXml already falls back to string identity for
// on a parse failure (see lib/equality.ts's `xmlEquivalent`).
import {
  createDefaultStyles,
  HORIZONTAL_RULE_STYLE_ID,
  type Block,
  type DefaultStyles,
  type DocumentModel,
  type DocumentProperties,
  type Image,
  type NamedStyle,
  type OpaqueBlock,
  type OpaquePart,
  type Paragraph,
  type Run,
  type RunFormatting,
  type Section,
  type SourceXml,
  type Table,
  type TableRow,
} from "../documentModel";
import { type ListFormatting } from "../core/lists";
import { type MediaPart } from "../core/media";
import { equalColumnWidthsCm } from "../tables";
import { textAreaWidthCm } from "../images";
import { fitImageSizeCm } from "../core/media";
import { decodeImageDimensions } from "../core/imageDimensions";
import { bytesToBase64 } from "../opaque";
import { isFetchableUrl, sanitizeHref } from "../links";
import { sanitizeXmlText } from "../core/xmlText";
import { CODE_CHAR_STYLE_ID, headingStyleId, mdBuiltInStyles, resolveNamedStyle, type ResolvedNamedStyle } from "../styles";

/** The parsed-package shape `assembleDocumentModel` (lib/openDocument.ts)
 * consumes, matching `ParsedDocxDocument`/`ParsedOdtDocument`'s shape. */
export interface ParsedMdDocument {
  sections: DocumentModel["sections"];
  defaultStyles: DocumentModel["defaultStyles"];
  styles: NamedStyle[];
  opaque: OpaquePart[];
  sourceXml: SourceXml;
  media: MediaPart[];
  properties: DocumentProperties;
  /** YAML front matter, verbatim, including its `---` delimiters — see
   * DocumentModel.frontMatter. */
  frontMatter?: string;
}

export interface ParseMdOptions {
  /** The base URL a relative image `src` resolves against (SPEC-050's
   * "Supported subset": "`src` resolved relative to the file"). Only
   * meaningful for a document opened from a URL (see
   * lib/openDocumentFromUrl.ts) — a file picked locally or dropped has no
   * such base, so a relative `src` there cannot be resolved and the image
   * is preserved as an opaque block instead of failing the whole open. */
  baseUrl?: string;
}

function decodeUtf8(data: ArrayBuffer | Uint8Array): string {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  return new TextDecoder("utf-8").decode(bytes);
}

// --- front matter ------------------------------------------------------------

/** Splits off a leading YAML front-matter block (`---` ... `---`/`...`),
 * returned verbatim (SPEC-050 AC-4 — never parsed). Returns the remaining
 * body lines and the front-matter text (undefined when the file has none). */
function splitFrontMatter(lines: string[]): { body: string[]; frontMatter: string | undefined } {
  if (lines[0] !== "---") {
    return { body: lines, frontMatter: undefined };
  }
  for (let i = 1; i < lines.length; i++) {
    if (lines[i] === "---" || lines[i] === "...") {
      return {
        body: lines.slice(i + 1),
        frontMatter: lines.slice(0, i + 1).join("\n"),
      };
    }
  }
  return { body: lines, frontMatter: undefined };
}

// --- inline parsing (bold/italic/strikethrough/code/links) ------------------

function currentFormatting(bold: boolean, italic: boolean, strikethrough: boolean): RunFormatting {
  const formatting: RunFormatting = {};
  if (bold) {
    formatting.bold = true;
  }
  if (italic) {
    formatting.italic = true;
  }
  if (strikethrough) {
    formatting.strikethrough = true;
  }
  return formatting;
}

/** The index just past a `(...)` target's closing paren, honouring one level
 * of nested parens (a URL can legitimately contain them); -1 if unterminated. */
function findClosingParen(text: string, start: number): number {
  let depth = 1;
  for (let i = start; i < text.length; i++) {
    if (text[i] === "(") {
      depth++;
    } else if (text[i] === ")") {
      depth--;
      if (depth === 0) {
        return i;
      }
    }
  }
  return -1;
}

/** Parses run-level Markdown (the "Supported subset" table's inline rows)
 * into `Run[]`: `\` escapes, `` `code` `` spans (the Code character style,
 * SPEC-050), `**bold**`/`*italic*`/`~~strikethrough~~` (toggled, so nested
 * combinations like `**bold *and italic***` resolve correctly), and
 * `[text](url)` hyperlinks (recursing on the link text so it can itself
 * carry formatting). Images are handled at the block level (see
 * `parseMd`) — the document model has no inline-image run shape, only a
 * block-level `Image` (mirroring how the .docx/.odt readers split a
 * paragraph around one), so an inline `![]()` mid-paragraph is left as
 * literal text (SPEC-050's supported subset covers a standalone image
 * paragraph — see "Fixtures"). `resolveCodeFormatting` is called (and only
 * called) once a `` `code` `` span is actually found — it seeds the
 * document's style catalogue with `CodeChar` on first use and returns its
 * resolved formatting, which is merged onto the code run's baked
 * formatting the same way a paragraph's own style contribution is baked
 * elsewhere in this reader (every caller must route through the same
 * lazily-seeding resolver so a code span renders monospace regardless of
 * which other constructs the document happens to contain). */
function parseInline(text: string, resolveCodeFormatting: () => RunFormatting): Run[] {
  const runs: Run[] = [];
  let bold = false;
  let italic = false;
  let strikethrough = false;
  let buffer = "";

  const flush = () => {
    if (buffer.length > 0) {
      runs.push({ text: buffer, formatting: currentFormatting(bold, italic, strikethrough) });
      buffer = "";
    }
  };

  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === "\\" && i + 1 < text.length) {
      buffer += text[i + 1];
      i += 2;
      continue;
    }
    if (ch === "`") {
      let fenceLen = 1;
      while (text[i + fenceLen] === "`") {
        fenceLen++;
      }
      const fence = "`".repeat(fenceLen);
      const close = text.indexOf(fence, i + fenceLen);
      if (close === -1) {
        buffer += fence;
        i += fenceLen;
        continue;
      }
      flush();
      const code = text.slice(i + fenceLen, close);
      runs.push({
        text: code,
        formatting: { ...currentFormatting(bold, italic, strikethrough), ...resolveCodeFormatting() },
        styleId: CODE_CHAR_STYLE_ID,
      });
      i = close + fenceLen;
      continue;
    }
    if (text.startsWith("**", i)) {
      flush();
      bold = !bold;
      i += 2;
      continue;
    }
    if (text.startsWith("~~", i)) {
      flush();
      strikethrough = !strikethrough;
      i += 2;
      continue;
    }
    if (ch === "*" || ch === "_") {
      // The "Supported subset" table names `*italic*`, but the writer's own
      // canonical spelling is `_italic_` (AC-7: "a file using *italic* ...
      // comes back as _italic_") — both are recognised on read so that
      // canonical output round-trips, and both toggle the same flag rather
      // than being tracked separately, since Markdown never nests italic
      // inside italic.
      flush();
      italic = !italic;
      i += 1;
      continue;
    }
    if (ch === "[" && text[i - 1] !== "!") {
      const closeBracket = text.indexOf("]", i + 1);
      if (closeBracket !== -1 && text[closeBracket + 1] === "(") {
        const closeParen = findClosingParen(text, closeBracket + 2);
        if (closeParen !== -1) {
          flush();
          const linkText = text.slice(i + 1, closeBracket);
          // Sanitized on the way in, exactly as the .docx and .odt readers do
          // (docx/reader.ts's hyperlinkHref, odf/reader.ts's linkFormatting):
          // a `javascript:`/`data:`/`vbscript:` target leaves the text
          // unlinked rather than reaching the `:href` a rendered run binds.
          const href = sanitizeHref(text.slice(closeBracket + 2, closeParen));
          const outer = currentFormatting(bold, italic, strikethrough);
          for (const run of parseInline(linkText, resolveCodeFormatting)) {
            const formatting = { ...outer, ...run.formatting };
            runs.push({ ...run, formatting: href !== null ? { ...formatting, linkHref: href } : formatting });
          }
          i = closeParen + 1;
          continue;
        }
      }
    }
    buffer += ch;
    i++;
  }
  flush();
  return runs;
}

/** The inline grammar above over one line of text, as `Run[]` — a thin,
 * additive wrapper for SPEC-106's slide Markdown (generate/slideMarkdown.ts),
 * which reuses this vocabulary rather than re-implementing it. A code span
 * resolves against the built-in `CodeChar` style exactly as `parseMd`'s own
 * spans do; no `parseInline` caller changes. */
export function parseInlineRuns(text: string): Run[] {
  const stub = { defaultStyles: createDefaultStyles(), styles: mdBuiltInStyles() } as DocumentModel;
  return parseInline(text, () => resolveNamedStyle(stub, CODE_CHAR_STYLE_ID).runFormatting);
}

/** Joins a block's physical source lines into one logical line, honouring
 * "Two-space or backslash line end -> LineBreak" (a literal `\n` embedded in
 * run text, per Run's own doc comment — this model has no separate
 * LineBreak node) versus an ordinary soft wrap (a single space). */
function joinLines(rawLines: string[]): string {
  const parts: string[] = [];
  const hardBreakAfter: boolean[] = [];
  rawLines.forEach((raw, index) => {
    const hard = / {2,}$/.test(raw) || raw.endsWith("\\");
    let line = raw.replace(/\\$/, "").replace(/[ \t]+$/, "");
    line = index === 0 ? line.replace(/^[ \t]+/, "") : line.trim();
    parts.push(line);
    hardBreakAfter.push(hard);
  });
  let result = parts[0] ?? "";
  for (let index = 1; index < parts.length; index++) {
    result += (hardBreakAfter[index - 1] ? "\n" : " ") + parts[index];
  }
  return result;
}

// --- images -------------------------------------------------------------

function guessContentType(src: string): string {
  const ext = src.split(/[?#]/)[0].split(".").pop()?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}

/** Resolves an image `src` to its bytes: a `data:` URI is decoded directly
 * (no network — always resolvable, and what SPEC-050's own kitchen-sink
 * fixture uses to stay hermetic); anything else is fetched, resolved against
 * `options.baseUrl` when given. Returns null when the src cannot be resolved
 * (a relative path with no base, a scheme we will not fetch, or a failed
 * fetch) — the caller preserves the image reference as an opaque block
 * instead of failing the open. */
async function resolveImageSrc(
  src: string,
  options: ParseMdOptions,
): Promise<{ contentType: string; data: string } | null> {
  const dataUri = /^data:([^;,]*)(;base64)?,([\s\S]*)$/.exec(src);
  if (dataUri) {
    const [, rawContentType, isBase64, payload] = dataUri;
    const contentType = rawContentType || "application/octet-stream";
    if (isBase64) {
      return { contentType, data: payload.replace(/\s+/g, "") };
    }
    return { contentType, data: bytesToBase64(new TextEncoder().encode(decodeURIComponent(payload))) };
  }
  let url: string;
  try {
    url = options.baseUrl ? new URL(src, options.baseUrl).toString() : new URL(src).toString();
  } catch {
    return null;
  }
  // Checked on the RESOLVED url, so a relative `src` still resolves against
  // an http(s) baseUrl and passes, while an absolute `src` naming any other
  // scheme does not: the document is untrusted, and this is the app issuing
  // a request on the reader's behalf rather than following a link they
  // clicked. `data:` never reaches here — it returned above.
  if (!isFetchableUrl(url)) {
    return null;
  }
  try {
    const response = await fetch(url);
    if (!response.ok) {
      return null;
    }
    const contentType = response.headers.get("content-type")?.split(";")[0]?.trim() || guessContentType(src);
    return { contentType, data: bytesToBase64(new Uint8Array(await response.arrayBuffer())) };
  } catch {
    return null;
  }
}

// --- table rows ---------------------------------------------------------

/** Splits one `| a | b |` row into trimmed cell strings, honouring `\|` as a
 * literal pipe. Leading/trailing pipes are optional, per GFM. */
function splitTableRow(line: string): string[] {
  let trimmed = line.trim();
  if (trimmed.startsWith("|")) {
    trimmed = trimmed.slice(1);
  }
  if (trimmed.endsWith("|") && !trimmed.endsWith("\\|")) {
    trimmed = trimmed.slice(0, -1);
  }
  const cells: string[] = [];
  let current = "";
  for (let i = 0; i < trimmed.length; i++) {
    if (trimmed[i] === "\\" && i + 1 < trimmed.length) {
      current += trimmed[i] + trimmed[i + 1];
      i++;
      continue;
    }
    if (trimmed[i] === "|") {
      cells.push(current);
      current = "";
      continue;
    }
    current += trimmed[i];
  }
  cells.push(current);
  return cells.map((cell) => cell.trim().replace(/\\\|/g, "|"));
}

/** Whether `line` is a GFM table delimiter row (`| --- | :--: |`, …) —
 * alignment colons are recognised as valid syntax but not modelled (the
 * document model has no per-column alignment, and SPEC-050's "Supported
 * subset" does not ask for one), so they are accepted here and ignored. */
function isTableDelimiterRow(line: string): boolean {
  const cells = splitTableRow(line);
  return cells.length > 0 && cells.every((cell) => /^:?-+:?$/.test(cell));
}

// --- block parsing --------------------------------------------------------

const FENCE_RE = /^ {0,3}(```+|~~~+)(.*)$/;
const HEADING_RE = /^(#{1,6})[ \t]+(.*?)[ \t]*#*$/;
const THEMATIC_BREAK_RE = /^ {0,3}([-*_])[ \t]*(?:\1[ \t]*){2,}$/;
const BLOCKQUOTE_RE = /^ {0,3}>[ \t]?/;
const LIST_ITEM_RE = /^( *)([-*+]|\d+[.)])[ \t]+(.*)$/;
const STANDALONE_IMAGE_RE = /^!\[([^\]]*)\]\(([^)]*)\)[ \t]*$/;
const HTML_BLOCK_RE = /^ {0,3}<[a-zA-Z!/]/;
const REFERENCE_DEF_RE = /^ {0,3}\[[^\]]+\]:\s*\S/;

function looksLikeBlockStart(line: string): boolean {
  return (
    line.trim() === "" ||
    FENCE_RE.test(line) ||
    HEADING_RE.test(line) ||
    THEMATIC_BREAK_RE.test(line) ||
    BLOCKQUOTE_RE.test(line) ||
    LIST_ITEM_RE.test(line) ||
    HTML_BLOCK_RE.test(line) ||
    REFERENCE_DEF_RE.test(line)
  );
}

function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

const bakeRun = (run: Run, base: RunFormatting): Run => ({ ...run, formatting: { ...base, ...run.formatting } });

const bakeParagraphFormatting = (base: ResolvedNamedStyle) =>
  Object.keys(base.paragraphFormatting).length > 0 ? base.paragraphFormatting : undefined;

/** Parses Markdown source into the document model (SPEC-050). Hand-written,
 * line-oriented: blocks are recognised by the "Supported subset" table's
 * patterns in priority order (fenced code, raw HTML, a reference
 * definition, an ATX heading, a thematic break, a blockquote, a GFM table, a
 * list item, a standalone image, then a plain paragraph as the fallback),
 * so nothing needs backtracking across block kinds. `no CommonMark/GFM
 * conformance` (explicitly out of scope) — this recognises exactly the
 * subset the spec names and preserves everything else verbatim rather than
 * guessing at it. */
export async function parseMd(
  data: ArrayBuffer | Uint8Array,
  options: ParseMdOptions = {},
): Promise<ParsedMdDocument> {
  // Sanitized at the decode, before anything is parsed out of it, so the
  // invariant holds for every payload this reader produces — run text,
  // preserved opaque blocks and front matter alike — rather than only the
  // paths that happen to route through insertText. Markdown is the one
  // non-XML source we read, so it is the one reader that can carry a scalar
  // XML cannot (core/xmlText.ts). SPEC-050 asserts `equalStructural` on a
  // `.md` round trip and explicitly not `equalBytes`, so normalising here is
  // within what the format's round-trip contract allows.
  const text = sanitizeXmlText(decodeUtf8(data)).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const { body, frontMatter } = splitFrontMatter(text.split("\n"));

  const defaultStyles: DefaultStyles = createDefaultStyles();
  const modelStub = { defaultStyles } as DocumentModel;

  let styles: NamedStyle[] = [];
  let stylesSeeded = false;
  const ensureMdStyles = (): NamedStyle[] => {
    if (!stylesSeeded) {
      styles = mdBuiltInStyles();
      stylesSeeded = true;
    }
    return styles;
  };
  const codeLanguageStyleIds = new Map<string, string>();
  const codeStyleIdFor = (language: string): string => {
    if (language.length === 0) {
      return "Code";
    }
    let id = codeLanguageStyleIds.get(language);
    if (!id) {
      id = `Code:${language}`;
      codeLanguageStyleIds.set(language, id);
      ensureMdStyles().push({ id, name: "Code", basedOn: "Code", language });
    }
    return id;
  };

  // Every reader in this codebase bakes a paragraph's *effective* (resolved)
  // formatting onto its runs, keeping the style reference alongside it
  // (equality.ts's `resolve()` relies on this — see specs/002's
  // "Implementation status"); a heading/blockquote/code-block paragraph
  // here is no different, even though the source markup only ever encodes
  // the direct layer (bold/italic/strikethrough marks) — the style's own
  // contribution (a heading's bold+size, Quote's indent+spacing) must be
  // merged in explicitly since there is no separate style-resolution pass
  // for a freshly parsed document the way docx/odf styles.xml gets one.
  const resolveStyle = (styleId: string): ResolvedNamedStyle => resolveNamedStyle({ ...modelStub, styles }, styleId);

  // Passed to every `parseInline` call so an inline code span seeds the
  // catalogue and gets its monospace formatting baked in regardless of
  // which block it's found in — see `parseInline`'s doc comment.
  const codeCharFormatting = (): RunFormatting => {
    ensureMdStyles();
    return resolveStyle(CODE_CHAR_STYLE_ID).runFormatting;
  };

  const media: MediaPart[] = [];
  const mediaBySrc = new Map<string, string>();
  let mediaCounter = 0;

  const blocks: Block[] = [];
  const opaqueBlocks: OpaqueBlock[] = [];
  const pushOpaque = (raw: string): void => {
    opaqueBlocks.push({ xml: raw, blockIndex: blocks.length });
  };

  const lines = body;
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === "") {
      i++;
      continue;
    }

    const fence = FENCE_RE.exec(line);
    if (fence) {
      const markerChar = fence[1][0];
      const markerLen = fence[1].length;
      const closeRe = new RegExp(`^ {0,3}\\${markerChar}{${markerLen},}[ \\t]*$`);
      const language = fence[2].trim();
      i++;
      const codeLines: string[] = [];
      while (i < lines.length && !closeRe.test(lines[i])) {
        codeLines.push(lines[i]);
        i++;
      }
      i++; // the closing fence (or EOF, harmlessly past the end)
      const styleId = codeStyleIdFor(language);
      const codeBase = resolveStyle(styleId);
      const codeParagraphFormatting = bakeParagraphFormatting(codeBase);
      if (codeLines.length === 0) {
        blocks.push({ type: "paragraph", runs: [], formatting: codeParagraphFormatting, styleId });
      } else {
        for (const codeLine of codeLines) {
          blocks.push({
            type: "paragraph",
            runs: [{ text: codeLine, formatting: { ...codeBase.runFormatting } }],
            formatting: codeParagraphFormatting,
            styleId,
          });
        }
      }
      continue;
    }

    if (HTML_BLOCK_RE.test(line)) {
      const htmlLines: string[] = [];
      while (i < lines.length && lines[i].trim() !== "") {
        htmlLines.push(lines[i]);
        i++;
      }
      pushOpaque(htmlLines.join("\n"));
      continue;
    }

    if (REFERENCE_DEF_RE.test(line)) {
      pushOpaque(line);
      i++;
      continue;
    }

    const heading = HEADING_RE.exec(line);
    if (heading) {
      const level = heading[1].length;
      const styleId = headingStyleId(level);
      const base = resolveStyle(styleId);
      blocks.push({
        type: "paragraph",
        runs: parseInline(heading[2].trim(), codeCharFormatting).map((run) => bakeRun(run, base.runFormatting)),
        formatting: bakeParagraphFormatting(base),
        styleId,
        headingLevel: level,
      });
      i++;
      continue;
    }

    if (THEMATIC_BREAK_RE.test(line)) {
      ensureMdStyles();
      blocks.push({ type: "paragraph", runs: [], styleId: HORIZONTAL_RULE_STYLE_ID });
      i++;
      continue;
    }

    if (BLOCKQUOTE_RE.test(line)) {
      const quoteLines: string[] = [];
      while (i < lines.length && BLOCKQUOTE_RE.test(lines[i])) {
        quoteLines.push(lines[i].replace(BLOCKQUOTE_RE, ""));
        i++;
      }
      const base = resolveStyle("Quote");
      blocks.push({
        type: "paragraph",
        runs: parseInline(joinLines(quoteLines), codeCharFormatting).map((run) => bakeRun(run, base.runFormatting)),
        formatting: bakeParagraphFormatting(base),
        styleId: "Quote",
      });
      continue;
    }

    if (line.includes("|") && i + 1 < lines.length && isTableDelimiterRow(lines[i + 1])) {
      const headerCells = splitTableRow(line);
      const columnCount = headerCells.length;
      i += 2;
      const bodyRows: string[][] = [];
      while (i < lines.length && lines[i].trim() !== "" && lines[i].includes("|")) {
        bodyRows.push(splitTableRow(lines[i]));
        i++;
      }
      const buildRow = (cells: string[]): TableRow => ({
        cells: Array.from({ length: columnCount }, (_, col) => ({
          blocks: [{ type: "paragraph", runs: parseInline(cells[col] ?? "", codeCharFormatting) }],
        })),
      });
      const table: Table = {
        type: "table",
        rows: [{ ...buildRow(headerCells), isHeader: true }, ...bodyRows.map(buildRow)],
        columnWidthsCm: equalColumnWidthsCm(modelStub, columnCount),
      };
      blocks.push(table);
      continue;
    }

    const listItem = LIST_ITEM_RE.exec(line);
    if (listItem) {
      const [, indent, marker, itemText] = listItem;
      const level = Math.min(9, Math.floor(indent.length / 2) + 1);
      const ordered = /\d/.test(marker);
      const list: ListFormatting = { type: ordered ? "number" : "bullet", level };
      if (ordered) {
        // Only meaningful on the first item of a contiguous run at this
        // level (lib/lists.ts's markersForBlocks only ever reads it there
        // — numbering is adjacency-derived, not identity-derived); setting
        // it redundantly on every item is harmless for a pure .md round
        // trip but is not what a .docx/.odt round trip of an ad-hoc
        // (unnumbered-definition) list preserves, so it is skipped here to
        // match.
        const previous = blocks[blocks.length - 1];
        const continuesRun =
          previous?.type === "paragraph" && previous.list?.type === "number" && previous.list.level === level;
        const startAt = Number.parseInt(marker, 10);
        if (!continuesRun && startAt !== 1) {
          list.startAt = startAt;
        }
      }
      blocks.push({ type: "paragraph", runs: parseInline(itemText.trim(), codeCharFormatting), list });
      i++;
      continue;
    }

    const standaloneImage = STANDALONE_IMAGE_RE.exec(line);
    if (standaloneImage && (i + 1 >= lines.length || lines[i + 1].trim() === "")) {
      const [, alt, src] = standaloneImage;
      const resolved = await resolveImageSrc(src.trim(), options);
      if (resolved) {
        let mediaId = mediaBySrc.get(src.trim());
        if (!mediaId) {
          mediaId = `md-image-${++mediaCounter}`;
          mediaBySrc.set(src.trim(), mediaId);
          media.push({ mediaId, contentType: resolved.contentType, data: resolved.data });
        }
        const bytes = Uint8Array.from(atob(resolved.data), (char) => char.charCodeAt(0));
        const dims = decodeImageDimensions(bytes, resolved.contentType) ?? { widthPx: 96, heightPx: 96 };
        const { widthCm, heightCm } = fitImageSizeCm(dims.widthPx, dims.heightPx, textAreaWidthCm(modelStub));
        const image: Image = {
          type: "image",
          mediaId,
          naturalWidthPx: dims.widthPx,
          naturalHeightPx: dims.heightPx,
          widthCm,
          heightCm,
          wrap: "inline",
        };
        if (alt) {
          image.altText = alt;
        }
        blocks.push(image);
      } else {
        pushOpaque(line);
      }
      i++;
      continue;
    }

    const paraLines: string[] = [line];
    i++;
    while (i < lines.length && !looksLikeBlockStart(lines[i])) {
      paraLines.push(lines[i]);
      i++;
    }
    blocks.push({ type: "paragraph", runs: parseInline(joinLines(paraLines), codeCharFormatting) });
  }

  if (blocks.length === 0) {
    blocks.push(emptyParagraph());
  }

  // The document-default font family/size "underlay" (styles.ts's
  // splitRunFormatting doc comment): every reader in this codebase bakes it
  // onto a run that doesn't otherwise get one from a style, so a plain run
  // survives a same-content .docx/.odt round trip unchanged. Applied once,
  // over every run, after the fact — simpler than threading it through each
  // of the block constructors above. A run with its own character-style
  // reference (an inline code span's Code character style) bakes *that*
  // style's own formatting instead of the bare document default — the code
  // span construction above only sets `styleId`, not its resolved
  // formatting, the same "reference kept alongside effective formatting"
  // deferral every other styleId in this file resolves right away. Layered
  // low to high: document default < the run's own character style < the
  // direct formatting already on the run.
  for (const block of blocks) {
    const paragraphs = block.type === "table" ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks)) : block.type === "paragraph" ? [block] : [];
    for (const paragraph of paragraphs) {
      for (const run of paragraph.runs) {
        const characterStyle = run.styleId ? resolveStyle(run.styleId).runFormatting : {};
        run.formatting = {
          fontFamily: defaultStyles.fontFamily,
          fontSizePt: defaultStyles.fontSizePt,
          ...characterStyle,
          ...run.formatting,
        };
      }
    }
  }

  const sections: Section[] = [{ blocks, ...(opaqueBlocks.length > 0 ? { opaqueBlocks } : {}) }];

  const parsed: ParsedMdDocument = {
    sections,
    defaultStyles,
    styles,
    opaque: [],
    sourceXml: {},
    media,
    properties: {},
  };
  if (frontMatter !== undefined) {
    parsed.frontMatter = frontMatter;
  }
  return parsed;
}
