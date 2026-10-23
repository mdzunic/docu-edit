// Slide Markdown (SPEC-106) — a generation input, never a file format the
// product opens (`.md` stays SPEC-050's text format). One deck per text:
// `---` between blocks starts a slide, `#`/`##` opening a slide is its title,
// `-`/`*`/`+`/`1.` rows are body bullets nested by two-space indent, a
// standalone `![alt](data:…)` is an image, `Notes:` opens speaker notes to
// the end of the slide, and YAML front matter sets `title`/`format`/`size`.
// Those rules and only those: anything else is dropped and REPORTED as a
// `DeckDropRecord`, never guessed at and never half-parsed.
//
// Inline formatting is SPEC-050's own parser, reached through md/reader.ts's
// one additive export `parseInlineRuns` — the single place this module (the
// only one in generate/'s deck path that does) touches the text product,
// with the `Run` type. `writeSlideMarkdown` is the inverse in canonical
// spelling, following md/writer.ts's escaping discipline, and reports every
// construct a deck carries that Markdown cannot. The round trip asserted is
// `.md` → deck → `.md` with `equalStructural` on the model, never
// `equalBytes` on the text (SPEC-050's rule).
import { SLIDE_SIZES } from "../core/pageSetup";
import { sanitizeXmlText } from "../core/xmlText";
import { mediaDataUri } from "../core/media";
import type { Run } from "../documentModel";
import { parseInlineRuns } from "../md/reader";
import type {
  PresentationModel,
  Slide,
  SlideFrame,
  SlideParagraph,
  SlideRun,
  TextFrame,
} from "../presentation/model";
import { slideOutline } from "../presentation/slideModel";
import { titleBodyFrames, titleSlideFrames } from "../presentation/slides";
import { MAX_INDENT_LEVEL } from "../presentation/text";
import { slideLocation, type DeckDropKind, type DeckDropRecord } from "./deckDrops";
import {
  createDeckContext,
  DECK_SIZE_NAMES,
  decodeImage,
  imageInBodyRect,
  landImage,
  PRESENTATION_FORMATS,
  setSlideNotes,
  type DeckContext,
  type DecodedImage,
  type DeckSizeName,
} from "./deckSlides";
import { DocumentGenerationError } from "./errors";
import type { LengthInput } from "./length";
import { expectArray, expectPlainObject, expectString, rejectUnknownKeys } from "./validate";

/** Fallbacks for what front matter does not set. */
export interface SlideMarkdownOptions {
  title?: string;
  format?: "pptx" | "odp";
  size?: DeckSizeName | { width: LengthInput; height: LengthInput };
}

/** A parse and everything it dropped. */
export interface SlideMarkdownParse {
  deck: PresentationModel;
  dropped: DeckDropRecord[];
}

// --- the grammar's patterns (md/reader.ts's, restated for the slide rules) ----

const SEPARATOR_RE = /^---[ \t]*$/;
const NOTES_RE = /^Notes:(.*)$/;
const FENCE_RE = /^ {0,3}(```+|~~~+)(.*)$/;
// md/reader.ts's heading pattern, with CommonMark's closing-sequence rule
// (a closing `#` run must follow whitespace) so a title like "C#" survives.
const HEADING_RE = /^(#{1,6})[ \t]+(.*?)(?:[ \t]+#+)?[ \t]*$/;
const EMPTY_HEADING_RE = /^(#{1,6})[ \t]*$/;
const THEMATIC_BREAK_RE = /^ {0,3}([-*_])[ \t]*(?:\1[ \t]*){2,}$/;
const BLOCKQUOTE_RE = /^ {0,3}>[ \t]?/;
const LIST_ITEM_RE = /^( *)([-*+]|\d+[.)])[ \t]+(.*)$/;
const STANDALONE_IMAGE_RE = /^!\[([^\]]*)\]\(([^)]*)\)[ \t]*$/;
const HTML_BLOCK_RE = /^ {0,3}<[a-zA-Z!/]/;
const REFERENCE_DEF_RE = /^ {0,3}\[[^\]]+\]:\s*\S/;
const DATA_URI_RE = /^data:([^;,]*)((?:;[^;,]*)*),([\s\S]*)$/;
const FRONT_MATTER_KEY_RE = /^([A-Za-z0-9_-]+)[ \t]*:(.*)$/;
const FRONT_MATTER_KEYS = ["title", "format", "size"];

function isTableDelimiterRow(line: string): boolean {
  const cells = line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|");
  return cells.length > 0 && cells.every((cell) => /^:?-+:?$/.test(cell.trim()));
}

function startsBlock(line: string): boolean {
  return (
    line.trim() === "" ||
    NOTES_RE.test(line) ||
    FENCE_RE.test(line) ||
    HEADING_RE.test(line) ||
    THEMATIC_BREAK_RE.test(line) ||
    BLOCKQUOTE_RE.test(line) ||
    LIST_ITEM_RE.test(line) ||
    STANDALONE_IMAGE_RE.test(line) ||
    HTML_BLOCK_RE.test(line) ||
    REFERENCE_DEF_RE.test(line)
  );
}

// --- parsing ------------------------------------------------------------------

/** md/reader.ts's `splitFrontMatter` rule: a first line `---` opens front
 * matter closed by the next `---`/`...` line. */
function splitFrontMatter(lines: string[]): { body: string[]; frontMatter?: string[] } {
  if (lines[0] !== "---") {
    return { body: lines };
  }
  for (let i = 1; i < lines.length; i++) {
    if (lines[i] === "---" || lines[i] === "...") {
      return { body: lines.slice(i + 1), frontMatter: lines.slice(1, i) };
    }
  }
  return { body: lines };
}

/** A YAML scalar's value: double-quoted (JSON-compatible escapes),
 * single-quoted (`''` escapes a quote) or plain (a ` #` comment stripped). */
function yamlScalar(raw: string): string {
  const value = raw.trim();
  if (value.startsWith('"') && value.endsWith('"') && value.length >= 2) {
    try {
      return String(JSON.parse(value));
    } catch {
      return value.slice(1, -1);
    }
  }
  if (value.startsWith("'") && value.endsWith("'") && value.length >= 2) {
    return value.slice(1, -1).replace(/''/g, "'");
  }
  return value.replace(/[ \t]+#.*$/, "");
}

/** Front matter sets `title`, `format` and `size` and nothing else: any
 * other key is reported once, and so is a value the deck cannot take (a
 * nested value, an unknown format or size) — the deck keeps its fallback. */
function applyFrontMatter(ctx: DeckContext, lines: string[]): void {
  const report = (reason: string): void => {
    ctx.dropped.push({ kind: "frontMatterKey", location: "front matter", reason });
  };
  const entries: { key: string; value: string; nested: boolean }[] = [];
  for (const line of lines) {
    if (line.trim() === "" || line.trimStart().startsWith("#")) {
      continue;
    }
    const match = FRONT_MATTER_KEY_RE.exec(line);
    if (match) {
      entries.push({ key: match[1], value: yamlScalar(match[2]), nested: false });
    } else if ((/^\s/.test(line) || line.startsWith("- ")) && entries.length > 0) {
      entries[entries.length - 1].nested = true;
    } else {
      report(`unreadable front-matter line "${line.trim()}": not kept`);
    }
  }
  const reported = new Set<string>();
  for (const { key, value, nested } of entries) {
    if (reported.has(key)) {
      continue;
    }
    if (!FRONT_MATTER_KEYS.includes(key)) {
      reported.add(key);
      report(`"${key}" is not a slide front-matter key (title, format, size): not kept`);
    } else if (nested) {
      reported.add(key);
      report(`"${key}" takes a single value, not a nested one: not kept`);
    } else if (key === "title") {
      ctx.deck.title = value;
    } else if (key === "format") {
      const format = PRESENTATION_FORMATS.find((candidate) => candidate === value);
      if (format) {
        ctx.deck.format = format;
      } else {
        reported.add(key);
        report(`format "${value}" is not "pptx" or "odp": not kept`);
      }
    } else {
      const size = DECK_SIZE_NAMES.find((candidate) => candidate === value);
      if (size) {
        ctx.deck.slideWidthCm = SLIDE_SIZES[size].widthCm;
        ctx.deck.slideHeightCm = SLIDE_SIZES[size].heightCm;
      } else {
        reported.add(key);
        report(`size "${value}" is not "16:9" or "4:3": not kept`);
      }
    }
  }
}

function isBlankChunk(chunk: string[]): boolean {
  return chunk.every((line) => line.trim() === "");
}

/** Splits the body at `---` lines into one line list per slide. A fenced
 * block is opaque to the split, and so are speaker notes (which run to the
 * next separator). A leading or trailing empty chunk mints no slide. */
function splitSlides(lines: string[]): string[][] {
  const chunks: string[][] = [[]];
  let fence: RegExp | undefined;
  let inNotes = false;
  for (const line of lines) {
    if (fence) {
      chunks[chunks.length - 1].push(line);
      if (fence.test(line)) {
        fence = undefined;
      }
      continue;
    }
    if (SEPARATOR_RE.test(line)) {
      chunks.push([]);
      inNotes = false;
      continue;
    }
    chunks[chunks.length - 1].push(line);
    if (inNotes) {
      continue;
    }
    if (NOTES_RE.test(line)) {
      inNotes = true;
      continue;
    }
    const open = FENCE_RE.exec(line);
    if (open) {
      fence = new RegExp(`^ {0,3}\\${open[1][0]}{${open[1].length},}[ \\t]*$`);
    }
  }
  if (chunks.length > 0 && isBlankChunk(chunks[0])) {
    chunks.shift();
  }
  if (chunks.length > 0 && isBlankChunk(chunks[chunks.length - 1])) {
    chunks.pop();
  }
  return chunks;
}

/** Field → how its drop reads, for the run mappings in both directions. */
const RUN_FIELD_NAMES: Record<string, string> = {
  strikethrough: "strikethrough",
  code: "code-span formatting",
  color: "colour",
  sizePt: "size",
  fontFamily: "font",
  underline: "underline",
  letterSpacingPt: "letter spacing",
  linkSlide: "slide link",
  linkTooltip: "link tooltip",
};

function reportRunDrops(
  dropped: DeckDropRecord[],
  slideIndex: number,
  counts: Map<string, number>,
  because: (field: string) => string,
): void {
  for (const [field, count] of counts) {
    dropped.push({
      kind: "runFormatting",
      location: slideLocation(slideIndex),
      reason: `${RUN_FIELD_NAMES[field] ?? field} dropped from ${count} run${count === 1 ? "" : "s"}: ${because(field)}`,
    });
  }
}

function tally(counts: Map<string, number>, field: string): void {
  counts.set(field, (counts.get(field) ?? 0) + 1);
}

/** Adjacent runs that read the same (a dropped mark no longer telling them
 * apart) become one, so the same text always parses to the same runs. */
function coalesce(runs: SlideRun[]): SlideRun[] {
  const out: SlideRun[] = [];
  for (const run of runs) {
    const previous = out[out.length - 1];
    const { text: _text, ...format } = run;
    if (previous) {
      const { text: _previousText, ...previousFormat } = previous;
      if (JSON.stringify(format) === JSON.stringify(previousFormat)) {
        previous.text += run.text;
        continue;
      }
    }
    out.push({ ...run });
  }
  return out;
}

/** SPEC-050's inline grammar over one line, mapped onto slide runs: text,
 * bold, italic, underline and a link carry over; strikethrough and a code
 * span's formatting (which a slide run has no field for) are counted. */
function slideRunsOf(text: string, counts: Map<string, number>): SlideRun[] {
  return coalesce(slideRunsFrom(text, counts));
}

function slideRunsFrom(text: string, counts: Map<string, number>): SlideRun[] {
  return parseInlineRuns(text)
    .filter((run: Run) => run.text.length > 0)
    .map((run: Run) => {
      const slideRun: SlideRun = { text: run.text };
      const { formatting } = run;
      if (formatting.bold) slideRun.bold = true;
      if (formatting.italic) slideRun.italic = true;
      if (formatting.underline) slideRun.underline = true;
      if (formatting.linkHref) slideRun.linkHref = formatting.linkHref;
      if (formatting.strikethrough) tally(counts, "strikethrough");
      if (run.styleId !== undefined) {
        tally(counts, "code");
      } else {
        for (const key of Object.keys(formatting)) {
          if (!["bold", "italic", "underline", "linkHref", "strikethrough"].includes(key)) {
            tally(counts, key);
          }
        }
      }
      return slideRun;
    });
}

/** Whether a line ends in a hard break: two trailing spaces, or an
 * unescaped trailing backslash. */
function hardBreak(raw: string): boolean {
  if (/ {2,}$/.test(raw)) {
    return true;
  }
  const backslashes = /\\*$/.exec(raw)![0].length;
  return backslashes % 2 === 1;
}

interface SlideSource {
  title?: SlideRun[];
  subtitle: SlideParagraph[];
  body: SlideParagraph[];
  images: DecodedImage[];
  notes?: string[];
}

/** Parses one slide's lines into its content, reporting every drop. */
function readSlide(lines: string[], slideIndex: number, dropped: DeckDropRecord[]): SlideSource {
  const location = slideLocation(slideIndex);
  const drop = (kind: DeckDropKind, reason: string): void => {
    dropped.push({ kind, location, reason });
  };
  const counts = new Map<string, number>();
  const source: SlideSource = { subtitle: [], body: [], images: [] };
  let seenBlock = false;
  let seenList = false;
  const addParagraph = (runs: SlideRun[]): void => {
    if (runs.length === 0) {
      return;
    }
    // Plain paragraphs before any list are the first slide's subtitle; every
    // other plain paragraph is a body paragraph.
    const target = slideIndex === 0 && source.title !== undefined && !seenList ? source.subtitle : source.body;
    target.push({ runs });
  };

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (source.notes !== undefined) {
      // Speaker notes run to the end of the slide, verbatim but for trailing
      // whitespace and a leading `\` escape.
      const text = line.replace(/[ \t]+$/, "");
      source.notes.push(text.startsWith("\\") ? text.slice(1) : text);
      i++;
      continue;
    }
    if (line.trim() === "") {
      i++;
      continue;
    }
    const notes = NOTES_RE.exec(line);
    if (notes) {
      source.notes = [];
      const first = notes[1].trim();
      if (first.length > 0) {
        source.notes.push(first.startsWith("\\") ? first.slice(1) : first);
      }
      i++;
      continue;
    }
    const fence = FENCE_RE.exec(line);
    if (fence) {
      const close = new RegExp(`^ {0,3}\\${fence[1][0]}{${fence[1].length},}[ \\t]*$`);
      i++;
      while (i < lines.length && !close.test(lines[i])) {
        i++;
      }
      i++;
      seenBlock = true;
      drop("codeFence", "fenced code block: code blocks are outside the slide grammar");
      continue;
    }
    if (HTML_BLOCK_RE.test(line)) {
      while (i < lines.length && lines[i].trim() !== "") {
        i++;
      }
      seenBlock = true;
      drop("rawHtml", "raw HTML block: HTML is outside the slide grammar");
      continue;
    }
    if (REFERENCE_DEF_RE.test(line)) {
      i++;
      seenBlock = true;
      drop("rawHtml", "link reference definition: reference definitions are outside the slide grammar");
      continue;
    }
    const heading = HEADING_RE.exec(line) ?? EMPTY_HEADING_RE.exec(line);
    if (heading) {
      const level = heading[1].length;
      if (level <= 2 && !seenBlock) {
        source.title = slideRunsOf((heading[2] ?? "").trim(), counts);
      } else if (level > 2) {
        drop("heading", `level-${level} heading: only # and ## (a slide's title) are in the slide grammar`);
      } else {
        drop("heading", "heading after a slide's first block: only a heading that opens a slide becomes its title");
      }
      seenBlock = true;
      i++;
      continue;
    }
    if (THEMATIC_BREAK_RE.test(line)) {
      i++;
      seenBlock = true;
      drop("rawHtml", "thematic break: only a `---` line is in the slide grammar, where it separates slides");
      continue;
    }
    if (BLOCKQUOTE_RE.test(line)) {
      while (i < lines.length && BLOCKQUOTE_RE.test(lines[i])) {
        i++;
      }
      seenBlock = true;
      drop("blockQuote", "block quote: quotes are outside the slide grammar");
      continue;
    }
    if (line.includes("|") && i + 1 < lines.length && isTableDelimiterRow(lines[i + 1])) {
      i += 2;
      while (i < lines.length && lines[i].trim() !== "" && lines[i].includes("|")) {
        i++;
      }
      seenBlock = true;
      drop("table", "table: tables are outside the slide grammar");
      continue;
    }
    const item = LIST_ITEM_RE.exec(line);
    if (item) {
      const [, indent, marker, text] = item;
      const runs = slideRunsOf(text.trim(), counts);
      if (runs.length > 0) {
        const paragraph: SlideParagraph = { runs, bullet: /\d/.test(marker) ? "number" : "disc" };
        const level = Math.min(MAX_INDENT_LEVEL, Math.floor(indent.length / 2));
        if (level > 0) {
          paragraph.level = level;
        }
        source.body.push(paragraph);
      }
      seenList = true;
      seenBlock = true;
      i++;
      continue;
    }
    const image = STANDALONE_IMAGE_RE.exec(line);
    if (image) {
      readImage(image[1], image[2].trim(), source, drop);
      seenBlock = true;
      i++;
      continue;
    }
    // A paragraph: lines up to the next block start, split into visual lines
    // at hard breaks (a slide paragraph is one rendered line).
    const visual: string[][] = [[]];
    while (i < lines.length && (visual.flat().length === 0 || !startsBlock(lines[i]))) {
      const raw = lines[i];
      visual[visual.length - 1].push(raw.replace(/\\$/, (match) => (hardBreak(raw) ? "" : match)).trim());
      if (hardBreak(raw)) {
        visual.push([]);
      }
      i++;
    }
    for (const group of visual) {
      addParagraph(slideRunsOf(group.join(" ").trim(), counts));
    }
    seenBlock = true;
  }
  reportRunDrops(dropped, slideIndex, counts, (field) => `a slide run has no ${RUN_FIELD_NAMES[field] ?? field}`);
  return source;
}

/** A standalone image: a `data:` URI of a supported type is decoded; a
 * remote `src` is never fetched (SPEC-088's self-containment rule). */
function readImage(alt: string, src: string, source: SlideSource, drop: (kind: DeckDropKind, reason: string) => void): void {
  const uri = DATA_URI_RE.exec(src);
  if (!uri) {
    drop("remoteImage", `image "${src}": only data: URIs are embedded — nothing is fetched`);
    return;
  }
  const [, contentType, params, payload] = uri;
  let data: string | Uint8Array;
  try {
    data = params.split(";").includes("base64")
      ? payload.replace(/\s+/g, "")
      : new TextEncoder().encode(decodeURIComponent(payload));
    source.images.push(decodeImage({ data, contentType: contentType || "application/octet-stream" }, "image"));
  } catch (error) {
    const detail = error instanceof DocumentGenerationError ? error.message.replace(/^image(\.\w+)?: /, "") : "unreadable payload";
    drop("media", `image (${contentType || "no type"}): ${detail} — supported images are PNG, JPEG and GIF`);
    return;
  }
  if (alt.length > 0) {
    drop("media", "image alt text not kept: a slide image frame has no alt-text field");
  }
}

/** Assembles a parsed slide from the editor's own placeholder geometry. */
function buildSlide(ctx: DeckContext, source: SlideSource, slideIndex: number): void {
  const w = ctx.deck.slideWidthCm;
  const h = ctx.deck.slideHeightCm;
  const frames: SlideFrame[] = [];
  const [, bodyFrame] = titleBodyFrames(w, h);
  bodyFrame.paragraphs = source.body;
  if (source.title !== undefined) {
    const titleParagraphs = source.title.length > 0 ? [{ runs: source.title }] : [];
    if (slideIndex === 0 && (source.subtitle.length > 0 || source.body.length === 0)) {
      // The first slide is a title slide (addTitleSlide's pair) — with a body
      // placeholder too when it also carries list rows.
      const [title, subtitle] = titleSlideFrames(w, h);
      title.paragraphs = titleParagraphs;
      subtitle.paragraphs = source.subtitle;
      frames.push(title, subtitle);
      if (source.body.length > 0) {
        frames.push(bodyFrame);
      }
    } else {
      // A title + body slide (addBulletSlide's pair); an image-only slide
      // keeps just the title, as addImageSlide does.
      const [title] = titleBodyFrames(w, h);
      title.paragraphs = titleParagraphs;
      frames.push(title);
      if (source.body.length > 0 || source.images.length === 0) {
        frames.push(bodyFrame);
      }
    }
  } else if (source.body.length > 0) {
    frames.push(bodyFrame);
  }
  const slide: Slide = { id: crypto.randomUUID(), frames };
  ctx.deck.slides.push(slide);
  for (const image of source.images) {
    landImage(ctx, slideIndex, image, imageInBodyRect(ctx.deck, image.natural));
  }
  setSlideNotes(slide, source.notes ?? []);
}

/** Parses slide Markdown into a deck, reporting everything it dropped. */
export function parseSlideMarkdownWithReport(text: string, options: SlideMarkdownOptions = {}): SlideMarkdownParse {
  const source = expectString(text, "text");
  const opts = expectPlainObject(options, "options");
  rejectUnknownKeys(opts, ["title", "format", "size"], "options");
  const ctx = createDeckContext(opts, "options");
  const lines = sanitizeXmlText(source).replace(/\r\n?/g, "\n").split("\n");
  const { body, frontMatter } = splitFrontMatter(lines);
  if (frontMatter) {
    applyFrontMatter(ctx, frontMatter);
  }
  splitSlides(body).forEach((chunk, index) => buildSlide(ctx, readSlide(chunk, index, ctx.dropped), index));
  return { deck: structuredClone(ctx.deck), dropped: ctx.dropped };
}

/** Parses slide Markdown into a deck (see `parseSlideMarkdownWithReport` for
 * what was dropped on the way). */
export function parseSlideMarkdown(text: string, options: SlideMarkdownOptions = {}): PresentationModel {
  return parseSlideMarkdownWithReport(text, options).deck;
}

// --- writing ------------------------------------------------------------------

/** md/writer.ts's inline escape set: the characters this grammar's own
 * inline marks use. */
function escapeInline(text: string): string {
  return text.replace(/[\\`*_~[\]]/g, (match) => `\\${match}`);
}

function escapeLinkTarget(href: string): string {
  return href.replace(/[()]/g, (match) => `\\${match}`);
}

/** md/writer.ts's line-start escape, plus the two starts only the slide
 * grammar reads: `Notes:` and a raw-HTML `<`. */
function escapeLineStart(line: string): string {
  const startsLikeBlock =
    /^(#{1,6}[ \t]|>|<|[-*+][ \t]|\d+[.)][ \t]|```|~~~|Notes:)/.test(line) || /^[-*_]{3,}$/.test(line.replace(/[ \t]/g, ""));
  return startsLikeBlock ? `\\${line}` : line;
}

/** One run's spelling: escaped, then italic/bold, then a link outermost —
 * the nesting md/writer.ts uses and `parseInlineRuns` recovers. */
function renderRun(run: SlideRun): string {
  let text = escapeInline(run.text);
  if (run.italic) {
    text = `_${text}_`;
  }
  if (run.bold) {
    text = `**${text}**`;
  }
  if (run.linkHref) {
    text = `[${text}](${escapeLinkTarget(run.linkHref)})`;
  }
  return text;
}

const CARRIED_RUN_KEYS = new Set(["text", "bold", "italic", "linkHref"]);

function renderRuns(runs: SlideRun[], counts: Map<string, number>): string {
  const carried = runs.filter((run) => run.text.length > 0);
  for (const run of carried) {
    for (const key of Object.keys(run)) {
      if (!CARRIED_RUN_KEYS.has(key) && run[key as keyof SlideRun] !== undefined) {
        tally(counts, key);
      }
    }
  }
  return carried.map(renderRun).join("");
}

/** Within the ±0.01 cm a write → read round trip allows. */
function near(a: number, b: number): boolean {
  return Math.abs(a - b) < 0.01;
}

function sameRect(frame: SlideFrame, rect: { xCm: number; yCm: number; widthCm: number; heightCm: number }): boolean {
  return (
    near(frame.xCm, rect.xCm) &&
    near(frame.yCm, rect.yCm) &&
    near(frame.widthCm, rect.widthCm) &&
    near(frame.heightCm, rect.heightCm) &&
    !frame.rotationDeg
  );
}

/** The frames a slide's Markdown re-parses into, by role — what a text
 * frame's geometry must match for Markdown to carry it. */
function canonicalPlaceholders(model: PresentationModel, slide: Slide, slideIndex: number): Map<string, SlideFrame> {
  const w = model.slideWidthCm;
  const h = model.slideHeightCm;
  const byRole = new Map<string, SlideFrame>();
  const hasSubtitle = slide.frames.some((frame) => frame.kind === "text" && frame.role === "subtitle");
  const [title, subtitle] = slideIndex === 0 && hasSubtitle ? titleSlideFrames(w, h) : titleBodyFrames(w, h);
  byRole.set("title", title);
  if (subtitle.role === "subtitle") {
    byRole.set("subtitle", subtitle);
  }
  byRole.set("body", titleBodyFrames(w, h)[1]);
  return byRole;
}

interface SlideWriter {
  blocks: string[];
  counts: Map<string, number>;
  drop: (kind: DeckDropKind, reason: string) => void;
}

/** A text frame's paragraphs as Markdown blocks: list rows (indented two
 * spaces a level, numbered rows counted per level) and plain paragraphs. */
function writeParagraphs(paragraphs: SlideParagraph[], out: SlideWriter): void {
  let list: string[] = [];
  const counters: number[] = [];
  const flushList = (): void => {
    if (list.length > 0) {
      out.blocks.push(list.join("\n"));
      list = [];
      counters.length = 0;
    }
  };
  for (const paragraph of paragraphs) {
    const text = renderRuns(paragraph.runs, out.counts);
    if (text.length === 0) {
      continue;
    }
    if (paragraph.bullet) {
      const level = Math.min(MAX_INDENT_LEVEL, Math.max(0, paragraph.level ?? 0));
      counters.length = level + 1;
      counters[level] = (counters[level] ?? 0) + 1;
      const marker = paragraph.bullet === "number" ? `${counters[level]}.` : "-";
      list.push(`${"  ".repeat(level)}${marker} ${escapeLineStart(text)}`);
    } else {
      flushList();
      out.blocks.push(escapeLineStart(text));
    }
  }
  flushList();
}

const FRAME_DROPS: Partial<Record<SlideFrame["kind"], DeckDropKind>> = {
  shape: "shape",
  group: "group",
  foreign: "foreignFrame",
  table: "table",
  chart: "chart",
  media: "media",
};

/** A frame's text a Markdown body can still carry (a shape's, a group's
 * children's), its own construct reported. */
function nestedParagraphs(frame: SlideFrame): SlideParagraph[] {
  if (frame.kind === "text" || frame.kind === "shape") {
    return frame.paragraphs ?? [];
  }
  if (frame.kind === "group") {
    return frame.children.flatMap(nestedParagraphs);
  }
  return [];
}

function writeSlide(model: PresentationModel, slide: Slide, slideIndex: number, dropped: DeckDropRecord[]): string {
  const location = slideLocation(slideIndex);
  const out: SlideWriter = {
    blocks: [],
    counts: new Map(),
    drop: (kind, reason) => dropped.push({ kind, location, reason }),
  };
  const canonical = canonicalPlaceholders(model, slide, slideIndex);
  const titleFrame = slide.frames.find((frame): frame is TextFrame => frame.kind === "text" && frame.role === "title");
  if (titleFrame) {
    // The outline's title says whether there is any title text; the frame's
    // runs give its formatting. A closing `#` run is escaped so the heading
    // pattern cannot read it as a closing sequence.
    const title =
      slideOutline(slide).title === undefined
        ? ""
        : titleFrame.paragraphs
            .map((paragraph) => renderRuns(paragraph.runs, out.counts))
            .filter((text) => text.length > 0)
            .join(" ")
            .replace(/([ \t])(#+)$/, "$1\\$2");
    out.blocks.push(`${slideIndex === 0 ? "#" : "##"}${title.length > 0 ? ` ${title}` : ""}`);
  }
  const images: string[] = [];
  for (const frame of slide.frames) {
    const kindDrop = FRAME_DROPS[frame.kind];
    if (kindDrop) {
      out.drop(kindDrop, `${frame.kind === "foreign" ? "foreign object" : frame.kind} frame: Markdown has no spelling for it`);
    }
    if ((frame.kind === "text" || frame.kind === "shape") && frame.fill) {
      out.drop("fill", `${frame.kind} frame fill: Markdown has no spelling for it`);
    }
    if ((frame.kind === "text" || frame.kind === "shape" || frame.kind === "image") && frame.stroke) {
      out.drop("stroke", `${frame.kind} frame outline: Markdown has no spelling for it`);
    }
    if (frame.kind === "image") {
      const part = model.media?.find((media) => media.mediaId === frame.mediaId);
      let natural: DecodedImage["natural"] | undefined;
      try {
        natural = part ? decodeImage({ data: part.data, contentType: part.contentType }, "image").natural : undefined;
      } catch {
        natural = undefined;
      }
      if (!part || !natural) {
        out.drop("media", "image frame: its picture is missing or not PNG, JPEG or GIF");
        continue;
      }
      if (!sameRect(frame, imageInBodyRect(model, natural))) {
        out.drop("slideGeometry", "image position and size: Markdown places an image in the body area");
      }
      images.push(`![](${mediaDataUri(part)})`);
      continue;
    }
    if (frame === titleFrame) {
      if (!sameRect(frame, canonical.get("title")!)) {
        out.drop("slideGeometry", "title position and size: Markdown places the title at the layout's geometry");
      }
      continue;
    }
    if (frame.kind === "text") {
      const expected = frame.role ? canonical.get(frame.role) : undefined;
      if (frame.paragraphs.length > 0 && (!expected || !sameRect(frame, expected))) {
        out.drop(
          "slideGeometry",
          `${frame.role ? `${frame.role} placeholder` : "text box"} position and size: Markdown places its text in the body`,
        );
      }
    }
    writeParagraphs(nestedParagraphs(frame), out);
  }
  out.blocks.push(...images);
  if (slide.transition) {
    out.drop("transition", "slide transition: Markdown has no spelling for it");
  }
  for (const animation of slide.animations ?? []) {
    out.drop("animation", `${animation.effect} animation: Markdown has no spelling for it`);
  }
  if (slide.background) {
    out.drop("background", "slide background: Markdown has no spelling for it");
  }
  reportRunDrops(dropped, slideIndex, out.counts, () => "Markdown has no spelling for it");
  if (slide.notes && slide.notes.length > 0) {
    // Raw lines to the end of the slide; a line that would read as a slide
    // separator (or starts with the escape itself) takes a leading `\`.
    const lines = slide.notes.map((line) => (line.startsWith("\\") || SEPARATOR_RE.test(line) ? `\\${line}` : line));
    out.blocks.push((lines[0].length > 0 ? [`Notes: ${lines[0]}`, ...lines.slice(1)] : ["Notes:", ...lines]).join("\n"));
  }
  return out.blocks.join("\n\n");
}

/** A YAML scalar for the front matter: plain when unambiguous, else
 * double-quoted with JSON escapes. */
function yamlValue(value: string): string {
  return /^[A-Za-z0-9(][^:#"'\n]*$/.test(value) && value === value.trim() ? value : JSON.stringify(value);
}

/** The deck as slide Markdown in canonical spelling, plus one record per
 * construct the text cannot carry. */
export function writeSlideMarkdown(model: PresentationModel): { text: string; dropped: DeckDropRecord[] } {
  const deck = expectPlainObject(model, "deck") as unknown as PresentationModel;
  const slides = expectArray(deck.slides, "deck.slides") as Slide[];
  const dropped: DeckDropRecord[] = [];
  const frontMatter = ["---", `title: ${yamlValue(deck.title ?? "")}`, `format: ${deck.format === "odp" ? "odp" : "pptx"}`];
  const size = DECK_SIZE_NAMES.find(
    (name) =>
      Math.abs(SLIDE_SIZES[name].widthCm - deck.slideWidthCm) < 0.001 &&
      Math.abs(SLIDE_SIZES[name].heightCm - deck.slideHeightCm) < 0.001,
  );
  if (size) {
    frontMatter.push(`size: ${size}`);
  } else {
    dropped.push({
      kind: "slideGeometry",
      location: "deck",
      reason: `slide size ${deck.slideWidthCm} × ${deck.slideHeightCm} cm: front matter spells only 16:9 and 4:3`,
    });
  }
  frontMatter.push("---");
  const bodies = slides.map((slide, index) => {
    const body = writeSlide(deck, slide, index, dropped);
    // A first or last slide with nothing on it would be an empty chunk the
    // parser skips; an empty notes line keeps it a slide.
    return body.length === 0 && (index === 0 || index === slides.length - 1) ? "Notes:" : body;
  });
  const text = [frontMatter.join("\n"), ...(bodies.length > 0 ? [bodies.join("\n\n---\n\n")] : [])].join("\n\n");
  return { text: `${text}\n`, dropped };
}
