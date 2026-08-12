// Hand-written Markdown serialiser (SPEC-050): the inverse of reader.ts,
// emitting the canonical spelling of every construct (`_italic_`, ATX
// headings, `-` bullets, `1.` numbers, …) regardless of what the source
// syntax was — a `.md` -> `.md` round trip is asserted with
// `equalStructural` on the *model*, never `equalBytes` on the file (see
// specs/050-markdown-open-edit-and-convert.md's "Equality"), which is what
// licenses always normalising rather than trying to echo input syntax.
import {
  DEFAULT_TITLE,
  HORIZONTAL_RULE_STYLE_ID,
  isBreak,
  isImage,
  isTable,
  isToc,
  type Block,
  type DocumentModel,
  type Paragraph,
  type Run,
  type Table,
  type TableCell,
} from "../documentModel";
import { collectDropRecords, type WriteResult } from "../opaque";
import { CODE_CHAR_STYLE_ID, styleById, styleChainIncludes } from "../styles";

/** Markdown's registered content type — the download's Blob type and the
 * `openDocumentFromUrl.ts` fetch-path's accepted content type both use it. */
export const MD_MIME_TYPE = "text/markdown";

/** The download file name for a Markdown save — mirrors
 * `docx/writer.ts`'s `docxFileName`/`odf/writer.ts`'s `odtFileName`. */
export function mdFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_TITLE}.md`;
}

/** Escapes the Markdown-significant characters this writer's own inline
 * marks use, so a run's literal text can never be misread as one on
 * re-parse: `\`, `` ` ``, `*`, `_`, `~`, `[`, `]`. */
function escapeInline(text: string): string {
  return text.replace(/[\\`*_~[\]]/g, (match) => `\\${match}`);
}

/** Escapes a hyperlink target's parens, the `(...)` delimiter's own
 * syntax character. */
function escapeLinkTarget(href: string): string {
  return href.replace(/[()]/g, (match) => `\\${match}`);
}

/** Escapes `|`, GFM's table-cell delimiter — applied after `escapeInline`,
 * table cells only. */
function escapePipes(text: string): string {
  return text.replace(/\|/g, "\\|");
}

/** Whether `line`'s own leading characters would be misread as a different
 * block's start on re-parse (a heading `#`, a blockquote `>`, a list
 * marker, a fence, a thematic break) — escaped with a leading backslash so
 * literal text starting the same way round-trips as plain text. */
function escapeLineStart(line: string): string {
  const startsLikeBlock =
    /^(#{1,6}[ \t]|>|[-*+][ \t]|\d+[.)][ \t]|```|~~~)/.test(line) ||
    /^[-*_]{3,}$/.test(line.replace(/[ \t]/g, ""));
  return startsLikeBlock ? `\\${line}` : line;
}

/** One run's Markdown spelling: an inline code span (the Code character
 * style, SPEC-050) picks its backtick fence length to beat any backtick run
 * already inside the text; everything else is escaped, then wrapped
 * strikethrough-out/italic/bold-in, then a hyperlink outermost — the exact
 * nesting order `md/reader.ts`'s `parseInline` recovers regardless of,
 * since it recognises open/close markers by toggling rather than by
 * position. A run's own embedded `\n` (SPEC-050's "Two-space or backslash
 * line end", carried as a literal character — see Run's own doc comment)
 * becomes a trailing backslash before the newline, the canonical spelling
 * `reader.ts`'s `joinLines` recognises on the way back in. */
function renderRun(run: Run): string {
  let text: string;
  if (run.styleId === CODE_CHAR_STYLE_ID) {
    const backtickRuns = run.text.match(/`+/g) ?? [];
    const fenceLen = backtickRuns.reduce((max, seq) => Math.max(max, seq.length), 0) + 1;
    const fence = "`".repeat(fenceLen);
    const pad = run.text.startsWith("`") || run.text.endsWith("`") || run.text.length === 0 ? " " : "";
    text = `${fence}${pad}${run.text}${pad}${fence}`;
  } else {
    text = escapeInline(run.text);
  }
  text = text.replace(/\n/g, "\\\n");
  if (run.formatting.strikethrough) {
    text = `~~${text}~~`;
  }
  if (run.formatting.italic) {
    text = `_${text}_`;
  }
  if (run.formatting.bold) {
    text = `**${text}**`;
  }
  if (run.formatting.linkHref) {
    text = `[${text}](${escapeLinkTarget(run.formatting.linkHref)})`;
  }
  return text;
}

/** A paragraph's runs as one Markdown-spelled line (may itself contain
 * embedded newlines for hard line breaks — see `renderRun`). Field/note
 * runs (a live page-number field, a footnote/endnote reference) are
 * dropped: Markdown has nowhere to keep a computed value (SPEC-050's "What
 * markdown cannot carry"), and `lib/opaque.ts`'s `withoutDropped` masks a
 * source model's field/note runs away entirely to match. */
function renderRuns(runs: Run[]): string {
  return runs
    .filter((run) => !run.field && !run.note)
    .map(renderRun)
    .join("");
}

// --- list rendering -----------------------------------------------------

const LIST_INDENT = "  ";

/** Renders one list-item paragraph, tracking a per-level ordinal counter
 * across the call sequence the same way `lib/lists.ts`'s `markersForBlocks`
 * derives displayed numbers: adjacency, not a stored identity. Bullets
 * always spell `-`; the level cycles this app's own render normally applies
 * (disc/circle/square, decimal/lowerAlpha/lowerRoman) have no GFM
 * equivalent, so nesting is expressed by indent alone. */
function renderListItem(paragraph: Paragraph, counters: Array<number | undefined>): string {
  const list = paragraph.list!;
  const level = Math.min(9, Math.max(1, Math.round(list.level)));
  for (let deeper = level + 1; deeper < counters.length; deeper++) {
    counters[deeper] = undefined;
  }
  const indent = LIST_INDENT.repeat(level - 1);
  let marker: string;
  if (list.type === "bullet") {
    marker = "-";
  } else {
    const ordinal = (counters[level] ?? (list.startAt ?? 1) - 1) + 1;
    counters[level] = ordinal;
    marker = `${ordinal}.`;
  }
  return `${indent}${marker} ${escapeLineStart(renderRuns(paragraph.runs))}`;
}

// --- fenced code blocks (grouped consecutive paragraphs) ----------------

function isCodeParagraph(model: DocumentModel, block: Block): block is Paragraph {
  return block.type === "paragraph" && styleChainIncludes(model, block.styleId, "Code");
}

function renderCodeGroup(model: DocumentModel, styleId: string, group: Paragraph[]): string {
  const language = styleById(model, styleId)?.language ?? "";
  const lines = group.map((paragraph) => paragraph.runs.map((run) => run.text).join(""));
  return `\`\`\`${language}\n${lines.join("\n")}\n\`\`\``;
}

// --- paragraphs, tables, images, TOC -------------------------------------

function renderQuote(paragraph: Paragraph): string {
  const rendered = renderRuns(paragraph.runs);
  return rendered
    .split("\n")
    .map((line) => `> ${escapeLineStart(line)}`)
    .join("\n");
}

function renderTable(table: Table): string {
  if (table.rows.length === 0) {
    return "";
  }
  const cellText = (cell: TableCell): string => escapePipes(renderRuns(cell.blocks[0]?.runs ?? []));
  const [header, ...body] = table.rows;
  const headerLine = `| ${header.cells.map(cellText).join(" | ")} |`;
  const delimiterLine = `| ${header.cells.map(() => "---").join(" | ")} |`;
  const bodyLines = body.map((row) => `| ${row.cells.map(cellText).join(" | ")} |`);
  return [headerLine, delimiterLine, ...bodyLines].join("\n");
}

/** A block's Markdown text, or null for one that carries no Markdown
 * representation at all (a manual page/column break — SPEC-050's "page
 * setup and section breaks" bucket). Fenced code blocks are handled by the
 * caller (`writeMd`'s main loop), which must group consecutive paragraphs
 * before this is reached. */
function renderBlock(
  model: DocumentModel,
  block: Block,
  listCounters: Array<number | undefined>,
): string | null {
  if (isBreak(block)) {
    return null;
  }
  if (isToc(block)) {
    listCounters.length = 0;
    // SPEC-050: "emitted as its rendered text, not as a field" — one plain
    // paragraph per entry, matching what `withoutDropped`'s "tableOfContents"
    // mask replaces the block with, so a docx -> md -> docx round trip's
    // masked source and reopened model agree.
    return block.entries.map((entry) => escapeLineStart(escapeInline(entry.text))).join("\n\n");
  }
  if (isImage(block)) {
    listCounters.length = 0;
    const media = (model.media ?? []).find((part) => part.mediaId === block.mediaId);
    if (!media) {
      return null;
    }
    const alt = block.altText ? escapeInline(block.altText) : "";
    return `![${alt}](data:${media.contentType};base64,${media.data})`;
  }
  if (isTable(block)) {
    listCounters.length = 0;
    return renderTable(block);
  }
  const paragraph = block;
  if (paragraph.styleId === HORIZONTAL_RULE_STYLE_ID) {
    listCounters.length = 0;
    return "---";
  }
  if (paragraph.headingLevel) {
    listCounters.length = 0;
    const level = Math.min(6, Math.max(1, paragraph.headingLevel));
    return `${"#".repeat(level)} ${escapeLineStart(renderRuns(paragraph.runs))}`;
  }
  if (styleChainIncludes(model, paragraph.styleId, "Quote")) {
    listCounters.length = 0;
    return renderQuote(paragraph);
  }
  if (paragraph.list) {
    return renderListItem(paragraph, listCounters);
  }
  listCounters.length = 0;
  return escapeLineStart(renderRuns(paragraph.runs));
}

export interface WriteMdOptions {}

/** Serialises a DocumentModel to Markdown (SPEC-050): the package-free
 * counterpart of `writeDocx`/`writeOdt` — a single UTF-8 text file, no ZIP,
 * no preserved parts. `dropped` (SPEC-008's `WriteResult`) is always
 * `collectDropRecords(model, "md")`, computed regardless of the model's own
 * source format — Markdown is not full-fidelity for content this editor can
 * itself create, so even a `.md`-sourced document being re-saved can have
 * something to report (a header added since it was opened, say). Every
 * section's blocks are concatenated into one continuous flow (Markdown has
 * no section concept), each section's own opaque blocks re-emitted verbatim
 * at their anchored position (SPEC-008), and front matter — if the model
 * carries any, only ever true for a `.md`-opened document — is emitted
 * first, unparsed (AC-4). */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
export async function writeMd(model: DocumentModel, _options: WriteMdOptions = {}): Promise<WriteResult> {
  const dropped = collectDropRecords(model, "md");
  const parts: string[] = [];
  if (model.frontMatter) {
    parts.push(model.frontMatter);
  }

  const listCounters: Array<number | undefined> = [];
  for (const section of model.sections) {
    const opaqueBlocks = [...(section.opaqueBlocks ?? [])].sort((a, b) => a.blockIndex - b.blockIndex);
    const emitOpaqueBefore = (index: number): void => {
      for (const opaque of opaqueBlocks) {
        if (Math.min(opaque.blockIndex, section.blocks.length) === index) {
          parts.push(opaque.xml);
        }
      }
    };

    let index = 0;
    while (index < section.blocks.length) {
      emitOpaqueBefore(index);
      const block = section.blocks[index];
      if (isCodeParagraph(model, block)) {
        const styleId = block.styleId as string;
        const group: Paragraph[] = [];
        while (index < section.blocks.length) {
          const candidate = section.blocks[index];
          if (!isCodeParagraph(model, candidate) || candidate.styleId !== styleId) {
            break;
          }
          group.push(candidate);
          index++;
        }
        parts.push(renderCodeGroup(model, styleId, group));
        listCounters.length = 0;
        continue;
      }
      const rendered = renderBlock(model, block, listCounters);
      if (rendered !== null) {
        parts.push(rendered);
      }
      index++;
    }
    emitOpaqueBefore(section.blocks.length);
  }

  const text = `${parts.filter((part) => part.length > 0).join("\n\n")}\n`;
  return { bytes: new TextEncoder().encode(text), dropped };
}
