// The deck designer's sources (SPEC-115), built here in the app layer — where
// the two products may meet — as lib/design's product-neutral OutlineSource:
// from a text document's model (SPEC-037's heading outline and the paragraphs
// between the headings), from a deck, or from pasted text. Reads its source
// and never writes to it; useDeckDesigner.spec asserts that against a frozen
// model.
import type { Block, DocumentModel, Paragraph } from "../../lib/documentModel";
import { buildOutline, type OutlineNode } from "../../lib/outline";
import { blockText } from "../../lib/textEditing";
import type { PresentationModel, TextFrame } from "../../lib/presentation/model";
import { currentLayoutKind, slideTitleText } from "../../lib/design/applyDesign";
import { MAX_OUTLINE_BULLET_WORDS, type OutlineBlock, type OutlineSource } from "../../lib/design/outlineDesigner";

function isParagraph(block: Block): block is Paragraph {
  return block.type === "paragraph";
}

function flatten(nodes: OutlineNode[], out: OutlineNode[] = []): OutlineNode[] {
  for (const node of nodes) {
    out.push(node);
    flatten(node.children, out);
  }
  return out;
}

/** A text document as an OutlineSource: its title, and its body in order —
 * each SPEC-037 outline heading as a heading, every other paragraph (list
 * items with their depth) as a paragraph. Tables, images and other blocks
 * carry no outline text and are skipped. */
export function outlineFromDocument(model: DocumentModel): OutlineSource {
  const headings = new Map(flatten(buildOutline(model)).map((node) => [`${node.section}:${node.block}`, node]));
  const blocks: OutlineBlock[] = [];
  model.sections.forEach((section, s) => {
    section.blocks.forEach((block, b) => {
      if (!isParagraph(block)) {
        return;
      }
      const heading = headings.get(`${s}:${b}`);
      if (heading) {
        blocks.push({ kind: "heading", level: heading.level, text: heading.text });
        return;
      }
      const text = blockText(block);
      if (text.trim() === "") {
        return;
      }
      blocks.push(block.list ? { kind: "paragraph", text, listLevel: Math.max(0, block.list.level - 1) } : { kind: "paragraph", text });
    });
  });
  return { title: model.title, blocks };
}

/** A deck as an OutlineSource: its title and theme, and per slide its title
 * (a section header slide's as a Heading 1, every other as a Heading 2)
 * followed by the paragraphs of its other text frames. Notes are not read. */
export function outlineFromDeck(model: PresentationModel): OutlineSource {
  const blocks: OutlineBlock[] = [];
  model.slides.forEach((slide, index) => {
    const level = currentLayoutKind(model, slide) === "sectionHeader" ? 1 : 2;
    blocks.push({ kind: "heading", level, text: slideTitleText(slide) || `Slide ${index + 1}` });
    for (const frame of slide.frames) {
      if (frame.kind !== "text" || frame.role === "title") {
        continue;
      }
      for (const paragraph of (frame as TextFrame).paragraphs) {
        const text = paragraph.runs.map((run) => run.text).join("");
        if (text.trim() !== "") {
          blocks.push({ kind: "paragraph", text, listLevel: paragraph.level ?? 0 });
        }
      }
    }
  });
  return { title: model.title, ...(model.theme ? { theme: model.theme.name } : {}), blocks };
}

const LIST_ITEM = /^(\s*)(?:[-*•]|\d+[.)])\s+(.*)$/;
const MARKED_HEADING = /^(#{1,6})\s+(.*)$/;

function words(text: string): number {
  return text.trim().split(/\s+/).filter((word) => word !== "").length;
}

/** Pasted text as an OutlineSource, read as blank-line-separated blocks:
 * a short single-line first block is the title; a line marked `#`…`######`
 * is a heading of that level; a block's first line is otherwise its heading
 * (level 2) and its other lines are paragraphs, `-`/`*`/`•`/`1.` items with
 * their indentation as list depth; a single line longer than twelve words is
 * a paragraph, not a heading. */
export function outlineFromText(text: string): OutlineSource {
  const chunks = text
    .replace(/\r\n?/g, "\n")
    .split(/\n[ \t]*\n/)
    .map((chunk) => chunk.split("\n").filter((line) => line.trim() !== ""))
    .filter((lines) => lines.length > 0);
  let title = "";
  const first = chunks[0];
  if (first && first.length === 1 && !LIST_ITEM.test(first[0]) && words(first[0]) <= MAX_OUTLINE_BULLET_WORDS) {
    title = first[0].replace(MARKED_HEADING, "$2").trim();
    chunks.shift();
  }
  const blocks: OutlineBlock[] = [];
  for (const lines of chunks) {
    lines.forEach((line, i) => {
      const marked = MARKED_HEADING.exec(line.trim());
      const item = LIST_ITEM.exec(line);
      if (marked) {
        blocks.push({ kind: "heading", level: marked[1].length, text: marked[2].trim() });
      } else if (item) {
        blocks.push({ kind: "paragraph", text: item[2].trim(), listLevel: Math.floor(item[1].replace(/\t/g, "  ").length / 2) });
      } else if (i === 0 && (lines.length > 1 || words(line) <= MAX_OUTLINE_BULLET_WORDS)) {
        blocks.push({ kind: "heading", level: 2, text: line.trim() });
      } else {
        blocks.push({ kind: "paragraph", text: line.trim() });
      }
    });
  }
  return { title, blocks };
}

/** An OutlineSource as the plain text a design request carries: headings
 * marked with `#`, list items with `-` at their depth. */
export function outlineText(source: OutlineSource): string {
  return source.blocks
    .map((block) =>
      block.kind === "heading"
        ? `${"#".repeat(Math.min(Math.max(block.level, 1), 6))} ${block.text}`
        : block.listLevel !== undefined
          ? `${"  ".repeat(block.listLevel)}- ${block.text}`
          : block.text,
    )
    .join("\n");
}
