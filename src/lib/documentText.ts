// A word-processing document's text for library search (SPEC-114). Walks a
// DocumentModel in document order — per section its headers, its body
// blocks and its footers, then the notes, then the comments — taking every
// paragraph's DISPLAYED text (a field shows its value, exactly as the find
// bar reads it): top-level paragraphs, table cells row by row, the text
// frames of shapes. Images, charts, connectors, breaks and generated
// tables of contents/figures contribute nothing. One line per paragraph.
//
// Each line is recorded as a span naming where it came from, so an offset
// into the text maps back to the caret position SPEC-017's matchSelection
// works in (documentTextCaret) — how a search result opens AT its match.
// Pure and framework-free: the model layer, plus core.
import {
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  isToc,
  isTof,
  type Block,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
} from "./documentModel";
import { cachedFieldText } from "./fields";
import type { CaretPosition } from "./textEditing";
import type { ExtractedText } from "./core/extractedText";

/** Which part of the document a line of text came from. */
export type DocumentTextPart = "header" | "body" | "footer" | "note" | "comment";

/** One line of the extracted text and its source. */
export interface DocumentTextSpan {
  /** [start, end) in the extracted text. */
  start: number;
  end: number;
  part: DocumentTextPart;
  /** The top-level block holding the line — for a header, footer or body
   * line; a table cell's or shape's line names its table or shape. */
  section?: number;
  block?: number;
  /** The line IS the top-level paragraph `block` (offsets map one to one,
   * fields aside); false for a line inside a table or shape. */
  paragraph?: boolean;
}

export interface DocumentText extends ExtractedText {
  spans: DocumentTextSpan[];
}

/** A paragraph's displayed text: each field run shows its value. */
function displayText(paragraph: Paragraph): string {
  let text = "";
  for (const run of paragraph.runs) {
    text += run.field ? cachedFieldText(run.field as FieldSpec) : run.text;
  }
  return text;
}

/** Every paragraph a block holds, in reading order. */
function paragraphsOf(block: Block): Paragraph[] {
  if (isTable(block)) {
    return block.rows.flatMap((row) => row.cells.filter((cell) => !cell.covered).flatMap((cell) => cell.blocks.flatMap(paragraphsOf)));
  }
  if (isShape(block)) {
    return (block.text?.blocks ?? []).flatMap(paragraphsOf);
  }
  if (isImage(block) || isChart(block) || isConnector(block) || isToc(block) || isTof(block) || block.type !== "paragraph") {
    return [];
  }
  return [block];
}

/** The document's text, in document order, with the span of every line. A
 * document model is read whole: its coverage is always `full`. */
export function documentText(model: DocumentModel): DocumentText {
  const lines: string[] = [];
  const spans: DocumentTextSpan[] = [];
  let length = 0;
  const add = (text: string, span: Omit<DocumentTextSpan, "start" | "end">): void => {
    const start = lines.length === 0 ? 0 : length + 1;
    lines.push(text);
    length = start + text.length;
    spans.push({ start, end: length, ...span });
  };
  const addBlocks = (blocks: Block[], part: DocumentTextPart, section?: number): void => {
    blocks.forEach((block, index) => {
      const topLevel = part !== "note" && block.type === "paragraph";
      for (const paragraph of paragraphsOf(block)) {
        add(displayText(paragraph), section === undefined ? { part } : { part, section, block: index, paragraph: topLevel });
      }
    });
  };
  model.sections.forEach((section, sectionIndex) => {
    for (const kind of ["default", "first", "even"] as const) {
      addBlocks(section.headers?.[kind]?.blocks ?? [], "header");
    }
    addBlocks(section.blocks, "body", sectionIndex);
    for (const kind of ["default", "first", "even"] as const) {
      addBlocks(section.footers?.[kind]?.blocks ?? [], "footer");
    }
  });
  for (const note of model.notes ?? []) {
    addBlocks(note.blocks, "note");
  }
  for (const thread of model.comments ?? []) {
    for (const message of thread.messages) {
      add(message.text, { part: "comment" });
    }
  }
  return { text: lines.join("\n"), coverage: "full", spans };
}

/** The span holding `offset` (spans are ascending and disjoint). */
function spanAt(spans: DocumentTextSpan[], offset: number): DocumentTextSpan | undefined {
  let lo = 0;
  let hi = spans.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const span = spans[mid];
    if (offset < span.start) {
      hi = mid - 1;
    } else if (offset > span.end) {
      lo = mid + 1;
    } else {
      return span;
    }
  }
  return undefined;
}

/** The caret an offset into `extraction` stands for — what matchSelection
 * and indexAtOrAfter compare. Inside a top-level body paragraph it is the
 * exact model offset (a field's displayed value maps to the field's own
 * position); inside a table or shape it is that block's start; in a
 * header, footer, note or comment, which the find bar does not search,
 * there is none. */
export function documentTextCaret(model: DocumentModel, extraction: DocumentText, offset: number): CaretPosition | null {
  const span = spanAt(extraction.spans, offset);
  if (!span || span.part !== "body" || span.section === undefined || span.block === undefined) {
    return null;
  }
  const block = model.sections[span.section]?.blocks[span.block];
  if (!block) {
    return null;
  }
  if (!span.paragraph || block.type !== "paragraph") {
    return { section: span.section, block: span.block, offset: 0 };
  }
  const target = offset - span.start;
  let shown = 0;
  let modelOffset = 0;
  for (const run of block.runs) {
    const text = run.field ? cachedFieldText(run.field as FieldSpec) : run.text;
    if (target < shown + text.length) {
      return { section: span.section, block: span.block, offset: run.field ? modelOffset : modelOffset + (target - shown) };
    }
    shown += text.length;
    modelOffset += run.text.length;
  }
  return { section: span.section, block: span.block, offset: modelOffset };
}
