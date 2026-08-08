// Table of contents (SPEC-033): builds a TOC block's entries from the
// document's Heading 1-6 paragraphs, detects staleness against the
// document's current headings and pagination, and the block's own
// insert/update commands. Framework-free like every lib module, and — like
// lib/links.ts — a command module OUTSIDE commands.ts with its own
// cloneContainers copy, so commands.undo.spec.ts's completeness guard
// (scoped to commands.ts's own exports) never has to account for these.
import {
  isToc,
  type Block,
  type DocumentModel,
  type TableOfContents,
  type TocEntry,
  type TocOptions,
} from "./documentModel";
import { listHeadings } from "./links";
import {
  deleteRange as deleteRangeInPlace,
  paragraphAt,
  splitParagraph as splitParagraphInPlace,
  type CaretPosition,
} from "./textEditing";
import {
  caretSelection,
  isCollapsed,
  orderedRange,
  type CommandResult,
  type ModelSelection,
} from "./commands";

/** Shown in the block when the document has no Heading 1-6 paragraphs to
 * list (SPEC-033 AC-23), and written as the .docx field's cached result /
 * the .odt index body's sole paragraph so the empty state round-trips too. */
export const TOC_PLACEHOLDER_TEXT = "No table of contents entries found.";

/** The heading levels a freshly inserted TOC includes by default — Word's
 * own default (top three levels). */
export const DEFAULT_TOC_LEVELS = [1, 2, 3];

/** A freshly inserted TOC's default options (SPEC-033 AC-5..7). */
export function defaultTocOptions(): TocOptions {
  return { levels: [...DEFAULT_TOC_LEVELS], showPageNumbers: true, hyperlinksOnly: false };
}

/** Resolves a page number for a heading anchor — supplied by the caller
 * (DocumentView), since only the pagination controller (DOM measurement)
 * knows current page numbers; this module stays framework-free by taking
 * the lookup as a plain function, the same shape lib/pagination.ts's
 * MetricsProvider uses for the caller-supplied part of a pure computation. */
export type TocPageNumberLookup = (anchor: string) => number | null;

/** A lookup that always returns null — used where no pagination is
 * available yet (e.g. building entries without a mounted canvas, as tests
 * do). */
export function noTocPageNumbers(): TocPageNumberLookup {
  return () => null;
}

/** The TOC entries a document's current headings produce for `options`
 * (SPEC-033 AC-1..4): every Heading 1-6 paragraph whose level is included,
 * in document order, with its derived anchor (lib/links.ts's headingAnchor
 * — the same scheme internal hyperlinks use) and, when `pageNumberFor` is
 * given, its current page number. Anchors are derived from the FULL heading
 * list before filtering, so they stay stable regardless of which levels a
 * particular TOC includes. */
export function tocEntriesFromHeadings(
  model: DocumentModel,
  options: TocOptions,
  pageNumberFor: TocPageNumberLookup = noTocPageNumbers(),
): TocEntry[] {
  const levels = new Set(options.levels);
  return listHeadings(model)
    .filter((heading) => levels.has(heading.level))
    .map((heading) => ({
      text: heading.text,
      level: heading.level,
      anchor: heading.anchor,
      pageNumber: pageNumberFor(heading.anchor),
    }));
}

/** Whether the document's headings have changed since `toc` was generated
 * (SPEC-033 AC-10): a different set of included headings, different text,
 * or a different order — page numbers are deliberately ignored here (that
 * is tocPageNumbersStale's job), so the two staleness conditions the block's
 * "Update table of contents" control reacts to stay independent. */
export function tocHeadingsStale(model: DocumentModel, toc: TableOfContents): boolean {
  const current = tocEntriesFromHeadings(model, toc.options);
  if (current.length !== toc.entries.length) {
    return true;
  }
  return current.some(
    (entry, index) =>
      entry.text !== toc.entries[index].text ||
      entry.level !== toc.entries[index].level ||
      entry.anchor !== toc.entries[index].anchor,
  );
}

/** Whether any entry's page number differs from what `pageNumberFor`
 * resolves now (SPEC-033 AC-11) — pagination having reflowed the document
 * since the TOC was last generated/updated. */
export function tocPageNumbersStale(toc: TableOfContents, pageNumberFor: TocPageNumberLookup): boolean {
  return toc.entries.some((entry) => pageNumberFor(entry.anchor) !== entry.pageNumber);
}

// Same container-cloning as commands.ts (see that file's immutability
// contract); a local copy for the same reason links.ts keeps one.
function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
  };
}

function result(model: DocumentModel, caret: CaretPosition): CommandResult {
  return { model, selection: caretSelection(caret) };
}

/** Inserts a table of contents at the selection (SPEC-033 AC-1, AC-22): a
 * block-level atom built from the document's current Heading 1-6 paragraphs
 * (or the empty-state placeholder when there are none), following the same
 * split-and-splice shape as commands.ts's insertBreak. Inert inside a table
 * cell — a TOC is a top-level block, like a manual break. The caret ends up
 * just after the inserted block. */
export function insertTableOfContents(
  model: DocumentModel,
  selection: ModelSelection,
  options: TocOptions,
  pageNumberFor: TocPageNumberLookup = noTocPageNumbers(),
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  if (base.cell) {
    return result(next, base);
  }
  const toc: TableOfContents = {
    type: "toc",
    options: { ...options, levels: [...options.levels].sort((a, b) => a - b) },
    entries: tocEntriesFromHeadings(next, options, pageNumberFor),
  };
  const blocks = next.sections[base.section].blocks;
  const atom = !paragraphAt(next, base);
  let at: number;
  if (!atom && base.offset > 0) {
    at = splitParagraphInPlace(next, base).block;
  } else if (atom && base.offset >= 1) {
    at = base.block + 1;
  } else {
    at = base.block;
  }
  blocks.splice(at, 0, toc);
  if (at + 1 >= blocks.length) {
    blocks.push({ type: "paragraph", runs: [] });
  }
  return result(next, { section: base.section, block: at + 1, offset: 0 });
}

function replaceTocAt(
  model: DocumentModel,
  section: number,
  blockIndex: number,
  next: TableOfContents,
): void {
  const blocks = [...model.sections[section].blocks];
  blocks[blockIndex] = next;
  model.sections[section] = { ...model.sections[section], blocks };
}

/** "Update entire table" (SPEC-033 AC-13): rebuilds a TOC block's entries
 * wholesale from the document's current headings, using its own stored
 * options — what a headings-changed staleness resolves. A no-op (still
 * returns a cloned model, never the input) when the address is not a TOC. */
export function updateTocEntirely(
  model: DocumentModel,
  selection: ModelSelection,
  section: number,
  blockIndex: number,
  pageNumberFor: TocPageNumberLookup = noTocPageNumbers(),
): CommandResult {
  const next = cloneContainers(model);
  const block: Block | undefined = next.sections[section]?.blocks[blockIndex];
  if (!isToc(block)) {
    return { model: next, selection };
  }
  replaceTocAt(next, section, blockIndex, {
    ...block,
    entries: tocEntriesFromHeadings(next, block.options, pageNumberFor),
  });
  return { model: next, selection };
}

/** "Update page numbers only" (SPEC-033 AC-12): keeps the current entries'
 * text/level/anchor and order exactly as they are, refreshing only each
 * entry's page number — what a pagination-only staleness resolves without
 * re-scanning headings (so a heading added/removed elsewhere is not
 * silently folded in by this command). */
export function updateTocPageNumbersOnly(
  model: DocumentModel,
  selection: ModelSelection,
  section: number,
  blockIndex: number,
  pageNumberFor: TocPageNumberLookup,
): CommandResult {
  const next = cloneContainers(model);
  const block: Block | undefined = next.sections[section]?.blocks[blockIndex];
  if (!isToc(block)) {
    return { model: next, selection };
  }
  replaceTocAt(next, section, blockIndex, {
    ...block,
    entries: block.entries.map((entry) => ({ ...entry, pageNumber: pageNumberFor(entry.anchor) })),
  });
  return { model: next, selection };
}
