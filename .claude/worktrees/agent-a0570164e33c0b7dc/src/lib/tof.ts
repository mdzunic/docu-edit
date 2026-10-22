// Table of figures (SPEC-055): builds a ToF block's entries from the
// document's Caption-styled paragraphs of one label, detects staleness against
// the document's current captions and pagination, and the block's own
// insert/update commands. Mirrors lib/toc.ts exactly (a ToF is to captions
// what a TOC is to headings); framework-free, and — like toc.ts — a command
// module OUTSIDE commands.ts so commands.undo.spec's completeness guard never
// accounts for these.
import {
  cloneContainers,
  isTof,
  type Block,
  type DocumentModel,
  type TableOfFigures,
  type TofEntry,
  type TofOptions,
} from "./documentModel";
import { listCaptions } from "./captions";
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

/** Shown in the block when the document has no captions of the chosen label
 * to list (mirrors TOC_PLACEHOLDER_TEXT), and written as the .docx field's
 * cached result / the .odt index body's sole paragraph so the empty state
 * round-trips too. */
export const TOF_PLACEHOLDER_TEXT = "No table of figures entries found.";

/** A freshly inserted table of figures's default options (SPEC-055). */
export function defaultTofOptions(): TofOptions {
  return { pageNumbers: true, dotLeader: true, hyperlinks: true };
}

/** The block a click "selected" — a ToF is contenteditable="false" like a
 * TOC, so the view tracks the selected one the same way (TocSelection). */
export interface TofSelection {
  section: number;
  blockIndex: number;
}

/** Typed `null` seed for a TofSelection ref in `.vue` scripts (same espree
 * constraint as noTocSelection). */
export function noTofSelection(): TofSelection | null {
  return null;
}

/** Resolves a page number for a caption anchor — supplied by the caller
 * (DocumentView), since only the pagination controller knows current page
 * numbers; keeps this module framework-free, the same shape as
 * TocPageNumberLookup. */
export type TofPageNumberLookup = (anchor: string) => number | null;

/** A lookup that always returns null — for building entries without a mounted
 * canvas (as tests do). */
export function noTofPageNumbers(): TofPageNumberLookup {
  return () => null;
}

/** The distinct caption labels the document currently has captions for, in
 * first-appearance order — the "Table of figures" dialog's label list. Falls
 * back to "Figure" so the dialog always offers something. */
export function availableTofLabels(model: DocumentModel): string[] {
  const seen: string[] = [];
  for (const caption of listCaptions(model)) {
    if (!seen.some((label) => label.toLowerCase() === caption.label.toLowerCase())) {
      seen.push(caption.label);
    }
  }
  return seen.length > 0 ? seen : ["Figure"];
}

/** The ToF entries a document's current captions of `label` produce (SPEC-055):
 * every body caption of that label, in document order, with its rendered text,
 * derived anchor and — when `pageNumberFor` is given — its current page
 * number. Case-insensitive on the label, so "figure" and "Figure" list
 * together. */
export function tofEntriesFromCaptions(
  model: DocumentModel,
  label: string,
  pageNumberFor: TofPageNumberLookup = noTofPageNumbers(),
): TofEntry[] {
  const lower = label.toLowerCase();
  return listCaptions(model)
    .filter((caption) => caption.label.toLowerCase() === lower)
    .map((caption) => ({
      text: caption.text,
      page: pageNumberFor(caption.anchor),
      targetBookmark: caption.anchor,
    }));
}

/** Whether the document's captions have changed since `tof` was generated
 * (mirrors tocHeadingsStale): a different set of listed captions, or different
 * text — compared by text only, never by anchor (a caption read from a
 * Word-authored fixture carries Word's own bookmark as its anchor). Page
 * numbers are ignored here (that is tofPageNumbersStale's job), so the two
 * staleness conditions stay independent. */
export function tofEntriesStale(model: DocumentModel, tof: TableOfFigures): boolean {
  const current = tofEntriesFromCaptions(model, tof.label);
  if (current.length !== tof.entries.length) {
    return true;
  }
  return current.some((entry, index) => entry.text !== tof.entries[index].text);
}

/** Whether any entry's page number differs from what `pageNumberFor` resolves
 * now (mirrors tocPageNumbersStale) — pagination having reflowed since the ToF
 * was last generated/updated. A freshly generated ToF is never stale here: its
 * entries were built with this same lookup. */
export function tofPageNumbersStale(tof: TableOfFigures, pageNumberFor: TofPageNumberLookup): boolean {
  return tof.entries.some((entry) => pageNumberFor(entry.targetBookmark) !== entry.page);
}

function result(model: DocumentModel, caret: CaretPosition): CommandResult {
  return { model, selection: caretSelection(caret) };
}

/** Inserts a table of figures at the selection (SPEC-055), mirroring
 * insertTableOfContents: a block-level atom built from the document's current
 * captions of `label` (or the empty-state placeholder). Inert inside a table
 * cell. The caret ends up just after the inserted block. */
export function insertTableOfFigures(
  model: DocumentModel,
  selection: ModelSelection,
  label: string,
  options: TofOptions,
  pageNumberFor: TofPageNumberLookup = noTofPageNumbers(),
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  if (base.cell) {
    return result(next, base);
  }
  const tof: TableOfFigures = {
    type: "tof",
    label,
    options: { ...options },
    entries: tofEntriesFromCaptions(next, label, pageNumberFor),
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
  blocks.splice(at, 0, tof);
  if (at + 1 >= blocks.length) {
    blocks.push({ type: "paragraph", runs: [] });
  }
  return result(next, { section: base.section, block: at + 1, offset: 0 });
}

function replaceTofAt(
  model: DocumentModel,
  section: number,
  blockIndex: number,
  next: TableOfFigures,
): void {
  const blocks = [...model.sections[section].blocks];
  blocks[blockIndex] = next;
  model.sections[section] = { ...model.sections[section], blocks };
}

/** "Update entire table" (SPEC-055): rebuilds a ToF block's entries wholesale
 * from the document's current captions of its own label — what a
 * captions-changed staleness resolves. Returns a cloned model even when the
 * address is not a ToF. */
export function updateTofEntirely(
  model: DocumentModel,
  selection: ModelSelection,
  section: number,
  blockIndex: number,
  pageNumberFor: TofPageNumberLookup = noTofPageNumbers(),
): CommandResult {
  const next = cloneContainers(model);
  const block: Block | undefined = next.sections[section]?.blocks[blockIndex];
  if (!isTof(block)) {
    return { model: next, selection };
  }
  replaceTofAt(next, section, blockIndex, {
    ...block,
    entries: tofEntriesFromCaptions(next, block.label, pageNumberFor),
  });
  return { model: next, selection };
}

/** "Update page numbers only" (SPEC-055): keeps the current entries' text and
 * order, refreshing only each entry's page number — what a pagination-only
 * staleness resolves without re-scanning captions. */
export function updateTofPageNumbersOnly(
  model: DocumentModel,
  selection: ModelSelection,
  section: number,
  blockIndex: number,
  pageNumberFor: TofPageNumberLookup,
): CommandResult {
  const next = cloneContainers(model);
  const block: Block | undefined = next.sections[section]?.blocks[blockIndex];
  if (!isTof(block)) {
    return { model: next, selection };
  }
  replaceTofAt(next, section, blockIndex, {
    ...block,
    entries: block.entries.map((entry) => ({ ...entry, page: pageNumberFor(entry.targetBookmark) })),
  });
  return { model: next, selection };
}

/** Rebuilds every table-of-figures block in the document from current captions
 * (mirrors how updateAllFields rebuilds every TOC) — used by "Update all
 * fields" so a ToF refreshes alongside the caption numbers and references. */
export function updateAllTablesOfFigures(
  model: DocumentModel,
  pageNumberFor: TofPageNumberLookup = noTofPageNumbers(),
): DocumentModel {
  let next = model;
  next.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isTof(block)) {
        next = updateTofEntirely(next, { anchor: { section: sectionIndex, block: blockIndex, offset: 0 }, focus: { section: sectionIndex, block: blockIndex, offset: 0 } }, sectionIndex, blockIndex, pageNumberFor).model;
      }
    });
  });
  return next;
}
