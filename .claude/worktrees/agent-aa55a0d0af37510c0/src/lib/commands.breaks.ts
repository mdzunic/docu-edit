// --- manual breaks and section breaks (SPEC-028) ----------------------------
import {
  isCollapsed,
  orderedRange,
  result,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import {
  cloneContainers,
  type BreakBlock,
  type BreakKind,
  type DocumentModel,
  type Section,
  type SectionStart,
} from "./documentModel";
import {
  deleteRange as deleteRangeInPlace,
  paragraphAt,
  splitParagraph as splitParagraphInPlace,
} from "./textEditing";

/** Inserts a manual break at the selection (SPEC-028 AC-4..7): the caret's
 * paragraph is split around a new break block, so the content after the
 * caret starts a new page (AC-5). At a paragraph start no split happens —
 * the break lands before the paragraph, leaving no stray empty half; the
 * one exception is a break with nothing before it at all, which gets an
 * empty paragraph so the page it ends stays reachable. A
 * non-collapsed selection is deleted first. Inert inside a table cell:
 * breaks are top-level blocks, and there is no page flow to break inside a
 * cell. The caret ends up at the start of the content after the break. */
export function insertBreak(
  model: DocumentModel,
  selection: ModelSelection,
  kind: BreakKind,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  if (base.cell) {
    return result(next, base);
  }
  const breakBlock: BreakBlock = { type: "break", kind };
  const blocks = next.sections[base.section].blocks;
  // Atom caret (table/image/break edge): the break lands on the side of the
  // atom the caret addresses. A paragraph caret splits its paragraph in two
  // and the break goes between the halves — unless it sits at the very
  // start, where there is no half to leave behind.
  const atom = !paragraphAt(next, base);
  let at: number;
  if (!atom && base.offset > 0) {
    at = splitParagraphInPlace(next, base).block;
  } else if (atom && base.offset >= 1) {
    at = base.block + 1;
  } else {
    at = base.block;
  }
  // Nothing precedes the break on the page it ends — it is the first block
  // of the document, or of a section that begins its own page. That page
  // would render empty AND unreachable: no paragraph on it for a click or
  // an arrow key to put the caret in, so its content could never be
  // written. Give it a home, the way insertSectionBreak backfills an empty
  // half of a split section.
  const startsOwnPage =
    base.section === 0 || next.sections[base.section].start !== "continuous";
  if (at === 0 && startsOwnPage) {
    blocks.splice(0, 0, { type: "paragraph", runs: [] });
    at = 1;
  }
  blocks.splice(at, 0, breakBlock);
  // …and a home on the new page when nothing follows.
  if (at + 1 >= blocks.length) {
    blocks.push({ type: "paragraph", runs: [] });
  }
  return result(next, { section: base.section, block: at + 1, offset: 0 });
}

/** Inserts a section break at the selection (SPEC-028 AC-8/9): the caret's
 * section is split in two at the caret (the paragraph splits too, unless
 * the caret sits at its start), and the new section — everything from the
 * caret on — begins per `sectionStart` ("nextPage" or "continuous"). The
 * new section inherits the split section's own page setup, so inserting
 * then deleting a break restores the original state. Inert inside a table
 * cell. The caret ends up at the start of the new section. */
export function insertSectionBreak(
  model: DocumentModel,
  selection: ModelSelection,
  sectionStart: SectionStart,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  if (base.cell) {
    return result(next, base);
  }
  const section = next.sections[base.section];
  let boundary: number;
  if (paragraphAt(next, base) && base.offset > 0) {
    boundary = splitParagraphInPlace(next, base).block;
  } else if (!paragraphAt(next, base) && base.offset >= 1) {
    boundary = base.block + 1;
  } else {
    boundary = base.block;
  }
  const before = section.blocks.slice(0, boundary);
  const after = section.blocks.slice(boundary);
  if (before.length === 0) {
    before.push({ type: "paragraph", runs: [] });
  }
  if (after.length === 0) {
    after.push({ type: "paragraph", runs: [] });
  }
  const created: Section = { blocks: after, start: sectionStart };
  if (section.pageSize) {
    created.pageSize = { ...section.pageSize };
  }
  if (section.margins) {
    created.margins = { ...section.margins };
  }
  // The first half keeps the original's preserved sectPr (it still begins
  // the way the original did); the created section's sectPr is regenerated
  // from its modelled geometry so its own `w:type` (sectionStart) is
  // emitted, never a stale copy. Splitting the final section moves the
  // preserved body-level sectPr onto the now-non-final first half.
  if (base.section === next.sections.length - 1 && next.sourceXml?.docxSectPr !== undefined) {
    section.docxSectPr = next.sourceXml.docxSectPr;
    const sourceXml = { ...next.sourceXml };
    delete sourceXml.docxSectPr;
    next.sourceXml = sourceXml;
  }
  const opaque = section.opaqueBlocks ?? [];
  const movedOpaque = opaque
    .filter((entry) => entry.blockIndex >= boundary)
    .map((entry) => ({ ...entry, blockIndex: entry.blockIndex - boundary }));
  if (movedOpaque.length > 0) {
    created.opaqueBlocks = movedOpaque;
    const kept = opaque.filter((entry) => entry.blockIndex < boundary);
    if (kept.length > 0) {
      section.opaqueBlocks = kept;
    } else {
      delete section.opaqueBlocks;
    }
  }
  section.blocks = before;
  next.sections.splice(base.section + 1, 0, created);
  return result(next, { section: base.section + 1, block: 0, offset: 0 });
}
