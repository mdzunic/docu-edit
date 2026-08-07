// Multi-column layout commands (SPEC-030): applying a column layout to a
// section, wrapping a selection in its own continuous-break section, dragging
// a column boundary on the ruler, and re-fitting a section's tables to the new
// column width. Framework-free like the other command modules (commands.ts,
// tables.ts); every command clones the containers it touches and returns a new
// model, so the snapshot-based undo history covers it for free.
import {
  MAX_COLUMN_COUNT,
  MIN_SECTION_COLUMN_WIDTH_CM,
  isTable,
  resolveColumnWidthsCm,
  sectionColumns,
  sectionContentWidthCm,
  type DocumentModel,
  type OpaqueBlock,
  type Section,
  type SectionColumns,
  type Table,
} from "./documentModel";
import { cmToTwips, twipsToCm } from "./docx/units";
import { caretSelection, orderedRange, type CommandResult, type ModelSelection } from "./commands";

/** Where a Columns dialog apply targets (AC-5): the caret's current section,
 * or the current selection (wrapped in its own continuous section). */
export type ColumnsScope = "thisSection" | "selection";

/** Typed pass-through for a scope coming from a `.vue` template's plain
 * string (same espree constraint as documentModel.asCharacterFormatKey). */
export function asColumnsScope(value: string): ColumnsScope {
  return value as ColumnsScope;
}

// Widths snap to the twip grid both file formats reproduce exactly (the same
// rule as commands.ts's snapCm / tables.ts's snapCm).
function snapCm(cm: number): number {
  return twipsToCm(cmToTwips(cm));
}

/** Normalises a requested column layout to a stored `SectionColumns`, or
 * `undefined` for a single column (count 1) — the model carries no `columns`
 * field for single-column sections, so they stay byte-identical to
 * pre-SPEC-030 documents. Counts clamp to 1..6 (AC-2), widths snap to the
 * twip grid and never fall below the minimum column width. */
export function normalizeColumns(columns: SectionColumns): SectionColumns | undefined {
  const count = Math.max(1, Math.min(MAX_COLUMN_COUNT, Math.round(columns.count)));
  if (count <= 1) {
    return undefined;
  }
  const gapCm = snapCm(Math.max(0, columns.gapCm));
  const equalWidth = columns.equalWidth !== false;
  const separator = columns.separator === true;
  const normalized: SectionColumns = { count, equalWidth, gapCm, separator };
  if (!equalWidth && columns.widthsCm && columns.widthsCm.length === count) {
    normalized.widthsCm = columns.widthsCm.map((w) =>
      snapCm(Math.max(MIN_SECTION_COLUMN_WIDTH_CM, w)),
    );
  } else {
    normalized.equalWidth = true;
  }
  return normalized;
}

/** The column layout the dialog/ruler should show for the caret's section —
 * always a concrete `SectionColumns` (a single column when the section has
 * none), so the dialog can seed its fields. */
export function sectionColumnsState(model: DocumentModel, selection: ModelSelection): SectionColumns {
  const section = orderedRange(model, selection).start.section;
  return sectionColumns(model, section);
}

// A section with its column layout replaced (or removed), its preserved
// `w:sectPr` dropped so the writer regenerates it with the new `w:cols` — the
// same rule commands.ts's withSectionGeometry applies for page geometry.
function withColumns(section: Section, columns: SectionColumns | undefined): Section {
  const updated: Section = { ...section };
  if (columns) {
    updated.columns = columns;
  } else {
    delete updated.columns;
  }
  delete updated.docxSectPr;
  return refitSectionTables(updated);
}

// Drops the document's preserved body-level `w:sectPr` when the edited section
// is the last one, so its regenerated sectPr (carrying the new columns) wins
// over the stale verbatim copy — mirrors commands.ts's clearPreservedSectPr.
function clearFinalPreservedSectPr(
  model: DocumentModel,
  sectionIndex: number,
): DocumentModel {
  if (sectionIndex !== model.sections.length - 1 || model.sourceXml?.docxSectPr === undefined) {
    return model;
  }
  const sourceXml = { ...model.sourceXml };
  delete sourceXml.docxSectPr;
  return { ...model, sourceXml };
}

/** Applies a column layout (AC-1..AC-6). `scope` "thisSection" replaces the
 * caret's section's columns; "selection" wraps the touched paragraphs in a new
 * continuous section first (AC-7). */
export function applyColumns(
  model: DocumentModel,
  selection: ModelSelection,
  columns: SectionColumns,
  scope: ColumnsScope,
): CommandResult {
  const normalized = normalizeColumns(columns);
  const { start } = orderedRange(model, selection);
  // Applying to a selection inside a table cell has no meaningful paragraph
  // range to wrap — fall back to the containing section.
  if (scope === "selection" && !start.cell) {
    return applyColumnsToSelection(model, selection, normalized);
  }
  const target = start.section;
  const next = clearFinalPreservedSectPr(
    {
      ...model,
      sections: model.sections.map((section, index) =>
        index === target ? withColumns(section, normalized) : section,
      ),
    },
    target,
  );
  const range = orderedRange(next, selection);
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

// Wraps the paragraphs the selection touches (whole paragraphs — AC-7b never
// splits one) in a new continuous-break section carrying `columns` (AC-7a),
// splitting the caret's section into up to three: the untouched blocks before,
// the wrapped middle, and the untouched blocks after (also continuous, so it
// resumes the section's original layout). An empty selection wraps just the
// caret's paragraph (AC-8).
function applyColumnsToSelection(
  model: DocumentModel,
  selection: ModelSelection,
  columns: SectionColumns | undefined,
): CommandResult {
  const { start, end } = orderedRange(model, selection);
  const sectionIndex = start.section;
  const section = model.sections[sectionIndex];
  const blocks = section.blocks;
  const firstBlock = Math.min(start.block, blocks.length - 1);
  // The last paragraph the selection touches: the end block, unless the range
  // ends at the very start of a later block (that block is not really touched)
  // or the range spills into a following section (then the section's last
  // block). A collapsed caret touches just its own block (AC-8).
  let lastBlock: number;
  if (end.section > sectionIndex) {
    lastBlock = blocks.length - 1;
  } else if (end.block > firstBlock && end.offset === 0) {
    lastBlock = end.block - 1;
  } else {
    lastBlock = Math.min(end.block, blocks.length - 1);
  }
  lastBlock = Math.max(lastBlock, firstBlock);

  const before = blocks.slice(0, firstBlock);
  const middle = blocks.slice(firstBlock, lastBlock + 1);
  const after = blocks.slice(lastBlock + 1);

  const geometry: Partial<Section> = {};
  if (section.pageSize) {
    geometry.pageSize = { ...section.pageSize };
  }
  if (section.margins) {
    geometry.margins = { ...section.margins };
  }

  const opaque = section.opaqueBlocks ?? [];
  const sliceOpaque = (lo: number, hi: number): OpaqueBlock[] =>
    opaque
      .filter((entry) => entry.blockIndex >= lo && entry.blockIndex < hi)
      .map((entry) => ({ ...entry, blockIndex: entry.blockIndex - lo }));

  const created: Section[] = [];
  // The section that keeps the original's identity (start, preserved sectPr,
  // geometry) is the "before" half when it has content, else the middle.
  if (before.length > 0) {
    const beforeSection: Section = { ...section, blocks: before };
    const beforeOpaque = sliceOpaque(0, firstBlock);
    if (beforeOpaque.length > 0) {
      beforeSection.opaqueBlocks = beforeOpaque;
    } else {
      delete beforeSection.opaqueBlocks;
    }
    created.push(beforeSection);
  }

  const middleSection: Section = refitSectionTables({
    blocks: middle,
    // A middle following a "before" half needs a continuous break to begin;
    // a middle at the section's start inherits the original's own start.
    start: before.length > 0 ? "continuous" : section.start,
    ...geometry,
  });
  if (columns) {
    middleSection.columns = columns;
  }
  const middleOpaque = sliceOpaque(firstBlock, lastBlock + 1);
  if (middleOpaque.length > 0) {
    middleSection.opaqueBlocks = middleOpaque;
  }
  created.push(middleSection);

  if (after.length > 0) {
    const afterSection: Section = { blocks: after, start: "continuous", ...geometry };
    // The trailing content resumes the section's original column layout.
    if (section.columns) {
      afterSection.columns = { ...section.columns };
    }
    const afterOpaque = sliceOpaque(lastBlock + 1, blocks.length + 1);
    if (afterOpaque.length > 0) {
      afterSection.opaqueBlocks = afterOpaque;
    }
    created.push(refitSectionTables(afterSection));
  }

  const sections = [
    ...model.sections.slice(0, sectionIndex),
    ...created,
    ...model.sections.slice(sectionIndex + 1),
  ];
  let next: DocumentModel = { ...model, sections };
  // Splitting the final section moves the preserved body-level sectPr onto the
  // now-non-final first half (as insertSectionBreak does); if the first half
  // is the wrapped middle itself we drop it, since the middle's sectPr must
  // regenerate to carry the new columns.
  if (sectionIndex === model.sections.length - 1 && model.sourceXml?.docxSectPr !== undefined) {
    const sourceXml = { ...model.sourceXml };
    delete sourceXml.docxSectPr;
    if (before.length > 0) {
      sections[sectionIndex] = { ...sections[sectionIndex], docxSectPr: model.sourceXml.docxSectPr };
    }
    next = { ...next, sourceXml };
  }
  const middleIndex = sectionIndex + (before.length > 0 ? 1 : 0);
  return { model: next, selection: caretSelection({ section: middleIndex, block: 0, offset: 0 }) };
}

/** Drag of a column boundary on the ruler (AC-14): resizes the two columns
 * adjacent to gutter `boundary` (0-based, the gutter to the right of column
 * `boundary`) by `deltaCm`, keeping the gutter and the total content width
 * fixed. Widths never fall below the minimum. A no-op outside a multi-column
 * section. */
export function setColumnBoundary(
  model: DocumentModel,
  selection: ModelSelection,
  boundary: number,
  deltaCm: number,
): CommandResult {
  const sectionIndex = orderedRange(model, selection).start.section;
  const columns = sectionColumns(model, sectionIndex);
  if (columns.count <= 1 || boundary < 0 || boundary >= columns.count - 1) {
    return { model, selection };
  }
  const widths = resolveColumnWidthsCm(columns, sectionContentWidthCm(model, sectionIndex));
  const left = widths[boundary];
  const right = widths[boundary + 1];
  // Clamp the drag so neither adjacent column drops below the minimum.
  const applied = Math.max(
    MIN_SECTION_COLUMN_WIDTH_CM - left,
    Math.min(right - MIN_SECTION_COLUMN_WIDTH_CM, deltaCm),
  );
  widths[boundary] = snapCm(left + applied);
  widths[boundary + 1] = snapCm(right - applied);
  const updated: SectionColumns = {
    count: columns.count,
    equalWidth: false,
    gapCm: columns.gapCm,
    separator: columns.separator,
    widthsCm: widths.map(snapCm),
  };
  const next = clearFinalPreservedSectPr(
    {
      ...model,
      sections: model.sections.map((section, index) =>
        index === sectionIndex ? withColumns(section, updated) : section,
      ),
    },
    sectionIndex,
  );
  const range = orderedRange(next, selection);
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

// Scales down any table in the section whose total width exceeds the section's
// column width, so a table already present when columns are applied lays out
// within the column rather than overflowing it (AC-15). Tables that already
// fit keep their exact object identity (v-memo).
function refitSectionTables(section: Section): Section {
  const columns = section.columns;
  const contentWidthCm = section.pageSize
    ? section.pageSize.widthCm - (section.margins?.leftCm ?? 0) - (section.margins?.rightCm ?? 0)
    : undefined;
  // Without the section's own geometry the column width is unknown here (it
  // depends on document defaults resolved at layout time); the insert path and
  // the pagination containment (overflow:hidden) handle those cases.
  if (!columns || columns.count <= 1 || contentWidthCm === undefined) {
    return section;
  }
  const columnWidth = Math.min(...resolveColumnWidthsCm(columns, Math.max(0.1, contentWidthCm)));
  let changed = false;
  const blocks = section.blocks.map((block) => {
    if (!isTable(block)) {
      return block;
    }
    const total = block.columnWidthsCm.reduce((sum, w) => sum + w, 0);
    if (total <= columnWidth) {
      return block;
    }
    changed = true;
    const scale = columnWidth / total;
    const scaled: Table = {
      ...block,
      columnWidthsCm: block.columnWidthsCm.map((w) => snapCm(w * scale)),
    };
    delete scaled.odtColumnStyleNames;
    return scaled;
  });
  return changed ? { ...section, blocks } : section;
}
