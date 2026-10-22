// --- SPEC-056 equation insert commands --------------------------------------
// "Insert → Equation" inserts an empty inline equation at the caret (AC-1);
// "Insert → Display equation" inserts one centred on its own paragraph (AC-2).
// Both build a one-character equation-marker run (createEquationRun) so the
// equation is a single caret atom, exactly like a field or note marker.
import {
  isCollapsed,
  orderedRange,
  result,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import {
  cloneContainers,
  createEquationRun,
  createFieldSpecRun,
  sectionContentWidthCm,
  type BookmarkAnchor,
  type DocumentModel,
  type EquationSpec,
  type FieldSpec,
  type MathNode,
  type Paragraph,
  type Run,
} from "./documentModel";
import {
  formattingAtOffset,
  insertRun as insertRunInPlace,
  paragraphAt,
  splitParagraph as splitParagraphInPlace,
} from "./textEditing";
import { deleteRange as deleteRangeInPlace } from "./textEditing";

/** Inserts an empty inline equation at the caret (AC-1). A non-collapsed
 * selection is replaced first (the typing-over rule), same as insertField. The
 * caret lands on the equation atom so the editor can enter equation-edit mode
 * "inside it". */
export function insertInlineEquation(
  model: DocumentModel,
  selection: ModelSelection,
  body: MathNode[] = [],
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  const paragraph = paragraphAt(next, base);
  if (!paragraph) {
    return result(next, base);
  }
  const formatting = { ...formattingAtOffset(paragraph, base.offset) };
  delete formatting.linkHref;
  delete formatting.linkAnchor;
  const equation: EquationSpec = { display: "inline", body };
  const caret = insertRunInPlace(next, base, createEquationRun(equation, formatting));
  return result(next, caret);
}

/** Inserts a display equation centred on its own paragraph (AC-2). The caret's
 * paragraph is split at the caret (unless it sits at the very start) so the
 * equation gets a paragraph to itself, Caption/number machinery (SPEC-055) can
 * attach to that paragraph, and the surrounding text is untouched. Inert
 * inside a table cell, like a break. */
export function insertDisplayEquation(
  model: DocumentModel,
  selection: ModelSelection,
  body: MathNode[] = [],
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  if (base.cell) {
    // No display-equation paragraph inside a cell — fall back to inline.
    return insertInlineEquation(model, selection, body);
  }
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
  const equation: EquationSpec = { display: "block", body };
  const paragraph: Paragraph = {
    type: "paragraph",
    runs: [createEquationRun(equation)],
    formatting: { alignment: "center" },
  };
  blocks.splice(at, 0, paragraph);
  // Leave a home to type into after the equation when nothing follows.
  if (at + 1 >= blocks.length) {
    blocks.push({ type: "paragraph", runs: [] });
  }
  return result(next, { section: base.section, block: at, offset: 0 });
}

/** The `SEQ` field code an equation number carries. Kept as an *unsupported*
 * field (SPEC-053) so it round-trips verbatim — the reader never reparses a
 * SEQ instruction into a lossy `{label, format}` model that drops its switches
 * (see docx/fields.spec.ts's fidelity guard). The label "Equation" is what
 * scopes the sequence, exactly as SPEC-055 captions do. */
export const EQUATION_SEQ_CODE = "SEQ Equation \\* ARABIC";

/** Every bookmark name already used anywhere in the model, so a new equation's
 * cross-reference anchor never collides (AC-18). */
function usedBookmarkNames(model: DocumentModel): Set<string> {
  const names = new Set<string>();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : block.type === "paragraph"
            ? [block]
            : [];
      for (const paragraph of paragraphs) {
        for (const bookmark of paragraph.bookmarks ?? []) names.add(bookmark.name);
      }
    }
  }
  return names;
}

/** How many numbered equations already exist, so the next one continues the
 * sequence. Counts `SEQ Equation` field runs in document order — the same
 * per-label counting SPEC-055 defines, with no equation-specific numbering
 * path. */
function equationNumberCount(model: DocumentModel): number {
  let count = 0;
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type !== "paragraph") continue;
      for (const run of block.runs) {
        if (run.field?.kind === "unsupported" && run.field.code.startsWith("SEQ Equation")) count += 1;
      }
    }
  }
  return count;
}

/** Inserts a *numbered* display equation (AC-17): a centred display equation on
 * its own paragraph, with the number pushed to the right margin by a right tab
 * stop (SPEC-054) as a caption carrying the built-in Equation label's `SEQ`
 * field (SPEC-055). The paragraph gets a bookmark so the equation is a valid
 * cross-reference target (AC-18). */
export function insertNumberedDisplayEquation(
  model: DocumentModel,
  selection: ModelSelection,
  body: MathNode[] = [],
): CommandResult {
  const inserted = insertDisplayEquation(model, selection, body);
  const next = inserted.model;
  const at = inserted.selection.focus.block;
  const section = inserted.selection.focus.section;
  const paragraph = next.sections[section].blocks[at];
  if (paragraph.type !== "paragraph" || inserted.selection.focus.cell) {
    // The cell fallback produced an inline equation; nothing to number.
    return inserted;
  }
  const number = equationNumberCount(next) + 1;
  const seqField: FieldSpec = { kind: "unsupported", code: EQUATION_SEQ_CODE, cached: String(number) };
  const numberRuns: Run[] = [
    { text: "\t", formatting: {} },
    { text: "(", formatting: {} },
    createFieldSpecRun(seqField),
    { text: ")", formatting: {} },
  ];
  paragraph.runs = [...paragraph.runs, ...numberRuns];

  // The bookmark that makes it a cross-reference target (AC-18), placed at the
  // number so a reference reads the equation's own number.
  const names = usedBookmarkNames(next);
  let bookmarkBase = number;
  let name = `Equation${bookmarkBase}`;
  while (names.has(name)) name = `Equation${++bookmarkBase}`;
  const numberOffset = paragraph.runs.slice(0, 1).reduce((n, run) => n + run.text.length, 0) + 1; // after the equation atom + tab
  const bookmark: BookmarkAnchor = { name, offset: numberOffset };
  paragraph.bookmarks = [...(paragraph.bookmarks ?? []), bookmark];

  // The right tab stop that pushes the number to the margin (SPEC-054).
  const rightCm = Math.max(1, sectionContentWidthCm(next, section));
  paragraph.formatting = {
    ...paragraph.formatting,
    tabStops: [...(paragraph.formatting?.tabStops ?? []), { positionCm: rightCm, alignment: "right" }],
  };
  return result(next, inserted.selection.focus);
}
