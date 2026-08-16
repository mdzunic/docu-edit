// --- clear formatting and format painter (SPEC-020) -------------------------
import {
  isCollapsed,
  orderedRange,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import { caretFormatting } from "./commands.formatting";
import { firstTouchedFormatting, touchedBlocks } from "./commands.paragraph";
import { restyledParagraph, runLayersFor } from "./commands.styles";
import {
  cloneContainers,
  type DocumentModel,
  type Paragraph,
  type ParagraphFormatting,
  type RunFormatting,
} from "./documentModel";
import { NORMAL_STYLE_ID, paragraphStyleBase, resolveNamedStyle } from "./styles";
import {
  blockLength,
  paragraphAt,
  replaceFormatting,
  replaceParagraphAt,
} from "./textEditing";

/** "Clear formatting" (AC-1, AC-2): every touched paragraph is re-baked
 * exactly like "Reset style" (see `restyledParagraph` above) but resolved
 * against Normal instead of the paragraph's own style, with every run's
 * direct-formatting diff forced empty — so bold/italic/font/colour/etc.
 * disappear regardless of whether the paragraph's style or its own direct
 * formatting carried them, and the paragraph's style reference becomes
 * Normal explicitly (the same value `handleApplyStyle("Normal")` writes).
 * `restyledParagraph` only ever replaces a run's `formatting`, never its
 * `text`, so content is untouched; `list` and `opaqueInlines` are carried
 * over by its `{ ...block }` spread, so list membership and unmodelled XML
 * survive too (AC-4, AC-6). */
export function clearFormatting(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const resolved = resolveNamedStyle(next, NORMAL_STYLE_ID);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const base = paragraphStyleBase(next, block);
    const replacement = restyledParagraph(
      block,
      resolved,
      {},
      runLayersFor(next, block, base).map((layers) => ({ underlay: layers.underlay, diff: {} })),
    );
    replacement.styleId = NORMAL_STYLE_ID;
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** The format painter's copy source (AC-9, AC-10): the caret's character
 * formatting, plus paragraph formatting (alignment, indents, spacing —
 * SPEC-014's ParagraphFormatting) when it was copied from a whole
 * paragraph. `paragraph` absent means the copy came from a partial
 * selection, so applying it only ever touches characters, never a
 * paragraph's own formatting. */
export interface PainterFormat {
  run: RunFormatting;
  paragraph?: ParagraphFormatting;
}

/** Typed `null` seed for painter-format refs in `.vue` scripts (same espree
 * constraint as noSelection). */
export function noPainterFormat(): PainterFormat | null {
  return null;
}

/** Whether `selection` spans exactly one (non-table) paragraph from its
 * first character to its last (AC-10's "a whole paragraph is selected") —
 * an empty paragraph's only possible span (a collapsed caret at offset 0)
 * counts too. Table cells are excluded (copying table formatting is out of
 * scope). */
export function isWholeParagraphSelected(model: DocumentModel, selection: ModelSelection): boolean {
  const { start, end } = orderedRange(model, selection);
  if (start.section !== end.section || start.block !== end.block || start.cell || end.cell) {
    return false;
  }
  const block = paragraphAt(model, { section: start.section, block: start.block });
  if (!block) {
    return false;
  }
  return start.offset === 0 && end.offset === blockLength(block);
}

/** Captures the format painter's copy source at the current selection
 * (AC-9, AC-10) — the toolbar button's click/dblclick handler. */
export function capturePainterFormat(model: DocumentModel, selection: ModelSelection): PainterFormat {
  const painter: PainterFormat = { run: caretFormatting(model, selection.focus) };
  if (isWholeParagraphSelected(model, selection)) {
    painter.paragraph = firstTouchedFormatting(model, selection);
  }
  return painter;
}

/** Applies the format painter's copied formatting to a new selection
 * (AC-12): character formatting replaces — never merges into — every run
 * the selection covers (a collapsed selection has no characters to
 * repaint, so this is skipped), and, only when formatting was captured from
 * a whole paragraph, every touched paragraph's own formatting is replaced
 * the same way. */
export function applyPainterFormat(
  model: DocumentModel,
  selection: ModelSelection,
  painter: PainterFormat,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  if (!isCollapsed(selection)) {
    replaceFormatting(next, start, end, painter.run);
  }
  if (painter.paragraph !== undefined) {
    for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
      const block = paragraphAt(next, address);
      if (!block) {
        continue;
      }
      const replacement: Paragraph = { ...block };
      if (Object.keys(painter.paragraph).length > 0) {
        replacement.formatting = { ...painter.paragraph };
      } else {
        delete replacement.formatting;
      }
      replaceParagraphAt(next, address, replacement);
    }
  }
  return { model: next, selection: { anchor: start, focus: end } };
}
