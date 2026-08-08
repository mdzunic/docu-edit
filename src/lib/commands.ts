// The editing command API (SPEC-011): every content mutation the editor
// performs goes through one of the commands below. Each command takes the
// current model plus a selection in model coordinates and returns a NEW
// model state — the input model is never modified, so callers can rely on
// reference inequality to know a command produced a change, and the DOM is
// never consulted (it is a projection of the model, never the source of
// truth).
//
// Immutability contract: commands clone the model's containers (the model
// object, its sections array, each section object and its blocks array) and
// delegate the actual edit to the SPEC-010 primitives in `textEditing`,
// which replace blocks rather than mutating them. Untouched Paragraph
// objects are therefore carried into the new model BY REFERENCE — that
// structural sharing is what keeps a command O(edit) instead of O(document),
// and it is safe precisely because nothing in the editing stack ever mutates
// a block in place. It is also what the canvas's `v-memo` keys on to skip
// re-rendering untouched paragraphs.
import {
  createEmptyHeaderFooterContent,
  createEmptyNoteContent,
  createFieldRun,
  createNoteRun,
  isTable,
  sectionHeaderFooter,
  type Block,
  type BreakBlock,
  type BreakKind,
  type CharacterFormatKey,
  type DocumentModel,
  type HeaderFooterKind,
  type HeaderFooterPart,
  type ListDefinition,
  type ListFormatting,
  type ListLevelDefinition,
  type ListMarkerStyle,
  type ListType,
  type Margins,
  type MeasurementUnit,
  type NamedStyle,
  type NoteKind,
  type NoteNumberFormat,
  type NoteRestart,
  type NoteSettings,
  type PageNumberFormat,
  type PageSize,
  type Paragraph,
  type ParagraphAlignment,
  type ParagraphDirection,
  type ParagraphFormatting,
  type PendingFormat,
  type Run,
  type RunFormatting,
  type Section,
  type SectionStart,
  type TabStop,
  type VerticalAlignment,
} from "./documentModel";
import { createNoteId } from "./notes";
import { cmToTwips, twipsToCm } from "./docx/units";
import { mergeParagraphFormatting, mergeRunFormatting } from "./docx/properties";
import {
  clampListLevel,
  isBulletStyle,
  listRunAround,
  markerStyleOf,
  markerType,
  planListLevels,
  type PlannedListLevel,
} from "./lists";
import {
  defaultDefinitionLevels,
  definedLevelType,
  headingLinkedLevels,
  headingStyleLinks,
  presetLevels,
} from "./listDefinitions";
import {
  builtInStyles,
  diffParagraphFormatting,
  namedStyleIdFor,
  NORMAL_STYLE_ID,
  paragraphStyleBase,
  resolveNamedStyle,
  splitRunFormatting,
  styleById,
  styleChainIncludes,
  styleDisplayName,
  type ResolvedNamedStyle,
} from "./styles";
import {
  blockLength,
  clampCaret,
  compareCaret,
  deleteBackward as deleteBackwardInPlace,
  deleteForward as deleteForwardInPlace,
  deleteRange as deleteRangeInPlace,
  formattingAtOffset,
  insertRun as insertRunInPlace,
  insertText as insertTextInPlace,
  paragraphAt,
  paragraphAtCaret,
  paragraphSlicesInRange,
  patchFormatting,
  replaceFormatting,
  replaceParagraphAt,
  runFormattingsInRange,
  setFormatting,
  splitParagraph as splitParagraphInPlace,
  type CaretPosition,
  type ParagraphAddress,
} from "./textEditing";

/** A selection in model coordinates. `anchor` is where the selection was
 * started, `focus` is the moving end (the caret); they are equal for a
 * collapsed selection (a plain caret) and may be in either document order. */
export interface ModelSelection {
  anchor: CaretPosition;
  focus: CaretPosition;
}

/** What every command returns: the new model state plus the selection after
 * the edit — collapsed to a caret at the natural post-edit position for
 * every command except `setFormat`/`toggleFormat`, which preserve the input
 * selection so the toolbar keeps reflecting the format's new state over the
 * same text (SPEC-012 AC-13, AC-14). */
export interface CommandResult {
  model: DocumentModel;
  selection: ModelSelection;
}

/** A collapsed selection (a caret) at a position. */
export function caretSelection(caret: CaretPosition): ModelSelection {
  return { anchor: caret, focus: caret };
}

/** Typed `null` seed for selection variables in `.vue` scripts, which espree
 * bars from writing the `ModelSelection | null` annotation themselves. */
export function noSelection(): ModelSelection | null {
  return null;
}

/** Typed `null` seed for command-result parameters in `.vue` scripts (same
 * espree constraint as noSelection). */
export function noCommandResult(): CommandResult | null {
  return null;
}

/** A command callback bound late to the current model/selection — how `.vue`
 * scripts hand paragraph commands to a shared dispatcher without writing
 * the function-type annotation espree can't parse. */
export type ParagraphCommand = (model: DocumentModel, selection: ModelSelection) => CommandResult;

/** Typed `null` seed for ParagraphCommand parameters in `.vue` scripts. */
export function noParagraphCommand(): ParagraphCommand | null {
  return null;
}

/** Typed `null` seed for pending-format refs in `.vue` scripts (same espree
 * constraint as noSelection). */
export function noPendingFormat(): PendingFormat | null {
  return null;
}

export function isCollapsed(selection: ModelSelection): boolean {
  return compareCaret(selection.anchor, selection.focus) === 0;
}

/** The selection's ends in document order, clamped into the model. */
export function orderedRange(
  model: DocumentModel,
  selection: ModelSelection,
): { start: CaretPosition; end: CaretPosition } {
  const a = clampCaret(model, selection.anchor);
  const b = clampCaret(model, selection.focus);
  return compareCaret(a, b) <= 0 ? { start: a, end: b } : { start: b, end: a };
}

// Clones the containers the primitives mutate (see the header comment);
// blocks and everything below them are shared with the input model. The
// clipboard commands (SPEC-016, lib/clipboard.ts) keep their own copy —
// exporting it here would trip commands.undo.spec.ts's completeness guard,
// which requires every export to be a command or a query.
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

/** Inserts text at the selection. A non-collapsed selection is replaced by
 * the inserted text (the "typing over a selection" rule); the text inherits
 * the formatting of the run it lands in, unless `formatOverride` forces
 * some of it on/off (SPEC-012's pending format, AC-9, AC-10). The caret
 * ends up after the inserted text. */
export function insertText(
  model: DocumentModel,
  selection: ModelSelection,
  text: string,
  formatOverride?: PendingFormat | null,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection)
    ? start
    : deleteRangeInPlace(next, start, end);
  return result(next, insertTextInPlace(next, base, text, formatOverride));
}

/** Inserts a forced line break (Shift+Enter) at the selection: a literal
 * "\n" inside the paragraph, which pre-wrap rendering displays as a new
 * line without creating a new paragraph. Replaces a non-collapsed
 * selection. */
export function insertLineBreak(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return insertText(model, selection, "\n");
}

/** Splits the paragraph at the selection in two (the Enter key), after
 * first deleting a non-collapsed selection. Both halves keep the source
 * paragraph's block-level formatting (alignment, indents, spacing, style
 * reference, heading level), so the new paragraph continues the old one's
 * paragraph formatting. The caret ends up at the start of the second
 * half. */
export function splitBlock(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection)
    ? start
    : deleteRangeInPlace(next, start, end);
  return result(next, splitParagraphInPlace(next, base));
}

// --- manual breaks and section breaks (SPEC-028) ----------------------------

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

/** Deletes the selection's range. Deleting across a paragraph boundary
 * merges the boundary paragraphs (the survivor keeps the first paragraph's
 * block formatting). A collapsed selection is a no-op that still returns a
 * fresh model state. The caret ends up at the deletion point. */
export function deleteRange(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  return result(next, deleteRangeInPlace(next, start, end));
}

/** Backspace: deletes a non-collapsed selection, else the character before
 * the caret; at a paragraph start it merges the paragraph into the previous
 * one. No-op at the very start of the document. */
export function deleteBackward(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  if (!isCollapsed(selection)) {
    return deleteRange(model, selection);
  }
  const next = cloneContainers(model);
  return result(next, deleteBackwardInPlace(next, selection.focus));
}

/** Forward delete: deletes a non-collapsed selection, else the character
 * after the caret; at a paragraph end it merges the next paragraph into
 * this one. No-op at the very end of the document. */
export function deleteForward(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  if (!isCollapsed(selection)) {
    return deleteRange(model, selection);
  }
  const next = cloneContainers(model);
  return result(next, deleteForwardInPlace(next, selection.focus));
}

// --- character formatting (SPEC-012) ----------------------------------------

/** Whether/how a character format is set: "on" only when every character
 * covered has it, "off" only when none do, "mixed" otherwise. */
export type FormatState = "on" | "off" | "mixed";

/** The formatting new text typed at a caret would inherit — the read-only
 * counterpart of insertText's inheritance rule (SPEC-012 AC-12). `{}` for a
 * table-atom caret, which addresses no paragraph (SPEC-022). */
export function caretFormatting(model: DocumentModel, caret: CaretPosition): RunFormatting {
  const c = clampCaret(model, caret);
  const paragraph = paragraphAtCaret(model, c);
  return paragraph ? formattingAtOffset(paragraph, c.offset) : {};
}

/** A character format's state for the toolbar (SPEC-012 AC-12..14). For a
 * collapsed selection (a caret) there is no text to inspect, so the state
 * reflects what the *next* typed character would carry: the run's
 * formatting the caret sits in, overridden by `pending` if given (the
 * caller's pending format, AC-9). For a real selection it is computed over
 * every run the selection covers. */
export function formatState(
  model: DocumentModel,
  selection: ModelSelection,
  key: CharacterFormatKey,
  pending: PendingFormat | null = null,
): FormatState {
  if (isCollapsed(selection)) {
    const base = caretFormatting(model, selection.focus)[key] === true;
    const value = pending && pending[key] !== undefined ? pending[key] : base;
    return value ? "on" : "off";
  }
  const { start, end } = orderedRange(model, selection);
  const formattings = runFormattingsInRange(model, start, end);
  if (formattings.length === 0) {
    return "off";
  }
  const onCount = formattings.filter((formatting) => formatting[key] === true).length;
  if (onCount === 0) {
    return "off";
  }
  return onCount === formattings.length ? "on" : "mixed";
}

/** Sets a character format across a real selection, splitting runs at the
 * boundaries and merging the result back down to the minimum run count
 * (SPEC-012 AC-7, AC-8, AC-20, AC-21). Unlike every other command, the
 * selection is not collapsed afterward — see `CommandResult`. A collapsed
 * selection is a no-op: collapsed-caret formatting is a pending format the
 * caller tracks outside the model (AC-9..11) and applies through
 * insertText's `formatOverride` instead. */
export function setFormat(
  model: DocumentModel,
  selection: ModelSelection,
  key: CharacterFormatKey,
  value: boolean,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  if (!isCollapsed(selection)) {
    setFormatting(next, start, end, key, value);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** Toggles a character format across a real selection: off or mixed turns
 * it fully on, on turns it fully off (SPEC-012 AC-16, AC-17). */
export function toggleFormat(
  model: DocumentModel,
  selection: ModelSelection,
  key: CharacterFormatKey,
): CommandResult {
  return setFormat(model, selection, key, formatState(model, selection, key) !== "on");
}

// --- font, colour, highlight, sub/superscript (SPEC-013) --------------------

/** Applies a run-format patch (font family/size, colour, highlight,
 * vertAlign — see PendingFormat) across a real selection, with the same
 * split/merge and selection-preserving behaviour as setFormat. A collapsed
 * selection is a no-op for the same reason setFormat's is: collapsed-caret
 * formatting is the caller's pending format, applied through insertText. */
export function applyFormat(
  model: DocumentModel,
  selection: ModelSelection,
  patch: PendingFormat,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  if (!isCollapsed(selection)) {
    patchFormatting(next, start, end, patch);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** The value-carrying run-format fields SPEC-013's toolbar reflects. */
export type RunFormatValueKey = "fontFamily" | "fontSizePt" | "color" | "highlight";

/** A value control's state for the toolbar (SPEC-013 AC): `mixed` when the
 * selection's runs disagree (the control renders blank/indeterminate),
 * otherwise the shared value — undefined meaning "not set", which the
 * caller resolves against the document defaults. For a collapsed selection
 * the value is what the next typed character would carry (the caret run's
 * formatting overridden by `pending`), mirroring formatState. */
export interface FormatValueState {
  mixed: boolean;
  value?: string | number | false;
}

// Highlight's explicit-off (`false`) and not-set (`undefined`) render
// identically (no highlight), so the control must not report a selection
// mixing only those two as indeterminate.
function normalizeFormatValue(
  key: RunFormatValueKey,
  value: string | number | false | undefined,
): string | number | false | undefined {
  return key === "highlight" && value === undefined ? false : value;
}

export function formatValue(
  model: DocumentModel,
  selection: ModelSelection,
  key: RunFormatValueKey,
  pending: PendingFormat | null = null,
): FormatValueState {
  if (isCollapsed(selection)) {
    const base = caretFormatting(model, selection.focus)[key];
    const value = pending && pending[key] !== undefined ? pending[key] : base;
    return { mixed: false, value: normalizeFormatValue(key, value) };
  }
  const { start, end } = orderedRange(model, selection);
  const formattings = runFormattingsInRange(model, start, end);
  if (formattings.length === 0) {
    return { mixed: false, value: normalizeFormatValue(key, undefined) };
  }
  const first = normalizeFormatValue(key, formattings[0][key]);
  for (const formatting of formattings) {
    if (normalizeFormatValue(key, formatting[key]) !== first) {
      return { mixed: true };
    }
  }
  return { mixed: false, value: first };
}

/** A sub/superscript button's tri-state, mirroring formatState: "on" only
 * when every covered character has exactly this vertical alignment. */
export function verticalAlignState(
  model: DocumentModel,
  selection: ModelSelection,
  which: VerticalAlignment,
  pending: PendingFormat | null = null,
): FormatState {
  if (isCollapsed(selection)) {
    const base = caretFormatting(model, selection.focus).vertAlign;
    const value = pending && pending.vertAlign !== undefined ? pending.vertAlign : base;
    return value === which ? "on" : "off";
  }
  const { start, end } = orderedRange(model, selection);
  const formattings = runFormattingsInRange(model, start, end);
  if (formattings.length === 0) {
    return "off";
  }
  const onCount = formattings.filter((formatting) => formatting.vertAlign === which).length;
  if (onCount === 0) {
    return "off";
  }
  return onCount === formattings.length ? "on" : "mixed";
}

/** Toggles sub/superscript across a real selection. Setting one always
 * replaces the other (they share the single vertAlign field), which is what
 * makes the two buttons mutually exclusive (SPEC-013 AC). */
export function toggleVerticalAlign(
  model: DocumentModel,
  selection: ModelSelection,
  which: VerticalAlignment,
): CommandResult {
  const on = verticalAlignState(model, selection, which) === "on";
  return applyFormat(model, selection, { vertAlign: on ? false : which });
}

/** What the colour picker shows for text with no explicit colour. */
export const DEFAULT_TEXT_COLOR = "#000000";

/** The four value controls' display state in one bag, with "not set"
 * resolved against the document defaults — shaped for a plain object ref
 * in DocumentView.vue, whose espree-parsed script cannot write the type
 * annotations itself. A mixed field renders blank/indeterminate:
 * fontFamily "", fontSizePt null, color "", highlight false, each beside
 * its explicit `*Mixed` flag (SPEC-013 AC). */
export interface ToolbarFormatValues {
  fontFamily: string;
  fontFamilyMixed: boolean;
  fontSizePt: number | null;
  fontSizeMixed: boolean;
  color: string;
  colorMixed: boolean;
  highlight: string | false;
  highlightMixed: boolean;
}

/** The initial/no-document state: everything blank, nothing mixed. */
export function emptyToolbarFormatValues(): ToolbarFormatValues {
  return {
    fontFamily: "",
    fontFamilyMixed: false,
    fontSizePt: null,
    fontSizeMixed: false,
    color: "",
    colorMixed: false,
    highlight: false,
    highlightMixed: false,
  };
}

export function toolbarFormatValues(
  model: DocumentModel,
  selection: ModelSelection,
  pending: PendingFormat | null = null,
): ToolbarFormatValues {
  const family = formatValue(model, selection, "fontFamily", pending);
  const size = formatValue(model, selection, "fontSizePt", pending);
  const color = formatValue(model, selection, "color", pending);
  const highlight = formatValue(model, selection, "highlight", pending);
  const highlightValue =
    highlight.mixed || highlight.value === undefined || typeof highlight.value === "number"
      ? false
      : highlight.value;
  return {
    fontFamily: family.mixed ? "" : String(family.value ?? model.defaultStyles.fontFamily),
    fontFamilyMixed: family.mixed,
    fontSizePt: size.mixed ? null : Number(size.value ?? model.defaultStyles.fontSizePt),
    fontSizeMixed: size.mixed,
    color: color.mixed ? "" : String(color.value ?? DEFAULT_TEXT_COLOR),
    colorMixed: color.mixed,
    highlight: highlightValue,
    highlightMixed: highlight.mixed,
  };
}

// --- paragraph formatting (SPEC-014) ----------------------------------------

/** The indent step the increase/decrease commands (and Tab/Shift+Tab) move
 * the start indent by (AC-6, AC-7) — Word's half inch. */
export const INDENT_STEP_CM = 1.27;

/** The explicit space-before the toggle writes on "Add" (AC-15) — the value
 * Word's spacing dropdown inserts. */
export const SPACE_BEFORE_TOGGLE_PT = 12;

// Edited lengths are snapped to the finest grid BOTH file formats can carry
// exactly, so a save + re-open reproduces the model bit-for-bit (AC-26):
// .docx stores whole twips / twentieths-of-a-point / 240ths-of-a-line, and
// the snapped doubles below are precisely what its reader computes back from
// those integers. .odt is no constraint — its writer (odf/units.ts) searches
// out an exact decimal for any double.
function snapCm(cm: number): number {
  return twipsToCm(cmToTwips(cm));
}

function snapPt(pt: number): number {
  return Math.round(pt * 20) / 20;
}

// Multiples snap to 20ths, which are both whole 240ths (the .docx grid) and
// whole percents (the .odt grid) — e.g. 1.15 = 276/240 = 115%.
function snapMultiple(value: number): number {
  return Math.round(value * 20) / 20;
}

/** A paragraph's address in the model — through `cell` for a paragraph
 * inside a table (SPEC-022; see textEditing.ParagraphAddress). */
export type BlockAddress = ParagraphAddress;

/** The paragraphs the selection touches (AC-4): the paragraph holding a
 * collapsed caret, or every paragraph the selection range intersects, in
 * document order — descending into table cells (SPEC-022): a selection
 * inside one cell touches that cell's paragraphs, and a top-level range
 * covering a table touches every paragraph of every cell. */
export function touchedBlocks(model: DocumentModel, selection: ModelSelection): BlockAddress[] {
  return paragraphSlicesInRange(model, selection.anchor, selection.focus).map(
    (slice) => slice.address,
  );
}

/** The first touched paragraph's formatting in document order — the label
 * source for the space-before toggle (AC-17) and what the toolbar/ruler
 * reflect. `{}` (never null) so `.vue` callers can read fields directly. */
export function firstTouchedFormatting(
  model: DocumentModel,
  selection: ModelSelection,
): ParagraphFormatting {
  const addresses = touchedBlocks(model, selection);
  if (addresses.length === 0) {
    return {};
  }
  return paragraphAt(model, addresses[0])?.formatting ?? {};
}

/** Typed empty-formatting seed for refs in `.vue` scripts (espree bars the
 * type annotation, same constraint as noSelection). */
export function emptyParagraphFormatting(): ParagraphFormatting {
  return {};
}

/** Runs a per-paragraph update over every touched paragraph (AC-5: each
 * command evaluates independently against each one). The updater returns
 * the paragraph's next formatting, or null for "leave this paragraph
 * untouched" — untouched paragraphs keep their object identity, which is
 * what the canvas's v-memo keys on. Like the character-format commands, the
 * selection is preserved (in document order) rather than collapsed, so the
 * toolbar keeps reflecting the touched range. */
function updateTouchedParagraphs(
  model: DocumentModel,
  selection: ModelSelection,
  update: (formatting: ParagraphFormatting) => ParagraphFormatting | null,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const updated = update(block.formatting ?? {});
    if (updated === null) {
      continue;
    }
    const replacement: Paragraph = { ...block };
    if (Object.keys(updated).length > 0) {
      replacement.formatting = updated;
    } else {
      delete replacement.formatting;
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** A partial paragraph-format update: an absent field is untouched, `null`
 * removes the field ("not specified on the paragraph" — a style value may
 * show through again), a value sets it. `direction: "ltr"` removes the
 * field rather than storing it — absent already means LTR, and storing it
 * would break writer round-trips (see ParagraphDirection). */
export interface ParagraphFormatPatch {
  alignment?: ParagraphAlignment;
  indentStartCm?: number | null;
  indentEndCm?: number | null;
  indentFirstLineCm?: number | null;
  spaceBeforePt?: number | null;
  spaceAfterPt?: number | null;
  lineSpacing?: number | null;
  lineSpacingPt?: number | null;
  lineSpacingRule?: "exact" | null;
  direction?: ParagraphDirection;
  /** SPEC-028 pagination flags. `false` removes keepNext/keepLines/
   * pageBreakBefore (absent already means off), while `true` removes
   * widowControl (absent means ON — see ParagraphFormatting), so the model
   * only ever stores the non-default value of each. */
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
}

type NumericParagraphKey =
  | "indentStartCm"
  | "indentEndCm"
  | "indentFirstLineCm"
  | "spaceBeforePt"
  | "spaceAfterPt"
  | "lineSpacing"
  | "lineSpacingPt";

const PATCH_SNAPPERS: Record<NumericParagraphKey, (value: number) => number> = {
  indentStartCm: snapCm,
  indentEndCm: snapCm,
  indentFirstLineCm: snapCm,
  spaceBeforePt: snapPt,
  spaceAfterPt: snapPt,
  lineSpacing: snapMultiple,
  lineSpacingPt: snapPt,
};

/** Applies a paragraph-format patch to every touched paragraph (AC-2, AC-5,
 * AC-13, AC-24). Numeric values are snapped to the cross-format grid (see
 * snapCm/snapPt above) so the edit survives a save + re-open exactly. */
export function applyParagraphFormat(
  model: DocumentModel,
  selection: ModelSelection,
  patch: ParagraphFormatPatch,
): CommandResult {
  return updateTouchedParagraphs(model, selection, (formatting) => {
    const nextFormatting: ParagraphFormatting = { ...formatting };
    if (patch.alignment !== undefined) {
      nextFormatting.alignment = patch.alignment;
    }
    if (patch.direction !== undefined) {
      if (patch.direction === "rtl") {
        nextFormatting.direction = "rtl";
      } else {
        delete nextFormatting.direction;
      }
    }
    if (patch.lineSpacingRule !== undefined) {
      if (patch.lineSpacingRule === null) {
        delete nextFormatting.lineSpacingRule;
      } else {
        nextFormatting.lineSpacingRule = patch.lineSpacingRule;
      }
    }
    for (const key of ["keepNext", "keepLines", "pageBreakBefore"] as const) {
      const value = patch[key];
      if (value === undefined) {
        continue;
      }
      if (value) {
        nextFormatting[key] = true;
      } else {
        delete nextFormatting[key];
      }
    }
    if (patch.widowControl !== undefined) {
      if (patch.widowControl) {
        delete nextFormatting.widowControl;
      } else {
        nextFormatting.widowControl = false;
      }
    }
    for (const key of Object.keys(PATCH_SNAPPERS) as NumericParagraphKey[]) {
      const value = patch[key];
      if (value === undefined) {
        continue;
      }
      if (value === null) {
        delete nextFormatting[key];
      } else {
        nextFormatting[key] = PATCH_SNAPPERS[key](value);
      }
    }
    return nextFormatting;
  });
}

/** Sets the alignment of every touched paragraph (AC-2). */
export function setAlignment(
  model: DocumentModel,
  selection: ModelSelection,
  alignment: ParagraphAlignment,
): CommandResult {
  return applyParagraphFormat(model, selection, { alignment });
}

/** Moves every touched paragraph's start indent by a signed delta, floored
 * at 0 (AC-6, AC-7). The one indent command — the toolbar buttons and
 * Tab/Shift+Tab all dispatch it with ±INDENT_STEP_CM, and it moves the
 * *logical* start indent, so there is no direction check: in an RTL
 * paragraph the right edge visibly moves (AC-25). A decrease on a paragraph
 * with no start indent leaves it untouched rather than writing an explicit
 * 0 it never had. */
export function adjustIndent(
  model: DocumentModel,
  selection: ModelSelection,
  deltaCm: number,
): CommandResult {
  return updateTouchedParagraphs(model, selection, (formatting) => {
    const current = formatting.indentStartCm ?? 0;
    if (current <= 0 && deltaCm < 0) {
      return null;
    }
    const next = snapCm(Math.max(0, current + deltaCm));
    if (next === formatting.indentStartCm) {
      return null;
    }
    return { ...formatting, indentStartCm: next };
  });
}

/** Whether the space-before toggle shows "Add" or "Remove" for a
 * paragraph's formatting, per the AC-15/AC-16 table: 0 or absent reads as
 * "Add space before paragraph". The label reflects only the first touched
 * paragraph (AC-17) — call with `firstTouchedFormatting`'s result. */
export function hasSpaceBefore(formatting: ParagraphFormatting): boolean {
  return (formatting.spaceBeforePt ?? 0) !== 0;
}

/** The space-before toggle's effect (AC-15..20): evaluated independently
 * against each touched paragraph's own spaceBeforePt — 0/absent gets an
 * explicit 12 pt, any non-zero value an explicit 0. Stateless by design:
 * "Remove" writes 0 rather than clearing the field (an absent field would
 * let a style's space-before reappear on reload, AC-21), and a removed
 * custom 6 pt comes back as 12, never 6 (AC-20). */
export function toggleSpaceBefore(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return updateTouchedParagraphs(model, selection, (formatting) => ({
    ...formatting,
    spaceBeforePt: hasSpaceBefore(formatting) ? 0 : SPACE_BEFORE_TOGGLE_PT,
  }));
}

/** What the paragraph pagination menu reflects (SPEC-028 AC-22..28): the
 * first touched paragraph's four pagination flags, with widow/orphan
 * control reading as ON when unset (the AC-28 default). */
export interface ParagraphPaginationState {
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
}

/** The initial/no-document state — a typed seed for `.vue` refs (espree
 * bars the annotation, same constraint as noSelection). */
export function emptyParagraphPaginationState(): ParagraphPaginationState {
  return { keepNext: false, keepLines: false, pageBreakBefore: false, widowControl: true };
}

export function paragraphPaginationState(
  model: DocumentModel,
  selection: ModelSelection,
): ParagraphPaginationState {
  const formatting = firstTouchedFormatting(model, selection);
  return {
    keepNext: formatting.keepNext === true,
    keepLines: formatting.keepLines === true,
    pageBreakBefore: formatting.pageBreakBefore === true,
    widowControl: formatting.widowControl !== false,
  };
}

/** Drops a stale preserved `w:sectPr` (SPEC-027): any command that changes
 * the page size or margins must not let the writer re-emit the source
 * .docx's original section properties over top of the new values — the
 * same rule tables.ts's clearPreservedTable/clearPreservedCell apply to
 * `w:tblPr`/`w:tcPr`. Absent on documents without one (.odt sources, blank
 * documents), which is also the common case, so callers only pay for the
 * copy when there is something to clear. */
function clearPreservedSectPr(model: DocumentModel): DocumentModel {
  if (model.sourceXml?.docxSectPr === undefined) {
    return model;
  }
  const sourceXml = { ...model.sourceXml };
  delete sourceXml.docxSectPr;
  return { ...model, sourceXml };
}

/** A section rewritten with new page geometry, its preserved sectPr dropped
 * so the writer regenerates it from the new values (SPEC-028) — the section
 * -level sibling of clearPreservedSectPr. */
function withSectionGeometry(
  section: Section,
  geometry: { pageSize?: PageSize; margins?: Margins },
): Section {
  const updated: Section = { ...section, ...geometry };
  delete updated.docxSectPr;
  return updated;
}

/** Sets the page's left/right margins (the ruler's margin markers, AC-23).
 * When the caret's section carries its own margins (SPEC-028), the drag
 * edits that section; otherwise it edits the document defaults, exactly as
 * before sections could diverge. */
export function setPageMargins(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { leftCm?: number; rightCm?: number },
): CommandResult {
  const { start } = orderedRange(model, selection);
  const section = model.sections[start.section];
  const margins = { ...(section.margins ?? model.defaultStyles.margins) };
  if (patch.leftCm !== undefined) {
    margins.leftCm = snapCm(Math.max(0, patch.leftCm));
  }
  if (patch.rightCm !== undefined) {
    margins.rightCm = snapCm(Math.max(0, patch.rightCm));
  }
  let next: DocumentModel;
  if (section.margins) {
    next = {
      ...model,
      sections: model.sections.map((entry, index) =>
        index === start.section ? withSectionGeometry(entry, { margins }) : entry,
      ),
    };
    if (start.section === model.sections.length - 1) {
      next = clearPreservedSectPr(next);
    }
  } else {
    next = clearPreservedSectPr({
      ...model,
      defaultStyles: { ...model.defaultStyles, margins },
    });
  }
  const range = orderedRange(next, selection);
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

// --- page setup (SPEC-027) ---------------------------------------------

/** What a Page setup apply targets (SPEC-028 AC-20/21): the caret's section
 * or the whole document. */
export type PageSetupScope = "thisSection" | "wholeDocument";

/** Typed pass-through for a scope coming from a `.vue` template's plain
 * string (see documentModel.asCharacterFormatKey for the espree constraint). */
export function asPageSetupScope(value: string): PageSetupScope {
  return value as PageSetupScope;
}

/** The Page setup dialog's OK (SPEC-027 AC-1..7, SPEC-028 AC-20/21): paper
 * size, orientation and the four margins are already resolved to a concrete
 * `PageSize`/`Margins` by the dialog (paper-size lookup and orientation
 * swap happen there — see lib/pageSetup.ts); this command just snaps them
 * onto the twip grid and commits. `applyTo` picks the target: the caret's
 * section (`"thisSection"`), or the document defaults (`"wholeDocument"`,
 * which also drops every section's own override so the whole document
 * genuinely follows the new values). `measurementUnit`, when given,
 * replaces the document's preference (AC-7). */
export function setPageSetup(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { pageSize: PageSize; margins: Margins; measurementUnit?: MeasurementUnit },
  applyTo: PageSetupScope = "wholeDocument",
): CommandResult {
  const pageSize: PageSize = {
    widthCm: snapCm(patch.pageSize.widthCm),
    heightCm: snapCm(patch.pageSize.heightCm),
  };
  const margins: Margins = {
    topCm: snapCm(patch.margins.topCm),
    rightCm: snapCm(patch.margins.rightCm),
    bottomCm: snapCm(patch.margins.bottomCm),
    leftCm: snapCm(patch.margins.leftCm),
  };
  let next: DocumentModel;
  if (applyTo === "thisSection") {
    const target = orderedRange(model, selection).start.section;
    next = {
      ...model,
      sections: model.sections.map((section, index) =>
        index === target ? withSectionGeometry(section, { pageSize, margins }) : section,
      ),
      measurementUnit: patch.measurementUnit ?? model.measurementUnit,
    };
    if (target === model.sections.length - 1) {
      next = clearPreservedSectPr(next);
    }
  } else {
    next = clearPreservedSectPr({
      ...model,
      defaultStyles: { ...model.defaultStyles, pageSize, margins },
      sections: model.sections.map((section) => {
        if (
          section.pageSize === undefined &&
          section.margins === undefined &&
          section.docxSectPr === undefined
        ) {
          return section;
        }
        const updated = { ...section };
        delete updated.pageSize;
        delete updated.margins;
        delete updated.docxSectPr;
        return updated;
      }),
      measurementUnit: patch.measurementUnit ?? model.measurementUnit,
    });
  }
  const { start, end } = orderedRange(next, selection);
  return { model: next, selection: { anchor: start, focus: end } };
}

// --- headers, footers and page numbers (SPEC-029) ---------------------------

/** Sets a section's header/footer options (SPEC-029 AC-11, AC-12, AC-22,
 * AC-23): `titlePage` is "Different first page", `evenAndOddHeaders` is
 * "Different odd and even pages", and `pageNumberStart` is "Start numbering
 * at" — `null` clears it back to "continue from the previous section".
 * `undefined` on any field leaves it unchanged, so a dialog can submit only
 * the fields it actually shows. The current selection passes through
 * untouched: these are section-wide settings, not a text edit. */
export function setHeaderFooterOptions(
  model: DocumentModel,
  selection: ModelSelection,
  sectionIndex: number,
  patch: { titlePage?: boolean; evenAndOddHeaders?: boolean; pageNumberStart?: number | null },
): CommandResult {
  const sections = model.sections.map((section, index) => {
    if (index !== sectionIndex) {
      return section;
    }
    const next: Section = { ...section };
    if (patch.titlePage !== undefined) {
      if (patch.titlePage) {
        next.titlePage = true;
      } else {
        delete next.titlePage;
      }
    }
    if (patch.evenAndOddHeaders !== undefined) {
      if (patch.evenAndOddHeaders) {
        next.evenAndOddHeaders = true;
      } else {
        delete next.evenAndOddHeaders;
      }
    }
    if (patch.pageNumberStart !== undefined) {
      if (patch.pageNumberStart === null) {
        delete next.pageNumberStart;
      } else {
        next.pageNumberStart = patch.pageNumberStart;
      }
    }
    delete next.docxSectPr;
    return next;
  });
  return { model: { ...model, sections }, selection };
}

/** Addresses a section's header/footer content as a throwaway single-section
 * document (section 0 holding exactly the header/footer's own blocks) so
 * every existing body-editing primitive (insertText, deleteBackward,
 * toggleFormat, table/image insertion, …) works unmodified against
 * header/footer content — SPEC-029 reuses them rather than duplicating a
 * second copy of every editing command. Absent content starts from one
 * empty paragraph, the same seed a blank document's own section starts
 * from. */
export function withHeaderFooterModel(
  model: DocumentModel,
  sectionIndex: number,
  part: HeaderFooterPart,
  kind: HeaderFooterKind,
): DocumentModel {
  const content = sectionHeaderFooter(model, sectionIndex, part, kind) ?? createEmptyHeaderFooterContent();
  return { ...model, sections: [{ blocks: content.blocks }] };
}

/** Folds a `withHeaderFooterModel` command result back into the real
 * document: the synthetic section's blocks become the section's
 * header/footer content for `kind`, while every other model-level change
 * the inner command made (new `media` entries from an inserted image,
 * updated `styles`, …) is kept from `result`. */
export function commitHeaderFooterModel(
  before: DocumentModel,
  sectionIndex: number,
  part: HeaderFooterPart,
  kind: HeaderFooterKind,
  result: DocumentModel,
): DocumentModel {
  const blocks = result.sections[0].blocks;
  const key: "headers" | "footers" = part === "header" ? "headers" : "footers";
  const sections = before.sections.map((section, index) => {
    if (index !== sectionIndex) {
      return section;
    }
    return { ...section, [key]: { ...section[key], [kind]: { blocks } } };
  });
  return { ...result, sections };
}

/** The header/footer kinds a section actually renders (SPEC-029 AC-9,
 * AC-20): `"default"` always, plus `"first"`/`"even"` when the section's
 * own toggles are on — the "Insert page number" control writes into every
 * one of them so the field shows up on every page (AC-20), not only the
 * ones already edited. */
function applicableHeaderFooterKinds(section: Section): HeaderFooterKind[] {
  const kinds: HeaderFooterKind[] = ["default"];
  if (section.titlePage) {
    kinds.push("first");
  }
  if (section.evenAndOddHeaders) {
    kinds.push("even");
  }
  return kinds;
}

/** The "Insert page number" control's settings (SPEC-029 AC-14..19):
 * `part`/`position` place it (header or footer, left/centre/right — the
 * paragraph's alignment); `format` is its numbering style; `pageXofY`
 * inserts "Page X of Y" (a PAGE field, literal " of ", a NUMPAGES field)
 * instead of a bare page number. */
export interface PageNumberInsertOptions {
  part: HeaderFooterPart;
  position: ParagraphAlignment;
  format: PageNumberFormat;
  pageXofY: boolean;
}

function pageNumberRuns(options: PageNumberInsertOptions): Run[] {
  if (options.pageXofY) {
    return [
      createFieldRun("page", options.format),
      { text: " of ", formatting: {} },
      createFieldRun("numpages", options.format),
    ];
  }
  return [createFieldRun("page", options.format)];
}

/** Inserts a page-number paragraph into a section's header or footer
 * (SPEC-029 AC-14..19): appended after any existing content so a repeated
 * insert never clobbers text already there. Creates the header/footer (and
 * every kind the section's own first-page/odd-even toggles call for) on
 * demand — AC's "does the control create one automatically" ambiguity
 * resolved in favour of "yes", the friendlier default. */
export function insertPageNumber(
  model: DocumentModel,
  sectionIndex: number,
  options: PageNumberInsertOptions,
): DocumentModel {
  const section = model.sections[sectionIndex];
  if (!section) {
    return model;
  }
  const paragraph: Paragraph = {
    type: "paragraph",
    runs: pageNumberRuns(options),
    formatting: { alignment: options.position },
  };
  const key: "headers" | "footers" = options.part === "header" ? "headers" : "footers";
  const updated = { ...section[key] };
  for (const kind of applicableHeaderFooterKinds(section)) {
    const existing = updated[kind];
    updated[kind] = existing ? { blocks: [...existing.blocks, paragraph] } : { blocks: [paragraph] };
  }
  const sections = model.sections.map((entry, index) =>
    index === sectionIndex ? { ...entry, [key]: updated } : entry,
  );
  return { ...model, sections };
}

/** `insertPageNumber` wrapped as a CommandResult (selection passes through
 * unchanged — inserting a page number never moves the body caret) so it
 * plugs into DocumentView's applyEdit the same way every other
 * dialog-mediated command does. */
export function insertPageNumberCommand(
  model: DocumentModel,
  selection: ModelSelection,
  sectionIndex: number,
  options: PageNumberInsertOptions,
): CommandResult {
  return { model: insertPageNumber(model, sectionIndex, options), selection };
}

// --- footnotes and endnotes (SPEC-031) --------------------------------------

/** Inserts a footnote/endnote reference mark at the selection (SPEC-031
 * AC-1, AC-3): a superscript marker run at the caret plus an empty note
 * content for it in `model.notes`. A non-collapsed selection is deleted
 * first (the typing-over rule); an atom caret (table/image/break edge) is
 * inert. The returned selection sits after the mark; the view then moves
 * the caret into the note's own area (AC-2, AC-4). Displayed numbers are
 * derived from mark order (lib/notes.ts), so every other mark renumbers by
 * construction (AC-8). Pass `noteId` to control the id (the readers do);
 * otherwise one is generated. */
export function insertNote(
  model: DocumentModel,
  selection: ModelSelection,
  kind: NoteKind,
  noteId: string = createNoteId(),
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  const paragraph = paragraphAt(next, base);
  if (!paragraph) {
    return result(next, base);
  }
  // The mark inherits the surrounding text's formatting (like insertText),
  // minus any link — a mark at a link's edge must not join the link.
  const formatting = { ...formattingAtOffset(paragraph, base.offset) };
  delete formatting.linkHref;
  delete formatting.linkAnchor;
  const caret = insertRunInPlace(next, base, createNoteRun(noteId, kind, formatting));
  if (!next.notes?.some((note) => note.noteId === noteId)) {
    next.notes = [...(next.notes ?? []), createEmptyNoteContent(noteId, kind)];
  }
  return result(next, caret);
}

/** Sets a note stream's numbering settings (SPEC-031 AC-17..20).
 * `undefined` leaves a field unchanged, so the dialog submits only what it
 * shows; the selection passes through — this is a document-wide setting,
 * not a text edit. */
export function setNoteSettings(
  model: DocumentModel,
  selection: ModelSelection,
  kind: NoteKind,
  patch: { numberFormat?: NoteNumberFormat; startAt?: number; restart?: NoteRestart },
): CommandResult {
  const key = kind === "footnote" ? "footnoteSettings" : "endnoteSettings";
  const settings: NoteSettings = { ...model[key] };
  if (patch.numberFormat !== undefined) {
    settings.numberFormat = patch.numberFormat;
  }
  if (patch.startAt !== undefined) {
    settings.startAt = Math.max(1, Math.floor(patch.startAt));
  }
  if (patch.restart !== undefined) {
    if (patch.restart === "continuous") {
      delete settings.restart;
    } else {
      settings.restart = patch.restart;
    }
  }
  return { model: { ...model, [key]: settings }, selection };
}

/** Addresses one note's content as a throwaway single-section document —
 * the withHeaderFooterModel seam reused for notes (SPEC-031 AC-6), so
 * every body-editing primitive works unmodified against a note's blocks.
 * A missing or empty note starts from one empty paragraph. */
export function withNoteModel(model: DocumentModel, noteId: string): DocumentModel {
  const note = model.notes?.find((entry) => entry.noteId === noteId);
  const blocks: Block[] =
    note && note.blocks.length > 0 ? note.blocks : [{ type: "paragraph", runs: [] }];
  return { ...model, sections: [{ blocks }] };
}

/** Folds a `withNoteModel` command result back into the real document: the
 * synthetic section's blocks become the note's content, the real sections
 * come back from `before`, and every other model-level change the inner
 * command made (new media, updated styles, …) is kept from `result`. */
export function commitNoteModel(
  before: DocumentModel,
  noteId: string,
  result: DocumentModel,
): DocumentModel {
  const blocks = result.sections[0].blocks;
  const notes = (before.notes ?? []).map((note) =>
    note.noteId === noteId ? { ...note, blocks } : note,
  );
  return { ...result, sections: before.sections, notes };
}

// --- bulleted and numbered lists (SPEC-018) ---------------------------------

/** Runs a per-paragraph list update over every touched paragraph, with the
 * same replace-don't-mutate and selection-preserving contract as
 * updateTouchedParagraphs. The updater returns the paragraph's next list
 * formatting, null to remove list formatting, or `undefined` for "leave
 * this paragraph untouched" (object identity kept for v-memo). */
function updateTouchedLists(
  model: DocumentModel,
  selection: ModelSelection,
  update: (paragraph: Paragraph) => ListFormatting | null | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const updated = update(block);
    if (updated === undefined) {
      continue;
    }
    const replacement: Paragraph = { ...block };
    if (updated === null) {
      delete replacement.list;
    } else {
      replacement.list = updated;
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** Whether every touched paragraph is a list item of the given type — the
 * AC-1/AC-2 toggle predicate, also the toolbar buttons' pressed state. */
export function allTouchedOfListType(
  model: DocumentModel,
  selection: ModelSelection,
  type: ListType,
): boolean {
  const addresses = touchedBlocks(model, selection);
  return addresses.every(
    (address) => paragraphAt(model, address)?.list?.type === type,
  );
}

/** The bulleted/numbered toolbar toggle (AC-1, AC-2): if every touched
 * paragraph is already a list item of `type`, list formatting is removed
 * from all of them; otherwise all of them become items of `type` — a
 * paragraph already in a list (of either type) keeps its level, a plain
 * paragraph joins at level 1. An explicit marker is dropped when the type
 * changes (it belonged to the other type's dropdown). */
export function toggleList(
  model: DocumentModel,
  selection: ModelSelection,
  type: ListType,
): CommandResult {
  const removeAll = allTouchedOfListType(model, selection, type);
  return updateTouchedLists(model, selection, (paragraph) => {
    if (removeAll) {
      return null;
    }
    const current = paragraph.list;
    if (!current) {
      return { type, level: 1 };
    }
    if (current.type === type) {
      return undefined;
    }
    const next: ListFormatting = { ...current, type };
    if (next.marker !== undefined && markerType(next.marker) !== type) {
      delete next.marker;
    }
    return next;
  });
}

/** Tab / Shift+Tab in a list item (AC-3, AC-4): moves every touched list
 * item's level by `delta`, clamped to 1..9 — a no-op at the bounds, and
 * non-list paragraphs in the selection are untouched. */
export function changeListLevel(
  model: DocumentModel,
  selection: ModelSelection,
  delta: number,
): CommandResult {
  return updateTouchedLists(model, selection, (paragraph) => {
    const current = paragraph.list;
    if (!current) {
      return undefined;
    }
    const level = clampListLevel(current.level + delta);
    if (level === current.level) {
      return undefined;
    }
    // The imported per-level indent belongs to the level the item is
    // leaving; the new level renders at the editor default.
    const next: ListFormatting = { ...current, level };
    delete next.indentCm;
    return next;
  });
}

/** A marker-style pick from a list button's dropdown (AC-12, AC-13):
 * applies the style — and the type it implies — to every item at the
 * current level of the current list, i.e. the contiguous run of list
 * paragraphs around the first touched block. The level is the first touched
 * list item's own. No-op when the selection is not in a list. */
export function setListMarker(
  model: DocumentModel,
  selection: ModelSelection,
  marker: ListMarkerStyle,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const result: CommandResult = { model: next, selection: { anchor: start, focus: end } };
  const startParagraph = paragraphAtCaret(next, start);
  const list = startParagraph?.list;
  // The containing block list: the cell's own blocks for a caret inside a
  // table (SPEC-022 — a cell is its own list context), else the section's.
  const table = next.sections[start.section].blocks[start.block];
  const container =
    start.cell && isTable(table)
      ? table.rows[start.cell.row].cells[start.cell.col].blocks
      : next.sections[start.section].blocks;
  const index = start.cell ? start.cell.block : start.block;
  const run = listRunAround(container, index);
  if (!list || !run) {
    return result;
  }
  const level = list.level;
  const type = markerType(marker);
  for (let blockIndex = run.first; blockIndex <= run.last; blockIndex++) {
    const block = container[blockIndex];
    if (block.type !== "paragraph" || block.list?.level !== level) {
      continue;
    }
    const address: BlockAddress = start.cell
      ? { section: start.section, block: start.block, cell: { ...start.cell, block: blockIndex } }
      : { section: start.section, block: blockIndex };
    replaceParagraphAt(next, address, { ...block, list: { ...block.list, type, marker } });
  }
  return result;
}

/** Removes list formatting from every touched paragraph without touching
 * its text — Backspace at the start of an item (AC-9) and Enter on an
 * empty level-1 item (AC-8). */
export function removeListFormat(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return updateTouchedLists(model, selection, (paragraph) =>
    paragraph.list ? null : undefined,
  );
}

/** The list-autoformat conversion (AC-14, AC-15): the caret paragraph —
 * which the caller has verified contains exactly a typed trigger prefix —
 * becomes an empty level-1 list item of `type`; the prefix text is
 * consumed. Recorded as its own undo step by the caller, so Ctrl/Cmd+Z
 * restores the literal prefix text (AC-16). */
export function applyListAutoformat(
  model: DocumentModel,
  selection: ModelSelection,
  type: ListType,
): CommandResult {
  const next = cloneContainers(model);
  const { start } = orderedRange(next, selection);
  const block = paragraphAtCaret(next, start);
  if (!block) {
    return result(next, start);
  }
  replaceParagraphAt(next, start, { ...block, runs: [], list: { type, level: 1 } });
  return result(next, { ...start, offset: 0 });
}

/** What the toolbar reflects about the selection's list state: the two
 * buttons' pressed states (every touched paragraph of that type, matching
 * the toggle predicate) and the first touched item's effective marker
 * style, "" when the selection is not in a list. */
export interface ToolbarListState {
  bulleted: boolean;
  numbered: boolean;
  marker: string;
}

/** The initial/no-document state — a typed seed for `.vue` refs (espree
 * bars the annotation, same constraint as noSelection). */
export function emptyToolbarListState(): ToolbarListState {
  return { bulleted: false, numbered: false, marker: "" };
}

export function toolbarListState(
  model: DocumentModel,
  selection: ModelSelection,
): ToolbarListState {
  const { start } = orderedRange(model, selection);
  const list = paragraphAtCaret(model, start)?.list;
  return {
    bulleted: allTouchedOfListType(model, selection, "bullet"),
    numbered: allTouchedOfListType(model, selection, "number"),
    marker: list ? markerStyleOf(list) : "",
  };
}

// --- multilevel lists and numbering control (SPEC-032) ----------------------

function newListObjectId(): string {
  return crypto.randomUUID();
}

/** The model's definitions array cloned for mutation (the listDefinitions
 * sibling of cloneContainers — the input model's array is never touched). */
function cloneDefinitions(model: DocumentModel): ListDefinition[] {
  return [...(model.listDefinitions ?? [])];
}

function definitionIn(model: DocumentModel, defId: string | undefined): ListDefinition | undefined {
  if (defId === undefined) {
    return undefined;
  }
  return model.listDefinitions?.find((def) => def.defId === defId);
}

/** The definition (and level) a paragraph style is linked to, scanning
 * newest-first so re-linking headings supersedes an older link (AC-13). */
function styleLinkFor(
  model: DocumentModel,
  styleId: string,
): { def: ListDefinition; level: number } | undefined {
  const defs = model.listDefinitions ?? [];
  for (let i = defs.length - 1; i >= 0; i--) {
    const level = defs[i].styleLinks?.[styleId];
    if (level !== undefined) {
      return { def: defs[i], level: clampListLevel(level) };
    }
  }
  return undefined;
}

/** Normalises the Define dialog's nine levels into canonical stored form:
 * defaults are elided (empty prefix/suffix, includeLevels 1, startAt 1,
 * left alignment, editor-default positions), lengths snap to the
 * cross-format grid — so an untouched save/re-open reproduces the
 * definition bit-for-bit. */
export function normalizeDefinitionLevels(levels: ListLevelDefinition[]): ListLevelDefinition[] {
  const out: ListLevelDefinition[] = [];
  for (let level = 1; level <= 9; level++) {
    const input = levels[level - 1] ?? { style: "decimal" };
    const cleaned: ListLevelDefinition = { style: input.style };
    if (input.prefix !== undefined && input.prefix !== "") {
      cleaned.prefix = input.prefix;
    }
    if (input.suffix !== undefined && input.suffix !== "") {
      cleaned.suffix = input.suffix;
    }
    const include = Math.round(input.includeLevels ?? 1);
    if (Number.isFinite(include) && include > 1) {
      cleaned.includeLevels = Math.min(include, level);
    }
    const startAt = Math.round(input.startAt ?? 1);
    if (Number.isFinite(startAt) && startAt >= 0 && startAt !== 1) {
      cleaned.startAt = startAt;
    }
    if (input.alignment !== undefined && input.alignment !== "left") {
      cleaned.alignment = input.alignment;
    }
    const defaultText = level * INDENT_STEP_CM;
    let textIndent = defaultText;
    if (input.textIndentCm !== undefined && Number.isFinite(input.textIndentCm)) {
      textIndent = snapCm(Math.max(0, input.textIndentCm));
      if (textIndent !== snapCm(defaultText)) {
        cleaned.textIndentCm = textIndent;
      } else {
        textIndent = defaultText;
      }
    }
    if (input.numberIndentCm !== undefined && Number.isFinite(input.numberIndentCm)) {
      const numberIndent = snapCm(Math.max(0, input.numberIndentCm));
      if (numberIndent !== snapCm(Math.max(0, textIndent - 0.635))) {
        cleaned.numberIndentCm = numberIndent;
      }
    }
    out.push(cleaned);
  }
  return out;
}

/** Every paragraph address in the document, in document order, descending
 * into table cells — the walk the whole-document list operations share. */
function allParagraphAddresses(model: DocumentModel): BlockAddress[] {
  return touchedBlocks(model, selectAll(model));
}

function sameAddress(a: BlockAddress, b: BlockAddress): boolean {
  if (a.section !== b.section || a.block !== b.block) {
    return false;
  }
  if (!a.cell || !b.cell) {
    return !a.cell && !b.cell;
  }
  return a.cell.row === b.cell.row && a.cell.col === b.cell.col && a.cell.block === b.cell.block;
}

/** Whether two addresses share a numbering context (SPEC-022's rule): both
 * top-level, or both inside the same table cell. */
function sameListContext(a: BlockAddress, b: BlockAddress): boolean {
  if (!a.cell && !b.cell) {
    return true;
  }
  if (!a.cell || !b.cell) {
    return false;
  }
  return (
    a.section === b.section &&
    a.block === b.block &&
    a.cell.row === b.cell.row &&
    a.cell.col === b.cell.col
  );
}

/** The addresses of the target's numbering context at or after (`forward`)
 * or strictly before, nearest first (`backward`) the target. */
function contextAddresses(
  model: DocumentModel,
  target: BlockAddress,
  direction: "forward" | "backward",
): BlockAddress[] {
  const all = allParagraphAddresses(model).filter((address) => sameListContext(address, target));
  const index = all.findIndex((address) => sameAddress(address, target));
  if (index === -1) {
    return [];
  }
  if (direction === "forward") {
    return all.slice(index);
  }
  return all.slice(0, index).reverse();
}

/** The BlockAddress a caret's paragraph lives at. */
function addressOfCaret(caret: CaretPosition): BlockAddress {
  if (caret.cell) {
    return { section: caret.section, block: caret.block, cell: { ...caret.cell } };
  }
  return { section: caret.section, block: caret.block };
}

function levelsFromPlanned(planned: PlannedListLevel[]): ListLevelDefinition[] {
  return planned.map((level, index) => {
    const def: ListLevelDefinition = { style: level.style };
    if (!isBulletStyle(level.style)) {
      def.suffix = ".";
    }
    if (level.startAt !== 1) {
      def.startAt = level.startAt;
    }
    if (cmToTwips(level.indentCm) !== (index + 1) * 720) {
      def.textIndentCm = level.indentCm;
    }
    return def;
  });
}

/** Upgrades the contiguous ad-hoc (SPEC-018) run around `target` into a
 * defined list: a definition planned from the run's own level styles plus a
 * fresh instance, so the numbering-control operations below always work on
 * instances. Returns the minted ids, or null when the target is not an
 * ad-hoc list paragraph. */
function upgradeAdHocRun(
  model: DocumentModel,
  target: BlockAddress,
): { defId: string; listId: string } | null {
  const block = paragraphAt(model, target);
  if (!block?.list || block.list.defId !== undefined) {
    return null;
  }
  const forward = contextAddresses(model, target, "forward");
  const backward = contextAddresses(model, target, "backward");
  const isAdHocItem = (address: BlockAddress) => {
    const paragraph = paragraphAt(model, address);
    return !!paragraph?.list && paragraph.list.defId === undefined;
  };
  // Contiguity within the context: the run extends while addresses are
  // consecutive blocks of the same block list and stay ad-hoc items.
  const run: BlockAddress[] = [target];
  let previous = target;
  for (const address of forward.slice(1)) {
    const at = address.cell ? address.cell.block : address.block;
    const prev = previous.cell ? previous.cell.block : previous.block;
    if (address.section !== previous.section || at !== prev + 1 || !isAdHocItem(address)) {
      break;
    }
    run.push(address);
    previous = address;
  }
  previous = target;
  for (const address of backward) {
    const at = address.cell ? address.cell.block : address.block;
    const prev = previous.cell ? previous.cell.block : previous.block;
    if (address.section !== previous.section || at !== prev - 1 || !isAdHocItem(address)) {
      break;
    }
    run.unshift(address);
    previous = address;
  }
  const items = run
    .map((address) => paragraphAt(model, address)?.list)
    .filter((list): list is ListFormatting => list !== undefined);
  const def: ListDefinition = {
    defId: newListObjectId(),
    levels: levelsFromPlanned(planListLevels(items)),
  };
  const listId = newListObjectId();
  model.listDefinitions = [...cloneDefinitions(model), def];
  for (const address of run) {
    const paragraph = paragraphAt(model, address);
    if (!paragraph?.list) {
      continue;
    }
    const list: ListFormatting = {
      type: paragraph.list.type,
      level: clampListLevel(paragraph.list.level),
      defId: def.defId,
      listId,
    };
    replaceParagraphAt(model, address, { ...paragraph, list });
  }
  return { defId: def.defId, listId };
}

/** Re-points `target` and every following item of its instance at new
 * instance ids — the shared core of restart/continue/set-value. `adoptDef`
 * (continue) also switches the definition, recomputing each item's type
 * from the adopted levels; `startAtFirst` (set value) restarts the target
 * item at an explicit ordinal. */
function reinstanceForward(
  model: DocumentModel,
  target: BlockAddress,
  oldKey: { defId: string; listId?: string },
  next: { defId: string; listId: string },
  startAtFirst?: number,
): void {
  const def = definitionIn(model, next.defId);
  let first = true;
  for (const address of contextAddresses(model, target, "forward")) {
    const paragraph = paragraphAt(model, address);
    const list = paragraph?.list;
    if (!paragraph || !list || list.defId !== oldKey.defId || list.listId !== oldKey.listId) {
      continue;
    }
    const level = clampListLevel(list.level);
    const updated: ListFormatting = {
      type: def ? definedLevelType(def, level) : list.type,
      level,
      defId: next.defId,
      listId: next.listId,
    };
    if (first && startAtFirst !== undefined) {
      updated.startAt = startAtFirst;
    }
    first = false;
    replaceParagraphAt(model, address, { ...paragraph, list: updated });
  }
}

/** Applies a gallery preset (AC-1..4): every touched paragraph becomes an
 * item of a NEW definition + instance minted from the preset — new, so
 * editing this list's levels later never touches another list started from
 * the same preset (AC-18). Paragraphs already in a list keep their level;
 * plain paragraphs join at level 1. */
export function applyMultilevelPreset(
  model: DocumentModel,
  selection: ModelSelection,
  presetKey: string,
): CommandResult {
  return applyDefinedList(model, selection, presetLevels(presetKey));
}

/** "Define new multilevel list…" OK (AC-5..12): with the selection in a
 * defined list, the list's own definition is updated in place — every item
 * of that list at every level re-renders from the new levels (AC-17), and
 * no other list changes (AC-18). Otherwise the touched paragraphs become a
 * new defined list with these levels (upgrading an ad-hoc SPEC-018 list or
 * starting one from plain paragraphs). */
export function applyListDefinition(
  model: DocumentModel,
  selection: ModelSelection,
  levels: ListLevelDefinition[],
): CommandResult {
  const cleaned = normalizeDefinitionLevels(levels);
  const { start } = orderedRange(model, selection);
  const target = paragraphAtCaret(model, start);
  const existing = definitionIn(model, target?.list?.defId);
  if (!existing) {
    return applyDefinedList(model, selection, cleaned);
  }
  const next = cloneContainers(model);
  const range = orderedRange(next, selection);
  const updated: ListDefinition = { ...existing, levels: cleaned };
  next.listDefinitions = cloneDefinitions(next).map((def) =>
    def.defId === existing.defId ? updated : def,
  );
  // A level's style may have crossed the bullet/number line: re-derive
  // every member item's type from the new levels.
  for (const address of allParagraphAddresses(next)) {
    const paragraph = paragraphAt(next, address);
    const list = paragraph?.list;
    if (!paragraph || !list || list.defId !== existing.defId) {
      continue;
    }
    const type = definedLevelType(updated, list.level);
    if (type === list.type) {
      continue;
    }
    replaceParagraphAt(next, address, { ...paragraph, list: { ...list, type } });
  }
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

function applyDefinedList(
  model: DocumentModel,
  selection: ModelSelection,
  levels: ListLevelDefinition[],
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const def: ListDefinition = { defId: newListObjectId(), levels };
  next.listDefinitions = [...cloneDefinitions(next), def];
  const listId = newListObjectId();
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const level = clampListLevel(block.list?.level ?? 1);
    const list: ListFormatting = {
      type: definedLevelType(def, level),
      level,
      defId: def.defId,
      listId,
    };
    replaceParagraphAt(next, address, { ...block, list });
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Number headings 1–3" (AC-13): creates a definition whose levels 1–3 are
 * linked to the built-in Heading 1–3 styles, then numbers every existing
 * heading paragraph. All linked headings share one instance (keyed by the
 * definition itself), so their numbering is continuous document-wide and
 * survives delete/move/demote (AC-29..31) — markers are derived, never
 * stored. Future paragraphs pick the link up through applyParagraphStyle. */
export function linkHeadingNumbering(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const def: ListDefinition = {
    defId: newListObjectId(),
    levels: headingLinkedLevels(),
    styleLinks: headingStyleLinks(),
  };
  next.listDefinitions = [...cloneDefinitions(next), def];
  for (const address of allParagraphAddresses(next)) {
    const paragraph = paragraphAt(next, address);
    if (!paragraph) {
      continue;
    }
    const namedId = namedStyleIdFor(next, paragraph.styleId);
    const level = namedId !== undefined ? def.styleLinks?.[namedId] : undefined;
    if (level === undefined) {
      continue;
    }
    const list: ListFormatting = {
      type: definedLevelType(def, level),
      level: clampListLevel(level),
      defId: def.defId,
      listId: def.defId,
    };
    replaceParagraphAt(next, address, { ...paragraph, list });
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Restart numbering at 1" (AC-14): the clicked item and every following
 * item of the same instance move to a fresh instance, whose counters start
 * from the definition's own start values. An ad-hoc SPEC-018 list is
 * upgraded to a defined one first, since only instances can restart
 * mid-run. */
export function restartListNumbering(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return reinstanceCommand(model, selection, undefined);
}

/** "Set numbering value…" (AC-16): like a restart, but the clicked item
 * carries an explicit start ordinal. */
export function setListNumberingValue(
  model: DocumentModel,
  selection: ModelSelection,
  value: number,
): CommandResult {
  const startAt = Math.max(0, Math.round(value));
  if (!Number.isFinite(startAt)) {
    return { model: cloneContainers(model), selection };
  }
  return reinstanceCommand(model, selection, startAt);
}

function reinstanceCommand(
  model: DocumentModel,
  selection: ModelSelection,
  startAtFirst: number | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const preserved = { model: next, selection: { anchor: start, focus: end } };
  const target = addressOfCaret(start);
  const paragraph = paragraphAt(next, target);
  if (!paragraph?.list) {
    return preserved;
  }
  if (paragraph.list.defId === undefined) {
    upgradeAdHocRun(next, target);
  }
  const list = paragraphAt(next, target)?.list;
  if (!list?.defId) {
    return preserved;
  }
  reinstanceForward(
    next,
    target,
    { defId: list.defId, listId: list.listId },
    { defId: list.defId, listId: newListObjectId() },
    startAtFirst,
  );
  return preserved;
}

/** "Continue previous numbering" (AC-15): the clicked item (and the rest of
 * its instance) rejoins the nearest preceding list's instance, continuing
 * its numbering across whatever separated them. Works from either side of
 * the SPEC-018 divide: an ad-hoc predecessor is upgraded to a defined list
 * first, and an ad-hoc target run adopts the predecessor's definition
 * outright. No-op when nothing numbered precedes the item. */
export function continueListNumbering(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const preserved = { model: next, selection: { anchor: start, focus: end } };
  const target = addressOfCaret(start);
  const targetParagraph = paragraphAt(next, target);
  if (!targetParagraph?.list) {
    return preserved;
  }
  // The nearest preceding list paragraph outside the target's own instance
  // (for an ad-hoc target: outside its contiguous run).
  const targetKey = targetParagraph.list.defId
    ? { defId: targetParagraph.list.defId, listId: targetParagraph.list.listId }
    : null;
  let predecessor: BlockAddress | undefined;
  let previous = target;
  let insideOwnRun = targetKey === null;
  for (const address of contextAddresses(next, target, "backward")) {
    const list = paragraphAt(next, address)?.list;
    if (!list) {
      insideOwnRun = false;
      previous = address;
      continue;
    }
    if (targetKey) {
      if (list.defId !== targetKey.defId || list.listId !== targetKey.listId) {
        predecessor = address;
        break;
      }
    } else {
      // Ad-hoc target: skip the target's own contiguous run, then take the
      // first list paragraph found.
      const at = address.cell ? address.cell.block : address.block;
      const prev = previous.cell ? previous.cell.block : previous.block;
      const contiguous = insideOwnRun && address.section === previous.section && at === prev - 1;
      if (contiguous && list.defId === undefined) {
        previous = address;
        continue;
      }
      predecessor = address;
      break;
    }
    previous = address;
  }
  if (!predecessor) {
    return preserved;
  }
  let adoptedList = paragraphAt(next, predecessor)?.list;
  if (adoptedList && adoptedList.defId === undefined) {
    upgradeAdHocRun(next, predecessor);
    adoptedList = paragraphAt(next, predecessor)?.list;
  }
  if (!adoptedList?.defId || !adoptedList.listId) {
    return preserved;
  }
  const adopted = { defId: adoptedList.defId, listId: adoptedList.listId };
  if (targetKey) {
    reinstanceForward(next, target, targetKey, adopted);
    return preserved;
  }
  // Ad-hoc target: the target and the rest of its contiguous run adopt the
  // predecessor's definition and instance directly.
  const def = definitionIn(next, adopted.defId);
  previous = target;
  for (const address of contextAddresses(next, target, "forward")) {
    const paragraph = paragraphAt(next, address);
    const list = paragraph?.list;
    const at = address.cell ? address.cell.block : address.block;
    const prev = previous.cell ? previous.cell.block : previous.block;
    const contiguous =
      sameAddress(address, target) || (address.section === previous.section && at === prev + 1);
    if (!paragraph || !list || list.defId !== undefined || !contiguous) {
      break;
    }
    const level = clampListLevel(list.level);
    replaceParagraphAt(next, address, {
      ...paragraph,
      list: {
        type: def ? definedLevelType(def, level) : list.type,
        level,
        defId: adopted.defId,
        listId: adopted.listId,
      },
    });
    previous = address;
  }
  return preserved;
}

/** The Define dialog's initial levels (a query): the selection's own
 * definition when it sits in a defined list, the levels planned from its
 * ad-hoc run when in a SPEC-018 list, else the decimal preset. */
export function selectionListDefinitionLevels(
  model: DocumentModel,
  selection: ModelSelection,
): ListLevelDefinition[] {
  const { start } = orderedRange(model, selection);
  const list = paragraphAtCaret(model, start)?.list;
  const def = definitionIn(model, list?.defId);
  if (def) {
    return structuredClone(def.levels);
  }
  if (list) {
    const target = addressOfCaret(start);
    const run = contextAddresses(model, target, "forward")
      .concat(contextAddresses(model, target, "backward"))
      .map((address) => paragraphAt(model, address)?.list)
      .filter((item): item is ListFormatting => !!item && item.defId === undefined);
    if (run.length > 0) {
      return levelsFromPlanned(planListLevels(run));
    }
  }
  return defaultDefinitionLevels();
}

// --- custom tab stops (SPEC-032 AC-19..25) ----------------------------------

/** Replaces the custom tab stops of every touched paragraph (the tab
 * dialog's OK, and the ruler's add/remove built on top of it). Stops are
 * snapped to the cross-format grid, sorted by position and deduplicated
 * (last one at a position wins, like Word); an empty list removes the
 * field. */
export function setTabStops(
  model: DocumentModel,
  selection: ModelSelection,
  stops: TabStop[],
): CommandResult {
  const byPosition = new Map<number, TabStop>();
  for (const stop of stops) {
    if (!Number.isFinite(stop.positionCm) || stop.positionCm < 0) {
      continue;
    }
    const positionCm = snapCm(stop.positionCm);
    const cleaned: TabStop = { positionCm, alignment: stop.alignment };
    if (stop.leader !== undefined) {
      cleaned.leader = stop.leader;
    }
    byPosition.set(positionCm, cleaned);
  }
  const normalized = [...byPosition.values()].sort((a, b) => a.positionCm - b.positionCm);
  return updateTouchedParagraphs(model, selection, (formatting) => {
    const nextFormatting: ParagraphFormatting = { ...formatting };
    if (normalized.length > 0) {
      nextFormatting.tabStops = normalized;
    } else {
      delete nextFormatting.tabStops;
    }
    return nextFormatting;
  });
}

/** The first touched paragraph's custom tab stops (a query) — what the tab
 * dialog opens with and the ruler paints. */
export function selectionTabStops(model: DocumentModel, selection: ModelSelection): TabStop[] {
  return (firstTouchedFormatting(model, selection).tabStops ?? []).map((stop) => ({ ...stop }));
}

// --- paragraph styles (SPEC-019) --------------------------------------------

/** A run's two layers relative to its old style base (see
 * styles.splitRunFormatting): the default-derived underlay and the genuine
 * direct-formatting diff. Recorded per run so a catalogue change can
 * re-bake the run as `underlay ⊕ new style ⊕ diff`. */
interface RunLayers {
  underlay: RunFormatting;
  diff: RunFormatting;
}

function runLayersFor(model: DocumentModel, block: Paragraph, base: ResolvedNamedStyle): RunLayers[] {
  return block.runs.map((run) => splitRunFormatting(model, run.formatting, base.runFormatting));
}

/** A paragraph rebuilt around a style resolution, with each run's layers
 * put back in order — underlay beneath the style, direct diff on top — and
 * the paragraph formatting as `resolved ⊕ diff`; the heading level follows
 * the resolution. Shared by apply (layers split against the old style) and
 * update (layers recorded before the catalogue changed). */
function restyledParagraph(
  block: Paragraph,
  resolved: ResolvedNamedStyle,
  paragraphDiff: ParagraphFormatting,
  runLayers: RunLayers[],
): Paragraph {
  const replacement: Paragraph = {
    ...block,
    runs: block.runs.map((run, index) => {
      const layers = runLayers[index] ?? { underlay: {}, diff: {} };
      return {
        ...run,
        formatting: mergeRunFormatting(
          mergeRunFormatting(layers.underlay, resolved.runFormatting),
          layers.diff,
        ),
      };
    }),
  };
  const formatting = mergeParagraphFormatting(resolved.paragraphFormatting, paragraphDiff);
  if (Object.keys(formatting).length > 0) {
    replacement.formatting = formatting;
  } else {
    delete replacement.formatting;
  }
  if (resolved.headingLevel !== undefined) {
    replacement.headingLevel = resolved.headingLevel;
  } else {
    delete replacement.headingLevel;
  }
  return replacement;
}

/** Applies a paragraph style to every paragraph the selection touches
 * (AC-3): each paragraph's baked formatting becomes the new style's
 * resolution with its *direct* formatting — the diff between its old baked
 * formatting and its old style's resolution — layered back on top, so
 * direct formatting survives the style change (AC-9). A built-in style
 * applied to a document that has a catalogue of its own is added to it, so
 * the picker lists it and a save declares it. */
export function applyParagraphStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  // Materialize the applied style into the model's catalogue: a document
  // without one gets the full built-in set (the picker's implicit list
  // becomes explicit), a document with its own gets the built-in definition
  // appended. This is what lets the writers declare the style — a paragraph
  // must never reference a style the saved file doesn't define (AC-15).
  if (!next.styles?.some((style) => style.id === styleId)) {
    const catalogue = next.styles ? [...next.styles] : builtInStyles();
    if (!catalogue.some((style) => style.id === styleId)) {
      const builtIn = builtInStyles().find((style) => style.id === styleId);
      if (builtIn) {
        catalogue.push(builtIn);
      }
    }
    next.styles = catalogue;
  }
  const resolved = resolveNamedStyle(next, styleId);
  // Heading numbering (SPEC-032 AC-13): applying a style linked to a list
  // definition makes the paragraph a defined item at the linked level (so a
  // new Heading 2 numbers itself, and re-styling Heading 1 -> Heading 2
  // demotes its number, AC-31); applying an unlinked style removes a
  // membership that only existed through such a link.
  const link = styleLinkFor(next, styleId);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const base = paragraphStyleBase(next, block);
    const replacement = restyledParagraph(
      block,
      resolved,
      diffParagraphFormatting(block.formatting ?? {}, base.paragraphFormatting),
      runLayersFor(next, block, base),
    );
    replacement.styleId = styleId;
    if (link) {
      replacement.list = {
        type: definedLevelType(link.def, link.level),
        level: link.level,
        defId: link.def.defId,
        listId: link.def.defId,
      };
    } else if (replacement.list?.defId !== undefined) {
      const memberOf = definitionIn(next, replacement.list.defId);
      if (memberOf?.styleLinks) {
        delete replacement.list;
      }
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Reset style" (AC-14): removes direct formatting from every touched
 * paragraph, leaving exactly what its style's resolution produces. A `.odt`
 * paragraph referencing an automatic style is re-pointed at the named style
 * underneath (the automatic style *was* the direct formatting). A heading
 * level the style chain doesn't set (a direct outline level) is kept — it
 * is structure, not appearance. */
export function resetParagraphStyle(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const namedId = namedStyleIdFor(next, block.styleId);
    const base = paragraphStyleBase(next, block);
    // Empty diffs: the underlay (the baked document defaults) survives the
    // reset, the direct formatting does not.
    const replacement = restyledParagraph(
      block,
      base,
      {},
      runLayersFor(next, block, base).map((layers) => ({ underlay: layers.underlay, diff: {} })),
    );
    if (base.headingLevel === undefined && block.headingLevel !== undefined) {
      replacement.headingLevel = block.headingLevel;
    }
    if (namedId !== undefined) {
      replacement.styleId = namedId;
    } else {
      delete replacement.styleId;
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Update <style> to match" (AC-12, AC-13): redefines `styleId` so its
 * resolution reproduces the current (first touched) paragraph's formatting,
 * then re-bakes every paragraph whose style chain passes through it —
 * keeping each one's direct formatting layered on the new resolution. A
 * document without a catalogue of its own gets one seeded from the
 * built-ins, so the picker keeps its full list. */
export function updateStyleFromSelection(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const preserved = { model: next, selection: { anchor: start, focus: end } };
  const target = styleById(next, styleId);
  const source = paragraphAtCaret(next, start);
  if (!target || !source) {
    return preserved;
  }

  // The style's new own formatting: the source paragraph's effective
  // formatting minus what the style's parents already provide. Run
  // formatting comes from the paragraph's first run; an empty paragraph
  // keeps the style's current run formatting.
  const parentResolved = resolveNamedStyle(next, target.basedOn);
  const ownParagraph = diffParagraphFormatting(source.formatting ?? {}, parentResolved.paragraphFormatting);
  const firstRun = source.runs[0];
  const ownRun = firstRun
    ? splitRunFormatting(next, firstRun.formatting, parentResolved.runFormatting).diff
    : (target.runFormatting ?? {});

  // Per-paragraph direct-formatting layers, recorded against the OLD
  // catalogue before it changes below.
  const affected: {
    address: BlockAddress;
    paragraphDiff: ParagraphFormatting;
    runLayers: RunLayers[];
  }[] = [];
  for (const address of touchedBlocks(next, selectAll(next))) {
    const block = paragraphAt(next, address);
    if (!block || !styleChainIncludes(next, block.styleId, styleId)) {
      continue;
    }
    const base = paragraphStyleBase(next, block);
    affected.push({
      address,
      paragraphDiff: diffParagraphFormatting(block.formatting ?? {}, base.paragraphFormatting),
      runLayers: runLayersFor(next, block, base),
    });
  }

  const updated: NamedStyle = { id: target.id };
  if (target.name !== undefined) {
    updated.name = target.name;
  }
  if (target.basedOn !== undefined) {
    updated.basedOn = target.basedOn;
  }
  if (Object.keys(ownParagraph).length > 0) {
    updated.paragraphFormatting = ownParagraph;
  }
  if (Object.keys(ownRun).length > 0) {
    updated.runFormatting = ownRun;
  }
  if (target.headingLevel !== undefined) {
    updated.headingLevel = target.headingLevel;
  }
  const catalogue = next.styles ? [...next.styles] : builtInStyles();
  const index = catalogue.findIndex((style) => style.id === styleId);
  if (index >= 0) {
    catalogue[index] = updated;
  } else {
    catalogue.push(updated);
  }
  next.styles = catalogue;

  // Re-bake with the new catalogue: resolution changed, the layers did not.
  for (const { address, paragraphDiff, runLayers } of affected) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    replaceParagraphAt(
      next,
      address,
      restyledParagraph(block, paragraphStyleBase(next, block), paragraphDiff, runLayers),
    );
  }
  return preserved;
}

/** What the style picker shows for the selection (AC-1): the first touched
 * paragraph's style — its resolved *named* style id and display name. */
export interface ToolbarStyleState {
  /** The named style id the current paragraph resolves to; "" when it
   * resolves to none (displays as Normal). */
  id: string;
  name: string;
}

/** The initial/no-document state — a typed seed for `.vue` refs (espree
 * bars the annotation, same constraint as noSelection). */
export function emptyToolbarStyleState(): ToolbarStyleState {
  return { id: "", name: "Normal" };
}

export function toolbarStyleState(
  model: DocumentModel,
  selection: ModelSelection,
): ToolbarStyleState {
  const { start } = orderedRange(model, selection);
  const paragraph = paragraphAtCaret(model, start);
  return {
    id: namedStyleIdFor(model, paragraph?.styleId) ?? "",
    name: styleDisplayName(model, paragraph?.styleId),
  };
}

// --- clear formatting and format painter (SPEC-020) -------------------------

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

// --- caret/selection queries (pure reads for keyboard navigation) ----------

/** The caret at the very start of the document. */
export function documentStart(): CaretPosition {
  return { section: 0, block: 0, offset: 0 };
}

/** The caret at the very end of the document (after the last block's last
 * character). */
export function documentEnd(model: DocumentModel): CaretPosition {
  const section = model.sections.length - 1;
  const blocks = model.sections[section].blocks;
  const block = blocks.length - 1;
  return { section, block, offset: blockLength(blocks[block]) };
}

/** The selection covering the whole document (Ctrl/Cmd+A). */
export function selectAll(model: DocumentModel): ModelSelection {
  return { anchor: documentStart(), focus: documentEnd(model) };
}

/** How many blocks the document has across all sections. */
export function blockCount(model: DocumentModel): number {
  let count = 0;
  for (const section of model.sections) {
    count += section.blocks.length;
  }
  return count;
}

/** The 1-based ordinal of a caret's block counting across all sections —
 * "paragraph 3 of 7" for screen-reader caret announcements. */
export function blockOrdinal(model: DocumentModel, caret: CaretPosition): number {
  const c = clampCaret(model, caret);
  let ordinal = c.block + 1;
  for (let s = 0; s < c.section; s += 1) {
    ordinal += model.sections[s].blocks.length;
  }
  return ordinal;
}
