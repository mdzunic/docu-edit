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
import type {
  CharacterFormatKey,
  DocumentModel,
  PendingFormat,
  RunFormatting,
  VerticalAlignment,
} from "./documentModel";
import {
  blockLength,
  clampCaret,
  compareCaret,
  deleteBackward as deleteBackwardInPlace,
  deleteForward as deleteForwardInPlace,
  deleteRange as deleteRangeInPlace,
  formattingAtOffset,
  insertText as insertTextInPlace,
  patchFormatting,
  runFormattingsInRange,
  setFormatting,
  splitParagraph as splitParagraphInPlace,
  type CaretPosition,
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
// blocks and everything below them are shared with the input model.
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
 * counterpart of insertText's inheritance rule (SPEC-012 AC-12). */
export function caretFormatting(model: DocumentModel, caret: CaretPosition): RunFormatting {
  const c = clampCaret(model, caret);
  return formattingAtOffset(model.sections[c.section].blocks[c.block], c.offset);
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
