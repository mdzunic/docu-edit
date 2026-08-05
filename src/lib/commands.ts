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
  isTable,
  type CharacterFormatKey,
  type DocumentModel,
  type ListFormatting,
  type ListMarkerStyle,
  type ListType,
  type Paragraph,
  type ParagraphAlignment,
  type ParagraphDirection,
  type ParagraphFormatting,
  type PendingFormat,
  type RunFormatting,
  type VerticalAlignment,
} from "./documentModel";
import { cmToTwips, twipsToCm } from "./docx/units";
import {
  clampListLevel,
  listRunAround,
  markerStyleOf,
  markerType,
} from "./lists";
import {
  blockLength,
  clampCaret,
  compareCaret,
  deleteBackward as deleteBackwardInPlace,
  deleteForward as deleteForwardInPlace,
  deleteRange as deleteRangeInPlace,
  formattingAtOffset,
  insertText as insertTextInPlace,
  paragraphAt,
  paragraphAtCaret,
  paragraphSlicesInRange,
  patchFormatting,
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

/** Sets the page's left/right margins (the ruler's margin markers, AC-23).
 * Margins live on defaultStyles, not a paragraph, but the update flows
 * through the same immutable command shape so applyEdit persists it. */
export function setPageMargins(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { leftCm?: number; rightCm?: number },
): CommandResult {
  const margins = { ...model.defaultStyles.margins };
  if (patch.leftCm !== undefined) {
    margins.leftCm = snapCm(Math.max(0, patch.leftCm));
  }
  if (patch.rightCm !== undefined) {
    margins.rightCm = snapCm(Math.max(0, patch.rightCm));
  }
  const next: DocumentModel = {
    ...model,
    defaultStyles: { ...model.defaultStyles, margins },
  };
  const { start, end } = orderedRange(next, selection);
  return { model: next, selection: { anchor: start, focus: end } };
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
