// --- character formatting (SPEC-012) ----------------------------------------
import { resolveColor, resolveFont, sameModelColor } from "./core/theme";
import type { DocumentTheme, ModelColor, ThemeColorRef } from "./core/theme";
import {
  isCollapsed,
  orderedRange,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import {
  cloneContainers,
  documentThemeOf,
  type CharacterFormatKey,
  type DocumentModel,
  type PendingFormat,
  type RunFormatting,
  type VerticalAlignment,
} from "./documentModel";
import {
  clampCaret,
  formattingAtOffset,
  paragraphAtCaret,
  patchFormatting,
  runFormattingsInRange,
  setFormatting,
  type CaretPosition,
} from "./textEditing";

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

/** Sets the document's default language (SPEC-035 AC-9): the language
 * selector's "no selection" case — a selection targets `applyFormat`'s
 * `{lang}` patch instead (AC-9.1). Touches no run/paragraph content, so no
 * cloneContainers: the selection passes through unchanged. */
export function setDocumentLanguage(
  model: DocumentModel,
  selection: ModelSelection,
  language: string,
): CommandResult {
  return { model: { ...model, defaultStyles: { ...model.defaultStyles, language } }, selection };
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
  value?: string | number | false | ThemeColorRef | { theme: "major" | "minor" };
}

/** What one run-format value field may hold after SPEC-077's widening. */
type RunFormatValue = FormatValueState["value"] | undefined;

// Highlight's explicit-off (`false`) and not-set (`undefined`) render
// identically (no highlight), so the control must not report a selection
// mixing only those two as indeterminate.
function normalizeFormatValue(key: RunFormatValueKey, value: RunFormatValue): RunFormatValue {
  return key === "highlight" && value === undefined ? false : value;
}

/** Formatting-value equality for the mixed-state fold (SPEC-077 AC-8): theme
 * references compare structurally through sameModelColor — two runs carrying
 * the same accent are one value, and a reference is never equal to its
 * resolved literal. Scalars keep `===`. */
function sameFormatValue(a: RunFormatValue, b: RunFormatValue): boolean {
  const aIsRef = typeof a === "object" && a !== null;
  const bIsRef = typeof b === "object" && b !== null;
  if (aIsRef || bIsRef) {
    return aIsRef && bIsRef && sameModelColor(a as ModelColor, b as ModelColor);
  }
  return a === b;
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
    if (!sameFormatValue(normalizeFormatValue(key, formatting[key]), first)) {
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
  // The toolbar controls display literals, so references resolve here —
  // through resolveColor/resolveFont alone (SPEC-077 AC-7).
  const theme = documentThemeOf(model);
  const highlightValue =
    highlight.mixed || highlight.value === undefined || typeof highlight.value === "number" || highlight.value === false
      ? false
      : displayFormatValue(highlight.value, theme);
  const familyValue = family.value ?? model.defaultStyles.fontFamily;
  const colorValue = color.value ?? DEFAULT_TEXT_COLOR;
  return {
    fontFamily: family.mixed ? "" : displayFormatValue(familyValue, theme),
    fontFamilyMixed: family.mixed,
    fontSizePt: size.mixed ? null : Number(size.value ?? model.defaultStyles.fontSizePt),
    fontSizeMixed: size.mixed,
    color: color.mixed ? "" : displayFormatValue(colorValue, theme),
    colorMixed: color.mixed,
    highlight: highlightValue,
    highlightMixed: highlight.mixed,
  };
}

/** A format value as the toolbar displays it: a literal string as-is, a
 * theme reference resolved against the document's theme, anything else
 * (number/false/undefined — not this control's shape) blank. */
function displayFormatValue(value: RunFormatValue, theme: DocumentTheme): string {
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "object" && value !== null) {
    return value.theme === "major" || value.theme === "minor"
      ? resolveFont(value, theme)
      : resolveColor(value, theme);
  }
  return "";
}
