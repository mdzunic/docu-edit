// Paragraph-style resolution (SPEC-019), shared by the picker, the style
// commands and both writers. Framework-free on purpose (no Vue imports).
//
// The document model keeps *effective* formatting baked on every paragraph
// and run (the readers resolve style chains at open time — see
// docx/reader.ts). This module is what recovers the two layers from that:
// a style id resolves through its basedOn chain to the formatting the style
// alone would produce, and the difference between a paragraph's baked
// formatting and that resolution is its direct formatting. Styles the
// document does not define fall back to the built-in set below, which is
// also the whole picker list for a blank document.
import {
  HORIZONTAL_RULE_STYLE_ID,
  mergeParagraphFormatting,
  mergeRunFormatting,
  type DocumentModel,
  type NamedStyle,
  type Paragraph,
  type ParagraphAlignment,
  type ParagraphFormatting,
  type RunFormatting,
  type StyleType,
} from "./documentModel";
import { type BorderLineStyle, type CellBorders } from "./core/tables";
import { runStyle } from "./renderFormatting";

/** The id of the style every unstyled paragraph implicitly has. Word's
 * default-style id, which is also what AC "Normal is always available"
 * resolves to when a document defines no Normal of its own. */
export const NORMAL_STYLE_ID = "Normal";

// The built-in styles (AC-2): Normal, Title, Subtitle, Heading 1-6, Quote
// and Code, under Word's built-in style ids — which is exactly what makes an
// applied heading "the built-in heading style" in a saved .docx (AC-15,
// w:pStyle w:val="Heading1"). Formatting values follow Word/Google Docs
// conventions closely enough to read as those styles.
const BUILT_IN_STYLES: readonly NamedStyle[] = [
  { id: NORMAL_STYLE_ID, name: "Normal" },
  {
    id: "Title",
    name: "Title",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { fontSizePt: 26 },
    paragraphFormatting: { spaceAfterPt: 6 },
  },
  {
    id: "Subtitle",
    name: "Subtitle",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { fontSizePt: 15, color: "#666666" },
    paragraphFormatting: { spaceAfterPt: 8 },
  },
  {
    id: "Heading1",
    name: "Heading 1",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { bold: true, fontSizePt: 20 },
    paragraphFormatting: { spaceBeforePt: 20, spaceAfterPt: 6 },
    headingLevel: 1,
  },
  {
    id: "Heading2",
    name: "Heading 2",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { bold: true, fontSizePt: 16 },
    paragraphFormatting: { spaceBeforePt: 18, spaceAfterPt: 6 },
    headingLevel: 2,
  },
  {
    id: "Heading3",
    name: "Heading 3",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { bold: true, fontSizePt: 14 },
    paragraphFormatting: { spaceBeforePt: 16, spaceAfterPt: 4 },
    headingLevel: 3,
  },
  {
    id: "Heading4",
    name: "Heading 4",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { bold: true, fontSizePt: 12 },
    paragraphFormatting: { spaceBeforePt: 14, spaceAfterPt: 4 },
    headingLevel: 4,
  },
  {
    id: "Heading5",
    name: "Heading 5",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { bold: true, fontSizePt: 11 },
    paragraphFormatting: { spaceBeforePt: 12, spaceAfterPt: 4 },
    headingLevel: 5,
  },
  {
    id: "Heading6",
    name: "Heading 6",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { bold: true, italic: true, fontSizePt: 11 },
    paragraphFormatting: { spaceBeforePt: 12, spaceAfterPt: 4 },
    headingLevel: 6,
  },
  {
    id: "Quote",
    name: "Quote",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { italic: true, color: "#666666" },
    paragraphFormatting: { indentStartCm: 1.27, spaceBeforePt: 10, spaceAfterPt: 10 },
  },
  {
    id: "Code",
    name: "Code",
    basedOn: NORMAL_STYLE_ID,
    runFormatting: { fontFamily: "Courier New" },
  },
];

// SPEC-050: Markdown-only styles a `.md` reader adds to a document's *own*
// catalogue (never here — BUILT_IN_STYLES is every document's picker
// fallback, and SPEC-019 AC-2 pins its exact contents) when the document
// actually uses the construct. `mdBuiltInStyles()` seeds the full built-in
// set alongside them so that document's picker still offers Normal/Heading/…
// instead of just the Markdown-specific additions (see stylePickerEntries's
// "document's own styles instead of the built-in set" fallback rule).

/** An inline code span (`` `code` ``) is a run styled with this character
 * style rather than direct RunFormatting, so a save to Markdown can tell
 * "this run came from a code span" apart from a run a user simply set to a
 * monospace font (formatting Markdown cannot otherwise carry — see
 * opaque.ts's "runFormatting" drop kind). */
export const CODE_CHAR_STYLE_ID = "CodeChar";

const MD_CODE_CHAR_STYLE: NamedStyle = {
  id: CODE_CHAR_STYLE_ID,
  name: "Code",
  type: "character",
  runFormatting: { fontFamily: "Courier New" },
};

/** A thematic break (`---`) becomes an empty paragraph in this style — a
 * pure marker with no formatting of its own, the same convention BreakBlock
 * uses for a break that carries no content. */
const MD_HORIZONTAL_RULE_STYLE: NamedStyle = {
  id: HORIZONTAL_RULE_STYLE_ID,
  name: "Horizontal Rule",
  basedOn: NORMAL_STYLE_ID,
};

/** Fresh copies of every built-in style plus the two Markdown-only ones
 * above — the catalogue seed for a `.md` document that uses a thematic
 * break, an inline code span or a fenced code block (which adds its own
 * `Code:<language>` entries on top, based on "Code"). */
export function mdBuiltInStyles(): NamedStyle[] {
  return [...builtInStyles(), structuredClone(MD_CODE_CHAR_STYLE), structuredClone(MD_HORIZONTAL_RULE_STYLE)];
}

/** Fresh copies of the built-in styles — the catalogue seed when a
 * document without one of its own has a style updated, so the picker keeps
 * offering the full built-in set afterwards. */
export function builtInStyles(): NamedStyle[] {
  return BUILT_IN_STYLES.map((style) => structuredClone(style));
}

/** The built-in heading style id for a 1-based level (clamped to 1..6) —
 * what the Ctrl/Cmd+Alt+1..6 shortcuts apply (AC-5). */
export function headingStyleId(level: number): string {
  return `Heading${Math.min(Math.max(Math.round(level), 1), 6)}`;
}

/** The style definition an id resolves to: the document's own catalogue
 * entry when it has one, else the built-in definition, else undefined. */
export function styleById(model: DocumentModel, styleId: string): NamedStyle | undefined {
  return (
    model.styles?.find((style) => style.id === styleId) ??
    BUILT_IN_STYLES.find((style) => style.id === styleId)
  );
}

/** A style's basedOn chain, root ancestor first, with a cycle guard —
 * the resolution order both merge loops below share. */
function styleChain(model: DocumentModel, styleId: string | undefined): NamedStyle[] {
  const chain: NamedStyle[] = [];
  const seen = new Set<string>();
  let currentId = styleId;
  while (currentId && !seen.has(currentId)) {
    seen.add(currentId);
    const style = styleById(model, currentId);
    if (!style) {
      break;
    }
    chain.unshift(style);
    currentId = style.basedOn;
  }
  return chain;
}

/** The formatting a style alone produces — the basedOn chain merged root
 * ancestor first, with a percentage font size (ODF styles) resolved against
 * the inherited size, falling back to the document default. Contains only
 * properties the chain actually sets; document defaults do not appear. */
export interface ResolvedNamedStyle {
  paragraphFormatting: ParagraphFormatting;
  runFormatting: RunFormatting;
  headingLevel?: number;
}

export function resolveNamedStyle(
  model: DocumentModel,
  styleId: string | undefined,
): ResolvedNamedStyle {
  let paragraphFormatting: ParagraphFormatting = {};
  let runFormatting: RunFormatting = {};
  let headingLevel: number | undefined;
  for (const style of styleChain(model, styleId)) {
    paragraphFormatting = mergeParagraphFormatting(
      paragraphFormatting,
      style.paragraphFormatting ?? {},
    );
    const merged = mergeRunFormatting(runFormatting, style.runFormatting ?? {});
    if (style.fontSizePercent !== undefined && style.runFormatting?.fontSizePt === undefined) {
      const base = runFormatting.fontSizePt ?? model.defaultStyles.fontSizePt;
      merged.fontSizePt = (base * style.fontSizePercent) / 100;
    }
    runFormatting = merged;
    if (style.headingLevel !== undefined) {
      headingLevel = style.headingLevel;
    }
  }
  // "ltr" exists in style definitions only to override an ancestor's RTL
  // during the merge; resolved formatting holds "rtl" or nothing, like the
  // paragraphs themselves (see ParagraphDirection).
  if (paragraphFormatting.direction === "ltr") {
    paragraphFormatting = { ...paragraphFormatting };
    delete paragraphFormatting.direction;
  }
  return { paragraphFormatting, runFormatting, headingLevel };
}

/** The nearest style in the id's chain that carries a display name — for a
 * `.odt` paragraph referencing an automatic style, the named style it is
 * based on. Undefined when the chain never reaches a named style. */
export function namedStyleIdFor(model: DocumentModel, styleId: string | undefined): string | undefined {
  const seen = new Set<string>();
  let currentId = styleId;
  while (currentId && !seen.has(currentId)) {
    seen.add(currentId);
    const style = styleById(model, currentId);
    if (style?.name !== undefined) {
      return currentId;
    }
    currentId = style?.basedOn;
  }
  return undefined;
}

/** The display name the picker shows for a paragraph's style reference —
 * "Normal" for unstyled paragraphs and for references that resolve to no
 * named style (AC-1). */
export function styleDisplayName(model: DocumentModel, styleId: string | undefined): string {
  const namedId = namedStyleIdFor(model, styleId);
  if (namedId === undefined) {
    return "Normal";
  }
  return styleById(model, namedId)?.name ?? namedId;
}

/** One picker entry: the style plus its resolved formatting, from which the
 * caller renders the preview (AC-2's "each preview rendered in its own
 * style"). */
export interface StylePickerEntry {
  id: string;
  name: string;
  headingLevel?: number;
  /** Inline CSS for the entry's preview text. */
  previewStyle: Record<string, string>;
}

function previewStyleFor(resolved: ResolvedNamedStyle): Record<string, string> {
  const style = runStyle(resolved.runFormatting);
  // Cap the preview size so Title/Heading 1 don't blow up the menu; the
  // relative size difference stays visible.
  const size = resolved.runFormatting.fontSizePt;
  if (size !== undefined) {
    style.fontSize = `${Math.min(size, 21)}pt`;
  }
  return style;
}

/** The styles the picker lists (AC-2, AC-8): the document's own named
 * styles when it has any — with a built-in Normal prepended if the document
 * doesn't define one ("Normal is always available") — otherwise the
 * built-in set. */
export function stylePickerEntries(model: DocumentModel): StylePickerEntry[] {
  const documentStyles = (model.styles ?? []).filter(
    (style): style is NamedStyle & { name: string } => style.name !== undefined,
  );
  const listed: NamedStyle[] =
    documentStyles.length > 0 ? documentStyles : BUILT_IN_STYLES.slice();
  const entries: StylePickerEntry[] = [];
  const hasNormal = listed.some(
    (style) => style.id === NORMAL_STYLE_ID || style.name === "Normal",
  );
  if (!hasNormal) {
    entries.push({ id: NORMAL_STYLE_ID, name: "Normal", previewStyle: {} });
  }
  for (const style of listed) {
    const resolved = resolveNamedStyle(model, style.id);
    const entry: StylePickerEntry = {
      id: style.id,
      name: style.name ?? style.id,
      previewStyle: previewStyleFor(resolved),
    };
    if (resolved.headingLevel !== undefined) {
      entry.headingLevel = resolved.headingLevel;
    }
    entries.push(entry);
  }
  return entries;
}

/** Typed empty seed for style-entry arrays in `.vue` scripts (espree bars
 * the annotation — see documentModel.asCharacterFormatKey). */
export function noStyleEntries(): StylePickerEntry[] {
  return [];
}

/** Typed pass-through for a `.vue` component's untyped Array prop (the
 * same pattern as lib/opaque's asDropRecords). */
export function asStyleEntries(value: unknown[]): StylePickerEntry[] {
  return value as StylePickerEntry[];
}

// --- direct-formatting diffs ------------------------------------------------
//
// The difference between baked (effective) formatting and a style's resolved
// formatting is the paragraph's direct formatting: exactly the layer that
// must survive applying another style (AC-9) and that "Reset style" removes
// (AC-14). A property the base sets but the effective formatting lacks is an
// explicit *off* where the model can express one (the boolean toggles and
// highlight), and unrepresentable otherwise — those rare cases fall back to
// the incoming style's value.

const BOOLEAN_RUN_KEYS = ["bold", "italic", "underline", "strikethrough"] as const;
const VALUE_RUN_KEYS = [
  "fontFamily",
  "fontSizePt",
  "color",
  "shadingFill",
  "vertAlign",
] as const;

export function diffRunFormatting(effective: RunFormatting, base: RunFormatting): RunFormatting {
  const diff: RunFormatting = {};
  for (const key of BOOLEAN_RUN_KEYS) {
    if (effective[key] !== undefined && effective[key] !== base[key]) {
      diff[key] = effective[key];
    } else if (effective[key] === undefined && base[key] === true) {
      diff[key] = false;
    }
  }
  for (const key of VALUE_RUN_KEYS) {
    const value = effective[key];
    if (value !== undefined && value !== base[key]) {
      // The union of the value keys' types; per-key assignment back into the
      // same key is type-safe but TypeScript cannot see it through the loop.
      (diff as Record<string, unknown>)[key] = value;
    }
  }
  if (effective.highlight !== undefined && effective.highlight !== base.highlight) {
    diff.highlight = effective.highlight;
  } else if (effective.highlight === undefined && typeof base.highlight === "string") {
    diff.highlight = false;
  }
  return diff;
}

/** The two layers inside a run's baked formatting relative to a style
 * base: `underlay` — the document defaults the readers bake into every run
 * (a font family/size equal to defaultStyles that the base itself doesn't
 * set) — and `diff`, the genuine direct formatting. They sit on opposite
 * sides of a style: defaults underneath it (a style's own font size must
 * beat the baked default), direct formatting on top (AC-10). */
export function splitRunFormatting(
  model: DocumentModel,
  effective: RunFormatting,
  base: RunFormatting,
): { underlay: RunFormatting; diff: RunFormatting } {
  const underlay: RunFormatting = {};
  const remainder = { ...effective };
  if (base.fontFamily === undefined && effective.fontFamily === model.defaultStyles.fontFamily) {
    underlay.fontFamily = effective.fontFamily;
    delete remainder.fontFamily;
  }
  if (base.fontSizePt === undefined && effective.fontSizePt === model.defaultStyles.fontSizePt) {
    underlay.fontSizePt = effective.fontSizePt;
    delete remainder.fontSizePt;
  }
  return { underlay, diff: diffRunFormatting(remainder, base) };
}

const PARAGRAPH_KEYS = [
  "alignment",
  "indentStartCm",
  "indentEndCm",
  "indentFirstLineCm",
  "spaceBeforePt",
  "spaceAfterPt",
  "lineSpacing",
  "lineSpacingPt",
  "lineSpacingRule",
  "direction",
] as const;

export function diffParagraphFormatting(
  effective: ParagraphFormatting,
  base: ParagraphFormatting,
): ParagraphFormatting {
  const diff: ParagraphFormatting = {};
  for (const key of PARAGRAPH_KEYS) {
    const value = effective[key];
    if (value !== undefined && value !== base[key]) {
      (diff as Record<string, unknown>)[key] = value;
    }
  }
  return diff;
}

/** Whether `styleId`'s resolution chain passes through `targetId` — which
 * paragraphs "Update <style> to match" must re-bake (AC-13), including
 * paragraphs whose own style is *based on* the updated one. */
export function styleChainIncludes(
  model: DocumentModel,
  styleId: string | undefined,
  targetId: string,
): boolean {
  const seen = new Set<string>();
  let currentId = styleId;
  while (currentId && !seen.has(currentId)) {
    if (currentId === targetId) {
      return true;
    }
    seen.add(currentId);
    currentId = styleById(model, currentId)?.basedOn;
  }
  return false;
}

/** The resolved formatting the paragraph's direct formatting sits on: the
 * nearest *named* style in its chain (for `.odt`, through the automatic
 * style to the named parent — an automatic style IS direct formatting), or
 * nothing for an unstyled paragraph. */
export function paragraphStyleBase(model: DocumentModel, paragraph: Paragraph): ResolvedNamedStyle {
  const namedId = namedStyleIdFor(model, paragraph.styleId);
  if (namedId === undefined) {
    return { paragraphFormatting: {}, runFormatting: {} };
  }
  return resolveNamedStyle(model, namedId);
}

// --- custom styles manager (SPEC-042) ---------------------------------------

/** A style's family: its own `type`, or `"paragraph"` for every style saved
 * before SPEC-042 (and every built-in). */
export function styleTypeOf(style: NamedStyle): StyleType {
  return style.type ?? "paragraph";
}

/** A character style's resolved run formatting — `resolveNamedStyle`'s
 * basedOn-chain merge already handles this (the chain does not care what
 * family the styles along it belong to); this is a thin, family-specific
 * name for the call sites that bake a run's character-style layer (run
 * parsing, `applyCharacterStyle`). */
export function resolveCharacterStyle(model: DocumentModel, styleId: string | undefined): RunFormatting {
  return resolveNamedStyle(model, styleId).runFormatting;
}

/** Every style in the document's catalogue, falling back to the built-in
 * set for a document without one of its own (mirrors `stylePickerEntries`'
 * fallback) — the Styles pane's full listing (AC-1), including styles with
 * no `name` (an id-only stub the writer round-trips but the picker hides). */
export function allStyles(model: DocumentModel): NamedStyle[] {
  return model.styles && model.styles.length > 0 ? model.styles : builtInStyles();
}

/** Whether `name` (trimmed) collides with an existing style's user-visible
 * name, case-insensitively — the check "New style"/"Rename" reject inline
 * (AC-7). `excludeId` is the style being renamed, so it doesn't collide with
 * itself. */
export function styleNameCollides(
  model: DocumentModel,
  name: string,
  excludeId?: string,
): boolean {
  const trimmed = name.trim().toLowerCase();
  if (trimmed.length === 0) {
    return false;
  }
  return allStyles(model).some(
    (style) =>
      style.id !== excludeId &&
      style.name !== undefined &&
      style.name.trim().toLowerCase() === trimmed,
  );
}

/** Whether at least one paragraph, run or table in the document references
 * `styleId` directly (AC-2's "in use" badge — a direct reference, not
 * merely reachable through an unused descendant's basedOn chain). */
export function isStyleInUse(model: DocumentModel, styleId: string): boolean {
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "paragraph") {
        if (block.styleId === styleId || block.runs.some((run) => run.styleId === styleId)) {
          return true;
        }
      } else if (block.type === "table") {
        if (block.styleId === styleId) {
          return true;
        }
        for (const row of block.rows) {
          for (const cell of row.cells) {
            for (const paragraph of cell.blocks) {
              if (
                paragraph.styleId === styleId ||
                paragraph.runs.some((run) => run.styleId === styleId)
              ) {
                return true;
              }
            }
          }
        }
      }
    }
  }
  return false;
}

/** Every style of one family, in catalogue order — the Styles pane's
 * per-group listing (AC-1). */
export function stylesOfType(model: DocumentModel, type: StyleType): NamedStyle[] {
  return allStyles(model).filter((style) => styleTypeOf(style) === type);
}

/** One row the Styles pane renders (AC-1, AC-2, AC-15): a style's identity,
 * family (for grouping), in-use state and whether the editor can offer
 * Modify for it. */
export interface StylesPaneEntry {
  id: string;
  name: string;
  type: StyleType;
  inUse: boolean;
  unmodeled: boolean;
}

/** Every named style in the document, as pane rows (AC-1, AC-2, AC-15) —
 * every named style, including one the editor cannot fully model (AC-15). */
export function stylesPaneEntries(model: DocumentModel): StylesPaneEntry[] {
  return allStyles(model)
    .filter((style): style is NamedStyle & { name: string } => style.name !== undefined)
    .map((style) => ({
      id: style.id,
      name: style.name,
      type: styleTypeOf(style),
      inUse: isStyleInUse(model, style.id),
      unmodeled: style.unmodeled === true,
    }));
}

/** Typed empty seed for pane-entry arrays in `.vue` scripts (espree bars
 * the annotation — see documentModel.asCharacterFormatKey). */
export function noStylesPaneEntries(): StylesPaneEntry[] {
  return [];
}

/** Typed pass-through for a `.vue` component's untyped Array prop (the
 * same pattern as noStyleEntries/asStyleEntries above). */
export function asStylesPaneEntries(value: unknown[]): StylesPaneEntry[] {
  return value as StylesPaneEntry[];
}

/** A `{ id, name }` pair for a style picker/select in the New/Modify style
 * dialog (the "based on"/"style for following paragraph" options). */
export interface StyleOption {
  id: string;
  name: string;
}

/** Typed empty seed for `StyleOption[]` props in `.vue` scripts. */
export function noStyleOptions(): StyleOption[] {
  return [];
}

/** Typed pass-through for a `.vue` component's untyped Array prop. */
export function asStyleOptions(value: unknown[]): StyleOption[] {
  return value as StyleOption[];
}

/** Typed pass-through for a plain string array arriving as a `.vue`
 * component's untyped Array prop (e.g. existing style names for a
 * collision check). */
export function asStringArray(value: unknown[]): string[] {
  return value as string[];
}

/** StyleEditorDialog's plain-string/number draft shape — every field a
 * `<select>`/`<input>` can bind to, one flat object so the dialog's own
 * script never needs to build a typed payload itself (SPEC-042): building
 * an object incrementally by key (`payload.foo = ...`) is exactly the kind
 * of TS-only-adjacent code the espree constraint on `.vue` scripts makes
 * awkward, so that assembly lives here instead — plain TypeScript, no
 * restriction — and the dialog just calls `styleFormInputFromDraft`. */
export interface StyleEditorDraft {
  name: string;
  type: string;
  basedOn: string;
  next: string;
  fontFamily: string;
  fontSizePt: number | undefined;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color: string;
  alignment: string;
  listPresetKey: string;
  cellShadingFill: string;
  cellBorderStyle: string;
  cellBorderWidthPt: number;
  cellBorderColor: string;
}

/** Builds a `commands.createStyle`/`modifyStyle` input from the dialog's
 * draft (structurally compatible with `StyleFormInput`; not imported by
 * name to avoid a styles.ts <-> commands.ts import cycle). */
export function styleFormInputFromDraft(draft: StyleEditorDraft) {
  const runFormatting: RunFormatting = {};
  if (draft.fontFamily.trim()) {
    runFormatting.fontFamily = draft.fontFamily.trim();
  }
  if (draft.fontSizePt !== undefined && draft.fontSizePt > 0) {
    runFormatting.fontSizePt = draft.fontSizePt;
  }
  if (draft.bold) {
    runFormatting.bold = true;
  }
  if (draft.italic) {
    runFormatting.italic = true;
  }
  if (draft.underline) {
    runFormatting.underline = true;
  }
  if (draft.color.trim()) {
    runFormatting.color = draft.color.trim();
  }

  const borderStyle = draft.cellBorderStyle as BorderLineStyle;
  const cellBorders: CellBorders | undefined =
    draft.type === "table" && draft.cellBorderStyle
      ? {
          top: { style: borderStyle, widthPt: draft.cellBorderWidthPt, color: draft.cellBorderColor },
          bottom: { style: borderStyle, widthPt: draft.cellBorderWidthPt, color: draft.cellBorderColor },
          left: { style: borderStyle, widthPt: draft.cellBorderWidthPt, color: draft.cellBorderColor },
          right: { style: borderStyle, widthPt: draft.cellBorderWidthPt, color: draft.cellBorderColor },
        }
      : undefined;

  return {
    name: draft.name.trim(),
    type: draft.type as StyleType,
    basedOn: draft.basedOn || undefined,
    next: draft.next || undefined,
    runFormatting,
    paragraphFormatting:
      draft.type === "paragraph" && draft.alignment
        ? { alignment: draft.alignment as ParagraphAlignment }
        : undefined,
    listPresetKey: draft.type === "list" ? draft.listPresetKey : undefined,
    cellShadingFill:
      draft.type === "table" && draft.cellShadingFill.trim() ? draft.cellShadingFill.trim() : undefined,
    cellBorders,
  };
}

/** Reclassifies re-opened "list" styles (SPEC-042): neither `.docx` nor
 * `.odt` has a first-class "list style" family — Word's own List styles are
 * ordinary paragraph styles whose paragraphs also carry numbering — so a
 * custom list style this app writes round-trips as a plain paragraph style
 * (`w:type="paragraph"` / `style:family="paragraph"`) with no persisted
 * marker of its own. This recovers the grouping from *usage*: a non-built-in
 * paragraph style every one of whose paragraphs is a defined list item of
 * the same numbering definition is reclassified as `type: "list"` linked to
 * that definition, so it lists under "List" again and Modify offers the
 * numbering it actually drives. Called once after a document's styles,
 * paragraphs and list definitions are all parsed (see lib/openDocument.ts). */
export function inferListStyleTypes(model: DocumentModel): DocumentModel {
  if (!model.styles || model.styles.length === 0) {
    return model;
  }
  const builtInIds = new Set(builtInStyles().map((style) => style.id));
  const usage = new Map<string, Set<string | undefined>>();
  let sawAny = false;
  for (const section of model.sections) {
    for (const block of section.blocks) {
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : block.type === "paragraph"
            ? [block]
            : [];
      for (const paragraph of paragraphs) {
        if (!paragraph.styleId || builtInIds.has(paragraph.styleId)) {
          continue;
        }
        sawAny = true;
        const defIds = usage.get(paragraph.styleId) ?? new Set<string | undefined>();
        defIds.add(paragraph.list?.defId);
        usage.set(paragraph.styleId, defIds);
      }
    }
  }
  if (!sawAny) {
    return model;
  }
  let changed = false;
  const nextStyles = model.styles.map((style) => {
    if (styleTypeOf(style) !== "paragraph") {
      return style;
    }
    const defIds = usage.get(style.id);
    if (!defIds || defIds.size !== 1) {
      return style;
    }
    const [defId] = [...defIds];
    if (defId === undefined) {
      return style;
    }
    changed = true;
    return { ...style, type: "list" as const, listDefinitionId: defId };
  });
  return changed ? { ...model, styles: nextStyles } : model;
}

/** The style id applied to a new paragraph typed after one in `styleId`
 * (AC-3/AC-4's "style for following paragraph"); undefined keeps the same
 * style, the common case for body-text styles. */
export function nextStyleIdFor(model: DocumentModel, styleId: string | undefined): string | undefined {
  if (styleId === undefined) {
    return undefined;
  }
  return styleById(model, styleId)?.next;
}
