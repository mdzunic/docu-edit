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
import type {
  DocumentModel,
  NamedStyle,
  Paragraph,
  ParagraphFormatting,
  RunFormatting,
} from "./documentModel";
import { mergeParagraphFormatting, mergeRunFormatting } from "./docx/properties";
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
