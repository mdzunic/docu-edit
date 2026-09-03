// --- paragraph formatting (SPEC-014) ----------------------------------------
import { orderedRange, type CommandResult, type ModelSelection } from "./commands.editing";
import type { ModelColor } from "./core/theme";
import { cmToTwips, twipsToCm } from "./core/units";
import {
  cloneContainers,
  type DocumentModel,
  type Paragraph,
  type ParagraphAlignment,
  type ParagraphBorders,
  type ParagraphDirection,
  type ParagraphFormatting,
} from "./documentModel";
import {
  bordersPaintAnyEdge,
  explicitNoBorders,
  normaliseParagraphBorders,
  paintedShadingFill,
  snapParagraphBorders,
  NO_SHADING_FILL,
} from "./paragraphBorders";
import { inheritedParagraphPaint } from "./styles";
import {
  paragraphAt,
  paragraphSlicesInRange,
  replaceParagraphAt,
  type ParagraphAddress,
} from "./textEditing";

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
// snapCm is shared with commands.pageSetup and commands.multilevel (the same
// grid applies to page geometry and list indents) — deliberately NOT
// re-exported by the commands.ts barrel; see the barrel's header.
export function snapCm(cm: number): number {
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
  return firstTouchedParagraph(model, selection)?.formatting ?? {};
}

/** The first touched paragraph itself, for the callers that need more than
 * its formatting — the Borders and shading dialog seeds from what the
 * paragraph *paints*, and only the paragraph carries the styleId that
 * resolves its inherited box (SPEC-074 AC-17). */
export function firstTouchedParagraph(
  model: DocumentModel,
  selection: ModelSelection,
): Paragraph | undefined {
  const addresses = touchedBlocks(model, selection);
  return addresses.length === 0 ? undefined : (paragraphAt(model, addresses[0]) ?? undefined);
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
// Shared with commands.multilevel (setTabStops) — deliberately NOT
// re-exported by the commands.ts barrel; see the barrel's header.
export function updateTouchedParagraphs(
  model: DocumentModel,
  selection: ModelSelection,
  update: (formatting: ParagraphFormatting, paragraph: Paragraph) => ParagraphFormatting | null,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const updated = update(block.formatting ?? {}, block);
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

// --- borders and shading (SPEC-074) ------------------------------------------

/** What the Borders and shading dialog applies. `null` is the dialog's "off"
 * — `borders: null` is its None preset, `shadingFill: null` its cleared Fill
 * checkbox ("no fill", which is distinct from `"#ffffff"` — AC-10). Off is
 * not the same as unset: see applyParagraphBorders. */
export interface ParagraphBorderPatch {
  borders?: ParagraphBorders | null;
  /** A literal "#rrggbb" or a theme reference (SPEC-077) — stored verbatim so
   * the paragraph's fill follows the document theme. */
  shadingFill?: ModelColor | null;
}

/** Applies borders and/or shading to every paragraph the selection touches
 * (SPEC-074 AC-11, using SPEC-014's "touched by the selection" rule — the
 * same `updateTouchedParagraphs` walk every other paragraph format uses).
 *
 * Padding is snapped to whole points because that is the finest grid OOXML's
 * `@w:space` can carry, so the edit survives a save + re-open exactly (the
 * same reasoning as snapCm/snapPt above).
 *
 * "Off" is spelled two different ways, and which one this writes depends on
 * the paragraph's style (AC-17). Borders and shading inherit from the style
 * until the paragraph sets its own, so on a paragraph whose style draws a
 * box, dropping the field would just re-inherit that box and leave the
 * dialog's None preset a silent no-op; there, off is an explicit value that
 * is present and paints nothing. Where the style contributes nothing there
 * is nothing to override, and off is simply the absent field — which keeps
 * an ordinary unbordered paragraph free of an empty husk. */
export function applyParagraphBorders(
  model: DocumentModel,
  selection: ModelSelection,
  patch: ParagraphBorderPatch,
): CommandResult {
  return updateTouchedParagraphs(model, selection, (formatting, paragraph) => {
    const next: ParagraphFormatting = { ...formatting };
    // What this paragraph would paint with no value of its own — the thing
    // an "off" has to override, if anything.
    const inherited = inheritedParagraphPaint(model, paragraph.styleId);
    if (patch.borders !== undefined) {
      const painted =
        patch.borders === null ? undefined : normaliseParagraphBorders(snapParagraphBorders(patch.borders));
      if (painted !== undefined) {
        next.borders = painted;
      } else if (bordersPaintAnyEdge(inherited.borders)) {
        next.borders = explicitNoBorders();
      } else {
        delete next.borders;
      }
    }
    if (patch.shadingFill !== undefined) {
      if (patch.shadingFill !== null) {
        next.shadingFill = patch.shadingFill;
      } else if (paintedShadingFill(inherited) !== undefined) {
        next.shadingFill = NO_SHADING_FILL;
      } else {
        delete next.shadingFill;
      }
    }
    return next;
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
