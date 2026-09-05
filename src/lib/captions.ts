// Captions (SPEC-055): the model logic behind Insert → Caption — caption
// detection, the per-label `seq` numbering, the positional caption↔object
// binding, and the insert command. Framework-free like every lib module, and
// — like lib/links.ts and lib/toc.ts — a command module OUTSIDE commands.ts
// built on documentModel.ts's shared cloneContainers, so commands.undo.spec's
// completeness guard (scoped to commands.ts's own exports) never accounts for
// these.
//
// A caption is an ordinary paragraph carrying the built-in Caption style whose
// runs are, in order: a label run ("Figure "), a `seq` field run (the number),
// a separator run (": ") and the user's editable text run — the label omitted
// when "Exclude label from caption" is set. There is no Caption block kind, so
// a caption is findable, counted, spell-checked and styleable for free.
import {
  cloneContainers,
  createFieldSpecRun,
  isChart,
  isImage,
  isTable,
  type Block,
  type DocumentModel,
  type NumberFormat,
  type Paragraph,
  type Run,
  type RunFormatting,
} from "./documentModel";
import { formatOrdinal } from "./core/lists";
import { CAPTION_STYLE, CAPTION_STYLE_ID } from "./styles";
import { clampCaret, paragraphAt, type ParagraphAddress } from "./textEditing";
import { caretSelection, type CommandResult } from "./commands";

/** The three built-in caption labels the dialog offers (SPEC-055 AC-5). Custom
 * labels are added by "New label…" and live only in the captions the user
 * creates — there is no separate label registry. English only, as scoped. */
export const BUILT_IN_CAPTION_LABELS = ["Figure", "Table", "Equation"];

/** A freshly inserted caption's default separator (Word's own). */
export const DEFAULT_CAPTION_SEPARATOR = ": ";

/** Where a caption sits relative to its object (SPEC-055 AC-7). */
export type CaptionPosition = "above" | "below";

/** The caption dialog's settings (SPEC-055 AC-2, AC-4..10). */
export interface CaptionOptions {
  label: string;
  position: CaptionPosition;
  separator: string;
  excludeLabel: boolean;
  format: NumberFormat;
}

/** A blank caption dialog's defaults for a given label. */
export function defaultCaptionOptions(label = "Figure"): CaptionOptions {
  return {
    label,
    position: "below",
    separator: DEFAULT_CAPTION_SEPARATOR,
    excludeLabel: false,
    format: "decimal",
  };
}

/** Whether "New label…" accepts `name` (SPEC-055 AC-11, AC-12): non-empty and
 * unique case-insensitively among the built-in and any custom labels. */
export function isValidNewLabel(name: string, existing: string[]): boolean {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    return false;
  }
  const lower = trimmed.toLowerCase();
  return !existing.some((label) => label.toLowerCase() === lower);
}

/** The `seq` field run of a paragraph, or null — the run whose number a
 * caption computes. */
export function captionSeqRun(paragraph: Paragraph): Run | null {
  return paragraph.runs.find((run) => run.field?.kind === "seq") ?? null;
}

/** Whether a block is a caption paragraph (SPEC-055): a Caption-styled
 * paragraph carrying a `seq` field run. Both conditions matter — the style
 * alone is a plain styled paragraph, the field alone a stray sequence. */
export function isCaptionParagraph(block: Block | undefined): block is Paragraph {
  return (
    block?.type === "paragraph" &&
    block.styleId === CAPTION_STYLE_ID &&
    captionSeqRun(block) !== null
  );
}

/** Whether a block is an object a caption can describe (SPEC-055): an image,
 * chart or table. Shapes are excluded, and equations do not exist yet. */
export function isCaptionableObject(block: Block | undefined): boolean {
  return isImage(block) || isChart(block) || isTable(block);
}

/** A caption's label, taken from its `seq` field. */
export function captionLabelOf(paragraph: Paragraph): string {
  return captionSeqRun(paragraph)?.field?.kind === "seq"
    ? (captionSeqRun(paragraph)!.field as { label: string }).label
    : "";
}

/** The text a caption displays: its runs concatenated, with the `seq` field's
 * cached number substituted for its one-character placeholder — e.g.
 * "Figure 3: Revenue by quarter". This is what a table of figures lists and
 * what a `full` cross-reference renders. */
export function captionDisplayText(paragraph: Paragraph): string {
  let text = "";
  for (const run of paragraph.runs) {
    text += run.field?.kind === "seq" ? run.field.cached : run.text;
  }
  return text;
}

/** The caption text past its label and number — the part the user typed
 * ("Revenue by quarter"): everything after the separator run that follows the
 * `seq` field. A `textOnly` cross-reference and the caption text itself. */
export function captionUserText(paragraph: Paragraph): string {
  const seqIndex = paragraph.runs.findIndex((run) => run.field?.kind === "seq");
  if (seqIndex < 0) {
    return captionDisplayText(paragraph);
  }
  // runs after the seq field are the separator run then the user's text runs.
  return paragraph.runs
    .slice(seqIndex + 2)
    .map((run) => run.text)
    .join("");
}

/** What a caption cross-reference renders for each caption `show` format
 * (SPEC-055 Reference formats): `full` the whole caption, `labelAndNumber`
 * "Figure 3", `numberOnly` "3", `textOnly` the user's text. */
export function captionRefText(
  paragraph: Paragraph,
  show: "full" | "labelAndNumber" | "numberOnly" | "textOnly",
): string {
  const seqRun = captionSeqRun(paragraph);
  const number = seqRun?.field?.kind === "seq" ? seqRun.field.cached : "";
  const seqIndex = paragraph.runs.findIndex((run) => run.field?.kind === "seq");
  const label = seqIndex > 0 ? paragraph.runs.slice(0, seqIndex).map((run) => run.text).join("") : "";
  switch (show) {
    case "full":
      return captionDisplayText(paragraph);
    case "labelAndNumber":
      return `${label}${number}`.trim();
    case "numberOnly":
      return number;
    case "textOnly":
      return captionUserText(paragraph);
  }
}

/** The derived anchor id of a caption — deterministic in its label and
 * per-label ordinal, so a table-of-figures entry and a cross-reference resolve
 * to the same caption after a reflow (both re-derive it), the same scheme
 * headingAnchor uses for TOC entries. Letters/digits/underscore only, valid as
 * a `w:anchor` bookmark name and an ODF `#ref`. */
export function captionAnchor(label: string, ordinal: number): string {
  const slug = label.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 30);
  return `_cap_${slug || "label"}_${ordinal}`;
}

/** One caption found in the document body (SPEC-055): its address, paragraph,
 * label, 1-based ordinal within its label, rendered text and derived anchor.
 * Body-only — headers, footers and footnotes (which are not in
 * `model.sections[].blocks`) are excluded from numbering and the table of
 * figures per §Numbering. */
export interface CaptionInfo {
  address: ParagraphAddress;
  paragraph: Paragraph;
  label: string;
  ordinal: number;
  text: string;
  anchor: string;
}

/** Every caption in the document body, in document order, numbered per label
 * (SPEC-055 §Numbering). */
export function listCaptions(model: DocumentModel): CaptionInfo[] {
  const counters = new Map<string, number>();
  const found: CaptionInfo[] = [];
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (!isCaptionParagraph(block)) {
        return;
      }
      const label = captionLabelOf(block);
      const key = label.toLowerCase();
      const ordinal = (counters.get(key) ?? 0) + 1;
      counters.set(key, ordinal);
      found.push({
        address: { section: sectionIndex, block: blockIndex },
        paragraph: block,
        label,
        ordinal,
        text: captionDisplayText(block),
        anchor: captionAnchor(label, ordinal),
      });
    });
  });
  return found;
}

/** Renumbers every caption's `seq` field per label, in document order,
 * starting at 1 (SPEC-055 AC-15..17, §Numbering) — recomputed over the whole
 * document rather than incrementally from the edit point. Returns the input by
 * reference when no number changed, so it is cheap to run on every edit
 * (typing never reorders captions) and leaves header/footer/footnote captions
 * untouched (they are not in `sections[].blocks`). */
export function renumberCaptions(model: DocumentModel): DocumentModel {
  const counters = new Map<string, number>();
  const updates: Array<{ section: number; block: number; run: number; cached: string }> = [];
  let changed = false;
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (block.type !== "paragraph") {
        return;
      }
      block.runs.forEach((run, runIndex) => {
        if (run.field?.kind !== "seq") {
          return;
        }
        const key = run.field.label.toLowerCase();
        const ordinal = (counters.get(key) ?? 0) + 1;
        counters.set(key, ordinal);
        const cached = formatOrdinal(run.field.format, ordinal);
        if (cached !== run.field.cached) {
          changed = true;
        }
        updates.push({ section: sectionIndex, block: blockIndex, run: runIndex, cached });
      });
    });
  });
  if (!changed) {
    return model;
  }
  const next = cloneContainers(model);
  for (const update of updates) {
    const paragraph = next.sections[update.section].blocks[update.block] as Paragraph;
    const runs = [...paragraph.runs];
    const run = runs[update.run];
    if (run.field?.kind === "seq") {
      runs[update.run] = { ...run, field: { ...run.field, cached: update.cached } };
    }
    next.sections[update.section].blocks[update.block] = { ...paragraph, runs };
  }
  return next;
}

/** The object a caption is bound to, or null (SPEC-055 §binding is
 * positional): the object in the paragraph immediately BEFORE it (a caption
 * positioned below), else immediately AFTER it (positioned above). Top-level
 * blocks only — the binding is between body blocks. */
export function objectForCaption(
  model: DocumentModel,
  caption: ParagraphAddress,
): ParagraphAddress | null {
  if (caption.cell) {
    return null;
  }
  const blocks = model.sections[caption.section]?.blocks ?? [];
  const before = blocks[caption.block - 1];
  if (isCaptionableObject(before)) {
    return { section: caption.section, block: caption.block - 1 };
  }
  const after = blocks[caption.block + 1];
  if (isCaptionableObject(after)) {
    return { section: caption.section, block: caption.block + 1 };
  }
  return null;
}

/** The caption bound to an object, or null (SPEC-055 AC-20..22, the reverse of
 * objectForCaption): the caption paragraph immediately AFTER the object (a
 * below caption), else immediately BEFORE it (an above caption). */
export function captionForObject(
  model: DocumentModel,
  object: ParagraphAddress,
): ParagraphAddress | null {
  const blocks = model.sections[object.section]?.blocks ?? [];
  const after = blocks[object.block + 1];
  if (isCaptionParagraph(after)) {
    return { section: object.section, block: object.block + 1 };
  }
  const before = blocks[object.block - 1];
  if (isCaptionParagraph(before)) {
    return { section: object.section, block: object.block - 1 };
  }
  return null;
}

/** The runs of a fresh caption (SPEC-055 AC-4): a label run (unless excluded),
 * a `seq` field run, a separator run and an empty editable text run. The
 * caption's baked run formatting comes from the Caption style so it renders
 * distinctly without a catalogue lookup at paint time. */
function captionRuns(options: CaptionOptions): { runs: Run[]; textOffset: number } {
  const runFormatting: RunFormatting = { ...CAPTION_STYLE.runFormatting };
  const runs: Run[] = [];
  let textOffset = 0;
  if (!options.excludeLabel) {
    const labelText = `${options.label} `;
    runs.push({ text: labelText, formatting: { ...runFormatting } });
    textOffset += labelText.length;
  }
  runs.push(
    createFieldSpecRun({ kind: "seq", label: options.label, format: options.format, cached: formatOrdinal(options.format, 1) }, {
      ...runFormatting,
    }),
  );
  textOffset += 1;
  runs.push({ text: options.separator, formatting: { ...runFormatting } });
  textOffset += options.separator.length;
  runs.push({ text: "", formatting: { ...runFormatting } });
  return { runs, textOffset };
}

/** Inserts a caption into a nested editing surface's synthetic model — the
 * header/footer band (SPEC-055 AC-37; SPEC-029 already puts images and tables
 * there): the same paragraph insertCaption builds, but with NO renumber pass.
 * Captions outside the body are excluded from the per-label sequence
 * (§Numbering, AC-38) — they keep the cached number they were created with
 * (the format's first ordinal, seeded by captionRuns), and the body counter
 * never sees them. The caret ends up in the editable text run. */
export function insertSurfaceCaption(
  model: DocumentModel,
  object: ParagraphAddress,
  options: CaptionOptions,
): CommandResult {
  const next = cloneContainers(model);
  const blocks = next.sections[object.section].blocks;
  const { runs, textOffset } = captionRuns(options);
  const caption: Paragraph = {
    type: "paragraph",
    runs,
    styleId: CAPTION_STYLE_ID,
    formatting: { ...CAPTION_STYLE.paragraphFormatting },
  };
  const at = options.position === "below" ? object.block + 1 : object.block;
  blocks.splice(at, 0, caption);
  return {
    model: next,
    selection: caretSelection({ section: object.section, block: at, offset: textOffset }),
  };
}

/** Inserts a caption paragraph for the object at `object`, positioned above or
 * below it per `options` (SPEC-055 AC-3, AC-4), then renumbers every caption
 * so the new one shows its real number. The caret ends up in the editable text
 * run. */
export function insertCaption(
  model: DocumentModel,
  object: ParagraphAddress,
  options: CaptionOptions,
): CommandResult {
  const inserted = insertSurfaceCaption(model, object, options);
  return { model: renumberCaptions(inserted.model), selection: inserted.selection };
}

/** Resolves an object's model address from a caption's, kept next to
 * paragraphAt so callers importing captions.ts do not also reach into
 * textEditing for the trivial narrowing. */
export function captionParagraphAt(model: DocumentModel, address: ParagraphAddress): Paragraph | undefined {
  const block = paragraphAt(model, address);
  return block ?? undefined;
}

/** Deletes both a captioned object and its bound caption in one edit (SPEC-055
 * AC-20, AC-21 — "delete its caption too"), then renumbers the remaining
 * captions. Both are top-level blocks in the same section; the higher index is
 * spliced first so the lower one stays valid. */
export function deleteObjectWithCaption(
  model: DocumentModel,
  object: ParagraphAddress,
  caption: ParagraphAddress,
): CommandResult {
  const next = cloneContainers(model);
  const blocks = next.sections[object.section].blocks;
  const indices = [object.block, caption.block].toSorted((a, b) => b - a);
  for (const index of indices) {
    blocks.splice(index, 1);
  }
  const renumbered = renumberCaptions(next);
  const caret = clampCaret(renumbered, {
    section: object.section,
    block: Math.min(object.block, caption.block),
    offset: 0,
  });
  return { model: renumbered, selection: caretSelection(caret) };
}
