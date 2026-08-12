// Mail-merge chip serialisation (SPEC-046): the one grammar shared by a
// chip's display label and its on-disk form in both formats — see "Chip
// serialisation" in specs/046-template-gallery-and-mail-merge.md. Kept
// framework-free like the rest of lib/ so both readers/writers and the
// editor's paint-time substitution can share it.
import { createChipRun, type MergeChipKind, type MergeFieldSpec, type Run, type RunFormatting } from "./documentModel";

/** A field or array name: letters/digits/underscore, first character not a
 * digit, dots traverse nested objects. Anything else is refused at
 * insertion and, on read, left as whatever it was rather than becoming a
 * chip. */
export const CHIP_NAME_PATTERN = /^[A-Za-z_][A-Za-z0-9_.]*$/;

export function isValidChipName(name: string): boolean {
  return CHIP_NAME_PATTERN.test(name);
}

/** The chip's canonical text — its display label and its on-disk form
 * (MERGEFIELD instruction / text:placeholder content) in both formats. */
export function chipLabel(field: MergeFieldSpec): string {
  switch (field.kind) {
    case "merge":
      return `{{${field.name}}}`;
    case "if":
      return `{{#if ${field.name}}}`;
    case "endif":
      return "{{/if}}";
    case "each":
      return `{{#each ${field.name}}}`;
  }
}

const MERGE_RE = /^\{\{([A-Za-z_][A-Za-z0-9_.]*)\}\}$/;
const IF_RE = /^\{\{#if ([A-Za-z_][A-Za-z0-9_.]*)\}\}$/;
const ENDIF_RE = /^\{\{\/if\}\}$/;
const EACH_RE = /^\{\{#each ([A-Za-z_][A-Za-z0-9_.]*)\}\}$/;

/** Parses a chip's canonical text back into a `MergeFieldSpec`, or `null`
 * if it does not match any chip grammar exactly — used by the .docx/.odt
 * readers (a `MERGEFIELD`/`text:placeholder` whose content doesn't parse
 * falls through to the opaque-inline path, per SPEC-008) and by "Convert
 * {{…}} text to merge fields". */
export function parseChipText(text: string): MergeFieldSpec | null {
  const merge = MERGE_RE.exec(text);
  if (merge) {
    return { kind: "merge", name: merge[1] };
  }
  const ifMatch = IF_RE.exec(text);
  if (ifMatch) {
    return { kind: "if", name: ifMatch[1] };
  }
  if (ENDIF_RE.test(text)) {
    return { kind: "endif" };
  }
  const eachMatch = EACH_RE.exec(text);
  if (eachMatch) {
    return { kind: "each", name: eachMatch[1] };
  }
  return null;
}

/** Every `{{…}}`-shaped substring of `text`, whether or not it parses as a
 * chip — used to find literal conversion candidates ("Convert {{…}} text to
 * merge fields") without requiring the substring to already be well-formed
 * chip syntax (a malformed one is just not offered). */
const LITERAL_CANDIDATE_RE = /\{\{[^{}]*\}\}/g;

export function findLiteralChipCandidates(text: string): string[] {
  return text.match(LITERAL_CANDIDATE_RE) ?? [];
}

export function createMergeChipRun(name: string, formatting: RunFormatting = {}): Run {
  return createChipRun({ kind: "merge", name }, formatting);
}

export function createIfChipRun(name: string, formatting: RunFormatting = {}): Run {
  return createChipRun({ kind: "if", name }, formatting);
}

export function createEndifChipRun(formatting: RunFormatting = {}): Run {
  return createChipRun({ kind: "endif" }, formatting);
}

export function createEachChipRun(name: string, formatting: RunFormatting = {}): Run {
  return createChipRun({ kind: "each", name }, formatting);
}

export const MERGE_CHIP_KINDS: readonly MergeChipKind[] = ["merge", "if", "endif", "each"];
