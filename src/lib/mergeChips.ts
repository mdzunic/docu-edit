// Mail-merge chip serialisation (SPEC-046): the one grammar shared by a
// chip's display label and its on-disk form in both formats — see "Chip
// serialisation" in specs/046-template-gallery-and-mail-merge.md. Kept
// framework-free like the rest of lib/ so both readers/writers and the
// editor's paint-time substitution can share it.
import {
  createChipRun,
  isTable,
  type Block,
  type DocumentModel,
  type HeaderFooterKind,
  type MergeChipKind,
  type MergeFieldSpec,
  type Paragraph,
  type Run,
  type RunFormatting,
} from "./documentModel";

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

// --- "Convert {{…}} text to merge fields" -----------------------------------
//
// SPEC-046 "Literal text is not a chip": typing/opening `{{client_name}}`
// produces text, never a chip, implicitly. Conversion is one explicit,
// undoable command that scans the body, tables, headers, footers and notes
// and converts every well-formed match in a single pass — offered when a
// document is opened as a template and literal candidates are found.

/** Matches exactly the chip grammar (not "any {{...}}" — a malformed one,
 * e.g. `{{client name}}` with a space, is left as text) so this scanner and
 * the actual conversion below always agree on what counts as a candidate. */
const CHIP_SCAN_PATTERN = "\\{\\{(?:#if [A-Za-z_][A-Za-z0-9_.]*|#each [A-Za-z_][A-Za-z0-9_.]*|/if|[A-Za-z_][A-Za-z0-9_.]*)\\}\\}";

function chipScanRegExp(): RegExp {
  // A fresh RegExp per call: a `g`-flagged instance is stateful (lastIndex),
  // and this is called once per run being scanned.
  return new RegExp(CHIP_SCAN_PATTERN, "g");
}

/** True when `model` contains at least one well-formed `{{…}}` literal
 * outside any existing chip run — what enables/offers the conversion
 * command. */
export function hasConvertibleLiteralChips(model: DocumentModel): boolean {
  const pattern = chipScanRegExp();
  return someTextRun(model, (run) => {
    pattern.lastIndex = 0;
    return pattern.test(run.text);
  });
}

function someTextRun(model: DocumentModel, predicate: (run: Run) => boolean): boolean {
  const paragraphs = allParagraphs(model);
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs) {
      if (!run.field && predicate(run)) {
        return true;
      }
    }
  }
  return false;
}

function allParagraphs(model: DocumentModel): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  const collectBlocks = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        paragraphs.push(block);
      } else if (isTable(block)) {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            paragraphs.push(...cell.blocks);
          }
        }
      }
    }
  };
  const HF_KINDS: HeaderFooterKind[] = ["default", "first", "even"];
  model.sections.forEach((section) => {
    collectBlocks(section.blocks);
    HF_KINDS.forEach((kind) => {
      if (section.headers?.[kind]) {
        collectBlocks(section.headers[kind].blocks);
      }
      if (section.footers?.[kind]) {
        collectBlocks(section.footers[kind].blocks);
      }
    });
  });
  (model.notes ?? []).forEach((note) => collectBlocks(note.blocks));
  return paragraphs;
}

/** Splits one text run's literal `{{…}}` matches into chip runs plus the
 * surrounding text, preserving formatting/styleId/tracked-change state on
 * every piece. A run already carrying a field (an existing chip) is left
 * alone — its text is the one-character placeholder, which never matches. */
function splitLiteralChipsInRun(run: Run): { runs: Run[]; converted: number } {
  const pattern = chipScanRegExp();
  const pieces: Run[] = [];
  let converted = 0;
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(run.text))) {
    const field = parseChipText(match[0]);
    if (!field) {
      continue;
    }
    if (match.index > lastIndex) {
      pieces.push({ ...run, text: run.text.slice(lastIndex, match.index) });
    }
    const chipRun = createChipRun(field, run.formatting);
    if (run.styleId !== undefined) {
      chipRun.styleId = run.styleId;
    }
    if (run.changes) {
      chipRun.changes = run.changes;
    }
    pieces.push(chipRun);
    converted += 1;
    lastIndex = match.index + match[0].length;
  }
  if (converted === 0) {
    return { runs: [run], converted: 0 };
  }
  if (lastIndex < run.text.length) {
    pieces.push({ ...run, text: run.text.slice(lastIndex) });
  }
  return { runs: pieces, converted };
}

/** Converts every well-formed `{{…}}` literal in the document (body,
 * tables, headers, footers, notes) into a chip run — the "Convert {{…}}
 * text to merge fields" command, applied as a single undo step by the
 * caller (it returns a new model; it does not mutate `model`). */
export function convertLiteralChipsToRuns(model: DocumentModel): { model: DocumentModel; converted: number } {
  const next = structuredClone(model);
  let converted = 0;
  const convertParagraph = (paragraph: Paragraph): void => {
    const runs: Run[] = [];
    for (const run of paragraph.runs) {
      if (run.field) {
        runs.push(run);
        continue;
      }
      const split = splitLiteralChipsInRun(run);
      converted += split.converted;
      runs.push(...split.runs);
    }
    paragraph.runs = runs;
  };
  const convertBlocks = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        convertParagraph(block);
      } else if (isTable(block)) {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            cell.blocks.forEach(convertParagraph);
          }
        }
      }
    }
  };
  const HF_KINDS: HeaderFooterKind[] = ["default", "first", "even"];
  next.sections.forEach((section) => {
    convertBlocks(section.blocks);
    HF_KINDS.forEach((kind) => {
      if (section.headers?.[kind]) {
        convertBlocks(section.headers[kind].blocks);
      }
      if (section.footers?.[kind]) {
        convertBlocks(section.footers[kind].blocks);
      }
    });
  });
  (next.notes ?? []).forEach((note) => convertBlocks(note.blocks));
  return { model: next, converted };
}
