// Plain-text editing operations over the document model (SPEC-010). The
// autosave spec needs a real typing surface — content changes, a caret and
// paragraph splits/merges — but rich-text editing proper is SPEC-011, so
// this module deliberately stays at the character level: every operation
// preserves the run formatting around the edit and never invents new
// formatting.
//
// Mutation contract: an edited block is REPLACED in its section's `blocks`
// array (never mutated in place). The renderer relies on this — `v-memo`
// compares block object identity to skip re-rendering untouched paragraphs,
// which is what keeps typing responsive on very large documents.
import type {
  CharacterFormatKey,
  DocumentModel,
  Paragraph,
  PendingFormat,
  Run,
  RunFormatting,
} from "./documentModel";

/** A caret location in model coordinates: a block address plus a character
 * offset into the block's concatenated run text (UTF-16 units). */
export interface CaretPosition {
  section: number;
  block: number;
  offset: number;
}

/** Typed `null` seed for caret refs in `.vue` scripts, which espree bars
 * from writing the `CaretPosition | null` annotation themselves. */
export function noCaret(): CaretPosition | null {
  return null;
}

/** The block's visible text: its runs' text concatenated. */
export function blockText(block: Paragraph): string {
  let text = "";
  for (const run of block.runs) {
    text += run.text;
  }
  return text;
}

export function blockLength(block: Paragraph): number {
  let length = 0;
  for (const run of block.runs) {
    length += run.text.length;
  }
  return length;
}

/** Negative when a precedes b, positive when a follows b, 0 when equal. */
export function compareCaret(a: CaretPosition, b: CaretPosition): number {
  if (a.section !== b.section) {
    return a.section - b.section;
  }
  if (a.block !== b.block) {
    return a.block - b.block;
  }
  return a.offset - b.offset;
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/** Snaps a caret into the model's valid address space (indices and offset
 * clamped), so stale carets — e.g. restored from a record whose document
 * has since shrunk — can never address outside the tree. */
export function clampCaret(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const section = clampNumber(caret.section, 0, model.sections.length - 1);
  const blocks = model.sections[section].blocks;
  const block = clampNumber(caret.block, 0, blocks.length - 1);
  const offset = clampNumber(caret.offset, 0, blockLength(blocks[block]));
  return { section, block, offset };
}

function sameFormatting(a: Run["formatting"], b: Run["formatting"]): boolean {
  const keysA = Object.keys(a) as Array<keyof Run["formatting"]>;
  const keysB = Object.keys(b) as Array<keyof Run["formatting"]>;
  return (
    keysA.length === keysB.length && keysA.every((key) => a[key] === b[key])
  );
}

/** Drops empty runs and coalesces adjacent runs with identical formatting,
 * so editing never fragments a paragraph into needless run boundaries.
 * Exported as the model's run normalisation: the clipboard fragment
 * machinery (SPEC-016) runs pasted/extracted runs through it so fragments
 * obey the same invariants as edited paragraphs. */
export function normalizeRuns(runs: Run[]): Run[] {
  const result: Run[] = [];
  for (const run of runs) {
    if (run.text.length === 0) {
      continue;
    }
    const last = result[result.length - 1];
    if (last && sameFormatting(last.formatting, run.formatting)) {
      result[result.length - 1] = { ...last, text: last.text + run.text };
    } else {
      result.push(run);
    }
  }
  return result;
}

/** Splits a run list at a character offset. Untouched runs are carried over
 * by reference (safe: operations replace runs, never mutate them); the run
 * straddling the offset is split into two runs sharing its formatting.
 * Exported for the clipboard fragment machinery (SPEC-016), which splits
 * boundary runs at the selection edges when extracting or splicing. */
export function splitRuns(runs: Run[], offset: number): [Run[], Run[]] {
  const before: Run[] = [];
  const after: Run[] = [];
  let consumed = 0;
  for (const run of runs) {
    const end = consumed + run.text.length;
    if (end <= offset) {
      before.push(run);
    } else if (consumed >= offset) {
      after.push(run);
    } else {
      const cut = offset - consumed;
      before.push({ text: run.text.slice(0, cut), formatting: { ...run.formatting } });
      after.push({ text: run.text.slice(cut), formatting: { ...run.formatting } });
    }
    consumed = end;
  }
  return [normalizeRuns(before), normalizeRuns(after)];
}

/** `formatting` with `key` forced to `value`. `false` deletes the key
 * instead of storing an explicit false, so the result merges with plain
 * runs that never carried the key at all — the same outcome as if the text
 * had never been formatted (SPEC-012 AC-20, AC-21). */
function withFormat(
  formatting: RunFormatting,
  key: CharacterFormatKey,
  value: boolean,
): RunFormatting {
  if (!value) {
    const rest = { ...formatting };
    delete rest[key];
    return rest;
  }
  return { ...formatting, [key]: true };
}

/** `base` with every key of `override` forced to its value — insertText's
 * pending-format merge (SPEC-012 AC-9, AC-10) and the per-run step of the
 * SPEC-013 selection-formatting command. Off semantics per PendingFormat's
 * contract: `false` deletes boolean toggles and `vertAlign` (so the result
 * merges with never-formatted runs), but is *stored* for `highlight` —
 * "No colour" is an explicit off that must survive saving. */
export function applyFormatPatch(base: RunFormatting, override: PendingFormat): RunFormatting {
  let formatting = base;
  for (const key of Object.keys(override) as Array<keyof PendingFormat>) {
    const value = override[key];
    if (value === undefined) {
      continue;
    }
    if (key === "vertAlign" && value === false) {
      const rest = { ...formatting };
      delete rest.vertAlign;
      formatting = rest;
    } else if (typeof value === "boolean" && key !== "highlight") {
      formatting = withFormat(formatting, key as CharacterFormatKey, value);
    } else {
      formatting = { ...formatting, [key]: value };
    }
  }
  return formatting;
}

function replaceBlock(
  model: DocumentModel,
  section: number,
  block: number,
  next: Paragraph,
): void {
  model.sections[section].blocks.splice(block, 1, next);
}

/** Inserts text at the caret. The inserted text inherits the formatting of
 * the run it lands in (the run to the left at a boundary, matching how word
 * processors extend formatting while typing) — or, with `formatOverride`
 * set, that formatting with the override's keys forced on/off (SPEC-012's
 * pending format, AC-9, AC-10). Returns the caret after the inserted text. */
export function insertText(
  model: DocumentModel,
  caret: CaretPosition,
  text: string,
  formatOverride?: PendingFormat | null,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = model.sections[c.section].blocks[c.block];
  const [before, after] = splitRuns(block.runs, c.offset);
  const baseFormatting =
    before.length > 0
      ? before[before.length - 1].formatting
      : after.length > 0
        ? after[0].formatting
        : {};
  const formatting = formatOverride
    ? applyFormatPatch(baseFormatting, formatOverride)
    : baseFormatting;
  const inserted: Run = { text, formatting };
  replaceBlock(model, c.section, c.block, {
    ...block,
    runs: normalizeRuns([...before, inserted, ...after]),
  });
  return { section: c.section, block: c.block, offset: c.offset + text.length };
}

/** Splits the caret's paragraph in two (the Enter key). Both halves keep the
 * paragraph's block-level formatting — including list membership, which is
 * what makes Enter in a list item create the next item at the same level
 * (SPEC-018 AC-6). Returns the caret at the start of the second half. */
export function splitParagraph(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = model.sections[c.section].blocks[c.block];
  const [before, after] = splitRuns(block.runs, c.offset);
  const first: Paragraph = { ...block, runs: before };
  const second: Paragraph = {
    ...block,
    runs: after,
    formatting: block.formatting ? { ...block.formatting } : undefined,
  };
  if (block.list) {
    second.list = { ...block.list };
  }
  model.sections[c.section].blocks.splice(c.block, 1, first, second);
  return { section: c.section, block: c.block + 1, offset: 0 };
}

/** Deletes everything between two carets (any order). Deleting across a
 * paragraph boundary merges the boundary paragraphs (the surviving one keeps
 * the first paragraph's block formatting); across a section boundary the
 * sections merge too. Returns the caret at the deletion point. */
export function deleteRange(
  model: DocumentModel,
  a: CaretPosition,
  b: CaretPosition,
): CaretPosition {
  let start = clampCaret(model, a);
  let end = clampCaret(model, b);
  if (compareCaret(start, end) > 0) {
    [start, end] = [end, start];
  }
  if (compareCaret(start, end) === 0) {
    return start;
  }
  const startBlock = model.sections[start.section].blocks[start.block];
  const endBlock = model.sections[end.section].blocks[end.block];
  if (startBlock === endBlock) {
    const [before] = splitRuns(startBlock.runs, start.offset);
    const [, after] = splitRuns(startBlock.runs, end.offset);
    replaceBlock(model, start.section, start.block, {
      ...startBlock,
      runs: normalizeRuns([...before, ...after]),
    });
    return start;
  }
  const [before] = splitRuns(startBlock.runs, start.offset);
  const [, after] = splitRuns(endBlock.runs, end.offset);
  const merged: Paragraph = {
    ...startBlock,
    runs: normalizeRuns([...before, ...after]),
  };
  if (start.section === end.section) {
    model.sections[start.section].blocks.splice(
      start.block,
      end.block - start.block + 1,
      merged,
    );
  } else {
    const startSection = model.sections[start.section];
    const endSection = model.sections[end.section];
    startSection.blocks = [
      ...startSection.blocks.slice(0, start.block),
      merged,
      ...endSection.blocks.slice(end.block + 1),
    ];
    model.sections.splice(start.section + 1, end.section - start.section);
  }
  return start;
}

/** The address one position past the end of a block, or null at the end of
 * the document. */
function nextBlockStart(
  model: DocumentModel,
  section: number,
  block: number,
): CaretPosition | null {
  if (block + 1 < model.sections[section].blocks.length) {
    return { section, block: block + 1, offset: 0 };
  }
  if (section + 1 < model.sections.length) {
    return { section: section + 1, block: 0, offset: 0 };
  }
  return null;
}

function previousBlockEnd(
  model: DocumentModel,
  section: number,
  block: number,
): CaretPosition | null {
  if (block > 0) {
    const prev = model.sections[section].blocks[block - 1];
    return { section, block: block - 1, offset: blockLength(prev) };
  }
  if (section > 0) {
    const prevSection = model.sections[section - 1];
    const prev = prevSection.blocks[prevSection.blocks.length - 1];
    return {
      section: section - 1,
      block: prevSection.blocks.length - 1,
      offset: blockLength(prev),
    };
  }
  return null;
}

/** Backspace: deletes the character before the caret, or merges with the
 * previous paragraph when the caret is at a paragraph start. No-op at the
 * very start of the document. */
export function deleteBackward(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const c = clampCaret(model, caret);
  if (c.offset > 0) {
    return deleteRange(model, { ...c, offset: c.offset - 1 }, c);
  }
  const prev = previousBlockEnd(model, c.section, c.block);
  if (!prev) {
    return c;
  }
  return deleteRange(model, prev, c);
}

/** Forward delete: deletes the character after the caret, or merges with the
 * next paragraph when the caret is at a paragraph end. No-op at the very end
 * of the document. */
export function deleteForward(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = model.sections[c.section].blocks[c.block];
  if (c.offset < blockLength(block)) {
    return deleteRange(model, c, { ...c, offset: c.offset + 1 });
  }
  const next = nextBlockStart(model, c.section, c.block);
  if (!next) {
    return c;
  }
  return deleteRange(model, c, next);
}

// --- character formatting (SPEC-012) ----------------------------------------

/** The formatting new content at `offset` would inherit — the run ending
 * there if any, else the run starting there, else `{}` for an empty
 * paragraph. The read-only counterpart of insertText's inheritance rule,
 * used to drive the toolbar's caret-formatting display (AC-12). */
export function formattingAtOffset(block: Paragraph, offset: number): RunFormatting {
  const [before, after] = splitRuns(block.runs, offset);
  if (before.length > 0) {
    return before[before.length - 1].formatting;
  }
  if (after.length > 0) {
    return after[0].formatting;
  }
  return {};
}

/** Every block address from `start`'s block up to and including `end`'s,
 * in document order. Callers pass an already-ordered range (`start` at or
 * before `end`). */
function blockAddressesInRange(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
): Array<{ section: number; block: number }> {
  const addresses: Array<{ section: number; block: number }> = [];
  for (let section = start.section; section <= end.section; section++) {
    const blocks = model.sections[section].blocks;
    const from = section === start.section ? start.block : 0;
    const to = section === end.section ? end.block : blocks.length - 1;
    for (let block = from; block <= to; block++) {
      addresses.push({ section, block });
    }
  }
  return addresses;
}

/** Every run's formatting strictly between `start` and `end` (already
 * ordered), including runs split at the range's boundaries. A paragraph the
 * range merely passes through without covering any of its characters (an
 * empty paragraph, or a boundary landing exactly between two of its runs)
 * contributes nothing. */
export function runFormattingsInRange(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
): RunFormatting[] {
  const formattings: RunFormatting[] = [];
  for (const { section, block } of blockAddressesInRange(model, start, end)) {
    const paragraph = model.sections[section].blocks[block];
    const from = section === start.section && block === start.block ? start.offset : 0;
    const to =
      section === end.section && block === end.block ? end.offset : blockLength(paragraph);
    if (from >= to) {
      continue;
    }
    const [, tail] = splitRuns(paragraph.runs, from);
    const [within] = splitRuns(tail, to - from);
    for (const run of within) {
      formattings.push(run.formatting);
    }
  }
  return formattings;
}

/** Applies a format patch to every run between `start` and `end` (already
 * ordered), splitting the runs straddling the range's boundaries first and
 * merging the result back down to the minimum number of runs afterward
 * (SPEC-012 AC-7, AC-8, AC-20, AC-21; SPEC-013 extends the patch beyond
 * boolean toggles — see PendingFormat for the per-key off semantics). */
export function patchFormatting(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
  patch: PendingFormat,
): void {
  for (const { section, block } of blockAddressesInRange(model, start, end)) {
    const paragraph = model.sections[section].blocks[block];
    const from = section === start.section && block === start.block ? start.offset : 0;
    const to =
      section === end.section && block === end.block ? end.offset : blockLength(paragraph);
    if (from >= to) {
      continue;
    }
    const [before, tail] = splitRuns(paragraph.runs, from);
    const [within, after] = splitRuns(tail, to - from);
    const changed = within.map((run) => ({
      text: run.text,
      formatting: applyFormatPatch(run.formatting, patch),
    }));
    replaceBlock(model, section, block, {
      ...paragraph,
      runs: normalizeRuns([...before, ...changed, ...after]),
    });
  }
}

/** Sets a single boolean character format across a range — SPEC-012's
 * original entry point, now a thin wrapper over the SPEC-013 patch form. */
export function setFormatting(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
  key: CharacterFormatKey,
  value: boolean,
): void {
  patchFormatting(model, start, end, { [key]: value });
}
