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
import type { DocumentModel, Paragraph, Run } from "./documentModel";

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
 * so editing never fragments a paragraph into needless run boundaries. */
function normalizeRuns(runs: Run[]): Run[] {
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
 * straddling the offset is split into two runs sharing its formatting. */
function splitRuns(runs: Run[], offset: number): [Run[], Run[]] {
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
 * processors extend formatting while typing). Returns the caret after the
 * inserted text. */
export function insertText(
  model: DocumentModel,
  caret: CaretPosition,
  text: string,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = model.sections[c.section].blocks[c.block];
  const [before, after] = splitRuns(block.runs, c.offset);
  let runs: Run[];
  if (before.length > 0) {
    const last = before[before.length - 1];
    runs = [...before.slice(0, -1), { ...last, text: last.text + text }, ...after];
  } else if (after.length > 0) {
    const first = after[0];
    runs = [{ ...first, text: text + first.text }, ...after.slice(1)];
  } else {
    runs = [{ text, formatting: {} }];
  }
  replaceBlock(model, c.section, c.block, { ...block, runs: normalizeRuns(runs) });
  return { section: c.section, block: c.block, offset: c.offset + text.length };
}

/** Splits the caret's paragraph in two (the Enter key). Both halves keep the
 * paragraph's block-level formatting. Returns the caret at the start of the
 * second half. */
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
