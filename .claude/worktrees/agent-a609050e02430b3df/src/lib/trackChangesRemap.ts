// Selection remapping for the track-changes diff engine (SPEC-039).
//
// `applyTrackChanges` rewrites an ordinary command's result so deletions are
// *retained* (struck through) instead of removed, and a paragraph merge is
// kept as two paragraphs — which shifts character offsets and block indexes
// relative to the untracked result the command actually produced. The
// command's post-edit selection is in the untracked model's coordinates, so
// it must be translated into the tracked model's space before it becomes
// the live caret. Without this, the caret after "type over a selection"
// points inside the resurrected deletion, and every following keystroke
// scatters its insertion into the middle of the strike — which accept-all
// then commits as scrambled text.
//
// The mapping never needs a diff of its own: the tracked model contains the
// untracked model's characters as an in-order subsequence (the diff only
// splices retained-deletion characters in; it never reorders or drops), and
// a character's `delete` mark survives identically on both sides — so a
// two-pointer walk that skips tracked-only (resurrected) characters aligns
// the two streams exactly.
import { isTable, type Block, type DocumentModel, type Paragraph, type Table } from "./documentModel";
import { windowOf } from "./core/arrayWindow";
import type { CaretPosition } from "./textEditing";
import type { ModelSelection } from "./commands";

interface CharStream {
  text: string;
  /** One flag per UTF-16 unit of `text` (caret offsets are UTF-16 units). */
  deleted: boolean[];
  /** Cumulative end offset of each paragraph's text within the stream. */
  ends: number[];
}

function streamOf(paragraphs: Paragraph[]): CharStream {
  let text = "";
  const deleted: boolean[] = [];
  const ends: number[] = [];
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs) {
      const isDeleted = run.changes?.delete !== undefined;
      text += run.text;
      for (let i = 0; i < run.text.length; i++) {
        deleted.push(isDeleted);
      }
    }
    ends.push(text.length);
  }
  return { text, deleted, ends };
}

/** Maps (paragraph, offset) through a changed paragraph window by aligning
 * the untracked character stream against the tracked one. A caret that sat
 * at a paragraph *start* (offset 0) stays at a paragraph start — an
 * Enter-split's caret must land at the head of the new paragraph, not at
 * the visually-different end of the previous one. */
function remapThroughParagraphWindow(
  uMid: Paragraph[],
  tMid: Paragraph[],
  paragraphIndex: number,
  offset: number,
): { paragraphIndex: number; offset: number } {
  const u = streamOf(uMid);
  const t = streamOf(tMid);
  const uStart = paragraphIndex === 0 ? 0 : (u.ends[paragraphIndex - 1] ?? 0);
  const target = Math.min(uStart + Math.max(0, offset), u.text.length);
  let uIdx = 0;
  let tIdx = 0;
  while (uIdx < target && tIdx < t.text.length) {
    if (u.text[uIdx] === t.text[tIdx] && u.deleted[uIdx] === t.deleted[tIdx]) {
      uIdx++;
      tIdx++;
    } else {
      // A character only the tracked side has — a deletion the diff
      // resurrected. Step over it; the caret stays before/after it exactly
      // as the untracked position dictates.
      tIdx++;
    }
  }
  // The walk stops as soon as the untracked offset is reached, so trailing
  // resurrected characters at the boundary stay *after* the caret — typing
  // continues right after the previous insertion, never inside the strike.
  const atParagraphStart = offset <= 0;
  let pIdx = 0;
  while (pIdx < t.ends.length - 1 && (atParagraphStart ? tIdx >= t.ends[pIdx] : tIdx > t.ends[pIdx])) {
    pIdx++;
  }
  const start = pIdx === 0 ? 0 : t.ends[pIdx - 1];
  return { paragraphIndex: pIdx, offset: tIdx - start };
}

/** Index mapping for a row/cell array the diff rewrote: untouched prefix and
 * suffix entries keep (shifted) positions, and inside the changed window the
 * diff places retained (deleted) entries ahead of the kept ones, so a kept
 * entry moves by the window's growth. */
function shiftIndexThroughWindow<T>(uArr: T[], tArr: T[], index: number): number {
  if (uArr === tArr) {
    return index;
  }
  const w = windowOf(uArr, tArr);
  if (index < w.prefixLen) {
    return index;
  }
  const delta = tArr.length - uArr.length;
  if (index >= uArr.length - w.suffixLen) {
    return index + delta;
  }
  return delta > 0 ? index + delta : index;
}

function remapThroughCellBlocks(
  uParas: Paragraph[],
  tParas: Paragraph[],
  blockIndex: number,
  offset: number,
): { paragraphIndex: number; offset: number } {
  if (uParas === tParas) {
    return { paragraphIndex: blockIndex, offset };
  }
  const w = windowOf(uParas, tParas);
  if (blockIndex < w.prefixLen) {
    return { paragraphIndex: blockIndex, offset };
  }
  if (blockIndex >= uParas.length - w.suffixLen) {
    return { paragraphIndex: blockIndex + (tParas.length - uParas.length), offset };
  }
  const mapped = remapThroughParagraphWindow(w.beforeMid, w.afterMid, blockIndex - w.prefixLen, offset);
  return { paragraphIndex: w.prefixLen + mapped.paragraphIndex, offset: mapped.offset };
}

function remapThroughTable(uTable: Table, tTable: Table, caret: CaretPosition): CaretPosition {
  const cell = caret.cell;
  if (!cell) {
    return caret;
  }
  const row = shiftIndexThroughWindow(uTable.rows, tTable.rows, cell.row);
  const uRow = uTable.rows[cell.row];
  const tRow = tTable.rows[row];
  if (!uRow || !tRow) {
    return caret;
  }
  const col = shiftIndexThroughWindow(uRow.cells, tRow.cells, cell.col);
  const uCell = uRow.cells[cell.col];
  const tCell = tRow.cells[col];
  if (!uCell || !tCell) {
    return caret;
  }
  const mapped = remapThroughCellBlocks(uCell.blocks, tCell.blocks, cell.block, caret.offset);
  if (row === cell.row && col === cell.col && mapped.paragraphIndex === cell.block && mapped.offset === caret.offset) {
    return caret;
  }
  return { ...caret, offset: mapped.offset, cell: { row, col, block: mapped.paragraphIndex } };
}

function isParagraphBlock(block: Block): block is Paragraph {
  return block.type === "paragraph";
}

function remapThroughBlocks(uBlocks: Block[], tBlocks: Block[], caret: CaretPosition): CaretPosition {
  if (uBlocks === tBlocks) {
    return caret;
  }
  const w = windowOf(uBlocks, tBlocks);
  if (caret.block < w.prefixLen) {
    return caret;
  }
  if (caret.block >= uBlocks.length - w.suffixLen) {
    const delta = tBlocks.length - uBlocks.length;
    return delta === 0 ? caret : { ...caret, block: caret.block + delta };
  }
  const indexInMid = caret.block - w.prefixLen;
  const uBlock = w.beforeMid[indexInMid];
  if (uBlock && isTable(uBlock)) {
    // Tables diff 1:1 in place (diffTable), so the block index is stable;
    // descend for a caret addressing a cell paragraph.
    const tBlock = w.afterMid[indexInMid];
    if (caret.cell && tBlock && isTable(tBlock)) {
      return remapThroughTable(uBlock, tBlock, caret);
    }
    return caret;
  }
  if (w.beforeMid.every(isParagraphBlock) && w.afterMid.every(isParagraphBlock)) {
    const mapped = remapThroughParagraphWindow(
      w.beforeMid as Paragraph[],
      w.afterMid as Paragraph[],
      indexInMid,
      caret.offset,
    );
    const block = w.prefixLen + mapped.paragraphIndex;
    if (block === caret.block && mapped.offset === caret.offset) {
      return caret;
    }
    return { ...caret, block, offset: mapped.offset };
  }
  // Mixed window (an image or table retained beside the edit): the diff
  // places resurrected blocks ahead of the kept ones, so the caret's block
  // shifts by the window's growth.
  const delta = w.afterMid.length - w.beforeMid.length;
  return delta > 0 ? { ...caret, block: caret.block + delta } : caret;
}

/** Translates one caret from the untracked command result's coordinates to
 * the tracked model `applyTrackChanges` produced from it. */
export function remapCaretThroughTracking(
  untracked: DocumentModel,
  tracked: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  if (untracked === tracked) {
    return caret;
  }
  const uSection = untracked.sections[caret.section];
  const tSection = tracked.sections[caret.section];
  if (!uSection || !tSection || uSection === tSection) {
    return caret;
  }
  return remapThroughBlocks(uSection.blocks, tSection.blocks, caret);
}

/** Translates a command's post-edit selection into tracked-model
 * coordinates — applied wherever `applyTrackChanges` rewrites a command's
 * result (DocumentView's applyEdit). Returns the selection object unchanged
 * when nothing shifted, so callers can keep cheap identity checks. */
export function remapSelectionThroughTracking(
  untracked: DocumentModel,
  tracked: DocumentModel,
  selection: ModelSelection,
): ModelSelection {
  if (untracked === tracked) {
    return selection;
  }
  const anchor = remapCaretThroughTracking(untracked, tracked, selection.anchor);
  const focus =
    selection.focus === selection.anchor ? anchor : remapCaretThroughTracking(untracked, tracked, selection.focus);
  if (anchor === selection.anchor && focus === selection.focus) {
    return selection;
  }
  return { anchor, focus };
}
