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
import { sameModelColor } from "./core/theme";
import type { ModelColor } from "./core/theme";
import {
  isBreak,
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  isToc,
  isTof,
  type Block,
  type BookmarkAnchor,
  type CharacterFormatKey,
  type CommentAnchor,
  type DocumentModel,
  type Paragraph,
  type PendingFormat,
  type Run,
  type RunFormatting,
  type Table,
} from "./documentModel";
import { nextGraphemeBoundary, previousGraphemeBoundary } from "./core/grapheme";
import { releaseAnchorOrdinal } from "./drawingAnchors";

/** The path from a table block down to one of its cell's paragraphs
 * (SPEC-022): the cell's row/column plus the paragraph index within the
 * cell's own block list. */
export interface CellPath {
  row: number;
  col: number;
  block: number;
}

/** A caret location in model coordinates: a block address plus a character
 * offset into the block's concatenated run text (UTF-16 units). When the
 * addressed block is a table (SPEC-022), `cell` descends to a paragraph
 * inside one of its cells and `offset` indexes that paragraph's text; a
 * table address *without* `cell` is the table treated as an atom — offset 0
 * is the position before it, offset 1 the position after — which is how
 * ranges cover a whole table without addressing its interior. */
export interface CaretPosition {
  section: number;
  block: number;
  offset: number;
  cell?: CellPath;
}

/** A paragraph's address in the model: a top-level block, or — through
 * `cell` — a paragraph inside a table cell (SPEC-022). The shared address
 * shape for every per-paragraph command (commands.ts's touchedBlocks). */
export interface ParagraphAddress {
  section: number;
  block: number;
  cell?: CellPath;
}

/** Typed `null` seed for caret refs in `.vue` scripts, which espree bars
 * from writing the `CaretPosition | null` annotation themselves. */
export function noCaret(): CaretPosition | null {
  return null;
}

/** The block's visible text: its runs' text concatenated. For a table, the
 * text of every cell paragraph in DOM order — which is exactly what the
 * rendered table element's `textContent` is, keeping the canvas's
 * model/DOM text equality check meaningful (SPEC-022). */
export function blockText(block: Block): string {
  if (isTable(block)) {
    let text = "";
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const paragraph of cell.blocks) {
          text += blockText(paragraph);
        }
      }
    }
    return text;
  }
  if (isImage(block) || isShape(block) || isChart(block) || isConnector(block) || isBreak(block) || isToc(block) || isTof(block)) {
    return "";
  }
  let text = "";
  for (const run of block.runs) {
    text += run.text;
  }
  return text;
}

/** The block's caret-offset space. A table, image or manual break is an atom
 * of length 1 (offset 0 = before it, 1 = after it — see CaretPosition),
 * never addressed by character. */
export function blockLength(block: Block): number {
  if (isTable(block) || isImage(block) || isShape(block) || isChart(block) || isConnector(block) || isBreak(block) || isToc(block) || isTof(block)) {
    return 1;
  }
  let length = 0;
  for (const run of block.runs) {
    length += run.text.length;
  }
  return length;
}

/** Negative when a precedes b, positive when a follows b, 0 when equal.
 * Within one table block, a cell path orders by row, column, paragraph and
 * offset; a table-atom caret (no cell) at offset 0 precedes the whole
 * interior, at offset 1 follows it. */
export function compareCaret(a: CaretPosition, b: CaretPosition): number {
  if (a.section !== b.section) {
    return a.section - b.section;
  }
  if (a.block !== b.block) {
    return a.block - b.block;
  }
  if (a.cell || b.cell) {
    if (!a.cell) {
      return a.offset <= 0 ? -1 : 1;
    }
    if (!b.cell) {
      return b.offset <= 0 ? 1 : -1;
    }
    if (a.cell.row !== b.cell.row) {
      return a.cell.row - b.cell.row;
    }
    if (a.cell.col !== b.cell.col) {
      return a.cell.col - b.cell.col;
    }
    if (a.cell.block !== b.cell.block) {
      return a.cell.block - b.cell.block;
    }
  }
  return a.offset - b.offset;
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/** Snaps a caret into the model's valid address space (indices and offset
 * clamped), so stale carets — e.g. restored from a record whose document
 * has since shrunk — can never address outside the tree. A cell path is
 * clamped into its table, dropped when the addressed block is not a table;
 * a table addressed without a cell path clamps to its atom range 0..1. */
export function clampCaret(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const section = clampNumber(caret.section, 0, model.sections.length - 1);
  const blocks = model.sections[section].blocks;
  const block = clampNumber(caret.block, 0, blocks.length - 1);
  const target = blocks[block];
  if (isTable(target)) {
    if (caret.cell) {
      const row = clampNumber(caret.cell.row, 0, target.rows.length - 1);
      const cells = target.rows[row].cells;
      const col = clampNumber(caret.cell.col, 0, cells.length - 1);
      const cellBlocks = cells[col].blocks;
      const cellBlock = clampNumber(caret.cell.block, 0, cellBlocks.length - 1);
      const offset = clampNumber(caret.offset, 0, blockLength(cellBlocks[cellBlock]));
      return { section, block, offset, cell: { row, col, block: cellBlock } };
    }
    return { section, block, offset: clampNumber(caret.offset, 0, 1) };
  }
  if (isImage(target) || isShape(target) || isChart(target) || isConnector(target) || isBreak(target) || isToc(target) || isTof(target)) {
    return { section, block, offset: clampNumber(caret.offset, 0, 1) };
  }
  const offset = clampNumber(caret.offset, 0, blockLength(target));
  return { section, block, offset };
}

/** The paragraph a caret addresses — through its cell path for a caret
 * inside a table — or null for a table-atom caret, which addresses the
 * table's edge rather than any paragraph. The caret must be clamped. */
export function paragraphAtCaret(
  model: DocumentModel,
  caret: CaretPosition,
): Paragraph | null {
  return paragraphAt(model, caret);
}

/** The paragraph at an address, or null when the address points at a table
 * block without descending into a cell. */
export function paragraphAt(
  model: DocumentModel,
  address: ParagraphAddress,
): Paragraph | null {
  const block = model.sections[address.section].blocks[address.block];
  if (isTable(block)) {
    if (!address.cell) {
      return null;
    }
    return block.rows[address.cell.row].cells[address.cell.col].blocks[address.cell.block];
  }
  if (isImage(block) || isShape(block) || isChart(block) || isConnector(block) || isBreak(block) || isToc(block) || isTof(block)) {
    return null;
  }
  return block;
}

/** `table` with the addressed cell's block list replaced — every container
 * along the path (rows array, row, cells array, cell) is cloned, everything
 * else shared, upholding the replace-don't-mutate contract one level down
 * from the section's own blocks array. */
export function withCellBlocks(
  table: Table,
  row: number,
  col: number,
  blocks: Paragraph[],
): Table {
  const rows = [...table.rows];
  const cells = [...rows[row].cells];
  cells[col] = { ...cells[col], blocks };
  rows[row] = { ...rows[row], cells };
  return { ...table, rows };
}

/** Splices a cell's block list (copy-on-write, see withCellBlocks) and
 * writes the resulting table back into the section's blocks array. */
function spliceCellBlocks(
  model: DocumentModel,
  section: number,
  blockIndex: number,
  row: number,
  col: number,
  start: number,
  deleteCount: number,
  ...items: Paragraph[]
): void {
  const table = model.sections[section].blocks[blockIndex];
  if (!isTable(table)) {
    return;
  }
  const blocks = [...table.rows[row].cells[col].blocks];
  blocks.splice(start, deleteCount, ...items);
  model.sections[section].blocks.splice(
    blockIndex,
    1,
    withCellBlocks(table, row, col, blocks),
  );
}

/** Replaces the paragraph at an address (copy-on-write through a table cell
 * when the address descends into one). The shared write primitive of every
 * per-paragraph command. */
export function replaceParagraphAt(
  model: DocumentModel,
  address: ParagraphAddress,
  next: Paragraph,
): void {
  spliceParagraphs(model, address, 1, next);
}

/** Splices the block list containing the addressed paragraph — the
 * section's blocks, or the cell's own list (copy-on-write) for an address
 * inside a table. The n-for-m generalisation of replaceParagraphAt, for
 * callers (paste) that replace one paragraph with several. */
export function spliceParagraphs(
  model: DocumentModel,
  address: ParagraphAddress,
  deleteCount: number,
  ...items: Paragraph[]
): void {
  if (address.cell) {
    spliceCellBlocks(
      model,
      address.section,
      address.block,
      address.cell.row,
      address.cell.col,
      address.cell.block,
      deleteCount,
      ...items,
    );
    return;
  }
  model.sections[address.section].blocks.splice(address.block, deleteCount, ...items);
}

function sameFormatting(a: Run["formatting"], b: Run["formatting"]): boolean {
  const keysA = Object.keys(a) as Array<keyof Run["formatting"]>;
  const keysB = Object.keys(b) as Array<keyof Run["formatting"]>;
  return (
    keysA.length === keysB.length &&
    keysA.every((key) => {
      const left = a[key];
      const right = b[key];
      // Theme colour/font references (SPEC-077 AC-8) are objects: they
      // compare structurally through sameModelColor so runs sharing a
      // reference still coalesce, and a reference never equals a literal.
      if (typeof left === "object" && left !== null && typeof right === "object" && right !== null) {
        return sameModelColor(left as ModelColor, right as ModelColor);
      }
      return left === right;
    })
  );
}

/** Whether two runs' tracked-change state (SPEC-039) is the same — a plain
 * reference/JSON comparison is enough since `Run.changes` is always built
 * fresh by lib/trackChangesDiff.ts with a stable key order, never hand-
 * edited. Absent on both sides (the overwhelming majority of runs, and
 * every run when tracking has never been used) short-circuits on `a === b`
 * so behaviour for a tracking-free document is unchanged. */
function sameChanges(a: Run["changes"], b: Run["changes"]): boolean {
  if (a === b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }
  return JSON.stringify(a) === JSON.stringify(b);
}

/** A run whose text is a placeholder for a computed value — a page-number
 * field (SPEC-029) or a note reference mark (SPEC-031). Never coalesced
 * with neighbouring runs: merging would silently turn the marker into
 * literal text. */
function isMarkerRun(run: Run): boolean {
  return run.field !== undefined || run.note !== undefined || run.equation !== undefined;
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
    if (
      last &&
      !isMarkerRun(last) &&
      !isMarkerRun(run) &&
      sameFormatting(last.formatting, run.formatting) &&
      sameChanges(last.changes, run.changes) &&
      last.styleId === run.styleId
    ) {
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
      before.push({ ...run, text: run.text.slice(0, cut), formatting: { ...run.formatting } });
      after.push({ ...run, text: run.text.slice(cut), formatting: { ...run.formatting } });
    }
    consumed = end;
  }
  return [normalizeRuns(before), normalizeRuns(after)];
}

// --- anchor maintenance (bookmarks SPEC-025, comment anchors SPEC-038) -----
// Bookmarks and comment-range anchors are both character-offset points on a
// paragraph (side-lists, like opaqueInlines), so the operations below that
// rebuild a paragraph's text keep both kinds pointing at the same spot,
// through one shared generic implementation. Every helper returns the input
// untouched when there is nothing to adjust, so pre-SPEC-025 behaviour is
// byte-identical for anchor-free documents.

/** Anchors after inserting `length` characters at `offset`: anchors after
 * the point shift right; one exactly at it stays put, so text typed at a
 * just-inserted anchor lands after it. Generic over `BookmarkAnchor` and
 * `CommentAnchor` — both are plain `{offset}` points. */
function anchorsAfterInsert<T extends { offset: number }>(
  anchors: T[] | undefined,
  offset: number,
  length: number,
): T[] | undefined {
  if (!anchors || length === 0) {
    return anchors;
  }
  return anchors.map((anchor) =>
    anchor.offset > offset ? { ...anchor, offset: anchor.offset + length } : anchor,
  );
}

/** Anchors after deleting [from, to) within one paragraph: anchors past the
 * range shift left; anchors inside it collapse to the deletion point rather
 * than vanishing, so deleting text around an anchor never silently loses it. */
function anchorsAfterDelete<T extends { offset: number }>(
  anchors: T[] | undefined,
  from: number,
  to: number,
): T[] | undefined {
  if (!anchors || from >= to) {
    return anchors;
  }
  return anchors.map((anchor) =>
    anchor.offset <= from
      ? anchor
      : { ...anchor, offset: Math.max(from, anchor.offset - (to - from)) },
  );
}

/** Anchors of two paragraphs merged by a cross-paragraph deletion: the
 * first keeps [0, from] (later anchors collapse to `from`), the second's
 * land at `from` plus their distance past `to`. */
function mergedAnchors<T extends { offset: number }>(
  first: T[] | undefined,
  second: T[] | undefined,
  from: number,
  to: number,
): T[] | undefined {
  if (!first && !second) {
    return undefined;
  }
  const merged = [
    ...(first ?? []).map((anchor) => (anchor.offset <= from ? anchor : { ...anchor, offset: from })),
    ...(second ?? []).map((anchor) => ({
      ...anchor,
      offset: from + Math.max(0, anchor.offset - to),
    })),
  ];
  return merged.sort((a, b) => a.offset - b.offset);
}

/** Writes a bookmark list onto a freshly built paragraph: assigns it when
 * non-empty, removes the key entirely when empty — an `undefined`-valued
 * key would read as a structural difference in round-trip comparison. */
function setBookmarks(paragraph: Paragraph, bookmarks: BookmarkAnchor[] | undefined): Paragraph {
  if (bookmarks && bookmarks.length > 0) {
    paragraph.bookmarks = bookmarks;
  } else {
    delete paragraph.bookmarks;
  }
  return paragraph;
}

/** Writes a comment-anchor list onto a freshly built paragraph — the same
 * empty-means-absent contract as setBookmarks (SPEC-038). */
function setCommentAnchors(
  paragraph: Paragraph,
  anchors: CommentAnchor[] | undefined,
): Paragraph {
  if (anchors && anchors.length > 0) {
    paragraph.commentAnchors = anchors;
  } else {
    delete paragraph.commentAnchors;
  }
  return paragraph;
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

/** Inserts text at the caret. The inserted text inherits the formatting of
 * the run it lands in (the run to the left at a boundary, matching how word
 * processors extend formatting while typing) — or, with `formatOverride`
 * set, that formatting with the override's keys forced on/off (SPEC-012's
 * pending format, AC-9, AC-10). Returns the caret after the inserted text.
 * A table-atom caret (no cell path) is inert: there is no paragraph there
 * to type into. */
export function insertText(
  model: DocumentModel,
  caret: CaretPosition,
  text: string,
  formatOverride?: PendingFormat | null,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = paragraphAt(model, c);
  if (!block) {
    return c;
  }
  const [before, after] = splitRuns(block.runs, c.offset);
  const baseFormatting =
    before.length > 0
      ? before[before.length - 1].formatting
      : after.length > 0
        ? after[0].formatting
        : {};
  let formatting = formatOverride
    ? applyFormatPatch(baseFormatting, formatOverride)
    : baseFormatting;
  if (formatting.linkHref !== undefined || formatting.linkAnchor !== undefined) {
    // Typing at a link's edge must not extend the link (SPEC-025): the text
    // keeps the link only when both neighbouring runs carry the same target,
    // i.e. the caret sits strictly inside it.
    const insideLink =
      before.length > 0 &&
      after.length > 0 &&
      before[before.length - 1].formatting.linkHref === formatting.linkHref &&
      before[before.length - 1].formatting.linkAnchor === formatting.linkAnchor &&
      after[0].formatting.linkHref === formatting.linkHref &&
      after[0].formatting.linkAnchor === formatting.linkAnchor;
    if (!insideLink) {
      const rest = { ...formatting };
      delete rest.linkHref;
      delete rest.linkAnchor;
      formatting = rest;
    }
  }
  const inserted: Run = { text, formatting };
  replaceParagraphAt(
    model,
    c,
    setCommentAnchors(
      setBookmarks(
        { ...block, runs: normalizeRuns([...before, inserted, ...after]) },
        anchorsAfterInsert(block.bookmarks, c.offset, text.length),
      ),
      anchorsAfterInsert(block.commentAnchors, c.offset, text.length),
    ),
  );
  return { ...c, offset: c.offset + text.length };
}

/** Inserts a complete run at the caret as-is (SPEC-031's note reference
 * marks): unlike insertText it never inherits formatting — the run carries
 * its own — but maintains bookmark anchors the same way. Meant for marker
 * runs, which normalizeRuns refuses to coalesce into neighbours. Returns
 * the caret after the run. A table-atom caret is inert. */
export function insertRun(
  model: DocumentModel,
  caret: CaretPosition,
  run: Run,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = paragraphAt(model, c);
  if (!block) {
    return c;
  }
  const [before, after] = splitRuns(block.runs, c.offset);
  replaceParagraphAt(
    model,
    c,
    setCommentAnchors(
      setBookmarks(
        { ...block, runs: normalizeRuns([...before, run, ...after]) },
        anchorsAfterInsert(block.bookmarks, c.offset, run.text.length),
      ),
      anchorsAfterInsert(block.commentAnchors, c.offset, run.text.length),
    ),
  );
  return { ...c, offset: c.offset + run.text.length };
}

/** Splits the caret's paragraph in two (the Enter key) — inside its cell
 * when the caret is in a table (SPEC-022), which is how a cell grows a
 * second paragraph. Both halves keep the paragraph's block-level formatting
 * — including list membership, which is what makes Enter in a list item
 * create the next item at the same level (SPEC-018 AC-6). Returns the caret
 * at the start of the second half. A table-atom caret is inert. */
export function splitParagraph(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const c = clampCaret(model, caret);
  const block = paragraphAt(model, c);
  if (!block) {
    return c;
  }
  const [before, after] = splitRuns(block.runs, c.offset);
  const first: Paragraph = { ...block, runs: before };
  const second: Paragraph = {
    ...block,
    runs: after,
    formatting: block.formatting ? { ...block.formatting } : undefined,
  };
  // A floating drawing's anchor pairing stays with the half that keeps the
  // paragraph's top edge (SPEC-078) — never duplicated onto both, same as
  // bookmarks below; ordinals are unique per paragraph document-wide.
  delete second.anchorOrdinal;
  if (block.list) {
    second.list = { ...block.list };
  }
  if (block.bookmarks) {
    // Anchors distribute by offset: one exactly at the split point stays at
    // the end of the first half (the location it named), never duplicated
    // onto both halves — names are unique document-wide (SPEC-025).
    setBookmarks(first, block.bookmarks.filter((bookmark) => bookmark.offset <= c.offset));
    setBookmarks(
      second,
      block.bookmarks
        .filter((bookmark) => bookmark.offset > c.offset)
        .map((bookmark) => ({ ...bookmark, offset: bookmark.offset - c.offset })),
    );
  }
  if (block.commentAnchors) {
    // Same distribution rule as bookmarks (SPEC-038): a comment boundary
    // exactly at the split point stays with the first half.
    setCommentAnchors(first, block.commentAnchors.filter((anchor) => anchor.offset <= c.offset));
    setCommentAnchors(
      second,
      block.commentAnchors
        .filter((anchor) => anchor.offset > c.offset)
        .map((anchor) => ({ ...anchor, offset: anchor.offset - c.offset })),
    );
  }
  if (c.cell) {
    spliceCellBlocks(
      model,
      c.section,
      c.block,
      c.cell.row,
      c.cell.col,
      c.cell.block,
      1,
      first,
      second,
    );
    return {
      section: c.section,
      block: c.block,
      offset: 0,
      cell: { ...c.cell, block: c.cell.block + 1 },
    };
  }
  model.sections[c.section].blocks.splice(c.block, 1, first, second);
  return { section: c.section, block: c.block + 1, offset: 0 };
}

function sameCell(a: CaretPosition, b: CaretPosition): boolean {
  return (
    !!a.cell &&
    !!b.cell &&
    a.section === b.section &&
    a.block === b.block &&
    a.cell.row === b.cell.row &&
    a.cell.col === b.cell.col
  );
}

/** Snaps an ordered range whose ends live in different containers onto
 * table-atom boundaries (SPEC-022): a range across two cells — or from a
 * cell out of its table — covers the WHOLE table, never a partial grid, so
 * range edits can treat every table as an indivisible unit. Ends already in
 * the same container (both top-level, or both in one cell) pass through
 * untouched. Exported for the range-consuming commands, which must agree
 * with deleteRange on what a cross-container selection covers. */
export function normalizeRange(
  model: DocumentModel,
  a: CaretPosition,
  b: CaretPosition,
): { start: CaretPosition; end: CaretPosition } {
  let start = clampCaret(model, a);
  let end = clampCaret(model, b);
  if (compareCaret(start, end) > 0) {
    [start, end] = [end, start];
  }
  if (sameCell(start, end)) {
    return { start, end };
  }
  if (start.cell) {
    start = { section: start.section, block: start.block, offset: 0 };
  }
  if (end.cell) {
    end = { section: end.section, block: end.block, offset: 1 };
  }
  return { start, end };
}

// A table-atom END of a still-open range is rephrased as its neighbouring
// paragraph's edge where one exists, so the block-splicing logic below only
// ever sees a table as either fully covered or not covered at all: a start
// sitting AFTER a table (offset 1) advances to the next block's start; an
// end sitting BEFORE one (offset 0) retreats to the previous block's end.
function resolveTableEdges(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
): { start: CaretPosition; end: CaretPosition } {
  const startBlock = model.sections[start.section].blocks[start.block];
  if ((isTable(startBlock) || isImage(startBlock) || isShape(startBlock) || isChart(startBlock) || isConnector(startBlock) || isBreak(startBlock) || isToc(startBlock) || isTof(startBlock)) && !start.cell && start.offset >= 1) {
    const next = nextBlockStart(model, start.section, start.block);
    if (next) {
      start = next;
    }
  }
  const endBlock = model.sections[end.section].blocks[end.block];
  if ((isTable(endBlock) || isImage(endBlock) || isShape(endBlock) || isChart(endBlock) || isConnector(endBlock) || isBreak(endBlock) || isToc(endBlock) || isTof(endBlock)) && !end.cell && end.offset <= 0) {
    const prev = previousBlockEnd(model, end.section, end.block);
    if (prev) {
      end = prev;
    }
  }
  return { start, end };
}

/** Deletes everything between two carets (any order). Deleting across a
 * paragraph boundary merges the boundary paragraphs (the surviving one keeps
 * the first paragraph's block formatting); across a section boundary the
 * sections merge too. A range that reaches into a table covers the whole
 * table (see normalizeRange) and deletes it whole — cells are never merged
 * with outside text. Returns the caret at the deletion point. */
export function deleteRange(
  model: DocumentModel,
  a: CaretPosition,
  b: CaretPosition,
): CaretPosition {
  let { start, end } = normalizeRange(model, a, b);
  if (compareCaret(start, end) === 0) {
    return start;
  }
  // Both ends inside one cell: the same-block/cross-paragraph logic below,
  // scoped to the cell's own block list.
  if (start.cell && end.cell) {
    const cell = start.cell;
    const table = model.sections[start.section].blocks[start.block];
    if (!isTable(table)) {
      return start;
    }
    const cellBlocks = table.rows[cell.row].cells[cell.col].blocks;
    const startPara = cellBlocks[cell.block];
    if (cell.block === end.cell!.block) {
      const [before] = splitRuns(startPara.runs, start.offset);
      const [, after] = splitRuns(startPara.runs, end.offset);
      replaceParagraphAt(
        model,
        start,
        setCommentAnchors(
          setBookmarks(
            { ...startPara, runs: normalizeRuns([...before, ...after]) },
            anchorsAfterDelete(startPara.bookmarks, start.offset, end.offset),
          ),
          anchorsAfterDelete(startPara.commentAnchors, start.offset, end.offset),
        ),
      );
      return start;
    }
    const endPara = cellBlocks[end.cell!.block];
    const [before] = splitRuns(startPara.runs, start.offset);
    const [, after] = splitRuns(endPara.runs, end.offset);
    spliceCellBlocks(
      model,
      start.section,
      start.block,
      cell.row,
      cell.col,
      cell.block,
      end.cell!.block - cell.block + 1,
      setCommentAnchors(
        setBookmarks(
          { ...startPara, runs: normalizeRuns([...before, ...after]) },
          mergedAnchors(startPara.bookmarks, endPara.bookmarks, start.offset, end.offset),
        ),
        mergedAnchors(startPara.commentAnchors, endPara.commentAnchors, start.offset, end.offset),
      ),
    );
    return start;
  }
  ({ start, end } = resolveTableEdges(model, start, end));
  if (compareCaret(start, end) >= 0) {
    return start;
  }
  const startBlock = model.sections[start.section].blocks[start.block];
  const endBlock = model.sections[end.section].blocks[end.block];
  if (startBlock === endBlock) {
    if (isTable(startBlock) || isImage(startBlock) || isShape(startBlock) || isChart(startBlock) || isConnector(startBlock) || isBreak(startBlock) || isToc(startBlock) || isTof(startBlock)) {
      // The range covers exactly this atom: remove it whole. The caret
      // lands where the atom was (SPEC-022 AC-23, SPEC-024 AC-30).
      const blocks = model.sections[start.section].blocks;
      blocks.splice(start.block, 1);
      // A deleted floating drawing releases its anchor pairing when it was
      // the last referrer (SPEC-078) — same rule as shapes.ts's deleteShape,
      // which is the toolbar path to the same removal.
      if (isShape(startBlock) || isChart(startBlock) || isConnector(startBlock)) {
        releaseAnchorOrdinal(model, startBlock.anchorOrdinal);
      }
      if (blocks.length === 0) {
        blocks.push({ type: "paragraph", runs: [] });
      }
      return {
        section: start.section,
        block: Math.min(start.block, blocks.length - 1),
        offset: 0,
      };
    }
    const [before] = splitRuns(startBlock.runs, start.offset);
    const [, after] = splitRuns(startBlock.runs, end.offset);
    replaceParagraphAt(
      model,
      start,
      setCommentAnchors(
        setBookmarks(
          { ...startBlock, runs: normalizeRuns([...before, ...after]) },
          anchorsAfterDelete(startBlock.bookmarks, start.offset, end.offset),
        ),
        anchorsAfterDelete(startBlock.commentAnchors, start.offset, end.offset),
      ),
    );
    return start;
  }
  // Cross-block: a boundary paragraph contributes its kept half; a boundary
  // table/image/break (always fully covered after resolveTableEdges)
  // contributes nothing and is removed whole.
  //
  // Floating drawings inside the removed span release their anchor pairing
  // (SPEC-078, deleteShape's rule) — collected before the splice below
  // rearranges the lists, released after it so the referrer scan sees the
  // post-delete document.
  const removedSpan: Block[] =
    start.section === end.section
      ? model.sections[start.section].blocks.slice(start.block, end.block + 1)
      : [
          ...model.sections[start.section].blocks.slice(start.block),
          ...model.sections.slice(start.section + 1, end.section).flatMap((section) => section.blocks),
          ...model.sections[end.section].blocks.slice(0, end.block + 1),
        ];
  const releasedOrdinals = removedSpan.flatMap((block) =>
    (isShape(block) || isChart(block) || isConnector(block)) && block.anchorOrdinal !== undefined ? [block.anchorOrdinal] : [],
  );
  const head =
    isTable(startBlock) || isImage(startBlock) || isShape(startBlock) || isChart(startBlock) || isConnector(startBlock) || isBreak(startBlock) || isToc(startBlock) || isTof(startBlock)
      ? null
      : splitRuns(startBlock.runs, start.offset)[0];
  const tail =
    isTable(endBlock) || isImage(endBlock) || isShape(endBlock) || isChart(endBlock) || isConnector(endBlock) || isBreak(endBlock) || isToc(endBlock) || isTof(endBlock)
      ? null
      : splitRuns(endBlock.runs, end.offset)[1];
  let replacement: Paragraph | null;
  if (head !== null && tail !== null) {
    const startPara = startBlock as Paragraph;
    const endPara = endBlock as Paragraph;
    replacement = setCommentAnchors(
      setBookmarks(
        { ...startPara, runs: normalizeRuns([...head, ...tail]) },
        mergedAnchors(startPara.bookmarks, endPara.bookmarks, start.offset, end.offset),
      ),
      mergedAnchors(startPara.commentAnchors, endPara.commentAnchors, start.offset, end.offset),
    );
  } else if (head !== null) {
    const startPara = startBlock as Paragraph;
    replacement = setCommentAnchors(
      setBookmarks(
        { ...startPara, runs: normalizeRuns(head) },
        anchorsAfterDelete(startPara.bookmarks, start.offset, blockLength(startPara)),
      ),
      anchorsAfterDelete(startPara.commentAnchors, start.offset, blockLength(startPara)),
    );
  } else if (tail !== null) {
    const endPara = endBlock as Paragraph;
    replacement = setCommentAnchors(
      setBookmarks(
        { ...endPara, runs: normalizeRuns(tail) },
        anchorsAfterDelete(endPara.bookmarks, 0, end.offset),
      ),
      anchorsAfterDelete(endPara.commentAnchors, 0, end.offset),
    );
  } else {
    replacement = null;
  }
  const inserted = replacement ? [replacement] : [];
  if (start.section === end.section) {
    const blocks = model.sections[start.section].blocks;
    blocks.splice(start.block, end.block - start.block + 1, ...inserted);
    for (const ordinal of releasedOrdinals) {
      releaseAnchorOrdinal(model, ordinal);
    }
    if (blocks.length === 0) {
      blocks.push({ type: "paragraph", runs: [] });
    }
    return {
      section: start.section,
      block: Math.min(start.block, blocks.length - 1),
      offset: head !== null ? start.offset : 0,
    };
  }
  const startSection = model.sections[start.section];
  const endSection = model.sections[end.section];
  startSection.blocks = [
    ...startSection.blocks.slice(0, start.block),
    ...inserted,
    ...endSection.blocks.slice(end.block + 1),
  ];
  model.sections.splice(start.section + 1, end.section - start.section);
  for (const ordinal of releasedOrdinals) {
    releaseAnchorOrdinal(model, ordinal);
  }
  if (startSection.blocks.length === 0) {
    startSection.blocks.push({ type: "paragraph", runs: [] });
  }
  return {
    section: start.section,
    block: Math.min(start.block, startSection.blocks.length - 1),
    offset: head !== null ? start.offset : 0,
  };
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

/** Merges a section into the one before it, keeping the earlier section's
 * own page setup (SPEC-028): the boundary between two sections IS the
 * section break, so deleting it dissolves the later section into the
 * earlier one. Blocks are concatenated; the later section's preserved
 * opaque blocks are re-anchored past the earlier section's blocks. Returns
 * the caret at the start of what was the later section's first block. */
function mergeSections(model: DocumentModel, sectionIndex: number): CaretPosition {
  const before = model.sections[sectionIndex - 1];
  const after = model.sections[sectionIndex];
  const joinIndex = before.blocks.length;
  before.blocks = [...before.blocks, ...after.blocks];
  if (after.opaqueBlocks && after.opaqueBlocks.length > 0) {
    before.opaqueBlocks = [
      ...(before.opaqueBlocks ?? []),
      ...after.opaqueBlocks.map((opaque) => ({
        ...opaque,
        blockIndex: opaque.blockIndex + joinIndex,
      })),
    ];
  }
  model.sections.splice(sectionIndex, 1);
  return { section: sectionIndex - 1, block: joinIndex, offset: 0 };
}

/** Backspace: deletes the character before the caret, or merges with the
 * previous paragraph when the caret is at a paragraph start. No-op at the
 * very start of the document — and at every table boundary (SPEC-022): at
 * the start of a cell's first paragraph, and just after a table, nothing is
 * deleted, so no keystroke ever merges cell content with outside text or
 * silently swallows a table. At the start of a section it deletes the
 * section break (the sections merge — SPEC-028 AC-10), and just after a
 * manual break block it deletes the break, rejoining the flow. */
export function deleteBackward(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const c = clampCaret(model, caret);
  const paragraph = paragraphAt(model, c);
  if (!paragraph) {
    return c;
  }
  if (c.offset > 0) {
    // Step by extended grapheme cluster so a surrogate pair or an emoji
    // ZWJ/variation-selector/skin-tone sequence deletes as a single unit
    // (SPEC-026).
    const prevOffset = previousGraphemeBoundary(blockText(paragraph), c.offset);
    return deleteRange(model, { ...c, offset: prevOffset }, c);
  }
  if (c.cell) {
    if (c.cell.block === 0) {
      return c;
    }
    const table = model.sections[c.section].blocks[c.block];
    if (!isTable(table)) {
      return c;
    }
    const prevBlocks = table.rows[c.cell.row].cells[c.cell.col].blocks;
    const prev: CaretPosition = {
      section: c.section,
      block: c.block,
      offset: blockLength(prevBlocks[c.cell.block - 1]),
      cell: { ...c.cell, block: c.cell.block - 1 },
    };
    return deleteRange(model, prev, c);
  }
  const prev = previousBlockEnd(model, c.section, c.block);
  if (!prev) {
    return c;
  }
  const prevBlock = model.sections[prev.section].blocks[prev.block];
  if (
    prev.section !== c.section &&
    (isTable(prevBlock) || isImage(prevBlock) || isShape(prevBlock) || isChart(prevBlock) || isConnector(prevBlock) || isBreak(prevBlock) || isToc(prevBlock) || isTof(prevBlock))
  ) {
    // Backspace at a section start whose previous section ends in an atom:
    // the section break itself is deleted (the sections merge), the atom
    // stays — one keystroke deletes one thing (SPEC-028 AC-10).
    return mergeSections(model, c.section);
  }
  if (isBreak(prevBlock)) {
    // Deleting a manual break rejoins the flow (SPEC-028 AC-10): the break
    // atom is removed whole, the caret's paragraph is kept as-is.
    return deleteRange(model, { section: prev.section, block: prev.block, offset: 0 }, c);
  }
  if (isTable(prevBlock) || isImage(prevBlock) || isShape(prevBlock) || isChart(prevBlock) || isConnector(prevBlock) || isToc(prevBlock) || isTof(prevBlock)) {
    return c;
  }
  return deleteRange(model, prev, c);
}

/** Forward delete: deletes the character after the caret, or merges with the
 * next paragraph when the caret is at a paragraph end. No-op at the very end
 * of the document — and at every table boundary (SPEC-022), mirroring
 * deleteBackward's guarantees, including its SPEC-028 break/section-break
 * deletion behaviour. */
export function deleteForward(
  model: DocumentModel,
  caret: CaretPosition,
): CaretPosition {
  const c = clampCaret(model, caret);
  const paragraph = paragraphAt(model, c);
  if (!paragraph) {
    return c;
  }
  if (c.offset < blockLength(paragraph)) {
    // Step by extended grapheme cluster (SPEC-026), mirroring deleteBackward.
    const nextOffset = nextGraphemeBoundary(blockText(paragraph), c.offset);
    return deleteRange(model, c, { ...c, offset: nextOffset });
  }
  if (c.cell) {
    const table = model.sections[c.section].blocks[c.block];
    if (!isTable(table)) {
      return c;
    }
    const cellBlocks = table.rows[c.cell.row].cells[c.cell.col].blocks;
    if (c.cell.block >= cellBlocks.length - 1) {
      return c;
    }
    const next: CaretPosition = {
      section: c.section,
      block: c.block,
      offset: 0,
      cell: { ...c.cell, block: c.cell.block + 1 },
    };
    return deleteRange(model, c, next);
  }
  const next = nextBlockStart(model, c.section, c.block);
  if (!next) {
    return c;
  }
  const nextBlock = model.sections[next.section].blocks[next.block];
  if (
    next.section !== c.section &&
    (isTable(nextBlock) || isImage(nextBlock) || isShape(nextBlock) || isChart(nextBlock) || isConnector(nextBlock) || isBreak(nextBlock) || isToc(nextBlock) || isTof(nextBlock))
  ) {
    // Delete at a section end whose next section starts with an atom: the
    // section break is deleted (the sections merge), the atom stays and the
    // caret does not move (SPEC-028 AC-10).
    mergeSections(model, next.section);
    return c;
  }
  if (isBreak(nextBlock)) {
    // Deleting a manual break just ahead rejoins the flow (SPEC-028 AC-10).
    return deleteRange(model, c, { section: next.section, block: next.block, offset: 1 });
  }
  if (isTable(nextBlock) || isImage(nextBlock) || isShape(nextBlock) || isChart(nextBlock) || isConnector(nextBlock) || isToc(nextBlock) || isTof(nextBlock)) {
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

/** The character style reference (SPEC-042) at a caret offset — the same
 * "run just before, else just after" rule as `formattingAtOffset`, used for
 * the Styles pane's caret readout (AC-10). */
export function runStyleIdAtOffset(block: Paragraph, offset: number): string | undefined {
  const [before, after] = splitRuns(block.runs, offset);
  if (before.length > 0) {
    return before[before.length - 1].styleId;
  }
  if (after.length > 0) {
    return after[0].styleId;
  }
  return undefined;
}

/** One paragraph's share of a range: its address plus the covered character
 * span within it (`from` may equal `to` for a paragraph the range passes
 * through without covering text — commands that address paragraphs rather
 * than characters still want those). */
export interface ParagraphSlice {
  address: ParagraphAddress;
  from: number;
  to: number;
}

function cellSlices(
  section: number,
  block: number,
  table: Table,
): ParagraphSlice[] {
  const slices: ParagraphSlice[] = [];
  for (let row = 0; row < table.rows.length; row++) {
    const cells = table.rows[row].cells;
    for (let col = 0; col < cells.length; col++) {
      const blocks = cells[col].blocks;
      for (let index = 0; index < blocks.length; index++) {
        slices.push({
          address: { section, block, cell: { row, col, block: index } },
          from: 0,
          to: blockLength(blocks[index]),
        });
      }
    }
  }
  return slices;
}

/** Every paragraph the range touches, with its covered span, in document
 * order — through table cells: a range confined to one cell yields that
 * cell's paragraphs, and a top-level range that covers a table yields every
 * paragraph of every cell (fully covered), which is what makes select-all
 * formatting reach table content (SPEC-022). Ends in different containers
 * are snapped by normalizeRange first, so a table is always either fully
 * covered or not touched at all. */
export function paragraphSlicesInRange(
  model: DocumentModel,
  a: CaretPosition,
  b: CaretPosition,
): ParagraphSlice[] {
  let { start, end } = normalizeRange(model, a, b);
  if (start.cell && end.cell) {
    const table = model.sections[start.section].blocks[start.block];
    if (!isTable(table)) {
      return [];
    }
    const blocks = table.rows[start.cell.row].cells[start.cell.col].blocks;
    const slices: ParagraphSlice[] = [];
    for (let index = start.cell.block; index <= end.cell.block; index++) {
      slices.push({
        address: {
          section: start.section,
          block: start.block,
          cell: { ...start.cell, block: index },
        },
        from: index === start.cell.block ? start.offset : 0,
        to: index === end.cell.block ? end.offset : blockLength(blocks[index]),
      });
    }
    return slices;
  }
  ({ start, end } = resolveTableEdges(model, start, end));
  if (compareCaret(start, end) > 0) {
    return [];
  }
  const slices: ParagraphSlice[] = [];
  for (let section = start.section; section <= end.section; section++) {
    const blocks = model.sections[section].blocks;
    const first = section === start.section ? start.block : 0;
    const last = section === end.section ? end.block : blocks.length - 1;
    for (let block = first; block <= last; block++) {
      const target = blocks[block];
      if (isTable(target)) {
        slices.push(...cellSlices(section, block, target));
        continue;
      }
      const from = section === start.section && block === start.block ? start.offset : 0;
      const to =
        section === end.section && block === end.block ? end.offset : blockLength(target);
      slices.push({ address: { section, block }, from, to });
    }
  }
  return slices;
}

/** Every run's formatting strictly between `start` and `end`, including
 * runs split at the range's boundaries. A paragraph the range merely passes
 * through without covering any of its characters (an empty paragraph, or a
 * boundary landing exactly between two of its runs) contributes nothing. */
export function runFormattingsInRange(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
): RunFormatting[] {
  const formattings: RunFormatting[] = [];
  for (const { address, from, to } of paragraphSlicesInRange(model, start, end)) {
    if (from >= to) {
      continue;
    }
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
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

/** Applies a format patch to every run between `start` and `end`, splitting
 * the runs straddling the range's boundaries first and merging the result
 * back down to the minimum number of runs afterward (SPEC-012 AC-7, AC-8,
 * AC-20, AC-21; SPEC-013 extends the patch beyond boolean toggles — see
 * PendingFormat for the per-key off semantics). */
export function patchFormatting(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
  patch: PendingFormat,
): void {
  for (const { address, from, to } of paragraphSlicesInRange(model, start, end)) {
    if (from >= to) {
      continue;
    }
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
      continue;
    }
    const [before, tail] = splitRuns(paragraph.runs, from);
    const [within, after] = splitRuns(tail, to - from);
    // Spread the run so a marker (field/note) keeps its identity when only
    // its formatting changes.
    const changed = within.map((run) => ({
      ...run,
      formatting: applyFormatPatch(run.formatting, patch),
    }));
    replaceParagraphAt(model, address, {
      ...paragraph,
      runs: normalizeRuns([...before, ...changed, ...after]),
    });
  }
}

/** Replaces every run's formatting outright between `start` and `end`
 * (SPEC-020): unlike `patchFormatting`, `formatting` fully replaces each
 * covered run's formatting rather than overlaying a patch onto it — what
 * "Clear formatting" (an empty object) and the format painter (a copied
 * `RunFormatting`) both need, since neither should inherit anything the
 * covered text already carried. */
export function replaceFormatting(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
  formatting: RunFormatting,
): void {
  for (const { address, from, to } of paragraphSlicesInRange(model, start, end)) {
    if (from >= to) {
      continue;
    }
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
      continue;
    }
    const [before, tail] = splitRuns(paragraph.runs, from);
    const [within, after] = splitRuns(tail, to - from);
    const changed = within.map((run) => ({ ...run, formatting: { ...formatting } }));
    replaceParagraphAt(model, address, {
      ...paragraph,
      runs: normalizeRuns([...before, ...changed, ...after]),
    });
  }
}

/** Replaces every run strictly between `start` and `end` with `mapper`'s
 * result, splitting the runs straddling the range's boundaries first and
 * merging back down afterward — the same split/merge shape as
 * `patchFormatting`/`replaceFormatting`, generalised to touch fields outside
 * `formatting` (SPEC-042's `Run.styleId`, which `applyCharacterStyle` sets
 * alongside a computed formatting layer `mapper` builds per run). */
export function mapRunsInRange(
  model: DocumentModel,
  start: CaretPosition,
  end: CaretPosition,
  mapper: (run: Run, paragraph: Paragraph) => Run,
): void {
  for (const { address, from, to } of paragraphSlicesInRange(model, start, end)) {
    if (from >= to) {
      continue;
    }
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
      continue;
    }
    const [before, tail] = splitRuns(paragraph.runs, from);
    const [within, after] = splitRuns(tail, to - from);
    const changed = within.map((run) => mapper(run, paragraph));
    replaceParagraphAt(model, address, {
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
