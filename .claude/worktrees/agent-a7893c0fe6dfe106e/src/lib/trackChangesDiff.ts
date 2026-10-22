// The track-changes diff engine (SPEC-039): converts an ordinary command's
// before/after model pair into tracked-change markup instead of touching
// each of commands.ts's ~40 editing commands individually. Every command
// already returns an immutable `{model, selection}` (see commands.ts's
// header comment) and leaves untouched blocks/rows/cells at the exact same
// array *reference* as `before` (cloneContainers only shallow-clones the
// containers on the path to the edit) — so this module finds "what changed"
// with cheap reference-equality trims, then character/formatting diffs only
// the handful of paragraphs (or rows) that actually differ.
//
// Applied once, centrally, in DocumentView's `applyEdit` when recording is
// on (see useTrackChanges.ts) — every command that inserts/deletes text,
// changes formatting, splits/merges a paragraph, or adds/removes a table
// row or image becomes tracked automatically, with no per-command wiring.
import {
  isImage,
  isTable,
  type Block,
  type CellPropertyChange,
  type DocumentModel,
  type FieldSpec,
  type Image,
  type NoteRef,
  type Paragraph,
  type RowPropertyChange,
  type Run,
  type RunFormatting,
  type Section,
  type Table,
  type TableCell,
  type TablePropertyChange,
  type TableRow,
  type TrackedChangeInfo,
  createChangeId,
  isMergeChipField,
} from "./documentModel";
import { windowOf } from "./core/arrayWindow";

function markInfo(author: string, date: string): TrackedChangeInfo {
  return { changeId: createChangeId(), author, date };
}

/** An LCS-matching key fragment for a field run — unique per field
 * "identity" so a diff never treats two different chips/page fields as the
 * same atom. Merge chips have no `format`; page/numpages fields have no
 * `name`, so the two branches build the tag from whichever fields the kind
 * actually carries. */
function fieldAtomTag(field: FieldSpec): string {
  if (isMergeChipField(field)) {
    return `${field.kind}:${field.kind === "endif" ? "" : field.name}`;
  }
  switch (field.kind) {
    case "page":
    case "numpages":
    case "date":
    case "time":
      return `${field.kind}:${field.format}`;
    case "docProperty":
      return `${field.kind}:${field.name}`;
    case "fileName":
      return `${field.kind}:${field.path ? "path" : ""}`;
    case "ref":
      return `${field.kind}:${field.target}:${field.show}`;
    case "pageRef":
      return `${field.kind}:${field.target}`;
    case "unsupported":
      return `${field.kind}:${field.code}`;
    default:
      return field.kind;
  }
}

/** Lazily-created marks shared by every atom a single `applyTrackChanges`
 * call touches: one edit event (a keystroke, a paste, a format toggle) is
 * one review-pane change, even when it spans many characters/runs. */
interface MarkPool {
  insert(): TrackedChangeInfo;
  delete(): TrackedChangeInfo;
  formatChange(): TrackedChangeInfo;
  /** One shared mark for every table/row/cell *property* change (SPEC-039
   * AC-15: a merge/split, a border/shading/padding/alignment edit, a row
   * height drag, a column-width/table-alignment/width-mode edit) a single
   * command produces — distinct from `formatChange`'s run/paragraph text
   * formatting so the two review-pane change kinds never collide within one
   * edit event. */
  propertyChange(): TrackedChangeInfo;
}

function createMarkPool(author: string, date: string): MarkPool {
  let insertMark: TrackedChangeInfo | undefined;
  let deleteMark: TrackedChangeInfo | undefined;
  let formatChangeMark: TrackedChangeInfo | undefined;
  let propertyChangeMark: TrackedChangeInfo | undefined;
  return {
    insert: () => (insertMark ??= markInfo(author, date)),
    delete: () => (deleteMark ??= markInfo(author, date)),
    formatChange: () => (formatChangeMark ??= markInfo(author, date)),
    propertyChange: () => (propertyChangeMark ??= markInfo(author, date)),
  };
}

function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) {
    return false;
  }
  const aKeys = Object.keys(a as Record<string, unknown>);
  const bKeys = Object.keys(b as Record<string, unknown>);
  if (aKeys.length !== bKeys.length) {
    return false;
  }
  return aKeys.every((key) =>
    deepEqual((a as Record<string, unknown>)[key], (b as Record<string, unknown>)[key]),
  );
}

function markRow(row: TableRow, mark: TrackedChangeInfo, kind: "insert" | "delete"): TableRow {
  return { ...row, changes: { ...row.changes, [kind]: mark } };
}

function markImage(image: Image, mark: TrackedChangeInfo, kind: "insert" | "delete"): Image {
  return { ...image, changes: { ...image.changes, [kind]: mark } };
}

// --- run/character-level diff (a single paragraph) --------------------------

interface Atom {
  /** LCS matching key: the literal character, or a synthetic tag for an
   * atomic field/note run (these never split — they are one caret unit). */
  key: string;
  text: string;
  formatting: RunFormatting;
  field?: FieldSpec;
  note?: NoteRef;
  changes?: Run["changes"];
}

function paragraphAtoms(runs: Run[]): Atom[] {
  const atoms: Atom[] = [];
  for (const run of runs) {
    if (run.field || run.note) {
      const tag = run.field
        ? `\u0000field:${fieldAtomTag(run.field)}`
        : `\u0000note:${run.note!.noteId}:${run.note!.kind}`;
      atoms.push({ key: tag, text: run.text, formatting: run.formatting, field: run.field, note: run.note, changes: run.changes });
      continue;
    }
    for (const ch of run.text) {
      atoms.push({ key: ch, text: ch, formatting: run.formatting, changes: run.changes });
    }
  }
  return atoms;
}

function atomsToRuns(atoms: Atom[]): Run[] {
  const runs: Run[] = [];
  for (const atom of atoms) {
    const last = runs[runs.length - 1];
    if (
      last &&
      !last.field &&
      !last.note &&
      !atom.field &&
      !atom.note &&
      deepEqual(last.formatting, atom.formatting) &&
      deepEqual(last.changes, atom.changes)
    ) {
      last.text += atom.text;
      continue;
    }
    const run: Run = { text: atom.text, formatting: atom.formatting };
    if (atom.field) run.field = atom.field;
    if (atom.note) run.note = atom.note;
    if (atom.changes) run.changes = atom.changes;
    runs.push(run);
  }
  return runs;
}

type AtomOp =
  | { type: "equal"; before: Atom; after: Atom }
  | { type: "delete"; atom: Atom }
  | { type: "insert"; atom: Atom };

// Above this before*after cell count, fall back to a plain whole-span
// replace instead of the O(n*m) LCS table — a large paste/find-replace-all
// still tracks correctly (as one delete + one insert span), just without
// fine-grained alignment of any incidentally-unchanged characters inside it.
const MAX_LCS_CELLS = 250_000;

function lcsMiddle(a: Atom[], b: Atom[]): AtomOp[] {
  const n = a.length;
  const m = b.length;
  if (n === 0 || m === 0 || n * m > MAX_LCS_CELLS) {
    return [...a.map((atom) => ({ type: "delete" as const, atom })), ...b.map((atom) => ({ type: "insert" as const, atom }))];
  }
  const dp: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i].key === b[j].key ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const ops: AtomOp[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i].key === b[j].key) {
      ops.push({ type: "equal", before: a[i], after: b[j] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      ops.push({ type: "delete", atom: a[i] });
      i++;
    } else {
      ops.push({ type: "insert", atom: b[j] });
      j++;
    }
  }
  while (i < n) ops.push({ type: "delete", atom: a[i++] });
  while (j < m) ops.push({ type: "insert", atom: b[j++] });
  return ops;
}

function diffAtoms(a: Atom[], b: Atom[]): AtomOp[] {
  const w = windowOf(a, b);
  const prefix: AtomOp[] = a.slice(0, w.prefixLen).map((atom, idx) => ({ type: "equal" as const, before: atom, after: b[idx] }));
  const middle = lcsMiddle(w.beforeMid, w.afterMid);
  const suffixStart = a.length - w.suffixLen;
  const bSuffixStart = b.length - w.suffixLen;
  const suffix: AtomOp[] = a.slice(suffixStart).map((atom, idx) => ({ type: "equal" as const, before: atom, after: b[bSuffixStart + idx] }));
  return [...prefix, ...middle, ...suffix];
}

/** Diffs one paragraph's runs (SPEC-039's core text-level tracking): real
 * insertions become `insert`-marked runs, real deletions are *retained* and
 * `delete`-marked instead of removed, and same-text formatting-only changes
 * become `formatChange` markers (`w:rPrChange`/`text:format-change`). */
function diffParagraphRuns(beforeRuns: Run[], afterRuns: Run[], author: string, date: string, pool: MarkPool): Run[] {
  const before = paragraphAtoms(beforeRuns);
  const after = paragraphAtoms(afterRuns);
  const ops = diffAtoms(before, after);
  const result: Atom[] = [];
  for (const op of ops) {
    if (op.type === "equal") {
      const { before: b, after: a } = op;
      if (deepEqual(b.formatting, a.formatting)) {
        result.push({ ...a, changes: a.changes });
        continue;
      }
      // Same text, different formatting: a formatting-only edit (bold
      // toggled, colour changed, …) — AC-5's "Formatted: Bold" case.
      const existing = b.changes?.formatChange;
      const mark = existing ?? pool.formatChange();
      result.push({
        ...a,
        changes: {
          ...b.changes,
          formatChange: { ...mark, previous: existing?.previous ?? b.formatting },
        },
      });
      continue;
    }
    if (op.type === "delete") {
      const atom = op.atom;
      const ownUnacceptedInsert = atom.changes?.insert && atom.changes.insert.author === author && !atom.changes.delete;
      if (ownUnacceptedInsert) {
        // Deleting text you yourself inserted (still pending) just removes
        // it — nothing was ever really "there" to redline.
        continue;
      }
      if (atom.changes?.delete) {
        // Already a pending deletion (e.g. re-selecting over already
        // struck-through text) — leave its existing mark untouched.
        result.push(atom);
        continue;
      }
      result.push({ ...atom, changes: { ...atom.changes, delete: pool.delete() } });
      continue;
    }
    // insert
    result.push({ ...op.atom, changes: { insert: pool.insert() } });
  }
  return atomsToRuns(mergeAdjacentMarks(result));
}

/** Relabels runs of consecutive same-author insert/delete marks to share
 * one `TrackedChangeInfo` (mutating each atom's `.changes` in place — safe,
 * these are fresh wrapper objects `paragraphAtoms` built for this call).
 * Without this, typing "XYZ" as three separate keystrokes — each its own
 * `applyTrackChanges` call — would give each character its own mark: three
 * single-character runs instead of one "XYZ" run, and three review-pane
 * entries instead of one. Every call re-scans the *whole* paragraph (old
 * marks carried over via the 'equal' path included), so marks from
 * different calls/keystrokes still merge as long as they end up adjacent —
 * covering both left-to-right typing and right-to-left backspacing. */
function mergeAdjacentMarks(atoms: Atom[]): Atom[] {
  mergeAdjacentMarksOfKind(atoms, "insert");
  mergeAdjacentMarksOfKind(atoms, "delete");
  return atoms;
}

function mergeAdjacentMarksOfKind(atoms: Atom[], kind: "insert" | "delete"): void {
  let current: TrackedChangeInfo | undefined;
  for (const atom of atoms) {
    const mark = atom.changes?.[kind];
    if (!mark) {
      current = undefined;
      continue;
    }
    if (current && current.author === mark.author) {
      atom.changes = { ...atom.changes, [kind]: current };
    } else {
      current = mark;
    }
  }
}

// --- paragraph-formatting diff -----------------------------------------------

function diffParagraphFormatting(before: Paragraph, after: Paragraph, pool: MarkPool): Paragraph["changes"] {
  const beforeFmt = before.formatting;
  const afterFmt = after.formatting;
  const listChanged = !deepEqual(before.list, after.list);
  const styleChanged = before.styleId !== after.styleId;
  const fmtChanged = !deepEqual(beforeFmt, afterFmt);
  if (!listChanged && !styleChanged && !fmtChanged) {
    return after.changes;
  }
  const existing = before.changes?.formatChange;
  const mark = existing ?? pool.formatChange();
  return {
    ...after.changes,
    formatChange: { ...mark, previous: existing?.previous ?? beforeFmt },
  };
}

// --- paragraph split/merge/insert/delete window resolution ------------------

function markParagraphRunsDeleted(paragraph: Paragraph, author: string, pool: MarkPool): Paragraph {
  const atoms = paragraphAtoms(paragraph.runs).map((atom) => {
    if (atom.changes?.delete) {
      return atom;
    }
    return { ...atom, changes: { ...atom.changes, delete: pool.delete() } };
  });
  return { ...paragraph, runs: atomsToRuns(atoms) };
}

function markParagraphRunsInserted(paragraph: Paragraph, pool: MarkPool): Paragraph {
  const mark = pool.insert();
  const atoms = paragraphAtoms(paragraph.runs).map((atom) => ({ ...atom, changes: { ...atom.changes, insert: mark } }));
  return { ...paragraph, runs: atomsToRuns(atoms), changes: { ...paragraph.changes, markInsert: mark } };
}

/** Resolves a window of paragraphs that changed between before/after
 * (always adjacent, since untouched paragraphs trim out by reference) into
 * tracked-change-aware paragraphs, recognising the two structural shapes
 * commands.ts actually produces: a 1:1 same-paragraph content edit, and a
 * split (1 -> 2, Enter) or merge (2 -> 1, Backspace at a paragraph start). */
function resolveParagraphWindow(beforeMid: Paragraph[], afterMid: Paragraph[], author: string, date: string, pool: MarkPool): Paragraph[] {
  if (beforeMid.length === 1 && afterMid.length === 1) {
    const before = beforeMid[0];
    const after = afterMid[0];
    const runs = diffParagraphRuns(before.runs, after.runs, author, date, pool);
    const changes = diffParagraphFormatting({ ...before, runs }, { ...after, runs }, pool);
    return [{ ...after, runs, changes }];
  }
  if (beforeMid.length === 0) {
    return afterMid.map((p) => markParagraphRunsInserted(p, pool));
  }
  if (afterMid.length === 0) {
    return beforeMid.map((p) => markParagraphRunsDeleted(p, author, pool));
  }
  if (beforeMid.length === 1 && afterMid.length === 2 && concatText(afterMid) === textOf(beforeMid[0])) {
    // Split (Enter): the first half's own trailing mark is the new one.
    const [head, tail] = afterMid;
    return [
      { ...head, changes: { ...head.changes, markInsert: pool.insert() } },
      { ...tail, changes: beforeMid[0].changes },
    ];
  }
  if (beforeMid.length === 2 && afterMid.length === 1 && concatText(beforeMid) === textOf(afterMid[0])) {
    // Merge (Backspace at a paragraph start): don't actually merge — keep
    // both paragraphs, marking the first's own mark pending-deleted.
    const [first, second] = beforeMid;
    return [{ ...first, changes: { ...first.changes, markDelete: pool.delete() } }, second];
  }
  // Fallback: an edit this engine doesn't specifically recognise (e.g. a
  // multi-paragraph paste). Never drop content — retain every removed
  // paragraph (marked deleted) ahead of every added one (marked inserted).
  return [...beforeMid.map((p) => markParagraphRunsDeleted(p, author, pool)), ...afterMid.map((p) => markParagraphRunsInserted(p, pool))];
}

function textOf(paragraph: Paragraph): string {
  return paragraph.runs.map((run) => run.text).join("");
}

function concatText(paragraphs: Paragraph[]): string {
  return paragraphs.map(textOf).join("");
}

function diffParagraphArray(before: Paragraph[], after: Paragraph[], author: string, date: string, pool: MarkPool): Paragraph[] {
  const w = windowOf(before, after);
  const middle = resolveParagraphWindow(w.beforeMid, w.afterMid, author, date, pool);
  return [...after.slice(0, w.prefixLen), ...middle, ...after.slice(after.length - w.suffixLen)];
}

// --- table diff --------------------------------------------------------------

function markCell(cell: TableCell, mark: TrackedChangeInfo, kind: "insert" | "delete"): TableCell {
  return { ...cell, changes: { ...cell.changes, [kind]: mark } };
}

/** A cell's pending property change (SPEC-039 AC-15): colSpan/rowSpan (a
 * merge or split), borders, shading, padding or vertical-alignment.
 * Restored verbatim if rejected, same shape as a run's formatChange. */
function diffCellPropertyChange(before: TableCell, after: TableCell, pool: MarkPool): CellPropertyChange | undefined {
  const changed =
    (before.colSpan ?? 1) !== (after.colSpan ?? 1) ||
    (before.rowSpan ?? 1) !== (after.rowSpan ?? 1) ||
    !deepEqual(before.borders, after.borders) ||
    before.shadingFill !== after.shadingFill ||
    before.paddingCm !== after.paddingCm ||
    before.verticalAlign !== after.verticalAlign;
  if (!changed) {
    return after.changes?.propertyChange;
  }
  const existing = before.changes?.propertyChange;
  const mark = existing ?? pool.propertyChange();
  return {
    ...mark,
    previous: existing?.previous ?? {
      colSpan: before.colSpan,
      rowSpan: before.rowSpan,
      borders: before.borders,
      shadingFill: before.shadingFill,
      paddingCm: before.paddingCm,
      verticalAlign: before.verticalAlign,
    },
  };
}

function diffCell(before: TableCell, after: TableCell, author: string, date: string, pool: MarkPool): TableCell {
  if (before === after) {
    return after;
  }
  const blocks = diffParagraphArray(before.blocks, after.blocks, author, date, pool);
  const propertyChange = diffCellPropertyChange(before, after, pool);
  const changes = propertyChange ? { ...after.changes, propertyChange } : after.changes;
  return { ...after, blocks, changes };
}

/** Windows a row's cell array for insert/delete (SPEC-039 AC-15's column
 * insert/delete) the same way `windowOf`/`diffParagraphArray` do for
 * paragraphs — untouched cells keep their exact reference (insertTableColumn/
 * deleteTableColumn only ever splice the affected column), so only the
 * genuinely added/removed cells land in the trimmed "middle" window. */
function diffRowCells(before: TableCell[], after: TableCell[], author: string, date: string, pool: MarkPool): TableCell[] {
  const w = windowOf(before, after);
  let middle: TableCell[];
  if (w.beforeMid.length === w.afterMid.length) {
    middle = w.beforeMid.map((cell, idx) => diffCell(cell, w.afterMid[idx], author, date, pool));
  } else if (w.beforeMid.length === 0) {
    middle = w.afterMid.map((cell) => markCell(cell, pool.insert(), "insert"));
  } else if (w.afterMid.length === 0) {
    middle = w.beforeMid.map((cell) => markCell(cell, pool.delete(), "delete"));
  } else {
    middle = [...w.beforeMid.map((cell) => markCell(cell, pool.delete(), "delete")), ...w.afterMid.map((cell) => markCell(cell, pool.insert(), "insert"))];
  }
  return [...after.slice(0, w.prefixLen), ...middle, ...after.slice(after.length - w.suffixLen)];
}

/** A row's pending property change (SPEC-039 AC-15): its minimum height. */
function diffRowPropertyChange(before: TableRow, after: TableRow, pool: MarkPool): RowPropertyChange | undefined {
  if (before.minHeightCm === after.minHeightCm) {
    return after.changes?.propertyChange;
  }
  const existing = before.changes?.propertyChange;
  const mark = existing ?? pool.propertyChange();
  return { ...mark, previous: existing?.previous ?? { minHeightCm: before.minHeightCm } };
}

function diffRow(before: TableRow, after: TableRow, author: string, date: string, pool: MarkPool): TableRow {
  if (before === after) {
    return after;
  }
  const cells = diffRowCells(before.cells, after.cells, author, date, pool);
  const propertyChange = diffRowPropertyChange(before, after, pool);
  const changes = propertyChange ? { ...after.changes, propertyChange } : after.changes;
  return { ...after, cells, changes };
}

/** A table's pending property change (SPEC-039 AC-15): alignment, width
 * mode/percent, or column widths — the last only when the column *count*
 * didn't change, so a column insert/delete's incidental width rescaling
 * (every other column narrows to keep the table's overall width) isn't
 * double-tracked alongside the cell insert/delete that is the real edit. */
function diffTablePropertyChange(before: Table, after: Table, pool: MarkPool): TablePropertyChange | undefined {
  const widthsChanged =
    before.columnWidthsCm.length === after.columnWidthsCm.length &&
    !deepEqual(before.columnWidthsCm, after.columnWidthsCm);
  const changed =
    widthsChanged ||
    before.alignment !== after.alignment ||
    before.widthMode !== after.widthMode ||
    before.widthPercent !== after.widthPercent;
  if (!changed) {
    return after.changes?.propertyChange;
  }
  const existing = before.changes?.propertyChange;
  const mark = existing ?? pool.propertyChange();
  return {
    ...mark,
    previous: existing?.previous ?? {
      columnWidthsCm: widthsChanged ? before.columnWidthsCm : undefined,
      alignment: before.alignment,
      widthMode: before.widthMode,
      widthPercent: before.widthPercent,
    },
  };
}

function diffTable(before: Table, after: Table, author: string, date: string, pool: MarkPool): Table {
  const w = windowOf(before.rows, after.rows);
  let middle: TableRow[];
  if (w.beforeMid.length === w.afterMid.length) {
    middle = w.beforeMid.map((row, idx) => diffRow(row, w.afterMid[idx], author, date, pool));
  } else if (w.beforeMid.length === 0) {
    middle = w.afterMid.map((row) => markRow(row, pool.insert(), "insert"));
  } else if (w.afterMid.length === 0) {
    middle = w.beforeMid.map((row) => markRow(row, pool.delete(), "delete"));
  } else {
    middle = [...w.beforeMid.map((row) => markRow(row, pool.delete(), "delete")), ...w.afterMid.map((row) => markRow(row, pool.insert(), "insert"))];
  }
  const rows = [...after.rows.slice(0, w.prefixLen), ...middle, ...after.rows.slice(after.rows.length - w.suffixLen)];
  const propertyChange = diffTablePropertyChange(before, after, pool);
  const changes = propertyChange ? { ...after.changes, propertyChange } : after.changes;
  return { ...after, rows, changes };
}

// --- section block diff -------------------------------------------------------

function resolveBlockWindow(beforeMid: Block[], afterMid: Block[], author: string, date: string, pool: MarkPool): Block[] {
  if (beforeMid.length === 1 && afterMid.length === 1) {
    const before = beforeMid[0];
    const after = afterMid[0];
    if (before.type === "paragraph" && after.type === "paragraph") {
      return resolveParagraphWindow([before], [after], author, date, pool);
    }
    if (isTable(before) && isTable(after)) {
      return [diffTable(before, after, author, date, pool)];
    }
    // Any other same-length pairing (an image replaced, a break edited, a
    // TOC regenerated, or a type change) is left as the after state
    // untracked — none of these are edit shapes SPEC-039 names, and doing
    // nothing here never loses content (the after state is what's kept).
    return [after];
  }
  if (beforeMid.every((b) => b.type === "paragraph") && afterMid.every((b) => b.type === "paragraph")) {
    return resolveParagraphWindow(beforeMid as Paragraph[], afterMid as Paragraph[], author, date, pool);
  }
  if (beforeMid.length === 0) {
    return afterMid.map((block) => markBlockInserted(block, pool));
  }
  if (afterMid.length === 0) {
    return beforeMid.map((block) => markBlockDeleted(block, author, pool));
  }
  // Fallback: retain everything, deletions first.
  return [...beforeMid.map((block) => markBlockDeleted(block, author, pool)), ...afterMid.map((block) => markBlockInserted(block, pool))];
}

function markBlockInserted(block: Block, pool: MarkPool): Block {
  if (isImage(block)) {
    return markImage(block, pool.insert(), "insert");
  }
  if (block.type === "paragraph") {
    return markParagraphRunsInserted(block, pool);
  }
  return block;
}

function markBlockDeleted(block: Block, author: string, pool: MarkPool): Block {
  if (isImage(block)) {
    return markImage(block, pool.delete(), "delete");
  }
  if (block.type === "paragraph") {
    return markParagraphRunsDeleted(block, author, pool);
  }
  return block;
}

function diffSection(before: Section, after: Section, author: string, date: string, pool: MarkPool): Section {
  if (before.blocks === after.blocks) {
    return after;
  }
  const w = windowOf(before.blocks, after.blocks);
  const middle = resolveBlockWindow(w.beforeMid, w.afterMid, author, date, pool);
  const blocks = [...after.blocks.slice(0, w.prefixLen), ...middle, ...after.blocks.slice(after.blocks.length - w.suffixLen)];
  return { ...after, blocks };
}

/** Rewrites `after` (the raw result of an ordinary command) so that every
 * real content change relative to `before` is tracked-change markup instead
 * of a plain edit: insertions are marked, deletions are retained (struck
 * through) rather than removed, and formatting-only changes become
 * `formatChange` markers. A no-op when `before === after` or neither model
 * has content differences the diff can find (nothing to track). */
export function applyTrackChanges(before: DocumentModel, after: DocumentModel, author: string, date: string): DocumentModel {
  if (before.sections === after.sections) {
    return after;
  }
  const pool = createMarkPool(author, date);
  const sections = before.sections.map((section, idx) => {
    const nextSection = after.sections[idx];
    if (!nextSection || section === nextSection) {
      return nextSection ?? section;
    }
    return diffSection(section, nextSection, author, date, pool);
  });
  // A section-count change (section break insert/merge, SPEC-028) isn't a
  // shape this diff engine resolves — pass the after sections through
  // untracked for any index beyond `before`'s length rather than lose them.
  const extra = after.sections.slice(before.sections.length);
  return { ...after, sections: [...sections, ...extra] };
}
