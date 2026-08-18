// Footnotes and endnotes (SPEC-031): pure helpers over the document model —
// collecting reference marks in document order, deriving displayed note
// numbers from the numbering settings, and addressing note content.
// Framework-free like the rest of lib/. Everything visible about notes
// (numbering, the note areas, both writers' output) derives from the
// reference marks in the body flow, never from `model.notes` alone — see
// NoteContent's "deleting a mark deletes its note" contract.
import {
  isBreak,
  isChart,
  isImage,
  isShape,
  isTable,
  isToc,
  isTof,
  type Block,
  type DocumentModel,
  type NoteContent,
  type NoteKind,
  type NoteNumberFormat,
  type NoteRestart,
  type NoteSettings,
  type Paragraph,
} from "./documentModel";
import { formatOrdinal } from "./core/lists";
import { blockText, type CellPath } from "./textEditing";

/** One reference mark's location in the body flow. `offset` is the mark's
 * character offset within its paragraph — the same coordinate space as
 * CaretPosition/BookmarkAnchor, which is what lets the page-placement
 * lookup reuse the pagination controller's char-windowed placements. */
export interface NoteRefLocation {
  noteId: string;
  kind: NoteKind;
  section: number;
  block: number;
  cell?: CellPath;
  offset: number;
}

function collectParagraphRefs(
  paragraph: Paragraph,
  kind: NoteKind,
  section: number,
  block: number,
  cell: CellPath | undefined,
  out: NoteRefLocation[],
): void {
  let offset = 0;
  for (const run of paragraph.runs) {
    if (run.note && run.note.kind === kind) {
      const location: NoteRefLocation = {
        noteId: run.note.noteId,
        kind,
        section,
        block,
        offset,
      };
      if (cell) {
        location.cell = cell;
      }
      out.push(location);
    }
    offset += run.text.length;
  }
}

/** Every reference mark of one note stream, in document order — through
 * table cells, the same walk order as paragraphSlicesInRange. Header/footer
 * content never holds note references (neither file format allows it), so
 * only the body flow is scanned. */
export function collectNoteRefs(model: DocumentModel, kind: NoteKind): NoteRefLocation[] {
  const refs: NoteRefLocation[] = [];
  for (let section = 0; section < model.sections.length; section += 1) {
    const blocks = model.sections[section].blocks;
    for (let block = 0; block < blocks.length; block += 1) {
      const target = blocks[block];
      if (isTable(target)) {
        for (let row = 0; row < target.rows.length; row += 1) {
          const cells = target.rows[row].cells;
          for (let col = 0; col < cells.length; col += 1) {
            const cellBlocks = cells[col].blocks;
            for (let index = 0; index < cellBlocks.length; index += 1) {
              collectParagraphRefs(
                cellBlocks[index],
                kind,
                section,
                block,
                { row, col, block: index },
                refs,
              );
            }
          }
        }
        continue;
      }
      if (isImage(target) || isShape(target) || isChart(target) || isBreak(target) || isToc(target) || isTof(target)) {
        continue;
      }
      collectParagraphRefs(target, kind, section, block, undefined, refs);
    }
  }
  return refs;
}

/** The note content a reference points at, or undefined for a dangling
 * reference (its note was never stored — a malformed source file). */
export function noteById(model: DocumentModel, noteId: string): NoteContent | undefined {
  return model.notes?.find((note) => note.noteId === noteId);
}

/** A note's plain text — the tooltip's body (AC-14). */
export function noteText(note: NoteContent): string {
  return note.blocks
    .map((block) => blockText(block))
    .join("\n")
    .trim();
}

/** One stream's numbering settings with every default resolved: footnotes
 * decimal, endnotes lowerRoman (Word's defaults), both from 1, continuous. */
export function effectiveNoteSettings(
  model: DocumentModel,
  kind: NoteKind,
): { numberFormat: NoteNumberFormat; startAt: number; restart: NoteRestart } {
  const settings: NoteSettings =
    (kind === "footnote" ? model.footnoteSettings : model.endnoteSettings) ?? {};
  return {
    numberFormat: settings.numberFormat ?? (kind === "footnote" ? "decimal" : "lowerRoman"),
    startAt: Math.max(1, Math.floor(settings.startAt ?? 1)),
    restart: settings.restart ?? "continuous",
  };
}

/** The traditional footnote symbol cycle (AC-17's "*"): * † ‡ §, doubling
 * every full cycle — the 5th note is **, the 9th ***. */
const NOTE_SYMBOLS = ["*", "†", "‡", "§"];

/** A note number rendered in a numbering format (1-based). */
export function formatNoteNumber(format: NoteNumberFormat, ordinal: number): string {
  const n = Math.max(1, Math.floor(ordinal));
  if (format === "symbol") {
    const symbol = NOTE_SYMBOLS[(n - 1) % NOTE_SYMBOLS.length];
    return symbol.repeat(Math.floor((n - 1) / NOTE_SYMBOLS.length) + 1);
  }
  return formatOrdinal(format, n);
}

/** The displayed number of every note of one stream, keyed by noteId —
 * derived, never persisted, like list numbering (AC-7..9). `refs` must be
 * collectNoteRefs' document-order list. Restart grouping: `"section"`
 * groups by the reference's section index; `"page"` groups by the page the
 * reference renders on, supplied by the caller (the pagination layer) —
 * without `pageOf` a page restart falls back to continuous numbering, so a
 * headless caller still gets stable labels. A note referenced twice (a
 * pasted copy of a mark) keeps its first reference's number. */
export function noteLabels(
  model: DocumentModel,
  kind: NoteKind,
  refs: NoteRefLocation[],
  pageOf?: (ref: NoteRefLocation, index: number) => number,
): Map<string, string> {
  const settings = effectiveNoteSettings(model, kind);
  const labels = new Map<string, string>();
  let counter = settings.startAt;
  let lastGroup: number | null = null;
  refs.forEach((ref, index) => {
    if (labels.has(ref.noteId)) {
      return;
    }
    let group: number | null = null;
    if (settings.restart === "section") {
      group = ref.section;
    } else if (settings.restart === "page" && pageOf) {
      group = pageOf(ref, index);
    }
    if (group !== null && lastGroup !== null && group !== lastGroup) {
      counter = settings.startAt;
    }
    lastGroup = group;
    labels.set(ref.noteId, formatNoteNumber(settings.numberFormat, counter));
    counter += 1;
  });
  return labels;
}

/** The ordered, referenced note contents of one stream — the notes a note
 * area (and both writers) actually emit, in first-reference order. Dangling
 * references (no stored content) are skipped; unreferenced contents are
 * invisible (see NoteContent). */
export function referencedNotes(
  model: DocumentModel,
  kind: NoteKind,
  refs: NoteRefLocation[] = collectNoteRefs(model, kind),
): NoteContent[] {
  const seen = new Set<string>();
  const notes: NoteContent[] = [];
  for (const ref of refs) {
    if (seen.has(ref.noteId)) {
      continue;
    }
    seen.add(ref.noteId);
    const note = noteById(model, ref.noteId);
    if (note) {
      notes.push(note);
    }
  }
  return notes;
}

/** `model` with one note's blocks replaced (copy-on-write on the notes
 * list) — the write primitive of note-content editing. */
export function withNoteBlocks(
  model: DocumentModel,
  noteId: string,
  blocks: Block[],
): DocumentModel {
  const notes = (model.notes ?? []).map((note) =>
    note.noteId === noteId ? { ...note, blocks } : note,
  );
  return { ...model, notes };
}

/** A fresh document-unique note id. Prefixed so ids read as what they are
 * in fixtures and saved JSON; uniqueness comes from the UUID. */
export function createNoteId(): string {
  return `note-${crypto.randomUUID()}`;
}

/** `paragraph` with every note reference run's placeholder text replaced by
 * its displayed label — the text the canvas actually renders (SPEC-031), so
 * text measurement wraps exactly like the rendered paragraph. Returns the
 * input object untouched (same identity) when it holds no note runs. */
export function paragraphWithNoteLabels(
  paragraph: Paragraph,
  labelOf: (noteId: string, kind: NoteKind) => string,
): Paragraph {
  if (!paragraph.runs.some((run) => run.note)) {
    return paragraph;
  }
  return {
    ...paragraph,
    runs: paragraph.runs.map((run) =>
      run.note ? { ...run, text: labelOf(run.note.noteId, run.note.kind) || run.text } : run,
    ),
  };
}

/** Maps an offset in a paragraph's rendered (label-substituted) text back
 * to model space — the inverse the measurement path needs so line-start
 * offsets stay model coordinates (see domCaret.renderedOffsetsToModel for
 * the DOM-side counterpart). An offset inside a label snaps to the mark's
 * start. */
export function renderedOffsetToModel(
  paragraph: Paragraph,
  labelOf: (noteId: string, kind: NoteKind) => string,
  renderedOffset: number,
): number {
  let rendered = 0;
  let model = 0;
  for (const run of paragraph.runs) {
    const renderedLength = run.note
      ? (labelOf(run.note.noteId, run.note.kind) || run.text).length
      : run.text.length;
    if (renderedOffset <= rendered + renderedLength) {
      const within = renderedOffset - rendered;
      return model + (run.note ? (within >= renderedLength ? run.text.length : 0) : within);
    }
    rendered += renderedLength;
    model += run.text.length;
  }
  return model;
}
