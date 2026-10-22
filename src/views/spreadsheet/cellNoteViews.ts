// What SPEC-135's note chrome shows, derived from a sheet's notes and threads
// (useCellNotes binds these to the sheet on show): each cell's indicators,
// the hover popover's content, the shown notes' boxes and the active cell's
// annotation text for assistive technology. Plain data — no Vue — so it is
// unit-tested directly; every text stays the model's plain string (the
// components interpolate it, never as markup).
import type { CellRef } from "../../lib/spreadsheet/address";
import { NOTE_DEFAULT_BOX, noteAt, sheetNotes, sheetThreads, threadAt } from "../../lib/spreadsheet/annotations";
import type { CellNote, CellThread, Sheet } from "../../lib/spreadsheet/model";

/** A cell's indicators: a note, and a thread (open or resolved). */
export interface AnnotationMarks {
  note: boolean;
  thread: "open" | "resolved" | null;
}

/** What a hovered cell's popover shows (either part may be absent). */
export interface NotePopoverView {
  sheetId: string;
  ref: CellRef;
  note: { noteId: string; author: string; text: string } | null;
  thread: { threadId: string; author: string; text: string; replies: number; resolved: boolean } | null;
}

/** A shown note's box. */
export interface NoteBoxView {
  noteId: string;
  ref: CellRef;
  author: string;
  text: string;
  widthPx: number;
  heightPx: number;
}

export const marksKey = (row: number, col: number): string => `${row}:${col}`;

/** A note's box size (NOTE_DEFAULT_BOX when the file gave none). */
export const noteBoxSize = (note: CellNote): { widthPx: number; heightPx: number } => note.box ?? NOTE_DEFAULT_BOX;

/** A note shows as a box: it is `visible`, or every note shows. */
const shownAsBox = (note: CellNote, showAll: boolean): boolean => showAll || note.visible === true;

/** A sheet's indicators by `marksKey(row, col)`. */
export function annotationMarks(sheet: Pick<Sheet, "notes" | "threads">): Map<string, AnnotationMarks> {
  const map = new Map<string, AnnotationMarks>();
  for (const note of sheetNotes(sheet)) {
    map.set(marksKey(note.ref.row, note.ref.col), { note: true, thread: null });
  }
  for (const thread of sheetThreads(sheet)) {
    const key = marksKey(thread.ref.row, thread.ref.col);
    map.set(key, { note: map.get(key)?.note === true, thread: thread.resolved ? "resolved" : "open" });
  }
  return map;
}

/** "1 reply" / "N replies". */
export function repliesLabel(count: number): string {
  return count === 1 ? "1 reply" : `${count} replies`;
}

/** "Note by Alex Kim: …" / "Comment by Sam Lee: … 2 replies" — a cell's
 * annotations as one plain-text description ("" for none). */
export function annotationDescription(note: CellNote | undefined, thread: CellThread | undefined): string {
  const parts: string[] = [];
  if (note) {
    parts.push(`${note.author ? `Note by ${note.author}` : "Note"}: ${note.text}`);
  }
  const first = thread?.messages[0];
  if (thread && first) {
    const replies = thread.messages.length - 1;
    const state = thread.resolved ? " (resolved)" : "";
    parts.push(`Comment by ${first.author}${state}: ${first.text}${replies > 0 ? ` ${repliesLabel(replies)}` : ""}`);
  }
  return parts.join(". ");
}

/** The popover of `ref` on `sheet`, or null (nothing to show — a note
 * already shown as a box has no popover of its own). */
export function notePopoverView(sheet: Sheet, ref: CellRef, showAll: boolean): NotePopoverView | null {
  const note = noteAt(sheet, ref);
  const thread = threadAt(sheet, ref);
  const first = thread?.messages[0];
  const popNote = note && !shownAsBox(note, showAll) ? note : undefined;
  if (!popNote && !(thread && first)) {
    return null;
  }
  return {
    sheetId: sheet.id,
    ref,
    note: popNote ? { noteId: popNote.noteId, author: popNote.author, text: popNote.text } : null,
    thread: thread && first ? { threadId: thread.threadId, author: first.author, text: first.text, replies: thread.messages.length - 1, resolved: thread.resolved } : null,
  };
}

/** The notes shown as boxes on `sheet` — all but `exceptNoteId` (the one in
 * the editor). */
export function noteBoxViews(sheet: Pick<Sheet, "notes">, showAll: boolean, exceptNoteId: string | null): NoteBoxView[] {
  return sheetNotes(sheet)
    .filter((note) => shownAsBox(note, showAll) && note.noteId !== exceptNoteId)
    .map((note) => ({ noteId: note.noteId, ref: note.ref, author: note.author, text: note.text, ...noteBoxSize(note) }));
}
