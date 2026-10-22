// The comments pane's rows (SPEC-135 "Threads"), framework-free: what
// SheetCommentsPanel shows for a workbook — its threads by sheet in tab
// order, each sheet's by (row, col), under the filter; in an .ods workbook,
// which holds notes only, its notes instead. useSheetComments computes them
// from the model, the filter and the comment identity. Texts stay plain
// strings; the panel renders them with text interpolation only.
import { canModifyMessage, commentInitials } from "../../lib/core/comments";
import { formatTimestamp } from "../../lib/format";
import { formatA1, type CellRef } from "../../lib/spreadsheet/address";
import { sheetNotes, sheetThreads } from "../../lib/spreadsheet/annotations";
import type { SpreadsheetModel, ThreadMessage } from "../../lib/spreadsheet/model";

export type CommentFilter = "all" | "unresolved";

/** One message row as the panel shows it. */
export interface CommentMessageView {
  messageId: string;
  author: string;
  initials: string;
  time: string;
  text: string;
  /** The identity wrote it: it may be edited and deleted. */
  own: boolean;
}

/** One thread or (in an .ods workbook) one note in the panel. */
export type CommentItemView =
  | { kind: "thread"; threadId: string; ref: CellRef; address: string; resolved: boolean; messages: CommentMessageView[] }
  | { kind: "note"; noteId: string; ref: CellRef; address: string; author: string; initials: string; time: string; text: string };

/** A sheet's threads (or notes) under its name. */
export interface CommentGroupView {
  sheetId: string;
  sheetName: string;
  items: CommentItemView[];
}

/** A message's time as the panel shows it ("" when it has none). */
export function messageTime(createdAt: string | undefined): string {
  const at = createdAt ? Date.parse(createdAt) : Number.NaN;
  return Number.isNaN(at) ? "" : formatTimestamp(new Date(at));
}

/** The panel's groups: each sheet with a thread (or, `notesOnly`, a note)
 * under the filter, in tab order; `author` is the comment identity (whose
 * own messages may be edited and deleted). */
export function commentGroupViews(model: Pick<SpreadsheetModel, "sheets">, options: { notesOnly: boolean; filter: CommentFilter; author: string }): CommentGroupView[] {
  const messageView = (message: ThreadMessage): CommentMessageView => ({
    messageId: message.messageId,
    author: message.author,
    initials: commentInitials(message.author),
    time: messageTime(message.createdAt),
    text: message.text,
    own: canModifyMessage(message, options.author),
  });
  return model.sheets.flatMap((sheet): CommentGroupView[] => {
    const items: CommentItemView[] = options.notesOnly
      ? sheetNotes(sheet).map((note) => ({
          kind: "note",
          noteId: note.noteId,
          ref: note.ref,
          address: formatA1(note.ref),
          author: note.author,
          initials: commentInitials(note.author),
          time: messageTime(note.createdAt),
          text: note.text,
        }))
      : sheetThreads(sheet)
          .filter((thread) => options.filter === "all" || !thread.resolved)
          .map((thread) => ({
            kind: "thread",
            threadId: thread.threadId,
            ref: thread.ref,
            address: formatA1(thread.ref),
            resolved: thread.resolved,
            messages: thread.messages.map(messageView),
          }));
    return items.length > 0 ? [{ sheetId: sheet.id, sheetName: sheet.name, items }] : [];
  });
}
