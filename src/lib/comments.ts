// Comment threads (SPEC-038): pure helpers over the document model — the
// insert/reply/edit/delete/resolve commands, deriving which threads are
// currently anchored (visible) from the live CommentAnchor points, and the
// document-order listing the comments-only view and margin pane both read.
// Framework-free like links.ts/notes.ts, and — like links.ts — a command
// module OUTSIDE commands.ts, with its own cloneContainers copy (see that
// file's comment on the split with commands.undo.spec.ts's completeness
// guard).
//
// Visibility is always DERIVED from the live commentAnchors side-lists,
// never cascade-deleted when text is edited: a thread survives in
// `model.comments` even once its anchor is gone, and reappears if undo (or
// a cut/paste of the anchored text) restores the anchor — the same
// contract lib/notes.ts's `referencedNotes`/`collectNoteRefs` implement for
// footnotes (SPEC-031). `visibleCommentThreads` is this module's
// `referencedNotes` equivalent.
import {
  isBreak,
  isImage,
  isShape,
  isTable,
  isToc,
  type CommentAnchor,
  type CommentMessage,
  type CommentThread,
  type DocumentModel,
  type Paragraph,
} from "./documentModel";
import {
  clampCaret,
  compareCaret,
  paragraphAt,
  replaceParagraphAt,
  type CaretPosition,
  type ParagraphAddress,
} from "./textEditing";
import {
  orderedRange,
  type CommandResult,
  type ModelSelection,
} from "./commands";

/** A fresh document-unique thread id. Prefixed so ids read as what they are
 * in fixtures and saved JSON; uniqueness comes from the UUID (see
 * lib/notes.ts's createNoteId for the same convention). */
export function createThreadId(): string {
  return `comment-${crypto.randomUUID()}`;
}

/** A fresh document-unique message id (a thread's opening comment or a
 * reply). */
export function createMessageId(): string {
  return `comment-msg-${crypto.randomUUID()}`;
}

// Same container-cloning as commands.ts/links.ts (see commands.ts's header
// comment for why this is duplicated rather than imported).
function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
  };
}

// Comments anchor only into the main body flow (section blocks and table
// cells) — the same scope lib/links.ts's bookmarks use, and for the same
// reason: header/footer content is edited through a throwaway synthetic
// single-section model (commands.ts's withHeaderFooterModel), so it is
// never reachable from `model.sections` walks like this one.
function forEachParagraph(
  model: DocumentModel,
  visit: (paragraph: Paragraph, address: ParagraphAddress) => void,
): void {
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isTable(block)) {
        block.rows.forEach((row, rowIndex) => {
          row.cells.forEach((cell, colIndex) => {
            cell.blocks.forEach((paragraph, cellBlockIndex) => {
              visit(paragraph, {
                section: sectionIndex,
                block: blockIndex,
                cell: { row: rowIndex, col: colIndex, block: cellBlockIndex },
              });
            });
          });
        });
      } else if (!isImage(block) && !isShape(block) && !isBreak(block) && !isToc(block)) {
        visit(block, { section: sectionIndex, block: blockIndex });
      }
    });
  });
}

// --- deriving visibility from live anchors ----------------------------------

/** One thread's resolved anchor range: its live start/end points, in model
 * coordinates — undefined for either end when the corresponding
 * CommentAnchor point is not present anywhere in the document (its range
 * was fully deleted, see the header comment). */
export interface CommentAnchorRange {
  threadId: string;
  start: CaretPosition;
  end: CaretPosition;
}

/** Every thread whose BOTH boundary points are still present in the
 * document AND still span at least one character, in document order (by
 * first-encountered start point) — the derived "this comment is still
 * anchored somewhere" set. A thread with only one of its two points left
 * (a partial deletion that missed the other boundary's paragraph) is not
 * returned: a range needs both ends to mean anything. Nor is a thread whose
 * start and end have collapsed onto the same point: deleting exactly the
 * commented text shifts an unmatched boundary anchor to the deletion point
 * rather than dropping it (the same "collapse, don't vanish" rule
 * textEditing.ts's anchorsAfterDelete applies to bookmarks), so both of a
 * fully-deleted range's anchors land on the same offset instead of
 * disappearing outright — without this check that would read as an
 * (empty, but present) comment surviving AC-17's "deleting the entire
 * range deletes the comment". */
export function collectCommentAnchors(model: DocumentModel): CommentAnchorRange[] {
  const starts = new Map<string, CaretPosition>();
  const ends = new Map<string, CaretPosition>();
  const order: string[] = [];
  forEachParagraph(model, (paragraph, address) => {
    for (const anchor of paragraph.commentAnchors ?? []) {
      const caret: CaretPosition = {
        section: address.section,
        block: address.block,
        offset: anchor.offset,
        cell: address.cell,
      };
      if (anchor.boundary === "start") {
        if (!starts.has(anchor.threadId)) {
          order.push(anchor.threadId);
        }
        starts.set(anchor.threadId, caret);
      } else {
        ends.set(anchor.threadId, caret);
      }
    }
  });
  const ranges: CommentAnchorRange[] = [];
  for (const threadId of order) {
    const start = starts.get(threadId);
    const end = ends.get(threadId);
    if (start && end && compareCaret(start, end) < 0) {
      ranges.push({ threadId, start, end });
    }
  }
  return ranges;
}

/** A visible (still-anchored) thread paired with its resolved range. */
export interface VisibleCommentThread {
  thread: CommentThread;
  start: CaretPosition;
  end: CaretPosition;
}

/** Every thread that is currently anchored somewhere in the document,
 * resolved or not, in document order — the set both the margin pane (which
 * additionally filters out resolved threads before painting a highlight,
 * AC-10) and the comments-only view (AC-22, which lists every thread and
 * lets the user filter) read from. A thread whose anchor was fully deleted
 * is simply absent — see the header comment; it is never removed from
 * `model.comments` itself by this function. */
export function visibleCommentThreads(model: DocumentModel): VisibleCommentThread[] {
  const found: VisibleCommentThread[] = [];
  for (const range of collectCommentAnchors(model)) {
    const thread = model.comments?.find((entry) => entry.threadId === range.threadId);
    if (thread) {
      found.push({ thread, start: range.start, end: range.end });
    }
  }
  return found;
}

/** The visible threads whose anchored range contains `caret` (inclusive of
 * both ends) — what "the caret/click landed inside a highlighted range"
 * resolves to (AC-7: clicking a highlighted range focuses its thread). */
export function commentThreadsAt(
  model: DocumentModel,
  caret: CaretPosition,
): VisibleCommentThread[] {
  const c = clampCaret(model, caret);
  return visibleCommentThreads(model).filter(
    (entry) => compareCaret(entry.start, c) <= 0 && compareCaret(c, entry.end) <= 0,
  );
}

/** A thread by id, regardless of whether it is currently anchored. */
export function commentThreadById(model: DocumentModel, threadId: string): CommentThread | undefined {
  return model.comments?.find((entry) => entry.threadId === threadId);
}

/** Whether `author` may edit/delete `message` (AC "author can edit/delete
 * their own comments") — a plain name comparison against the settable
 * identity (lib/commentIdentity.ts); there is no server-side identity in
 * this app to check against instead. */
export function canModifyMessage(message: CommentMessage, author: string): boolean {
  return message.author === author;
}

// --- reply text encoding (shared by the .docx and .odt writers/readers) ----
//
// Neither `.docx`'s `w:comment` nor `.odt`'s `office:annotation` has a
// native "this is a reply, by a different author, at a different time"
// concept without format-specific extensions (Word's w15 commentsExtended
// part; no ODF equivalent at all) — and comments are plain text only (rich
// text is out of scope), so both writers represent a thread as ONE comment
// element whose body holds one paragraph per message: the opening message's
// text verbatim, then each reply as plain text prefixed with its author and
// timestamp. This is the one place that encoding is defined, so both
// formats' readers/writers parse/produce it identically.

/** A reply message's plain-text paragraph encoding: `"{author} ({createdAt}): {text}"`. */
export function replyParagraphText(message: CommentMessage): string {
  return `${message.author} (${message.createdAt}): ${message.text}`;
}

const REPLY_TEXT_PATTERN = /^(.+) \(([^()]+)\): ([\s\S]*)$/;

/** Reverses `replyParagraphText`, or null when `text` doesn't match the
 * shape (including every paragraph of a comment this app did not write,
 * since a foreign multi-paragraph comment has no reply structure to
 * recover) — the caller folds a non-matching paragraph's text into the
 * previous message instead of losing it. */
export function parseReplyParagraphText(
  text: string,
): { author: string; createdAt: string; text: string } | null {
  const match = REPLY_TEXT_PATTERN.exec(text);
  if (!match) {
    return null;
  }
  const createdAt = match[2];
  if (Number.isNaN(Date.parse(createdAt))) {
    return null;
  }
  return { author: match[1], createdAt, text: match[3] };
}

// --- commands -----------------------------------------------------------

function addCommentAnchorPoint(
  model: DocumentModel,
  address: ParagraphAddress,
  anchor: CommentAnchor,
): boolean {
  const paragraph = paragraphAt(model, address);
  if (!paragraph) {
    return false;
  }
  const anchors = [...(paragraph.commentAnchors ?? []), anchor].sort((a, b) => a.offset - b.offset);
  replaceParagraphAt(model, address, { ...paragraph, commentAnchors: anchors });
  return true;
}

/** Anchors a new comment thread to the selection's range and adds its
 * opening message (AC "Selecting text... highlights the range"/"Submitting
 * adds the comment with an author name and timestamp"). A collapsed
 * selection, or one that lands on a table atom rather than inside a
 * paragraph, is a no-op (the model is still cloned so the caller always
 * gets a fresh state to hand to applyEdit). The selection returned covers
 * the same range that was passed in, so the caller's highlight/composer
 * keep tracking the just-commented text. Pass `threadId`/`now` to control
 * them (tests do); otherwise a fresh id/timestamp is generated. */
export function insertComment(
  model: DocumentModel,
  selection: ModelSelection,
  author: string,
  text: string,
  threadId: string = createThreadId(),
  now: string = new Date().toISOString(),
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  if (compareCaret(start, end) === 0) {
    return { model: next, selection };
  }
  if (!addCommentAnchorPoint(next, start, { threadId, boundary: "start", offset: start.offset })) {
    return { model: next, selection };
  }
  if (!addCommentAnchorPoint(next, end, { threadId, boundary: "end", offset: end.offset })) {
    return { model: next, selection };
  }
  const message: CommentMessage = { messageId: createMessageId(), author, text, createdAt: now };
  const thread: CommentThread = { threadId, resolved: false, messages: [message] };
  next.comments = [...(next.comments ?? []), thread];
  return { model: next, selection: { anchor: start, focus: end } };
}

/** Appends a reply to a thread (AC "Comment threads support replies").
 * Unknown thread ids are a no-op. The selection passes through unchanged —
 * replying never moves the body caret. */
export function addReply(
  model: DocumentModel,
  selection: ModelSelection,
  threadId: string,
  author: string,
  text: string,
  messageId: string = createMessageId(),
  now: string = new Date().toISOString(),
): CommandResult {
  const message: CommentMessage = { messageId, author, text, createdAt: now };
  const comments = (model.comments ?? []).map((thread) =>
    thread.threadId === threadId ? { ...thread, messages: [...thread.messages, message] } : thread,
  );
  return { model: { ...model, comments }, selection };
}

/** Replaces one message's text in place (AC "author can edit... their own
 * comments") — the caller is responsible for checking `canModifyMessage`
 * first. Unknown thread/message ids are a no-op. */
export function editMessage(
  model: DocumentModel,
  selection: ModelSelection,
  threadId: string,
  messageId: string,
  text: string,
): CommandResult {
  const comments = (model.comments ?? []).map((thread) =>
    thread.threadId === threadId
      ? {
          ...thread,
          messages: thread.messages.map((message) =>
            message.messageId === messageId ? { ...message, text } : message,
          ),
        }
      : thread,
  );
  return { model: { ...model, comments }, selection };
}

/** Removes one message from a thread (AC "author can... delete their own
 * comments") — the caller is responsible for checking `canModifyMessage`
 * first. Deleting a thread's last remaining message deletes the whole
 * thread, anchor included (there is nothing left to anchor). Unknown
 * thread/message ids are a no-op. */
export function deleteMessage(
  model: DocumentModel,
  selection: ModelSelection,
  threadId: string,
  messageId: string,
): CommandResult {
  const thread = model.comments?.find((entry) => entry.threadId === threadId);
  if (!thread) {
    return { model, selection };
  }
  const messages = thread.messages.filter((message) => message.messageId !== messageId);
  if (messages.length === 0) {
    return deleteThread(model, selection, threadId);
  }
  const comments = (model.comments ?? []).map((entry) =>
    entry.threadId === threadId ? { ...entry, messages } : entry,
  );
  return { model: { ...model, comments }, selection };
}

function setThreadResolved(model: DocumentModel, threadId: string, resolved: boolean): DocumentModel {
  const comments = (model.comments ?? []).map((thread) =>
    thread.threadId === threadId ? { ...thread, resolved } : thread,
  );
  return { ...model, comments };
}

/** Resolves a thread (AC "A thread can be resolved"): its highlight
 * disappears and it moves to the comments UI's "Resolved" section. The
 * selection passes through unchanged. */
export function resolveThread(
  model: DocumentModel,
  selection: ModelSelection,
  threadId: string,
): CommandResult {
  return { model: setThreadResolved(model, threadId, true), selection };
}

/** Reopens a resolved thread (AC "A resolved thread can be reopened."). */
export function reopenThread(
  model: DocumentModel,
  selection: ModelSelection,
  threadId: string,
): CommandResult {
  return { model: setThreadResolved(model, threadId, false), selection };
}

/** Removes both boundary anchors of `threadId` from every paragraph that
 * carries one, and drops the thread from `model.comments` (AC "deleting a
 * thread removes its anchor"). Unlike a text edit fully swallowing a
 * range (which leaves the thread behind, invisible — see the header
 * comment), this is the explicit "delete this comment" action: nothing is
 * left to resurrect on undo beyond the ordinary undo stack. */
export function deleteThread(
  model: DocumentModel,
  selection: ModelSelection,
  threadId: string,
): CommandResult {
  const next = cloneContainers(model);
  const updates: Array<{ address: ParagraphAddress; anchors: CommentAnchor[] }> = [];
  forEachParagraph(next, (paragraph, address) => {
    const anchors = paragraph.commentAnchors;
    if (!anchors?.some((anchor) => anchor.threadId === threadId)) {
      return;
    }
    updates.push({ address, anchors: anchors.filter((anchor) => anchor.threadId !== threadId) });
  });
  for (const { address, anchors } of updates) {
    const paragraph = paragraphAt(next, address);
    if (!paragraph) {
      continue;
    }
    const replacement: Paragraph = { ...paragraph };
    if (anchors.length > 0) {
      replacement.commentAnchors = anchors;
    } else {
      delete replacement.commentAnchors;
    }
    replaceParagraphAt(next, address, replacement);
  }
  next.comments = (next.comments ?? []).filter((thread) => thread.threadId !== threadId);
  return { model: next, selection };
}
