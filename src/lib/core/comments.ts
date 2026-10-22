// Comment messages and their plain-text reply encoding (SPEC-038), moved down
// into core by SPEC-101's phase one — performed by SPEC-135, whose cell
// threads and notes need them in the spreadsheet product, which may import
// only core. The two-phase move documentModel.ts used for MediaPart and
// SPEC-097 used for core/urls.ts: the definitions live here and
// documentModel.ts / comments.ts re-export them, so every existing import
// path keeps working unchanged. Neutral string logic only: no document model.

/** One message in a comment thread (SPEC-038): the thread's opening comment
 * (`CommentThread.messages[0]`) or a reply. Plain text only — rich text
 * inside comments is out of scope — so `text` is rendered as plain text
 * everywhere (never interpreted as HTML/markup, AC "no markup injection").
 * Deliberately not named `id` — same equality.ts reason as BookmarkAnchor. */
export interface CommentMessage {
  messageId: string;
  /** The commenter's display name, from the settable identity
   * (core/commentIdentity.ts) at the time the message was written. */
  author: string;
  text: string;
  /** ISO 8601 timestamp. */
  createdAt: string;
}

/** Whether `author` may edit/delete `message` (AC "author can edit/delete
 * their own comments") — a plain name comparison against the settable
 * identity (core/commentIdentity.ts); there is no server-side identity in
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
// formats' readers/writers parse/produce it identically (and SPEC-135's
// .xlsx → .ods copy turns a cell thread into one note the same way).

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

/** A commenter's initials (the `.docx` writer's `w:initials` rule, SPEC-038):
 * the uppercased first letter of each whitespace-separated word of the
 * trimmed name, at most three — "" for a blank name. */
export function commentInitials(author: string): string {
  return author
    .trim()
    .split(/\s+/)
    .map((word) => word[0]?.toUpperCase() ?? "")
    .join("")
    .slice(0, 3);
}
