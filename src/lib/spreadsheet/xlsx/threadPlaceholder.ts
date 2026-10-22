// The thread placeholder (SPEC-135 "The thread placeholder" — THIS FILE is
// where its bytes live). Excel 365 writes every thread twice: in the
// threaded-comments part, and as a legacy `<comment ref xr:uid="{threadId}">`
// whose author is `tc={threadId}` and whose text is the preamble below, then
// "Comment:\n    " + the opening message's text and "\nReply:\n    " + each
// reply's, with a hidden VML note shape — so an older Excel shows the thread
// read-only and Excel 365 raises no repair prompt.
//
// The reader recognises a placeholder by its `tc=` author, never by its text
// (an older Excel may have edited it); the writer regenerates every
// placeholder from its thread on every write, so they cannot drift.
//
// THREAD_PLACEHOLDER_PREAMBLE and the `tc=` author form are provisional until
// pinned from the human-captured `e2e/fixtures/spec135-excel365-threads.xlsx`
// (SPEC-086's golden-first rule): where that file's bytes differ, the file
// wins and these constants follow it (xlsx/annotations.spec.ts compares them
// once the file is committed).
import type { CellThread } from "../model";
import { encodeXlsxAttribute, encodeXlsxText } from "./sharedStrings";

export const THREAD_PLACEHOLDER_PREAMBLE =
  "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file " +
  "is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n";

/** The author prefix of a thread's placeholder comment. */
export const PLACEHOLDER_AUTHOR_PREFIX = "tc=";

/** The author a thread's placeholder comment is written under. */
export function placeholderAuthor(threadId: string): string {
  return `${PLACEHOLDER_AUTHOR_PREFIX}${threadId}`;
}

/** Whether a legacy comment author marks a thread's placeholder. */
export function isPlaceholderAuthor(author: string): boolean {
  return author.startsWith(PLACEHOLDER_AUTHOR_PREFIX);
}

/** The thread id a placeholder author names (undefined: not a placeholder). */
export function placeholderThreadId(author: string): string | undefined {
  return isPlaceholderAuthor(author) ? author.slice(PLACEHOLDER_AUTHOR_PREFIX.length) : undefined;
}

/** A thread's placeholder text: the preamble, the opening message, one
 * "Reply:" block per reply. */
export function placeholderText(thread: Pick<CellThread, "messages">): string {
  const [opening, ...replies] = thread.messages;
  return `${THREAD_PLACEHOLDER_PREAMBLE}Comment:\n    ${opening?.text ?? ""}${replies.map((reply) => `\nReply:\n    ${reply.text}`).join("")}`;
}

/** A thread's placeholder `<comment>` as Excel 365 writes it: `ref`, the
 * placeholder author's `authorId`, `shapeId="0"`, `xr:uid` = the thread id,
 * one `<t>` (`xml:space="preserve"` only for leading or trailing
 * whitespace). `prefix` is the part's element prefix ("" or "x:"), `xr` the
 * prefix bound to the 2014 revision namespace. */
export function placeholderCommentXml(thread: Pick<CellThread, "threadId" | "messages">, ref: string, authorId: number, prefix = "", xr = "xr"): string {
  const p = prefix;
  const text = placeholderText(thread);
  const preserve = /^[ \t\n\r]|[ \t\n\r]$/.test(text) ? ` xml:space="preserve"` : "";
  return (
    `<${p}comment ref="${ref}" authorId="${authorId}" shapeId="0" ${xr}:uid="${encodeXlsxAttribute(thread.threadId)}">` +
    `<${p}text><${p}t${preserve}>${encodeXlsxText(text)}</${p}t></${p}text></${p}comment>`
  );
}
