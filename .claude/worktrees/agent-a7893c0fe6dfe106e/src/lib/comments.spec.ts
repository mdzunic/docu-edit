// SPEC-038 — comment thread commands and derived-visibility helpers.
import { describe, expect, it } from "vitest";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { caretSelection, type ModelSelection } from "./commands";
import { deleteRange as deleteRangeCommand, insertText as insertTextCommand } from "./commands";
import type { CaretPosition } from "./textEditing";
import {
  addReply,
  canModifyMessage,
  collectCommentAnchors,
  commentThreadById,
  commentThreadsAt,
  deleteMessage,
  deleteThread,
  editMessage,
  insertComment,
  reopenThread,
  resolveThread,
  visibleCommentThreads,
} from "./comments";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("comments-test");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function selection(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

describe("insertComment (AC-1..4)", () => {
  it("anchors a new thread to the selection and adds the opening message", () => {
    const model = docWithText("hello world");
    const result = insertComment(
      model,
      selection(caret(0, 0), caret(0, 5)),
      "Ada",
      "needs work",
      "t1",
      "2026-01-01T00:00:00.000Z",
    );
    expect(result.model.comments).toEqual([
      {
        threadId: "t1",
        resolved: false,
        messages: [
          { messageId: expect.any(String), author: "Ada", text: "needs work", createdAt: "2026-01-01T00:00:00.000Z" },
        ],
      },
    ]);
    const paragraphOut = result.model.sections[0].blocks[0] as Paragraph;
    expect(paragraphOut.commentAnchors).toEqual([
      { threadId: "t1", boundary: "start", offset: 0 },
      { threadId: "t1", boundary: "end", offset: 5 },
    ]);
    // The selection returned still covers the commented text.
    expect(result.selection).toEqual(selection(caret(0, 0), caret(0, 5)));
  });

  it("is a no-op for a collapsed selection", () => {
    const model = docWithText("hello");
    const result = insertComment(model, caretSelection(caret(0, 2)), "Ada", "x");
    expect(result.model.comments ?? []).toEqual([]);
  });

  it("defaults the thread/message ids and timestamp when not supplied", () => {
    const model = docWithText("hello");
    const result = insertComment(model, selection(caret(0, 0), caret(0, 3)), "Ada", "x");
    expect(result.model.comments?.[0].threadId).toEqual(expect.any(String));
    expect(result.model.comments?.[0].messages[0].createdAt).toEqual(expect.any(String));
  });
});

describe("visibleCommentThreads / collectCommentAnchors (document order)", () => {
  it("lists threads in document order by their start point", () => {
    let model = docWithText("aaa bbb", "ccc ddd");
    model = insertComment(model, selection(caret(1, 0), caret(1, 3)), "Ada", "second").model;
    model = insertComment(model, selection(caret(0, 0), caret(0, 3)), "Ada", "first").model;
    const threads = visibleCommentThreads(model);
    expect(threads.map((entry) => entry.thread.messages[0].text)).toEqual(["first", "second"]);
    // Document order is by position, not insertion order.
    const anchors = collectCommentAnchors(model);
    expect(anchors[0].start).toEqual(caret(0, 0));
    expect(anchors[1].start).toEqual(caret(1, 0));
  });

  it("commentThreadsAt finds threads whose range contains the caret", () => {
    const model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "x").model;
    expect(commentThreadsAt(model, caret(0, 2))).toHaveLength(1);
    expect(commentThreadsAt(model, caret(0, 8))).toHaveLength(0);
    // Inclusive of both boundary offsets.
    expect(commentThreadsAt(model, caret(0, 0))).toHaveLength(1);
    expect(commentThreadsAt(model, caret(0, 5))).toHaveLength(1);
  });
});

describe("replies, edit, delete (AC-5, AC-13, AC-14)", () => {
  function threaded(): DocumentModel {
    return insertComment(
      docWithText("hello world"),
      selection(caret(0, 0), caret(0, 5)),
      "Ada",
      "first",
      "t1",
      "2026-01-01T00:00:00.000Z",
    ).model;
  }

  it("appends a reply", () => {
    const model = threaded();
    const result = addReply(model, caretSelection(caret(0, 0)), "t1", "Bea", "reply text", "m2");
    const thread = commentThreadById(result.model, "t1");
    expect(thread?.messages.map((m) => m.text)).toEqual(["first", "reply text"]);
    expect(thread?.messages[1].author).toBe("Bea");
  });

  it("canModifyMessage only allows the message's own author", () => {
    const model = threaded();
    const message = commentThreadById(model, "t1")!.messages[0];
    expect(canModifyMessage(message, "Ada")).toBe(true);
    expect(canModifyMessage(message, "Bea")).toBe(false);
  });

  it("edits a message's text in place", () => {
    const model = threaded();
    const messageId = commentThreadById(model, "t1")!.messages[0].messageId;
    const result = editMessage(model, caretSelection(caret(0, 0)), "t1", messageId, "edited");
    expect(commentThreadById(result.model, "t1")?.messages[0].text).toBe("edited");
  });

  it("deletes one reply without removing the thread", () => {
    let model = threaded();
    model = addReply(model, caretSelection(caret(0, 0)), "t1", "Bea", "reply", "m2").model;
    const result = deleteMessage(model, caretSelection(caret(0, 0)), "t1", "m2");
    expect(commentThreadById(result.model, "t1")?.messages.map((m) => m.messageId)).toEqual([
      commentThreadById(model, "t1")!.messages[0].messageId,
    ]);
  });

  it("deleting the last message deletes the whole thread and its anchor", () => {
    const model = threaded();
    const messageId = commentThreadById(model, "t1")!.messages[0].messageId;
    const result = deleteMessage(model, caretSelection(caret(0, 0)), "t1", messageId);
    expect(result.model.comments ?? []).toEqual([]);
    expect((result.model.sections[0].blocks[0] as Paragraph).commentAnchors).toBeUndefined();
  });
});

describe("resolve / reopen (AC-9..12)", () => {
  it("resolving hides the thread from an open-only view but keeps it anchored", () => {
    const model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "x", "t1").model;
    const resolved = resolveThread(model, caretSelection(caret(0, 0)), "t1").model;
    expect(commentThreadById(resolved, "t1")?.resolved).toBe(true);
    // Still anchored/visible for the comments-only view — resolution only
    // affects the margin-pane highlight, not derived visibility.
    expect(visibleCommentThreads(resolved)).toHaveLength(1);
    const reopened = reopenThread(resolved, caretSelection(caret(0, 0)), "t1").model;
    expect(commentThreadById(reopened, "t1")?.resolved).toBe(false);
  });
});

describe("deleteThread (AC-15)", () => {
  it("removes both boundary anchors and the thread record", () => {
    const model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "x", "t1").model;
    const result = deleteThread(model, caretSelection(caret(0, 0)), "t1");
    expect(result.model.comments ?? []).toEqual([]);
    expect((result.model.sections[0].blocks[0] as Paragraph).commentAnchors).toBeUndefined();
  });

  it("only removes the named thread's anchors, leaving others intact", () => {
    let model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "a", "t1").model;
    model = insertComment(model, selection(caret(0, 6), caret(0, 11)), "Ada", "b", "t2").model;
    const result = deleteThread(model, caretSelection(caret(0, 0)), "t1");
    expect(result.model.comments?.map((t) => t.threadId)).toEqual(["t2"]);
    const anchors = (result.model.sections[0].blocks[0] as Paragraph).commentAnchors ?? [];
    expect(anchors.every((a) => a.threadId === "t2")).toBe(true);
  });
});

describe("anchors survive edits inside the range (AC-16) and vanish when the whole range is removed (AC-17)", () => {
  it("keeps the comment anchored when text is inserted inside the range", () => {
    let model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 11)), "Ada", "x", "t1").model;
    // Insert text in the middle of the commented range.
    const edited = insertTextCommand(model, caretSelection(caret(0, 5)), "-").model;
    expect(visibleCommentThreads(edited)).toHaveLength(1);
    const range = collectCommentAnchors(edited)[0];
    expect(range.start).toEqual(caret(0, 0));
    expect(range.end).toEqual(caret(0, 12)); // shifted right by the inserted character
  });

  it("keeps the comment anchored (shrunk) when part of the range is deleted", () => {
    let model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 11)), "Ada", "x", "t1").model;
    const edited = deleteRangeCommand(model, selection(caret(0, 2), caret(0, 4))).model;
    expect(visibleCommentThreads(edited)).toHaveLength(1);
    const range = collectCommentAnchors(edited)[0];
    expect(range.start).toEqual(caret(0, 0));
    expect(range.end).toEqual(caret(0, 9));
  });

  it("makes the thread invisible once the entire anchored range is deleted, but keeps the thread record", () => {
    let model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "x", "t1").model;
    const edited = deleteRangeCommand(model, selection(caret(0, 0), caret(0, 5))).model;
    expect(visibleCommentThreads(edited)).toEqual([]);
    // The thread record itself is untouched — undo (or a paste-back) can
    // still find it, mirroring lib/notes.ts's NoteContent contract.
    expect(commentThreadById(edited, "t1")).toBeDefined();
  });
});
