// SPEC-038 — comment threads through the .docx reader and writer:
// word/comments.xml with w:commentRangeStart/End and w:commentReference
// marks in the body.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "./zip";
import { createDocument, type DocumentModel, type Paragraph } from "../documentModel";
import { addReply, insertComment, resolveThread, visibleCommentThreads } from "../comments";
import { caretSelection, type ModelSelection } from "../commands";
import type { CaretPosition } from "../textEditing";

const decoder = new TextDecoder();

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("spec-038-docx");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function selection(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

async function writeAndReadBack(model: DocumentModel): Promise<{
  parsed: ParsedDocxDocument;
  parts: Map<string, Uint8Array>;
}> {
  const { bytes } = await writeDocx(model);
  const parts = await unzip(bytes);
  const parsed = await parseDocx(bytes);
  return { parsed, parts };
}

function parsedModel(parsed: ParsedDocxDocument): DocumentModel {
  const model = createDocument("spec-038-docx-reparsed");
  model.sections = parsed.sections;
  model.comments = parsed.comments;
  return model;
}

describe("comment round trip (AC-18, AC-21)", () => {
  it("writes comments.xml with commentRangeStart/End + commentReference in the body and reads it back", async () => {
    const model = insertComment(
      docWithText("hello world"),
      selection(caret(0, 0), caret(0, 5)),
      "Ada Lovelace",
      "needs work",
      "t1",
      "2026-01-01T00:00:00.000Z",
    ).model;
    const { parsed, parts } = await writeAndReadBack(model);

    expect(parts.has("word/comments.xml")).toBe(true);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain('<w:commentRangeStart w:id="1"/>');
    expect(documentXml).toContain('<w:commentRangeEnd w:id="1"/>');
    expect(documentXml).toContain('<w:commentReference w:id="1"/>');
    const commentsXml = decoder.decode(parts.get("word/comments.xml")!);
    expect(commentsXml).toContain('w:author="Ada Lovelace"');
    expect(commentsXml).toContain('w:date="2026-01-01T00:00:00.000Z"');
    expect(commentsXml).toContain("needs work");

    const reparsed = parsedModel(parsed);
    const threads = visibleCommentThreads(reparsed);
    expect(threads).toHaveLength(1);
    expect(threads[0].thread.messages).toEqual([
      {
        messageId: expect.any(String),
        author: "Ada Lovelace",
        text: "needs work",
        createdAt: "2026-01-01T00:00:00.000Z",
      },
    ]);
    expect(threads[0].start).toEqual(caret(0, 0));
    expect(threads[0].end).toEqual(caret(0, 5));
  });

  it("round-trips replies (AC-20)", async () => {
    let model = insertComment(
      docWithText("hello world"),
      selection(caret(0, 0), caret(0, 5)),
      "Ada",
      "first",
      "t1",
      "2026-01-01T00:00:00.000Z",
    ).model;
    model = addReply(model, caretSelection(caret(0, 0)), "t1", "Bea", "second reply", "m2", "2026-01-02T00:00:00.000Z").model;
    model = addReply(model, caretSelection(caret(0, 0)), "t1", "Cid", "third reply", "m3", "2026-01-03T00:00:00.000Z").model;

    const { parsed, parts } = await writeAndReadBack(model);
    const commentsXml = decoder.decode(parts.get("word/comments.xml")!);
    // One w:comment per thread (not per message) — replies are additional
    // paragraphs inside the same comment body.
    expect((commentsXml.match(/<w:comment /g) ?? []).length).toBe(1);
    expect(commentsXml).toContain("Bea (2026-01-02T00:00:00.000Z): second reply");
    expect(commentsXml).toContain("Cid (2026-01-03T00:00:00.000Z): third reply");

    const reparsed = parsedModel(parsed);
    const threads = visibleCommentThreads(reparsed);
    expect(threads).toHaveLength(1);
    expect(threads[0].thread.messages.map((m) => ({ author: m.author, text: m.text, createdAt: m.createdAt }))).toEqual([
      { author: "Ada", text: "first", createdAt: "2026-01-01T00:00:00.000Z" },
      { author: "Bea", text: "second reply", createdAt: "2026-01-02T00:00:00.000Z" },
      { author: "Cid", text: "third reply", createdAt: "2026-01-03T00:00:00.000Z" },
    ]);
  });

  it("does not preserve resolved state (documented .docx fidelity limit)", async () => {
    let model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "x", "t1").model;
    model = resolveThread(model, caretSelection(caret(0, 0)), "t1").model;
    const { parsed } = await writeAndReadBack(model);
    const reparsed = parsedModel(parsed);
    expect(visibleCommentThreads(reparsed)[0].thread.resolved).toBe(false);
  });

  it("drops a comment whose anchor was fully deleted (mirrors NoteContent's unreferenced-note drop)", async () => {
    const model = docWithText("hello world");
    model.comments = [
      { threadId: "orphan", resolved: false, messages: [{ messageId: "m1", author: "Ada", text: "gone", createdAt: "2026-01-01T00:00:00.000Z" }] },
    ];
    const { parts } = await writeAndReadBack(model);
    expect(parts.has("word/comments.xml")).toBe(false);
  });

  it("emits no comments part for a comment-free document", async () => {
    const { parts } = await writeAndReadBack(docWithText("plain text"));
    expect(parts.has("word/comments.xml")).toBe(false);
  });

  it("survives a second round trip without duplicating threads", async () => {
    const model = insertComment(docWithText("hello world"), selection(caret(0, 0), caret(0, 5)), "Ada", "x", "t1").model;
    const first = await writeAndReadBack(model);
    const again = await writeAndReadBack(parsedModel(first.parsed));
    const threads = visibleCommentThreads(parsedModel(again.parsed));
    expect(threads).toHaveLength(1);
    expect(threads[0].thread.messages).toHaveLength(1);
  });
});
