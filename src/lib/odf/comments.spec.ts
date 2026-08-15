// SPEC-038 — comment threads through the .odt reader and writer:
// office:annotation/office:annotation-end inline in content.xml.
import { describe, it, expect } from "vitest";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../core/zip";
import { createDocument, type DocumentModel, type Paragraph } from "../documentModel";
import { addReply, insertComment, resolveThread, visibleCommentThreads } from "../comments";
import { caretSelection, type ModelSelection } from "../commands";
import type { CaretPosition } from "../textEditing";

const decoder = new TextDecoder();

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("spec-038-odt");
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
  parsed: ParsedOdtDocument;
  contentXml: string;
}> {
  const { bytes } = await writeOdt(model);
  const parts = await unzip(bytes);
  const parsed = await parseOdt(bytes);
  return { parsed, contentXml: decoder.decode(parts.get("content.xml")!) };
}

function parsedModel(parsed: ParsedOdtDocument): DocumentModel {
  const model = createDocument("spec-038-odt-reparsed");
  model.sections = parsed.sections;
  model.comments = parsed.comments;
  return model;
}

describe("comment round trip (AC-19, AC-21)", () => {
  it("writes office:annotation/office:annotation-end and reads them back", async () => {
    const model = insertComment(
      docWithText("hello world"),
      selection(caret(0, 0), caret(0, 5)),
      "Ada Lovelace",
      "needs work",
      "t1",
      "2026-01-01T00:00:00.000Z",
    ).model;
    const { parsed, contentXml } = await writeAndReadBack(model);

    expect(contentXml).toContain("<office:annotation");
    expect(contentXml).toContain("office:annotation-end");
    expect(contentXml).toContain("<dc:creator>Ada Lovelace</dc:creator>");
    expect(contentXml).toContain("<dc:date>2026-01-01T00:00:00.000Z</dc:date>");
    expect(contentXml).toContain("needs work");

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

    const { parsed, contentXml } = await writeAndReadBack(model);
    // One annotation per thread (not per message) — replies are additional
    // paragraphs inside the same annotation body.
    expect((contentXml.match(/<office:annotation /g) ?? []).length).toBe(1);
    expect(contentXml).toContain("Bea (2026-01-02T00:00:00.000Z): second reply");
    expect(contentXml).toContain("Cid (2026-01-03T00:00:00.000Z): third reply");

    const reparsed = parsedModel(parsed);
    const threads = visibleCommentThreads(reparsed);
    expect(threads).toHaveLength(1);
    expect(threads[0].thread.messages.map((m) => ({ author: m.author, text: m.text, createdAt: m.createdAt }))).toEqual([
      { author: "Ada", text: "first", createdAt: "2026-01-01T00:00:00.000Z" },
      { author: "Bea", text: "second reply", createdAt: "2026-01-02T00:00:00.000Z" },
      { author: "Cid", text: "third reply", createdAt: "2026-01-03T00:00:00.000Z" },
    ]);
  });

  it("does not preserve resolved state (documented .odt fidelity limit)", async () => {
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
    const { contentXml } = await writeAndReadBack(model);
    expect(contentXml).not.toContain("office:annotation");
  });

  it("emits no annotation for a comment-free document", async () => {
    const { contentXml } = await writeAndReadBack(docWithText("plain text"));
    expect(contentXml).not.toContain("office:annotation");
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
