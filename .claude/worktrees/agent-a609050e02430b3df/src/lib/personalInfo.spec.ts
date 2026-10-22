// SPEC-041 AC-12a..14 — "Remove personal information on save".
import { describe, expect, it } from "vitest";
import { caretSelection, insertText as insertTextCommand, type ModelSelection } from "./commands";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { addReply, insertComment } from "./comments";
import { applyTrackChanges } from "./trackChangesDiff";
import type { CaretPosition } from "./textEditing";
import { ANONYMIZED_DATE, stripPersonalInfo } from "./personalInfo";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("personal-info-test");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function selection(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

describe("stripPersonalInfo (AC-13)", () => {
  it("removes the document author and last-modified-by properties", () => {
    const model = docWithText("Hello");
    model.properties = { author: "Ada Lovelace", lastModifiedBy: "Grace Hopper", title: "Report" };
    const stripped = stripPersonalInfo(model);
    expect(stripped.properties?.author).toBeUndefined();
    expect(stripped.properties?.lastModifiedBy).toBeUndefined();
    // Everything else in `properties` survives untouched.
    expect(stripped.properties?.title).toBe("Report");
  });

  it("does not mutate the source model", () => {
    const model = docWithText("Hello");
    model.properties = { author: "Ada Lovelace" };
    stripPersonalInfo(model);
    expect(model.properties.author).toBe("Ada Lovelace");
  });

  it("leaves a document with no properties untouched", () => {
    const model = docWithText("Hello");
    const stripped = stripPersonalInfo(model);
    expect(stripped.properties).toBeUndefined();
  });
});

describe("stripPersonalInfo — comment authors and timestamps (AC-13, AC-14)", () => {
  it("replaces every comment author with a consistent 'Author N' label and strips the opening message's timestamp", () => {
    let model = docWithText("Hello world", "Second paragraph");
    const first = insertComment(model, selection(caret(0, 0), caret(0, 5)), "Ada Lovelace", "First");
    model = first.model;
    const second = insertComment(model, selection(caret(1, 0), caret(1, 6)), "Grace Hopper", "Second");
    model = second.model;
    // Ada comments again — same author, same label expected both times.
    model = addReply(model, selection(caret(0, 0), caret(0, 0)), model.comments![0].threadId, "Ada Lovelace", "Reply")
      .model;

    const stripped = stripPersonalInfo(model);
    const threads = stripped.comments!;
    expect(threads[0].messages[0].author).toBe("Author 1");
    expect(threads[0].messages[1].author).toBe("Author 1"); // same original author, same label
    expect(threads[1].messages[0].author).toBe("Author 2");
    expect(threads[0].messages[0].createdAt).toBe(ANONYMIZED_DATE);

    // The real names never leaked into the stripped copy anywhere.
    expect(JSON.stringify(stripped)).not.toContain("Ada Lovelace");
    expect(JSON.stringify(stripped)).not.toContain("Grace Hopper");

    // The original model is untouched.
    expect(model.comments![0].messages[0].author).toBe("Ada Lovelace");
  });
});

describe("stripPersonalInfo — tracked-change authors and timestamps (AC-13, AC-14)", () => {
  it("replaces tracked-change authors with 'Author N' labels, shared with comment authors when the name matches", () => {
    let model = docWithText("Hello world");
    // Ada comments first (becomes "Author 1")...
    model = insertComment(model, selection(caret(0, 0), caret(0, 5)), "Ada Lovelace", "note").model;
    // ...then Grace makes a tracked edit ("Author 2").
    const before = model;
    const afterInsert = insertTextCommand(model, caretSelection(caret(0, 5)), " there").model;
    model = applyTrackChanges(before, afterInsert, "Grace Hopper", "2026-01-01T00:00:00.000Z");
    // ...then Ada makes a tracked edit too — same label as her comment. The
    // inserted text ("XYZ") shares no characters with the existing text, so
    // the diff engine can only place it as its own run (nothing to align
    // with the existing "Hello there world" from the changeset).
    const before2 = model;
    const afterInsert2 = insertTextCommand(model, caretSelection(caret(0, 0)), "XYZ ").model;
    model = applyTrackChanges(before2, afterInsert2, "Ada Lovelace", "2026-01-02T00:00:00.000Z");

    const stripped = stripPersonalInfo(model);
    const run = (stripped.sections[0].blocks[0] as Paragraph).runs;
    const insertedByGrace = run.find((r) => r.text.includes("there"));
    const insertedByAda = run.find((r) => r.text.includes("XYZ"));
    expect(insertedByGrace?.changes?.insert?.author).toBe("Author 2");
    expect(insertedByAda?.changes?.insert?.author).toBe("Author 1");
    expect(insertedByGrace?.changes?.insert?.date).toBe(ANONYMIZED_DATE);

    expect(JSON.stringify(stripped)).not.toContain("Ada Lovelace");
    expect(JSON.stringify(stripped)).not.toContain("Grace Hopper");
  });
});
