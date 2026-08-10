// SPEC-039 — tracked changes through the .docx reader and writer: w:ins,
// w:del/w:delText and w:rPrChange in the body, preserving author and date.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "./zip";
import { createDocument, type DocumentModel, type Paragraph, type Table } from "../documentModel";
import { firstStructuralDifference } from "../equality";

const decoder = new TextDecoder();

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("spec-039-docx");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

async function writeAndReadBack(model: DocumentModel): Promise<{
  parsed: ParsedDocxDocument;
  documentXml: string;
}> {
  const { bytes } = await writeDocx(model);
  const parts = await unzip(bytes);
  const parsed = await parseDocx(bytes);
  return { parsed, documentXml: decoder.decode(parts.get("word/document.xml")!) };
}

function parsedModel(parsed: ParsedDocxDocument): DocumentModel {
  const model = createDocument("spec-039-docx-reparsed");
  model.sections = parsed.sections;
  return model;
}

describe("tracked insertion round trip (AC-13)", () => {
  it("writes w:ins wrapping the run and reads it back with author and date", async () => {
    const model = docWithText("hello world");
    (model.sections[0].blocks[0] as Paragraph).runs = [
      { text: "hello ", formatting: {} },
      {
        text: "there ",
        formatting: {},
        changes: { insert: { changeId: "c1", author: "Ada Lovelace", date: "2026-01-01T00:00:00.000Z" } },
      },
      { text: "world", formatting: {} },
    ];
    const { parsed, documentXml } = await writeAndReadBack(model);
    expect(documentXml).toContain("<w:ins");
    expect(documentXml).toContain('w:author="Ada Lovelace"');
    expect(documentXml).toContain('w:date="2026-01-01T00:00:00.000Z"');

    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.text).join("")).toBe("hello there world");
    const inserted = runs.find((r) => r.text === "there ");
    expect(inserted?.changes?.insert).toMatchObject({ author: "Ada Lovelace", date: "2026-01-01T00:00:00.000Z" });
  });
});

describe("tracked deletion round trip (AC-13)", () => {
  it("writes w:del with w:delText (not w:t) and reads the text back unchanged", async () => {
    const model = docWithText("hello world");
    (model.sections[0].blocks[0] as Paragraph).runs = [
      {
        text: "hello ",
        formatting: {},
        changes: { delete: { changeId: "c2", author: "Grace Hopper", date: "2026-02-02T00:00:00.000Z" } },
      },
      { text: "world", formatting: {} },
    ];
    const { parsed, documentXml } = await writeAndReadBack(model);
    expect(documentXml).toContain("<w:del");
    expect(documentXml).toContain("<w:delText");
    expect(documentXml).not.toContain("<w:t>hello");

    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.text).join("")).toBe("hello world");
    const deleted = runs.find((r) => r.text === "hello ");
    expect(deleted?.changes?.delete).toMatchObject({ author: "Grace Hopper", date: "2026-02-02T00:00:00.000Z" });
  });
});

describe("tracked formatting change round trip (AC-13)", () => {
  it("writes w:rPrChange with the previous formatting and reads it back", async () => {
    const model = docWithText("hello world");
    (model.sections[0].blocks[0] as Paragraph).runs = [
      {
        text: "hello",
        formatting: { bold: true },
        changes: {
          formatChange: { changeId: "c3", author: "Ada", date: "2026-03-03T00:00:00.000Z", previous: {} },
        },
      },
      { text: " world", formatting: {} },
    ];
    const { parsed, documentXml } = await writeAndReadBack(model);
    expect(documentXml).toContain("<w:rPrChange");

    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    const changed = runs.find((r) => r.text === "hello");
    expect(changed?.formatting.bold).toBe(true);
    expect(changed?.changes?.formatChange).toMatchObject({
      author: "Ada",
      date: "2026-03-03T00:00:00.000Z",
      previous: {},
    });
  });
});

describe("nested tracked change round trip (AC-12, AC-13)", () => {
  it("writes nested w:ins/w:del and reads back both marks", async () => {
    const model = docWithText("hello world");
    (model.sections[0].blocks[0] as Paragraph).runs = [
      { text: "hello ", formatting: {} },
      {
        text: "there ",
        formatting: {},
        changes: {
          insert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" },
          delete: { changeId: "c2", author: "Grace", date: "2026-01-02T00:00:00.000Z" },
        },
      },
      { text: "world", formatting: {} },
    ];
    const { parsed, documentXml } = await writeAndReadBack(model);
    // Nesting: w:ins outermost, w:del inside it.
    expect(documentXml).toMatch(/<w:ins[^>]*>\s*<w:del/);

    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    const nested = runs.find((r) => r.text === "there ");
    expect(nested?.changes?.insert?.author).toBe("Ada");
    expect(nested?.changes?.delete?.author).toBe("Grace");
  });
});

describe("same-format save of an already-tracked document (AC-14)", () => {
  it("leaves an untouched tracked change structurally identical after a re-save", async () => {
    const model = docWithText("hello world");
    (model.sections[0].blocks[0] as Paragraph).runs = [
      { text: "hello ", formatting: {} },
      {
        text: "there ",
        formatting: {},
        changes: { insert: { changeId: "c1", author: "Ada Lovelace", date: "2026-01-01T00:00:00.000Z" } },
      },
      { text: "world", formatting: {} },
    ];
    const { parsed: firstParsed } = await writeAndReadBack(model);
    const reopened = parsedModel(firstParsed);
    const { parsed: secondParsed } = await writeAndReadBack(reopened);
    const resaved = parsedModel(secondParsed);
    expect(firstStructuralDifference(reopened, resaved)).toBeNull();
  });
});

describe("tracked paragraph mark round trip (AC-15)", () => {
  it("writes a split's markInsert as w:pPr/w:rPr/w:ins and reads it back", async () => {
    const model = docWithText("hello", "world");
    (model.sections[0].blocks[0] as Paragraph).changes = {
      markInsert: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" },
    };
    const { parsed, documentXml } = await writeAndReadBack(model);
    expect(documentXml).toMatch(/<w:pPr>.*<w:rPr><w:ins/);

    const reparsed = parsedModel(parsed);
    const first = reparsed.sections[0].blocks[0] as Paragraph;
    expect(first.changes?.markInsert).toMatchObject({ author: "Ada", date: "2026-01-01T00:00:00.000Z" });
  });

  it("writes a pending merge's markDelete as w:pPr/w:rPr/w:del and reads it back", async () => {
    const model = docWithText("hello", "world");
    (model.sections[0].blocks[0] as Paragraph).changes = {
      markDelete: { changeId: "c2", author: "Grace", date: "2026-02-02T00:00:00.000Z" },
    };
    const { parsed, documentXml } = await writeAndReadBack(model);
    expect(documentXml).toMatch(/<w:pPr>.*<w:rPr><w:del/);

    const reparsed = parsedModel(parsed);
    const first = reparsed.sections[0].blocks[0] as Paragraph;
    expect(first.changes?.markDelete).toMatchObject({ author: "Grace", date: "2026-02-02T00:00:00.000Z" });
  });
});

describe("tracked table row round trip (AC-15)", () => {
  function tableModel(): DocumentModel {
    const model = createDocument("spec-039-docx-table");
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        { cells: [{ blocks: [paragraph("a1")] }, { blocks: [paragraph("a2")] }] },
        {
          cells: [{ blocks: [paragraph("b1")] }, { blocks: [paragraph("b2")] }],
          changes: { insert: { changeId: "c3", author: "Ada", date: "2026-03-03T00:00:00.000Z" } },
        },
      ],
    };
    model.sections = [{ blocks: [table] }];
    return model;
  }

  it("writes a tracked-inserted row as w:trPr/w:ins and reads it back", async () => {
    const { parsed, documentXml } = await writeAndReadBack(tableModel());
    expect(documentXml).toContain("<w:trPr><w:ins");

    const reparsed = parsedModel(parsed);
    const table = reparsed.sections[0].blocks[0] as Table;
    expect(table.rows).toHaveLength(2);
    expect(table.rows[1].changes?.insert).toMatchObject({ author: "Ada", date: "2026-03-03T00:00:00.000Z" });
    expect(table.rows[0].changes).toBeUndefined();
  });
});

describe("persisted recording flag round trip (AC-3)", () => {
  it("writes w:trackChanges in settings.xml when the flag is set, and reads it back", async () => {
    const model = docWithText("hello world");
    model.trackChangesRecording = true;
    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    const settingsXml = decoder.decode(parts.get("word/settings.xml")!);
    expect(settingsXml).toContain("<w:trackChanges");

    const reparsed = await parseDocx(bytes);
    expect(reparsed.trackChangesRecording).toBe(true);
  });

  it("does not create a settings part when the flag is unset", async () => {
    const model = docWithText("hello world");
    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    expect(parts.has("word/settings.xml")).toBe(false);
  });
});
