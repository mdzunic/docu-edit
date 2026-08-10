// SPEC-039 — tracked changes through the .odt reader and writer:
// text:tracked-changes/text:changed-region with text:change-start/-end
// marks in the body, preserving author and date.
import { describe, it, expect } from "vitest";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../docx/zip";
import { createDocument, type DocumentModel, type Paragraph, type Table } from "../documentModel";
import { firstStructuralDifference } from "../equality";

const decoder = new TextDecoder();

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("spec-039-odt");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
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
  const model = createDocument("spec-039-odt-reparsed");
  model.sections = parsed.sections;
  return model;
}

describe("tracked insertion round trip (AC-13)", () => {
  it("writes text:change-start/-end plus a changed-region and reads it back with author and date", async () => {
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
    const { parsed, contentXml } = await writeAndReadBack(model);
    expect(contentXml).toContain("<text:tracked-changes");
    expect(contentXml).toContain("<text:insertion>");
    expect(contentXml).toContain("<text:change-start");
    expect(contentXml).toContain("<dc:creator>Ada Lovelace</dc:creator>");
    expect(contentXml).toContain("<dc:date>2026-01-01T00:00:00.000Z</dc:date>");

    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.text).join("")).toBe("hello there world");
    const inserted = runs.find((r) => r.text === "there ");
    expect(inserted?.changes?.insert).toMatchObject({ author: "Ada Lovelace", date: "2026-01-01T00:00:00.000Z" });
  });
});

describe("tracked deletion round trip (AC-13)", () => {
  it("writes text:change-start/-end and reads the deleted text back unchanged", async () => {
    const model = docWithText("hello world");
    (model.sections[0].blocks[0] as Paragraph).runs = [
      {
        text: "hello ",
        formatting: {},
        changes: { delete: { changeId: "c2", author: "Grace Hopper", date: "2026-02-02T00:00:00.000Z" } },
      },
      { text: "world", formatting: {} },
    ];
    const { parsed, contentXml } = await writeAndReadBack(model);
    expect(contentXml).toContain("<text:deletion>");

    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.text).join("")).toBe("hello world");
    const deleted = runs.find((r) => r.text === "hello ");
    expect(deleted?.changes?.delete).toMatchObject({ author: "Grace Hopper", date: "2026-02-02T00:00:00.000Z" });
  });
});

describe("nested tracked change round trip (AC-12, AC-13)", () => {
  it("writes both marks and reads back both", async () => {
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
    const { parsed } = await writeAndReadBack(model);
    const reparsed = parsedModel(parsed);
    const runs = (reparsed.sections[0].blocks[0] as Paragraph).runs;
    const nested = runs.find((r) => r.text === "there ");
    expect(nested?.changes?.insert?.author).toBe("Ada");
    expect(nested?.changes?.delete?.author).toBe("Grace");
  });
});

describe("same-format save of an already-tracked document (AC-14 analog)", () => {
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

describe("tracked paragraph mark round trip (AC-16)", () => {
  it("writes a split's markInsert as a trailing text:change and reads it back", async () => {
    const model = docWithText("hello", "world");
    (model.sections[0].blocks[0] as Paragraph).changes = {
      markInsert: { changeId: "pm1", author: "Ada", date: "2026-01-01T00:00:00.000Z" },
    };
    const { parsed, contentXml } = await writeAndReadBack(model);
    expect(contentXml).toContain("<text:insertion>");
    expect(contentXml).toMatch(/hello<\/text:span><text:change text:change-id="pm1"\/>/);

    const reparsed = parsedModel(parsed);
    const first = reparsed.sections[0].blocks[0] as Paragraph;
    expect(first.changes?.markInsert).toMatchObject({ author: "Ada", date: "2026-01-01T00:00:00.000Z" });
  });

  it("writes a pending merge's markDelete as a trailing text:change and reads it back", async () => {
    const model = docWithText("hello", "world");
    (model.sections[0].blocks[0] as Paragraph).changes = {
      markDelete: { changeId: "pm2", author: "Grace", date: "2026-02-02T00:00:00.000Z" },
    };
    const { parsed, contentXml } = await writeAndReadBack(model);
    expect(contentXml).toContain("<text:deletion>");

    const reparsed = parsedModel(parsed);
    const first = reparsed.sections[0].blocks[0] as Paragraph;
    expect(first.changes?.markDelete).toMatchObject({ author: "Grace", date: "2026-02-02T00:00:00.000Z" });
  });
});

describe("tracked table row round trip (AC-16)", () => {
  function tableModel(): DocumentModel {
    const model = createDocument("spec-039-odt-table");
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        { cells: [{ blocks: [paragraph("a1")] }, { blocks: [paragraph("a2")] }] },
        {
          cells: [{ blocks: [paragraph("b1")] }, { blocks: [paragraph("b2")] }],
          changes: { insert: { changeId: "row1", author: "Ada", date: "2026-03-03T00:00:00.000Z" } },
        },
      ],
    };
    model.sections = [{ blocks: [table] }];
    return model;
  }

  it("writes a tracked-inserted row as a leading table:change-scope=row text:change and reads it back", async () => {
    const { parsed, contentXml } = await writeAndReadBack(tableModel());
    expect(contentXml).toContain('table:change-scope="row"');

    const reparsed = parsedModel(parsed);
    const table = reparsed.sections[0].blocks[0] as Table;
    expect(table.rows).toHaveLength(2);
    expect(table.rows[1].changes?.insert).toMatchObject({ author: "Ada", date: "2026-03-03T00:00:00.000Z" });
    expect(table.rows[0].changes).toBeUndefined();
    // The row's own content is unaffected by the marker.
    expect(table.rows[1].cells[0].blocks[0].runs[0].text).toBe("b1");
  });
});

describe("tracked table cell round trip (AC-16)", () => {
  function tableModel(): DocumentModel {
    const model = createDocument("spec-039-odt-cell");
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        {
          cells: [
            { blocks: [paragraph("a1")] },
            {
              blocks: [paragraph("a2")],
              changes: { insert: { changeId: "cell1", author: "Ada", date: "2026-04-04T00:00:00.000Z" } },
            },
          ],
        },
      ],
    };
    model.sections = [{ blocks: [table] }];
    return model;
  }

  it("writes a tracked-inserted cell as a leading (unscoped) text:change and reads it back", async () => {
    const { parsed, contentXml } = await writeAndReadBack(tableModel());
    expect(contentXml).toContain("<text:change");

    const reparsed = parsedModel(parsed);
    const table = reparsed.sections[0].blocks[0] as Table;
    expect(table.rows[0].cells[1].changes?.insert).toMatchObject({ author: "Ada", date: "2026-04-04T00:00:00.000Z" });
    expect(table.rows[0].cells[0].changes).toBeUndefined();
    expect(table.rows[0].cells[1].blocks[0].runs[0].text).toBe("a2");
  });
});
