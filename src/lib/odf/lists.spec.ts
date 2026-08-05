import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { buildContentXml, writeOdt } from "./writer";
import { parseOdt } from "./reader";
import { unzip } from "../docx/zip";
import { createDocument, type DocumentModel, type ListFormatting, type Paragraph } from "../documentModel";

function paragraph(text: string, list?: ListFormatting): Paragraph {
  const p: Paragraph = {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: {} }] : [],
  };
  if (list) {
    p.list = list;
  }
  return p;
}

function docWith(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("odf-lists-fixture");
  model.sections = [{ blocks }];
  return model;
}

const decoder = new TextDecoder();

describe("SPEC-018 AC-19 — text:list nesting and list styles", () => {
  it("wraps a run in one text:list with a text:list-item per item", () => {
    const model = docWith(
      paragraph("one", { type: "bullet", level: 1 }),
      paragraph("two", { type: "bullet", level: 1 }),
    );
    const xml = buildContentXml(model);
    expect(xml).toContain('<text:list text:style-name="L1">');
    expect(xml.match(/<text:list-item>/g)).toHaveLength(2);
    expect(xml).toContain('<text:list-style style:name="L1">');
    expect(xml).toContain('text:bullet-char="•"');
  });

  it("nests a sub-level list inside the parent's text:list-item", () => {
    const model = docWith(
      paragraph("parent", { type: "bullet", level: 1 }),
      paragraph("child", { type: "bullet", level: 2 }),
      paragraph("sibling", { type: "bullet", level: 1 }),
    );
    const xml = buildContentXml(model);
    // parent item stays open, nested list inside it, then closed before the
    // sibling item.
    const compact = xml
      .slice(xml.indexOf("<text:list "))
      .replace(/<text:span[^>]*>|<\/text:span>/g, "");
    expect(compact).toMatch(
      /<text:list-item><text:p[^>]*>parent<\/text:p><text:list><text:list-item><text:p[^>]*>child<\/text:p><\/text:list-item><\/text:list><\/text:list-item><text:list-item><text:p[^>]*>sibling<\/text:p>/,
    );
  });

  it("emits number levels with format, suffix and start value", () => {
    const model = docWith(paragraph("five", { type: "number", level: 1, startAt: 5, marker: "upperRoman" }));
    const xml = buildContentXml(model);
    expect(xml).toContain('style:num-format="I"');
    expect(xml).toContain('style:num-suffix="."');
    expect(xml).toContain('text:start-value="5"');
  });

  it("separates two runs interrupted by a plain paragraph into two text:list elements (AC-11)", () => {
    const model = docWith(
      paragraph("one", { type: "number", level: 1 }),
      paragraph("plain"),
      paragraph("two", { type: "number", level: 1 }),
    );
    const xml = buildContentXml(model);
    expect(xml.match(/<text:list text:style-name=/g)).toHaveLength(2);
  });
});

describe("SPEC-018 AC-21/AC-22/AC-23 — .odt list round trip", () => {
  const listedModel = () =>
    docWith(
      paragraph("first", { type: "number", level: 1, startAt: 3 }),
      paragraph("sub", { type: "number", level: 2, marker: "upperRoman" }),
      paragraph("deep bullet", { type: "bullet", level: 4, indentCm: 6.35 }),
      paragraph("plain"),
      paragraph("bullet", { type: "bullet", level: 1 }),
    );

  it("re-imports with identical markers, levels, indents and start values", async () => {
    const model = listedModel();
    const { bytes, dropped } = await writeOdt(model);
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    expect(reopened.sections[0].blocks.map((block) => block.list)).toEqual(
      model.sections[0].blocks.map((block) => block.list),
    );
    expect(reopened.sections[0].blocks.map((block) => block.runs.map((run) => run.text).join(""))).toEqual([
      "first", "sub", "deep bullet", "plain", "bullet",
    ]);
  });

  it("stays stable across a second same-format save", async () => {
    const first = await writeOdt(listedModel());
    const opened = await parseOdt(first.bytes);
    const editedModel: DocumentModel = {
      ...createDocument("odf-lists-fixture-2"),
      sections: opened.sections,
      defaultStyles: opened.defaultStyles,
      format: "odt",
      opaque: opened.opaque,
      sourceXml: opened.sourceXml,
    };
    const second = await writeOdt(editedModel);
    expect(second.dropped).toEqual([]);
    const reopened = await parseOdt(second.bytes);
    expect(reopened.sections[0].blocks.map((block) => block.list)).toEqual(
      opened.sections[0].blocks.map((block) => block.list),
    );
  });

  it("imports the SPEC-018 e2e fixture with its markers, levels, indents and start (AC-21, AC-22)", async () => {
    const fixture = new Uint8Array(
      readFileSync(
        join(dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "e2e", "fixtures", "lists.odt"),
      ),
    );
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks;
    expect(blocks.map((block) => block.runs.map((run) => run.text).join(""))).toEqual([
      "Groceries", "Apples", "Bananas", "Green ones", "Wrapped ones",
      "Steps", "First", "Second", "Detail",
      "Continue from five", "Fifth", "Sixth",
    ]);
    expect(blocks.map((block) => block.list)).toEqual([
      undefined,
      { type: "bullet", level: 1, indentCm: 1.905 }, // non-default indent from WL1
      { type: "bullet", level: 1, indentCm: 1.905 },
      { type: "bullet", level: 2 },
      { type: "bullet", level: 3 },
      undefined,
      { type: "number", level: 1 },
      { type: "number", level: 1 },
      { type: "number", level: 2 },
      undefined,
      { type: "number", level: 1, startAt: 5 },
      { type: "number", level: 1, startAt: 5 },
    ]);
  });

  it("keeps the manifest and package shape valid (content.xml still first-class)", async () => {
    const { bytes } = await writeOdt(listedModel());
    const parts = await unzip(bytes);
    expect(decoder.decode(parts.get("mimetype")!)).toBe("application/vnd.oasis.opendocument.text");
    expect(decoder.decode(parts.get("content.xml")!)).toContain("<text:list ");
  });
});
