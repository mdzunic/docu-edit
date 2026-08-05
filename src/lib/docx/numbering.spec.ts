import { describe, expect, it } from "vitest";
import {
  mergeNumberingXml,
  nextNumberingIds,
  parseNumberingCatalog,
  planNumbering,
} from "./numbering";
import { parseXml } from "./xml";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { unzip } from "./zip";
import { firstStructuralDifference } from "../equality";
import { createDocument, type DocumentModel, type ListFormatting, type Paragraph } from "../documentModel";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

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
  const model = createDocument("numbering-fixture");
  model.sections = [{ blocks }];
  return model;
}

const decoder = new TextDecoder();

describe("parseNumberingCatalog", () => {
  const numberingXml = `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/>
        <w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
    </w:abstractNum>
    <w:abstractNum w:abstractNumId="1">
      <w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
    <w:num w:numId="3"><w:abstractNumId w:val="1"/>
      <w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride></w:num>
  </w:numbering>`;
  const catalog = parseNumberingCatalog(parseXml(numberingXml));

  it("maps producer bullet characters (Wingdings, Courier 'o') onto the supported styles", () => {
    expect(catalog.listFormattingFor("1", 0)).toEqual({ type: "bullet", level: 1, marker: "square" });
    // Level 2's "o" maps to circle — the level default, so no marker stored;
    // its 2880-twip indent differs from the level default (1440).
    expect(catalog.listFormattingFor("1", 1)).toEqual({ type: "bullet", level: 2, indentCm: 5.08 });
  });

  it("carries numFmt, start and startOverride into the model (AC-22)", () => {
    expect(catalog.listFormattingFor("2", 0)).toEqual({
      type: "number", level: 1, marker: "upperRoman", startAt: 5,
    });
    expect(catalog.listFormattingFor("3", 0)).toEqual({
      type: "number", level: 1, marker: "upperRoman", startAt: 10,
    });
  });

  it("returns undefined for numId 0 and unknown numIds", () => {
    expect(catalog.listFormattingFor("0", 0)).toBeUndefined();
    expect(catalog.listFormattingFor("99", 0)).toBeUndefined();
  });
});

describe("planNumbering", () => {
  it("assigns one num per contiguous run, so an interruption restarts numbering (AC-11)", () => {
    const model = docWith(
      paragraph("one", { type: "number", level: 1 }),
      paragraph("two", { type: "number", level: 2 }),
      paragraph("plain"),
      paragraph("second list", { type: "number", level: 1 }),
    );
    const plan = planNumbering(model);
    expect(plan.refs.get("0:0")).toEqual({ numId: 1, ilvl: 0 });
    expect(plan.refs.get("0:1")).toEqual({ numId: 1, ilvl: 1 });
    expect(plan.refs.get("0:2")).toBeUndefined();
    expect(plan.refs.get("0:3")).toEqual({ numId: 2, ilvl: 0 });
    // Identical level definitions share one abstractNum.
    expect(plan.abstractNumsXml.match(/<w:abstractNum /g)).toHaveLength(1);
    expect(plan.numsXml.match(/<w:num /g)).toHaveLength(2);
  });

  it("emits levels with the run's markers, starts and indents", () => {
    const model = docWith(
      paragraph("five", { type: "number", level: 1, startAt: 5, marker: "upperAlpha" }),
      paragraph("bullet", { type: "bullet", level: 2, indentCm: 5.08 }),
    );
    const xml = planNumbering(model).abstractNumsXml;
    expect(xml).toContain('<w:start w:val="5"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/>');
    expect(xml).toContain('<w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/>');
    expect(xml).toContain('<w:ind w:left="2880" w:hanging="360"/>');
  });
});

describe("merge into a preserved numbering part", () => {
  it("keeps original definitions and allocates ids above them", () => {
    const original = `<?xml version="1.0"?><w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="4"><w:abstractNumId w:val="7"/></w:num></w:numbering>';
    expect(nextNumberingIds(original)).toEqual({ abstractNum: 8, num: 5 });

    const model = docWith(paragraph("item", { type: "bullet", level: 1 }));
    const plan = planNumbering(model, nextNumberingIds(original));
    const merged = mergeNumberingXml(original, plan);
    expect(merged).toContain('w:abstractNumId="7"');
    expect(merged).toContain('w:numId="4"');
    expect(merged).toContain('w:abstractNumId="8"');
    expect(merged).toContain('w:numId="5"');
    // Schema order: every abstractNum before every num.
    const firstNum = merged.indexOf("<w:num ");
    expect(merged.lastIndexOf("<w:abstractNum ")).toBeLessThan(firstNum);
  });
});

describe("SPEC-018 AC-18/AC-23 — .docx list round trip", () => {
  const listedModel = () =>
    docWith(
      paragraph("first", { type: "number", level: 1, startAt: 3 }),
      paragraph("sub", { type: "number", level: 2, marker: "upperRoman" }),
      paragraph("deep bullet", { type: "bullet", level: 4, indentCm: 6.35 }),
      paragraph("plain"),
      paragraph("bullet", { type: "bullet", level: 1 }),
    );

  it("writes a valid package whose numbering part has abstractNum + referencing num (AC-18)", async () => {
    const { bytes } = await writeDocx(listedModel());
    expect(await validateDocx(bytes)).toEqual([]);
    const parts = await unzip(bytes);
    const numbering = decoder.decode(parts.get("word/numbering.xml")!);
    expect(numbering).toContain("<w:abstractNum ");
    expect(numbering).toContain("<w:num ");
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>");
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).toContain("/word/numbering.xml");
    expect(decoder.decode(parts.get("word/_rels/document.xml.rels")!)).toContain("numbering.xml");
  });

  it("re-imports with identical markers, levels, indents and start values (AC-23)", async () => {
    const model = listedModel();
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    expect(reopened.sections[0].blocks.map((block) => block.list)).toEqual(
      model.sections[0].blocks.map((block) => block.list),
    );
  });

  it("stays stable across a second same-format save (merge path)", async () => {
    const model = listedModel();
    const first = await writeDocx(model);
    const opened = await parseDocx(first.bytes);
    const editedModel: DocumentModel = {
      ...createDocument("numbering-fixture-2"),
      sections: opened.sections,
      defaultStyles: opened.defaultStyles,
      format: "docx",
      opaque: opened.opaque,
      sourceXml: opened.sourceXml,
    };
    const second = await writeDocx(editedModel);
    expect(second.dropped).toEqual([]);
    const reopened = await parseDocx(second.bytes);
    expect(firstStructuralDifference(opened.sections, reopened.sections)).toBeNull();
    // The merged part still holds the first save's definitions verbatim.
    const numbering = decoder.decode((await unzip(second.bytes)).get("word/numbering.xml")!);
    expect(numbering.match(/<w:num /g)!.length).toBeGreaterThanOrEqual(4);
  });
});
