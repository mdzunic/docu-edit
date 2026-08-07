// SPEC-028 — breaks and sections through the .docx reader, writer and
// validator: the w:pPr pagination toggles (keepNext/keepLines/
// pageBreakBefore/widowControl), manual page/column breaks as BreakBlocks
// (AC-30) and multi-section documents via paragraph-level w:sectPr (AC-31).
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { buildDocumentXml, writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { unzip } from "./zip";
import { zip } from "./zipWriter";
import { WORD_NAMESPACE } from "./xml";
import { firstStructuralDifference } from "../equality";
import {
  asParagraph,
  createDocument,
  isBreak,
  type DocumentModel,
  type Paragraph,
  type ParagraphFormatting,
  type Section,
} from "../documentModel";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** A minimal package around a hand-written w:body — the reader.spec pattern
 * for markup this project's writer would not itself produce. */
async function docxFromBody(
  bodyXml: string,
  extraParts: { name: string; data: Uint8Array }[] = [],
): Promise<Uint8Array> {
  const documentXml =
    `${XML_DECLARATION}<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${bodyXml}</w:body></w:document>`;
  return zip([{ name: "word/document.xml", data: encoder.encode(documentXml) }, ...extraParts]);
}

function paragraph(text: string, formatting?: ParagraphFormatting): Paragraph {
  const p: Paragraph = { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
  if (formatting) {
    p.formatting = formatting;
  }
  return p;
}

function modelWith(sections: Section[]): DocumentModel {
  const model = createDocument("spec-028");
  model.sections = sections;
  return model;
}

/** The parsed document as a writable same-format model, carrying the
 * preserved sectPr XML — what an opened .docx actually holds. */
function rewrap(parsed: ParsedDocxDocument): DocumentModel {
  return {
    id: "spec-028",
    title: "Sections",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    sourceXml: parsed.sourceXml,
  };
}

async function documentXmlOf(model: DocumentModel): Promise<string> {
  const parts = await unzip((await writeDocx(model)).bytes);
  return decoder.decode(parts.get("word/document.xml")!);
}

function blockText(section: Section, index: number): string {
  return asParagraph(section.blocks[index]).runs.map((run) => run.text).join("");
}

describe("w:pPr pagination flags (SPEC-028 AC-22..29)", () => {
  it("reads keepNext/keepLines/pageBreakBefore as true and explicit-off as absent", async () => {
    const bytes = await docxFromBody(
      "<w:p><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/></w:pPr>" +
        "<w:r><w:t>flagged</w:t></w:r></w:p>" +
        '<w:p><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="false"/>' +
        '<w:pageBreakBefore w:val="off"/></w:pPr><w:r><w:t>off</w:t></w:r></w:p>',
    );
    const parsed = await parseDocx(bytes);
    const flagged = asParagraph(parsed.sections[0].blocks[0]).formatting;
    expect(flagged?.keepNext).toBe(true);
    expect(flagged?.keepLines).toBe(true);
    expect(flagged?.pageBreakBefore).toBe(true);
    const off = asParagraph(parsed.sections[0].blocks[1]).formatting;
    expect(off?.keepNext).toBeUndefined();
    expect(off?.keepLines).toBeUndefined();
    expect(off?.pageBreakBefore).toBeUndefined();
  });

  it("reads widowControl only when explicitly off — present-on and absent both mean on", async () => {
    const bytes = await docxFromBody(
      '<w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r><w:t>off</w:t></w:r></w:p>' +
        "<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t>on</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>default</w:t></w:r></w:p>",
    );
    const parsed = await parseDocx(bytes);
    expect(asParagraph(parsed.sections[0].blocks[0]).formatting?.widowControl).toBe(false);
    expect(asParagraph(parsed.sections[0].blocks[1]).formatting?.widowControl).toBeUndefined();
    expect(asParagraph(parsed.sections[0].blocks[2]).formatting?.widowControl).toBeUndefined();
  });

  it("lets direct explicit-off/on override the style chain, normalised to the canonical form", async () => {
    const stylesXml =
      `${XML_DECLARATION}<w:styles xmlns:w="${WORD_NAMESPACE}">` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Keeps"><w:name w:val="Keeps"/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:widowControl w:val="0"/></w:pPr></w:style>' +
      "</w:styles>";
    const relsXml =
      `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      "</Relationships>";
    const bytes = await docxFromBody(
      '<w:p><w:pPr><w:pStyle w:val="Keeps"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="Keeps"/><w:keepNext w:val="0"/><w:widowControl/></w:pPr>' +
        "<w:r><w:t>overridden</w:t></w:r></w:p>",
      [
        { name: "word/_rels/document.xml.rels", data: encoder.encode(relsXml) },
        { name: "word/styles.xml", data: encoder.encode(stylesXml) },
      ],
    );
    const parsed = await parseDocx(bytes);
    const styled = asParagraph(parsed.sections[0].blocks[0]).formatting;
    expect(styled?.keepNext).toBe(true);
    expect(styled?.keepLines).toBe(true);
    expect(styled?.widowControl).toBe(false);
    // The direct off/on wins over the style, and the winning default values
    // are absent (never an explicit false keepNext / true widowControl).
    const overridden = asParagraph(parsed.sections[0].blocks[1]).formatting;
    expect(overridden?.keepNext).toBeUndefined();
    expect(overridden?.keepLines).toBe(true);
    expect(overridden?.widowControl).toBeUndefined();
  });

  it("writes the flags in CT_PPr order: after w:pStyle, before w:numPr", () => {
    const model = modelWith([
      {
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: "x", formatting: {} }],
            styleId: "Heading1",
            headingLevel: 1,
            list: { type: "bullet", level: 1 },
            formatting: {
              keepNext: true,
              keepLines: true,
              pageBreakBefore: true,
              widowControl: false,
              alignment: "center",
            },
          },
        ],
      },
    ]);
    const xml = buildDocumentXml(model);
    const order = [
      "<w:pStyle",
      "<w:keepNext/>",
      "<w:keepLines/>",
      "<w:pageBreakBefore/>",
      '<w:widowControl w:val="0"/>',
      "<w:numPr>",
      "<w:jc",
    ].map((tag) => xml.indexOf(tag));
    expect(order.every((index) => index >= 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
  });

  it("round-trips every flag, including the widowControl explicit-off, and validates", async () => {
    const formatting: ParagraphFormatting = {
      keepNext: true,
      keepLines: true,
      pageBreakBefore: true,
      widowControl: false,
    };
    const model = modelWith([{ blocks: [paragraph("kept", formatting), paragraph("plain")] }]);
    const { bytes } = await writeDocx(model);
    expect(await validateDocx(bytes)).toEqual([]);
    const reopened = await parseDocx(bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).formatting).toEqual(formatting);
    const plain = asParagraph(reopened.sections[0].blocks[1]).formatting;
    expect(plain?.keepNext).toBeUndefined();
    expect(plain?.widowControl).toBeUndefined();
  });

  it("the validator flags a bad ST_OnOff value and an out-of-order flag", async () => {
    const model = modelWith([
      { blocks: [paragraph("x", { keepNext: true, widowControl: false, alignment: "center" })] },
    ]);
    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    const retamper = async (edit: (xml: string) => string): Promise<Uint8Array> =>
      zip(
        Array.from(parts, ([name, data]) => ({
          name,
          data:
            name === "word/document.xml" ? encoder.encode(edit(decoder.decode(data))) : data,
        })),
      );
    const badValue = await retamper((xml) =>
      xml.replace('<w:widowControl w:val="0"/>', '<w:widowControl w:val="perhaps"/>'),
    );
    expect((await validateDocx(badValue)).some((v) => v.includes("ST_OnOff"))).toBe(true);
    const outOfOrder = await retamper((xml) =>
      xml
        .replace("<w:keepNext/>", "")
        .replace('<w:jc w:val="center"/>', '<w:jc w:val="center"/><w:keepNext/>'),
    );
    expect(
      (await validateDocx(outOfOrder)).some((v) => v.includes("out of schema sequence order")),
    ).toBe(true);
  });
});

describe("manual breaks (SPEC-028 AC-30)", () => {
  it("writes a BreakBlock as a paragraph holding just the typed w:br run", async () => {
    const model = modelWith([
      { blocks: [paragraph("one"), { type: "break", kind: "page" }, paragraph("two")] },
    ]);
    const xml = await documentXmlOf(model);
    expect(xml).toContain('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    expect(await validateDocx((await writeDocx(model)).bytes)).toEqual([]);
  });

  it("reads the writer's break shape back as exactly one BreakBlock — a round-trip fixed point", async () => {
    const model = modelWith([
      { blocks: [paragraph("one"), { type: "break", kind: "page" }, paragraph("two")] },
    ]);
    const read1 = await parseDocx((await writeDocx(model)).bytes);
    expect(read1.sections[0].blocks).toHaveLength(3);
    expect(blockText(read1.sections[0], 0)).toBe("one");
    expect(read1.sections[0].blocks[1]).toEqual({ type: "break", kind: "page" });
    expect(blockText(read1.sections[0], 2)).toBe("two");
    const read2 = await parseDocx((await writeDocx(rewrap(read1))).bytes);
    expect(firstStructuralDifference(read1.sections, read2.sections)).toBeNull();
  });

  it("round-trips a column break as its own block", async () => {
    const model = modelWith([
      { blocks: [paragraph("left"), { type: "break", kind: "column" }, paragraph("right")] },
    ]);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    expect(reopened.sections[0].blocks[1]).toEqual({ type: "break", kind: "column" });
    expect(await validateDocx(bytes)).toEqual([]);
  });

  it("splits a Word-style mid-run page break into paragraph, break, paragraph", async () => {
    // Word puts the break inside a content-carrying run; both text pieces
    // keep the paragraph's own formatting.
    const bytes = await docxFromBody(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r></w:p>',
    );
    const parsed = await parseDocx(bytes);
    const blocks = parsed.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(parsed.sections[0], 0)).toBe("Before");
    expect(blocks[1]).toEqual({ type: "break", kind: "page" });
    expect(blockText(parsed.sections[0], 2)).toBe("After");
    expect(asParagraph(blocks[0]).formatting?.alignment).toBe("center");
    expect(asParagraph(blocks[2]).formatting?.alignment).toBe("center");
  });

  it("keeps untyped and textWrapping w:br as inline line breaks", async () => {
    const bytes = await docxFromBody(
      '<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t><w:br w:type="textWrapping"/><w:t>c</w:t></w:r></w:p>',
    );
    const parsed = await parseDocx(bytes);
    expect(parsed.sections[0].blocks).toHaveLength(1);
    expect(blockText(parsed.sections[0], 0)).toBe("a\nb\nc");
  });

  it("leaves a typed break inside an unmodelled run in that run's opaque XML", async () => {
    const bytes = await docxFromBody(
      "<w:p><w:r><w:t>x</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="begin"/><w:br w:type="page"/></w:r></w:p>',
    );
    const parsed = await parseDocx(bytes);
    expect(parsed.sections[0].blocks).toHaveLength(1);
    const p = asParagraph(parsed.sections[0].blocks[0]);
    expect(p.opaqueInlines).toHaveLength(1);
    expect(p.opaqueInlines![0].xml).toContain("fldChar");
    expect(p.opaqueInlines![0].xml).toContain("br");
  });
});

describe("multi-section round trip (SPEC-028 AC-31)", () => {
  // Twip-exact geometry (US Letter, 0.5" margins) so the exact-double
  // assertions below hold with no unit drift.
  const letter = { widthCm: 21.59, heightCm: 27.94 };
  const halfInch = { topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 };

  function twoSectionModel(): DocumentModel {
    return modelWith([
      { blocks: [paragraph("one")], pageSize: { ...letter }, margins: { ...halfInch } },
      { blocks: [paragraph("two")] },
    ]);
  }

  it("writes a non-final section's sectPr inside its last paragraph's pPr", async () => {
    const model = twoSectionModel();
    const xml = await documentXmlOf(model);
    // The first section's geometry, embedded in the paragraph…
    expect(xml).toContain(
      '<w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"' +
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr>',
    );
    // …the final section's as the body-level sectPr (the model default, A4).
    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(xml.match(/<w:sectPr>/g)).toHaveLength(2);
    expect(xml).not.toContain("<w:type");
    expect(await validateDocx((await writeDocx(model)).bytes)).toEqual([]);
  });

  it("reads its own two-section output back with per-section geometry", async () => {
    const parsed = await parseDocx((await writeDocx(twoSectionModel())).bytes);
    expect(parsed.sections).toHaveLength(2);
    const [first, second] = parsed.sections;
    expect(blockText(first, first.blocks.length - 1)).toBe("one");
    expect(blockText(second, 0)).toBe("two");
    expect(first.pageSize).toEqual(letter);
    expect(first.margins).toEqual(halfInch);
    expect(first.start).toBeUndefined();
    expect(first.docxSectPr).toContain("pgSz");
    // The final section stays on defaultStyles, like a single-section read.
    expect(second.pageSize).toBeUndefined();
    expect(second.margins).toBeUndefined();
    expect(parsed.defaultStyles.pageSize.widthCm).toBeCloseTo(21, 1);
  });

  function threeSectionModel(): DocumentModel {
    return modelWith([
      { blocks: [paragraph("first")] },
      {
        blocks: [paragraph("second")],
        start: "continuous",
        pageSize: { widthCm: 27.94, heightCm: 21.59 },
      },
      { blocks: [paragraph("third")], start: "continuous" },
    ]);
  }

  it("writes w:type continuous first in the sectPr and reads it back — final section included", async () => {
    const { bytes } = await writeDocx(threeSectionModel());
    const xml = decoder.decode((await unzip(bytes)).get("word/document.xml")!);
    const sectPrs = xml.match(/<w:sectPr>.*?<\/w:sectPr>/g)!;
    expect(sectPrs).toHaveLength(3);
    expect(sectPrs[0]).not.toContain("<w:type");
    // w:type precedes w:pgSz (CT_SectPr order), and the landscape section
    // is marked as such.
    expect(sectPrs[1]).toContain('<w:type w:val="continuous"/><w:pgSz');
    expect(sectPrs[1]).toContain('w:orient="landscape"');
    expect(sectPrs[2]).toContain('<w:type w:val="continuous"/>');
    expect(await validateDocx(bytes)).toEqual([]);

    const parsed = await parseDocx(bytes);
    expect(parsed.sections).toHaveLength(3);
    expect(parsed.sections[0].start).toBeUndefined();
    expect(parsed.sections[1].start).toBe("continuous");
    expect(parsed.sections[1].pageSize).toEqual({ widthCm: 27.94, heightCm: 21.59 });
    expect(parsed.sections[2].start).toBe("continuous");
  });

  it("write -> read -> write -> read is a structural fixed point", async () => {
    const read1 = await parseDocx((await writeDocx(threeSectionModel())).bytes);
    const read2 = await parseDocx((await writeDocx(rewrap(read1))).bytes);
    expect(firstStructuralDifference(read1.sections, read2.sections)).toBeNull();
    expect(read1.defaultStyles).toEqual(read2.defaultStyles);
  });

  it("reads a Word-authored empty-paragraph section break and preserves its sectPr verbatim", async () => {
    // Word's own shape: the section break is an empty paragraph whose pPr
    // holds the sectPr — with properties this model doesn't express
    // (w:cols) that must survive a same-format save.
    const bytes = await docxFromBody(
      "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
        '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>' +
        '<w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>' +
        '<w:cols w:num="2" w:space="708"/></w:sectPr></w:pPr></w:p>' +
        "<w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
    );
    const read1 = await parseDocx(bytes);
    expect(read1.sections).toHaveLength(2);
    const [first, second] = read1.sections;
    expect(first.blocks).toHaveLength(2);
    expect(blockText(first, 0)).toBe("First");
    expect(asParagraph(first.blocks[1]).runs).toEqual([]);
    expect(first.start).toBe("continuous");
    expect(first.pageSize).toEqual(letter);
    expect(first.margins).toEqual(halfInch);
    expect(first.docxSectPr).toContain("cols");
    expect(blockText(second, 0)).toBe("Second");

    // A same-format save re-emits the preserved sectPr (w:cols included).
    const rewritten = await writeDocx(rewrap(read1));
    const xml = decoder.decode((await unzip(rewritten.bytes)).get("word/document.xml")!);
    expect(xml).toContain('<w:cols w:num="2" w:space="708"/>');
    // The first save canonicalises the model (the styles-less fixture gains
    // docDefault run formatting, as in the SPEC-026 tests); from there the
    // round trip is a fixed point, preserved sectPr included.
    const read2 = await parseDocx(rewritten.bytes);
    expect(read2.sections[0].start).toBe("continuous");
    expect(read2.sections[0].docxSectPr).toContain("cols");
    const read3 = await parseDocx((await writeDocx(rewrap(read2))).bytes);
    expect(firstStructuralDifference(read2.sections, read3.sections)).toBeNull();
  });

  it("appends an empty section-break paragraph when a section ends in a non-paragraph block", async () => {
    const model = modelWith([
      { blocks: [paragraph("x"), { type: "break", kind: "page" }] },
      { blocks: [paragraph("y")] },
    ]);
    const { bytes } = await writeDocx(model);
    const xml = decoder.decode((await unzip(bytes)).get("word/document.xml")!);
    expect(xml).toContain('<w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:sectPr>');
    expect(await validateDocx(bytes)).toEqual([]);
    const parsed = await parseDocx(bytes);
    expect(parsed.sections).toHaveLength(2);
    expect(isBreak(parsed.sections[0].blocks[1])).toBe(true);
    expect(blockText(parsed.sections[1], 0)).toBe("y");
  });

  it("keeps single-section documents in the pre-SPEC-028 shape: geometry on defaultStyles only", async () => {
    const parsed = await parseDocx((await writeDocx(createDocument("spec-028"))).bytes);
    expect(parsed.sections).toHaveLength(1);
    const section = parsed.sections[0];
    expect(section.pageSize).toBeUndefined();
    expect(section.margins).toBeUndefined();
    expect(section.start).toBeUndefined();
    expect(section.docxSectPr).toBeUndefined();
  });
});
