// SPEC-029 — headers, footers and page numbers through the .docx reader and
// writer: header1.xml/footer1.xml parts referenced via w:headerReference/
// w:footerReference, w:titlePg/w:pgNumType, PAGE/NUMPAGES fields as
// w:fldSimple, and w:evenAndOddHeaders in word/settings.xml.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { buildDocumentXml, writeDocx } from "./writer";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { WORD_NAMESPACE } from "./xml";
import {
  createDocument,
  createFieldRun,
  type DocumentModel,
  type Paragraph,
  type Section,
} from "../documentModel";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// A round trip through this writer's from-scratch styles.xml resolves the
// document's default font/size onto every run when reading back (the same
// resolution body content goes through) — matched here so equality checks
// below compare the field that changed, not this pre-existing behaviour.
const RESOLVED_DEFAULT_FORMATTING = { fontFamily: "Calibri", fontSizePt: 11 };

function paragraph(text: string): Paragraph {
  return {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: { ...RESOLVED_DEFAULT_FORMATTING } }] : [],
    formatting: {},
  };
}

function modelWith(sections: Section[]): DocumentModel {
  const model = createDocument("spec-029");
  model.sections = sections;
  return model;
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

describe("header/footer content round trip (SPEC-029 AC-3..6, AC-9, AC-24)", () => {
  it("writes a header1.xml part with a w:headerReference and reads it back", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body text")],
        headers: { default: { blocks: [paragraph("Confidential")] } },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    expect(parts.has("word/header1.xml")).toBe(true);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain("w:headerReference");
    expect(documentXml).toContain('w:type="default"');
    expect(parsed.sections[0].headers?.default?.blocks).toEqual([paragraph("Confidential")]);
  });

  it("round-trips footer content, formatting and a table", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        footers: {
          default: {
            blocks: [
              { type: "paragraph", runs: [{ text: "Bold footer", formatting: { bold: true } }] },
              {
                type: "table",
                rows: [{ cells: [{ blocks: [paragraph("Cell A")] }, { blocks: [paragraph("Cell B")] }] }],
                columnWidthsCm: [5, 5],
              },
            ],
          },
        },
      },
    ]);
    const { parsed } = await writeAndReadBack(model);
    const blocks = parsed.sections[0].footers?.default?.blocks;
    expect(blocks).toHaveLength(2);
    expect((blocks![0] as Paragraph).runs[0]).toEqual({
      text: "Bold footer",
      formatting: { bold: true, ...RESOLVED_DEFAULT_FORMATTING },
    });
    expect(blocks?.[1].type).toBe("table");
  });

  it("does not carry a header/footer part when a section has none", async () => {
    const model = modelWith([{ blocks: [paragraph("Body")] }]);
    const { parsed, parts } = await writeAndReadBack(model);
    expect(parts.has("word/header1.xml")).toBe(false);
    expect(parsed.sections[0].headers).toBeUndefined();
    expect(parsed.sections[0].footers).toBeUndefined();
  });
});

describe("titlePg / pgNumType round trip (SPEC-029 AC-11, AC-22)", () => {
  it("round-trips titlePage and pageNumberStart", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        titlePage: true,
        pageNumberStart: 5,
        headers: {
          default: { blocks: [paragraph("Default header")] },
          first: { blocks: [paragraph("First page header")] },
        },
      },
    ]);
    const { parsed } = await writeAndReadBack(model);
    expect(parsed.sections[0].titlePage).toBe(true);
    expect(parsed.sections[0].pageNumberStart).toBe(5);
    expect(parsed.sections[0].headers?.default?.blocks).toEqual([paragraph("Default header")]);
    expect(parsed.sections[0].headers?.first?.blocks).toEqual([paragraph("First page header")]);
  });

  it("omits pgNumType/titlePg when unset", async () => {
    const model = modelWith([{ blocks: [paragraph("Body")], headers: { default: { blocks: [paragraph("H")] } } }]);
    const { parts } = await writeAndReadBack(model);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).not.toContain("w:titlePg");
    expect(documentXml).not.toContain("w:pgNumType");
  });
});

describe("w:evenAndOddHeaders round trip (SPEC-029 AC-12)", () => {
  it("writes word/settings.xml with the flag when any section wants it", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        evenAndOddHeaders: true,
        headers: { default: { blocks: [paragraph("H")] }, even: { blocks: [paragraph("Even H")] } },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    expect(parts.has("word/settings.xml")).toBe(true);
    expect(decoder.decode(parts.get("word/settings.xml")!)).toContain("w:evenAndOddHeaders");
    expect(parsed.sections[0].evenAndOddHeaders).toBe(true);
    expect(parsed.sections[0].headers?.even?.blocks).toEqual([paragraph("Even H")]);
  });

  it("writes no settings.xml when nothing wants it", async () => {
    const model = modelWith([{ blocks: [paragraph("Body")], headers: { default: { blocks: [paragraph("H")] } } }]);
    const { parts } = await writeAndReadBack(model);
    expect(parts.has("word/settings.xml")).toBe(false);
  });
});

describe("PAGE/NUMPAGES fields (SPEC-029 AC-16..18, AC-25)", () => {
  it("writes a PAGE field as w:fldSimple and reads it back as a field run", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        footers: {
          default: {
            blocks: [{ type: "paragraph", runs: [createFieldRun("page", "decimal")] }],
          },
        },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    const footerXml = decoder.decode(parts.get("word/footer1.xml")!);
    expect(footerXml).toContain("w:fldSimple");
    expect(footerXml).toMatch(/w:instr="PAGE/);
    const run = (parsed.sections[0].footers!.default!.blocks[0] as Paragraph).runs[0];
    expect(run.field).toEqual({ kind: "page", format: "decimal" });
  });

  it("round-trips 'Page X of Y' as PAGE, literal ' of ', NUMPAGES", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        footers: {
          default: {
            blocks: [
              {
                type: "paragraph",
                runs: [
                  createFieldRun("page", "lowerRoman"),
                  { text: " of ", formatting: {} },
                  createFieldRun("numpages", "lowerRoman"),
                ],
              },
            ],
          },
        },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    const footerXml = decoder.decode(parts.get("word/footer1.xml")!);
    expect(footerXml).toMatch(/w:instr="PAGE \\\* roman"/);
    expect(footerXml).toMatch(/w:instr="NUMPAGES \\\* roman"/);
    const runs = (parsed.sections[0].footers!.default!.blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.field?.kind ?? r.text)).toEqual(["page", " of ", "numpages"]);
    expect((runs[0].field as { format?: string } | undefined)?.format).toBe("lowerRoman");
    expect((runs[2].field as { format?: string } | undefined)?.format).toBe("lowerRoman");
  });

  it("reads a hand-authored Word-shaped fldSimple with the alphabetic switch", async () => {
    const documentXml =
      `${XML_DECLARATION}<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${R_NS}"><w:body>` +
      "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
      `<w:sectPr xmlns:r="${R_NS}"><w:headerReference w:type="default" r:id="rId1"/>` +
      '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
      "</w:sectPr></w:body></w:document>";
    const headerXml =
      `${XML_DECLARATION}<w:hdr xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>` +
      '<w:fldSimple w:instr="PAGE \\* alphabetic"><w:r><w:t>a</w:t></w:r></w:fldSimple></w:p></w:hdr>';
    const bytes = await zip([
      { name: "word/document.xml", data: encoder.encode(documentXml) },
      { name: "word/header1.xml", data: encoder.encode(headerXml) },
      {
        name: "word/_rels/document.xml.rels",
        data: encoder.encode(
          `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
            "</Relationships>",
        ),
      },
    ]);
    const parsed = await parseDocx(bytes);
    const runs = (parsed.sections[0].headers!.default!.blocks[0] as Paragraph).runs;
    expect(runs[0].text).toBe("Page ");
    expect(runs[1].field).toEqual({ kind: "page", format: "lowerAlpha" });
  });
});

describe("stale preserved sectPr is dropped for a section gaining headers (SPEC-029)", () => {
  it("buildDocumentXml regenerates sectPr instead of reusing a preserved verbatim one lacking the new reference", () => {
    // A section carrying a preserved docxSectPr with no header reference at
    // all (as if opened before any header existed) that has since gained
    // header content through editing — the writer must not reuse the stale
    // XML verbatim, since its relationship ids would dangle.
    const model: DocumentModel = {
      id: "spec-029-stale",
      title: "Stale",
      format: "docx",
      sections: [
        {
          blocks: [paragraph("Body")],
          headers: { default: { blocks: [paragraph("New header")] } },
          docxSectPr: '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
        },
        { blocks: [paragraph("Second")] },
      ],
      defaultStyles: createDocument("x").defaultStyles,
    };
    const documentXml = buildDocumentXml(model);
    expect(documentXml).toContain("w:headerReference");
  });
});
