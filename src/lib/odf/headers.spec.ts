// SPEC-029 — headers, footers and page numbers through the .odt reader and
// writer: style:header/style:footer (and -left variants) on the section's
// master page, style:page-number restart, and text:page-number/
// text:page-count fields.
import { describe, it, expect } from "vitest";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../docx/zip";
import { createDocument, createFieldRun, type DocumentModel, type Paragraph, type Section } from "../documentModel";

const decoder = new TextDecoder();

// A round trip through this writer's from-scratch styles.xml resolves the
// document's default font/size onto every run when reading back — matched
// here so equality checks below compare the field that changed.
const RESOLVED_DEFAULT_FORMATTING = { fontFamily: "Calibri", fontSizePt: 11 };

function paragraph(text: string): Paragraph {
  return {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: { ...RESOLVED_DEFAULT_FORMATTING } }] : [],
    formatting: {},
  };
}

function modelWith(sections: Section[]): DocumentModel {
  const model = createDocument("spec-029-odt");
  model.sections = sections;
  return model;
}

async function writeAndReadBack(model: DocumentModel): Promise<{
  parsed: ParsedOdtDocument;
  parts: Map<string, Uint8Array>;
}> {
  const { bytes } = await writeOdt(model);
  const parts = await unzip(bytes);
  const parsed = await parseOdt(bytes);
  return { parsed, parts };
}

describe("header/footer content round trip (SPEC-029 AC-3, AC-4, AC-9, AC-26)", () => {
  it("writes style:header/style:footer on the Standard master page and reads them back", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body text")],
        headers: { default: { blocks: [paragraph("Confidential")] } },
        footers: { default: { blocks: [paragraph("Page footer")] } },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    const stylesXml = decoder.decode(parts.get("styles.xml")!);
    expect(stylesXml).toContain("<style:header>");
    expect(stylesXml).toContain("<style:footer>");
    expect(parsed.sections[0].headers?.default?.blocks).toEqual([paragraph("Confidential")]);
    expect(parsed.sections[0].footers?.default?.blocks).toEqual([paragraph("Page footer")]);
  });

  it("round-trips bold formatting inside a header paragraph", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        headers: {
          default: { blocks: [{ type: "paragraph", runs: [{ text: "Bold header", formatting: { bold: true } }] }] },
        },
      },
    ]);
    const { parsed } = await writeAndReadBack(model);
    const run = (parsed.sections[0].headers!.default!.blocks[0] as Paragraph).runs[0];
    expect(run).toEqual({ text: "Bold header", formatting: { bold: true, ...RESOLVED_DEFAULT_FORMATTING } });
  });

  it("does not carry header/footer content when a section has none", async () => {
    const model = modelWith([{ blocks: [paragraph("Body")] }]);
    const { parsed } = await writeAndReadBack(model);
    expect(parsed.sections[0].headers).toBeUndefined();
    expect(parsed.sections[0].footers).toBeUndefined();
  });

  it("round-trips headers/footers on a second section's own master page", async () => {
    const model = modelWith([
      { blocks: [paragraph("First section")] },
      {
        blocks: [paragraph("Second section")],
        headers: { default: { blocks: [paragraph("Second header")] } },
      },
    ]);
    const { parsed } = await writeAndReadBack(model);
    expect(parsed.sections[0].headers).toBeUndefined();
    expect(parsed.sections[1].headers?.default?.blocks).toEqual([paragraph("Second header")]);
  });
});

describe("different odd/even pages (SPEC-029 AC-12)", () => {
  it("writes style:header-left/style:footer-left only when evenAndOddHeaders is set", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        evenAndOddHeaders: true,
        headers: { default: { blocks: [paragraph("Odd")] }, even: { blocks: [paragraph("Even")] } },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    expect(decoder.decode(parts.get("styles.xml")!)).toContain("<style:header-left>");
    expect(parsed.sections[0].evenAndOddHeaders).toBe(true);
    expect(parsed.sections[0].headers?.even?.blocks).toEqual([paragraph("Even")]);
  });

  it("omits header-left when evenAndOddHeaders is not set, even with 'even' content present", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        headers: { default: { blocks: [paragraph("Odd")] }, even: { blocks: [paragraph("Even")] } },
      },
    ]);
    const { parts } = await writeAndReadBack(model);
    expect(decoder.decode(parts.get("styles.xml")!)).not.toContain("style:header-left");
  });
});

describe("style:page-number restart (SPEC-029 AC-22)", () => {
  it("round-trips pageNumberStart on the first section", async () => {
    // Section 0's own paragraphs live in content.xml, not styles.xml — its
    // style:page-number rides on the first paragraph's own automatic style
    // there, unlike a later section's (which switches master page).
    const model = modelWith([{ blocks: [paragraph("Body")], pageNumberStart: 5 }]);
    const { parsed, parts } = await writeAndReadBack(model);
    expect(decoder.decode(parts.get("content.xml")!)).toContain('style:page-number="5"');
    expect(parsed.sections[0].pageNumberStart).toBe(5);
  });

  it("round-trips pageNumberStart on a later section", async () => {
    const model = modelWith([
      { blocks: [paragraph("First")] },
      { blocks: [paragraph("Second")], pageNumberStart: 1 },
    ]);
    const { parsed } = await writeAndReadBack(model);
    expect(parsed.sections[1].pageNumberStart).toBe(1);
    expect(parsed.sections[0].pageNumberStart).toBeUndefined();
  });

  it("omits style:page-number when unset (continue from previous section)", async () => {
    const model = modelWith([{ blocks: [paragraph("Body")] }]);
    const { parts } = await writeAndReadBack(model);
    expect(decoder.decode(parts.get("styles.xml")!)).not.toContain("style:page-number");
  });
});

describe("text:page-number / text:page-count fields (SPEC-029 AC-16..18, AC-25, AC-26)", () => {
  it("writes a page field and reads it back", async () => {
    const model = modelWith([
      {
        blocks: [paragraph("Body")],
        footers: { default: { blocks: [{ type: "paragraph", runs: [createFieldRun("page", "decimal")] }] } },
      },
    ]);
    const { parsed, parts } = await writeAndReadBack(model);
    expect(decoder.decode(parts.get("styles.xml")!)).toContain("<text:page-number>");
    const run = (parsed.sections[0].footers!.default!.blocks[0] as Paragraph).runs[0];
    expect(run.field).toEqual({ kind: "page", format: "decimal" });
  });

  it("round-trips 'Page X of Y' as page-number, literal ' of ', page-count", async () => {
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
    expect(decoder.decode(parts.get("styles.xml")!)).toMatch(/<text:page-number style:num-format="i">/);
    const runs = (parsed.sections[0].footers!.default!.blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.field?.kind ?? r.text)).toEqual(["page", " of ", "numpages"]);
    expect(runs[0].field?.format).toBe("lowerRoman");
  });
});

describe("same-format save (opening a real .odt, editing, saving again)", () => {
  async function realOdtModel(): Promise<{ model: DocumentModel; parsed: ParsedOdtDocument }> {
    // A from-scratch save is itself a valid, minimal real .odt package —
    // used here as the "opened source" to exercise the same-format
    // (preserved-parts) path deliberately, not the from-scratch one.
    const seed = modelWith([{ blocks: [paragraph("Body")] }]);
    const { bytes } = await writeOdt(seed);
    const parsed = await parseOdt(bytes);
    const model: DocumentModel = {
      id: "spec-029-odt-same-format",
      title: "Same format",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "odt",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
    };
    return { model, parsed };
  }

  it("persists a header added to section 0 of a previously headerless .odt", async () => {
    const { model } = await realOdtModel();
    model.sections[0] = {
      ...model.sections[0],
      headers: { default: { blocks: [paragraph("Added later")] } },
    };
    const { bytes } = await writeOdt(model);
    const parsed = await parseOdt(bytes);
    expect(parsed.sections[0].headers?.default?.blocks).toEqual([paragraph("Added later")]);
  });

  it("removes a header from section 0 when the model no longer has one", async () => {
    const { model } = await realOdtModel();
    model.sections[0] = {
      ...model.sections[0],
      headers: { default: { blocks: [paragraph("Temporary")] } },
    };
    const first = await writeOdt(model);
    const reopened = await parseOdt(first.bytes);
    expect(reopened.sections[0].headers?.default).toBeDefined();

    const modelWithoutHeader: DocumentModel = {
      id: "spec-029-odt-same-format",
      title: "Same format",
      sections: reopened.sections.map((section) => {
        const { headers: _headers, ...rest } = section;
        return rest;
      }),
      defaultStyles: reopened.defaultStyles,
      format: "odt",
      opaque: reopened.opaque,
      sourceXml: reopened.sourceXml,
    };
    const second = await writeOdt(modelWithoutHeader);
    const final = await parseOdt(second.bytes);
    expect(final.sections[0].headers).toBeUndefined();
  });
});
