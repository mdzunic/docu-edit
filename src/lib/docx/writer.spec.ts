import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { parseOdt } from "../odf/reader";
import {
  buildCorePropertiesXml,
  buildDocumentXml,
  DOCX_MIME_TYPE,
  docxFileName,
  writeDocx,
} from "./writer";
import { unzip } from "./zip";
import { parseXml } from "./xml";
import { createDocument, DEFAULT_TITLE, type DocumentModel, type Paragraph } from "../documentModel";

const here = dirname(fileURLToPath(import.meta.url));
const docxFixturesDir = join(here, "fixtures");
const odfFixturesDir = join(here, "..", "odf", "fixtures");

function loadFixture(dir: string, name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(dir, name)));
}

function modelFrom(parsed: ParsedDocxDocument, title = "Fixture"): DocumentModel {
  return {
    id: "test-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
  };
}

const decoder = new TextDecoder();

async function partText(bytes: Uint8Array, name: string): Promise<string> {
  const parts = await unzip(bytes);
  const raw = parts.get(name);
  if (!raw) {
    throw new Error(`part ${name} missing from written package`);
  }
  return decoder.decode(raw);
}

function paragraphTexts(model: { sections: { blocks: Paragraph[] }[] }): string[] {
  return model.sections.flatMap((section) =>
    section.blocks.map((block) => block.runs.map((run) => run.text).join("")),
  );
}

describe("writeDocx — package structure (AC-3..AC-8)", () => {
  it("writes a valid OPC ZIP containing every required part", async () => {
    const bytes = (await writeDocx(createDocument("id-1"))).bytes;
    const parts = await unzip(bytes);
    for (const required of [
      "[Content_Types].xml",
      "_rels/.rels",
      "word/document.xml",
      "word/styles.xml",
      "docProps/core.xml",
    ]) {
      expect(parts.has(required), `expected ${required} in the package`).toBe(true);
    }
  });

  it("every written part is well-formed XML", async () => {
    const bytes = (await writeDocx(createDocument("id-1"))).bytes;
    const parts = await unzip(bytes);
    for (const [name, raw] of parts) {
      expect(() => parseXml(decoder.decode(raw)), `${name} should parse`).not.toThrow();
    }
  });
});

describe("writeDocx — docProps/core.xml (AC-14, AC-15)", () => {
  const modified = new Date("2026-08-03T12:34:56Z");

  it("carries the document title", async () => {
    const model = { ...createDocument("id-1"), title: "Quarterly Report" };
    const core = parseXml(await partText((await writeDocx(model, { modified })).bytes, "docProps/core.xml"));
    const title = core.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0];
    expect(title.textContent).toBe("Quarterly Report");
  });

  it("carries a dcterms:modified W3CDTF timestamp", async () => {
    const core = parseXml(
      await partText((await writeDocx(createDocument("id-1"), { modified })).bytes, "docProps/core.xml"),
    );
    const el = core.getElementsByTagNameNS("http://purl.org/dc/terms/", "modified")[0];
    expect(el.textContent).toBe("2026-08-03T12:34:56Z");
    expect(el.getAttributeNS("http://www.w3.org/2001/XMLSchema-instance", "type")).toBe(
      "dcterms:W3CDTF",
    );
  });

  it("escapes XML metacharacters in the title", () => {
    const xml = buildCorePropertiesXml('Q3 <Report> & "Friends"', modified);
    const doc = parseXml(xml);
    expect(
      doc.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0].textContent,
    ).toBe('Q3 <Report> & "Friends"');
  });

  it("falls back to the default title when the model title is blank", async () => {
    const model = { ...createDocument("id-1"), title: "   " };
    const core = await partText((await writeDocx(model, { modified })).bytes, "docProps/core.xml");
    expect(core).toContain(`<dc:title>${DEFAULT_TITLE}</dc:title>`);
  });
});

describe.each(["word-authored.docx", "libreoffice-authored.docx", "google-docs-exported.docx"])(
  "writeDocx — %s round trip (AC-9..AC-13, AC-20, AC-21)",
  (fixtureName) => {
    it("open -> save -> re-open reproduces the model exactly", async () => {
      const first = await parseDocx(loadFixture(docxFixturesDir, fixtureName));
      const written = (await writeDocx(modelFrom(first))).bytes;
      const second = await parseDocx(written);
      // Text, run formatting, paragraph formatting, style references and
      // heading levels all live inside sections; page setup and default
      // font live in defaultStyles. Together these are the whole parsed
      // model.
      expect(second.sections).toEqual(first.sections);
      expect(second.defaultStyles).toEqual(first.defaultStyles);
    });

    it("a second save of the re-opened model is byte-identical (stable output)", async () => {
      const modified = new Date("2026-08-03T12:34:56Z");
      const first = await parseDocx(loadFixture(docxFixturesDir, fixtureName));
      const written = (await writeDocx(modelFrom(first), { modified })).bytes;
      const reopened = await parseDocx(written);
      const rewritten = (await writeDocx(modelFrom(reopened), { modified })).bytes;
      expect(rewritten).toEqual(written);
    });
  },
);

describe("writeDocx — content details", () => {
  it("round-trips tabs and forced line breaks as w:tab/w:br (AC-9)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "col1\tcol2\nnext line", formatting: {} }] },
    ];
    const documentXml = await partText((await writeDocx(model)).bytes, "word/document.xml");
    expect(documentXml).toContain("<w:tab/>");
    expect(documentXml).toContain("<w:br/>");
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual(["col1\tcol2\nnext line"]);
  });

  it("round-trips XML metacharacters and non-ASCII text (AC-9)", async () => {
    const text = `5 < 6 & 7 > 2, "quoted", 'single' — héllo π`;
    const model = createDocument("id-1");
    model.sections[0].blocks = [{ type: "paragraph", runs: [{ text, formatting: {} }] }];
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual([text]);
  });

  it("preserves leading/trailing spaces via xml:space='preserve' (AC-9)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "  spaced out  ", formatting: {} }] },
    ];
    const documentXml = await partText((await writeDocx(model)).bytes, "word/document.xml");
    expect(documentXml).toContain('<w:t xml:space="preserve">  spaced out  </w:t>');
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual(["  spaced out  "]);
  });

  it("writes explicitly-disabled run formatting as w:val='0'/'none' (AC-10)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "plain", formatting: { bold: false, underline: false } }],
      },
    ];
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    const run = (reopened.sections[0].blocks[0] as Paragraph).runs[0];
    expect(run.formatting.bold).toBe(false);
    expect(run.formatting.underline).toBe(false);
  });

  it("writes a hanging indent as w:hanging (AC-11)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "hung", formatting: {} }],
        formatting: { indentFirstLineCm: -1.27 },
      },
    ];
    const documentXml = await partText((await writeDocx(model)).bytes, "word/document.xml");
    expect(documentXml).toContain('<w:ind w:hanging="720"/>');
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect((reopened.sections[0].blocks[0] as Paragraph).formatting?.indentFirstLineCm).toBeCloseTo(
      -1.27,
      10,
    );
  });

  it("declares every referenced paragraph style in styles.xml (AC-12)", async () => {
    const first = await parseDocx(loadFixture(docxFixturesDir, "google-docs-exported.docx"));
    const stylesXml = await partText((await writeDocx(modelFrom(first))).bytes, "word/styles.xml");
    for (const id of ["Heading1", "Heading2", "Heading3"]) {
      expect(stylesXml).toContain(`w:styleId="${id}"`);
    }
    expect(stylesXml).toContain('w:default="1" w:styleId="Normal"');
  });

  it("marks a landscape page with w:orient (AC-13)", async () => {
    const first = await parseDocx(loadFixture(docxFixturesDir, "libreoffice-authored.docx"));
    const documentXml = await partText((await writeDocx(modelFrom(first))).bytes, "word/document.xml");
    expect(documentXml).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
  });
});

describe("writeDocx — empty document (AC-24)", () => {
  it("a freshly created blank document produces a valid one-paragraph file", async () => {
    const reopened = await parseDocx((await writeDocx(createDocument("id-1"))).bytes);
    expect(reopened.sections[0].blocks).toHaveLength(1);
    expect((reopened.sections[0].blocks[0] as Paragraph).runs).toEqual([]);
  });

  it("a model with no blocks at all still writes one paragraph", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [];
    const documentXml = await partText((await writeDocx(model)).bytes, "word/document.xml");
    expect(documentXml).toContain("<w:p/>");
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(reopened.sections[0].blocks).toHaveLength(1);
  });
});

describe("writeDocx — from a document opened as .odt (AC-23)", () => {
  it("produces a converted WordprocessingML package, not a renamed ODF file", async () => {
    const parsed = await parseOdt(loadFixture(odfFixturesDir, "libreoffice-writer.odt"));
    const model: DocumentModel = {
      id: "odt-doc",
      title: "Project Plan",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "odt",
    };
    const bytes = (await writeDocx(model)).bytes;
    const parts = await unzip(bytes);
    expect(parts.has("word/document.xml")).toBe(true);
    expect(parts.has("mimetype")).toBe(false);
    expect(parts.has("content.xml")).toBe(false);
    const documentXml = decoder.decode(parts.get("word/document.xml"));
    expect(documentXml).not.toContain("urn:oasis");
  });

  it("all text and headings survive the conversion", async () => {
    const parsed = await parseOdt(loadFixture(odfFixturesDir, "libreoffice-writer.odt"));
    const model: DocumentModel = {
      id: "odt-doc",
      title: "Project Plan",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "odt",
    };
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual(paragraphTexts(parsed));
    const originalHeadings = parsed.sections[0].blocks.map(
      (block) => (block as Paragraph).headingLevel,
    );
    const reopenedHeadings = reopened.sections[0].blocks.map(
      (block) => (block as Paragraph).headingLevel,
    );
    expect(reopenedHeadings).toEqual(originalHeadings);
  });
});

describe("docxFileName (AC-2)", () => {
  it("is <title>.docx", () => {
    expect(docxFileName("Quarterly Report")).toBe("Quarterly Report.docx");
  });

  it("falls back to the default title for a blank title", () => {
    expect(docxFileName("   ")).toBe(`${DEFAULT_TITLE}.docx`);
  });
});

describe("buildDocumentXml", () => {
  it("emits pPr children in the schema's sequence order", () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "x", formatting: {} }],
        styleId: "Heading1",
        headingLevel: 1,
        formatting: {
          alignment: "center",
          indentStartCm: 1.27,
          spaceBeforePt: 12,
        },
      },
    ];
    const xml = buildDocumentXml(model);
    const pPr = /<w:pPr>(.*?)<\/w:pPr>/.exec(xml)?.[1] ?? "";
    const order = ["w:pStyle", "w:spacing", "w:ind", "w:jc", "w:outlineLvl"].map((tag) =>
      pPr.indexOf(`<${tag}`),
    );
    expect(order.every((index) => index >= 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
  });

  it("exposes the .docx MIME type for the download blob", () => {
    expect(DOCX_MIME_TYPE).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
  });
});
