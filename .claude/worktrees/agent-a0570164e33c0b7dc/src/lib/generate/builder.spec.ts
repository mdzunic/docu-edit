import { describe, expect, it } from "vitest";
import { DocumentBuilder } from "./builder";
import { DocumentGenerationError } from "./errors";
import { DOCX_MIME_TYPE } from "../docx/writer";
import { ODT_MIMETYPE } from "../odf/writer";

describe("DocumentBuilder — fluent construction (AC-1)", () => {
  it("chains every add*/set* method and returns itself", () => {
    const builder = new DocumentBuilder({ title: "Doc" });
    const result = builder
      .addHeading("Title", 1)
      .addParagraph("Body")
      .addList(["a", "b"])
      .addTable({ rows: [["A"]] })
      .addPageBreak()
      .addTableOfContents()
      .setHeader("Header")
      .setFooter("Footer");
    expect(result).toBe(builder);
  });

  it("appends blocks to the document's single section in order", () => {
    const builder = new DocumentBuilder();
    builder.addHeading("H", 1).addParagraph("P").addPageBreak();
    const doc = builder.toDoc();
    expect(doc.sections).toHaveLength(1);
    expect(doc.sections[0].blocks.map((b) => b.type)).toEqual(["paragraph", "paragraph", "break"]);
  });

  it("addList appends one paragraph per item", () => {
    const builder = new DocumentBuilder();
    builder.addList(["a", "b", "c"]);
    expect(builder.toDoc().sections[0].blocks).toHaveLength(3);
  });

  it("uses the given title, defaulting otherwise", () => {
    expect(new DocumentBuilder({ title: "My Doc" }).toDoc().title).toBe("My Doc");
    expect(new DocumentBuilder().toDoc().title).toBe("Untitled document");
  });

  it("rejects an unknown constructor option (AC-9)", () => {
    expect(() => new DocumentBuilder({ bogus: true } as never)).toThrow(/unknown property/);
  });

  it("setHeader/setFooter default to the 'default' kind and can target others", () => {
    const builder = new DocumentBuilder();
    builder.addParagraph("body").setHeader("H1").setHeader("H-first", "first").setFooter("F1");
    const section = builder.toDoc().sections[0];
    expect(section.headers?.default).toBeDefined();
    expect(section.headers?.first).toBeDefined();
    expect(section.footers?.default).toBeDefined();
  });

  it("addTableOfContents resolves entries from headings already added, at output time", () => {
    const builder = new DocumentBuilder();
    builder.addTableOfContents(); // before any heading exists
    builder.addHeading("Chapter", 1);
    const toc = builder.toDoc().sections[0].blocks[0];
    expect(toc.type).toBe("toc");
    if (toc.type === "toc") {
      expect(toc.entries).toHaveLength(1);
      expect(toc.entries[0].text).toBe("Chapter");
    }
  });
});

describe("DocumentBuilder — validation (AC-6, AC-7, AC-8)", () => {
  it("throws before writing anything for invalid input", () => {
    const builder = new DocumentBuilder();
    expect(() => builder.addHeading("H", 9)).toThrow(DocumentGenerationError);
    expect(() => builder.addTable({ rows: [] })).toThrow(DocumentGenerationError);
  });

  it("names the offending path and constraint (AC-7)", () => {
    const builder = new DocumentBuilder();
    try {
      builder.addHeading("H", 9);
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(DocumentGenerationError);
      expect((error as Error).message).toMatch(/body\[0\]\.level/);
      expect((error as Error).message).toMatch(/1 and 6/);
    }
  });
});

describe("DocumentBuilder — output (AC-2)", () => {
  it("toBytes produces non-empty bytes for both formats", async () => {
    const builder = new DocumentBuilder();
    builder.addParagraph("hello");
    expect((await builder.toBytes("docx")).length).toBeGreaterThan(0);
    expect((await builder.toBytes("odt")).length).toBeGreaterThan(0);
  });

  it("toBlob produces a Blob with the right MIME type", async () => {
    const builder = new DocumentBuilder();
    builder.addParagraph("hello");
    const docxBlob = await builder.toBlob("docx");
    expect(docxBlob).toBeInstanceOf(Blob);
    expect(docxBlob.type).toBe(DOCX_MIME_TYPE);
    const odtBlob = await builder.toBlob("odt");
    expect(odtBlob.type).toBe(ODT_MIMETYPE);
  });

  it("toBuffer produces a Node Buffer", async () => {
    const builder = new DocumentBuilder();
    builder.addParagraph("hello");
    const buffer = await builder.toBuffer("docx");
    expect(Buffer.isBuffer(buffer)).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it("toDoc returns a structural clone independent of further mutation", () => {
    const builder = new DocumentBuilder();
    builder.addParagraph("hello");
    const first = builder.toDoc();
    builder.addParagraph("world");
    expect(first.sections[0].blocks).toHaveLength(1);
  });
});
