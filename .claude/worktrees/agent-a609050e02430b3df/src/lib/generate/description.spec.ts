import { describe, expect, it } from "vitest";
import { docFromDescription } from "./description";
import { DocumentBuilder } from "./builder";
import { DocumentGenerationError } from "./errors";
import { equalStructural } from "../equality";

describe("docFromDescription (AC-4)", () => {
  it("builds a document from title + body", () => {
    const doc = docFromDescription({
      title: "Invoice",
      body: [
        { type: "heading", text: "Invoice #1", level: 1 },
        { type: "paragraph", text: "Thank you for your business." },
      ],
    });
    expect(doc.title).toBe("Invoice");
    expect(doc.sections[0].blocks).toHaveLength(2);
  });

  it("an object with no fields describes an empty document (AC-10)", () => {
    const doc = docFromDescription({});
    expect(doc.sections[0].blocks).toEqual([]);
  });

  it("sets header/footer content", () => {
    const doc = docFromDescription({
      body: [{ type: "paragraph", text: "body" }],
      header: "Company Letterhead",
      footer: [{ type: "paragraph", text: "Page footer" }],
    });
    expect(doc.sections[0].headers?.default?.blocks).toHaveLength(1);
    expect(doc.sections[0].footers?.default?.blocks).toHaveLength(1);
  });

  it("rejects an unknown top-level property (AC-9)", () => {
    expect(() => docFromDescription({ bogus: true })).toThrow(/unknown property/);
  });

  it("rejects a non-object description", () => {
    expect(() => docFromDescription("not an object")).toThrow(DocumentGenerationError);
    expect(() => docFromDescription(null)).toThrow(DocumentGenerationError);
    expect(() => docFromDescription([1, 2])).toThrow(DocumentGenerationError);
  });

  it("names the offending path for an invalid block (AC-7)", () => {
    expect(() => docFromDescription({ body: [{ type: "table", rows: [] }] })).toThrow(/body\[0\]\.rows/);
  });

  it("produces the same model as the equivalent fluent builder calls (1:1 mapping)", () => {
    const fromJson = docFromDescription({
      title: "Report",
      body: [
        { type: "heading", text: "Chapter One", level: 1 },
        { type: "paragraph", runs: [{ text: "bold ", bold: true }, "plain"] },
        { type: "list", items: ["a", "b"], style: "number" },
        { type: "table", rows: [["X", "Y"]], header: true },
        { type: "pageBreak" },
        { type: "toc", levels: [1], hyperlinksOnly: true, showPageNumbers: false },
      ],
    });
    const fromBuilder = new DocumentBuilder({ title: "Report" })
      .addHeading("Chapter One", 1)
      .addParagraph([{ text: "bold ", bold: true }, "plain"])
      .addList(["a", "b"], { style: "number" })
      .addTable({ rows: [["X", "Y"]], header: true })
      .addPageBreak()
      .addTableOfContents({ levels: [1], hyperlinksOnly: true, showPageNumbers: false })
      .toDoc();
    expect(equalStructural(fromJson.sections, fromBuilder.sections)).toBe(true);
  });
});
