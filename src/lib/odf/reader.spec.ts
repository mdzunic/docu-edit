import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import {
  detectOdf,
  parseOdt,
  FlatOdfError,
  OdtOpenError,
  FLAT_ODF_MESSAGE,
  INVALID_ODT_MESSAGE,
} from "./reader";
import type { Paragraph } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "docx", "fixtures");

function loadFixture(dir: string, name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(dir, name)));
}

function text(paragraph: Paragraph): string {
  return paragraph.runs.map((run) => run.text).join("");
}

// A minimal flat XML ODF (.fodt) document — the whole package as one file.
const FLAT_ODF = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3">
 <office:body><office:text><text:p>Flat content</text:p></office:text></office:body>
</office:document>`;

// The two fixtures are hand-built ODF packages reproducing LibreOffice
// Writer (ODF 1.3) and OpenOffice.org-era (ODF 1.1) output respectively —
// see scripts/build-odt-fixtures.mjs for the rationale and exact markup;
// this environment has no network access to fetch genuine exports.
describe("parseOdt — libreoffice-writer.odt", () => {
  const fixture = loadFixture(fixturesDir, "libreoffice-writer.odt");

  it("maps every text:p and text:h to a paragraph (AC-3, AC-5)", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.sections[0].blocks).toHaveLength(8);
  });

  it("maps text:h with text:outline-level to heading levels (AC-6, AC-32)", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks.map((b) => b.headingLevel)).toEqual([
      1,
      undefined,
      undefined,
      undefined,
      2,
      undefined,
      undefined,
      undefined,
    ]);
  });

  it("keeps the raw style name as styleId", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks[0].styleId).toBe("Heading_20_1");
    expect(blocks[6].styleId).toBe("P1");
  });

  it("resolves heading formatting through the parent chain, including percentage font sizes (AC-12, AC-13, AC-14, AC-19)", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    // Heading_20_1: 130% of Heading's 14pt Liberation Sans, bold.
    expect(blocks[0].runs[0].formatting.bold).toBe(true);
    expect(blocks[0].runs[0].formatting.fontSizePt).toBeCloseTo(18.2, 2);
    expect(blocks[0].runs[0].formatting.fontFamily).toBe("Liberation Sans");
    // Heading_20_2: 115% of 14pt.
    expect(blocks[4].runs[0].formatting.fontSizePt).toBeCloseTo(16.1, 2);
    expect(blocks[4].runs[0].formatting.bold).toBe(true);
  });

  it("maps spans with automatic text styles to formatted runs (AC-7, AC-11, AC-14..17, AC-33)", async () => {
    const parsed = await parseOdt(fixture);
    const runs = (parsed.sections[0].blocks[1] as Paragraph).runs;
    expect(text(parsed.sections[0].blocks[1] as Paragraph)).toBe(
      "Body text with bold and italic and underlined and struck words.",
    );
    expect(runs.some((r) => r.text === "bold" && r.formatting.bold === true)).toBe(true);
    expect(runs.some((r) => r.text === "italic" && r.formatting.italic === true)).toBe(true);
    expect(runs.some((r) => r.text === "underlined" && r.formatting.underline === true)).toBe(true);
    expect(runs.some((r) => r.text === "struck" && r.formatting.strikethrough === true)).toBe(true);
  });

  it("expands text:s with text:c into explicit spaces and text:tab into a tab (AC-8, AC-9)", async () => {
    const parsed = await parseOdt(fixture);
    expect(text(parsed.sections[0].blocks[2] as Paragraph)).toBe("Wide    gap and tab\tstop.");
  });

  it("maps text:line-break to a newline (AC-10)", async () => {
    const parsed = await parseOdt(fixture);
    expect(text(parsed.sections[0].blocks[3] as Paragraph)).toBe("Line one\nLine two");
  });

  it("resolves font family (via style:font-name and font-face-decls), size and colour (AC-18, AC-19, AC-20)", async () => {
    const parsed = await parseOdt(fixture);
    const run = (parsed.sections[0].blocks[5] as Paragraph).runs[0];
    expect(run.text).toBe("Red Sans 14pt");
    expect(run.formatting.fontFamily).toBe("Liberation Sans");
    expect(run.formatting.fontSizePt).toBe(14);
    expect(run.formatting.color).toBe("#c9211e");
  });

  it("maps fo:background-color to the OOXML highlight name the .docx reader uses (AC-21)", async () => {
    const parsed = await parseOdt(fixture);
    const runs = (parsed.sections[0].blocks[5] as Paragraph).runs;
    expect(runs.some((r) => r.text === "highlighted" && r.formatting.highlight === "yellow")).toBe(true);
  });

  it("resolves paragraph alignment, indents and spacing from automatic styles (AC-11, AC-22, AC-23, AC-24)", async () => {
    const parsed = await parseOdt(fixture);
    const formatting = (parsed.sections[0].blocks[6] as Paragraph).formatting;
    expect(formatting?.alignment).toBe("center");
    expect(formatting?.indentStartCm).toBeCloseTo(1.27, 3);
    expect(formatting?.indentEndCm).toBeCloseTo(1.27, 3);
    expect(formatting?.indentFirstLineCm).toBeCloseTo(0.635, 3);
    expect(formatting?.spaceBeforePt).toBeCloseTo(12, 1);
    expect(formatting?.spaceAfterPt).toBeCloseTo(12, 1);
    // Inherited from the automatic style's parent, Text_20_body (AC-13).
    expect(formatting?.lineSpacing).toBeCloseTo(1.15, 6);
  });

  it("resolves named-style paragraph spacing and proportional line height (AC-12, AC-24)", async () => {
    const parsed = await parseOdt(fixture);
    const formatting = (parsed.sections[0].blocks[1] as Paragraph).formatting;
    expect(formatting?.spaceBeforePt).toBe(0);
    expect(formatting?.spaceAfterPt).toBeCloseTo(7, 1);
    expect(formatting?.lineSpacing).toBeCloseTo(1.15, 6);
  });

  it("merges nested spans and keeps hyperlink text (AC-7)", async () => {
    const parsed = await parseOdt(fixture);
    const paragraph = parsed.sections[0].blocks[7] as Paragraph;
    expect(text(paragraph)).toBe("Nested bold bold-italic spans and a link text.");
    const nested = paragraph.runs.find((r) => r.text === "bold-italic");
    expect(nested?.formatting.bold).toBe(true);
    expect(nested?.formatting.italic).toBe(true);
  });

  it("applies the master page's page layout: size, orientation and margins (AC-25, AC-26, AC-27)", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.defaultStyles.pageSize.widthCm).toBeCloseTo(21, 1);
    expect(parsed.defaultStyles.pageSize.heightCm).toBeCloseTo(29.7, 1);
    expect(parsed.defaultStyles.pageSize.heightCm).toBeGreaterThan(parsed.defaultStyles.pageSize.widthCm);
    expect(parsed.defaultStyles.margins).toEqual({ topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 });
  });

  it("takes the document default font from styles.xml's default-style (AC-4)", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.defaultStyles.fontFamily).toBe("Liberation Serif");
    expect(parsed.defaultStyles.fontSizePt).toBe(12);
  });
});

describe("parseOdt — openoffice-era.odt", () => {
  const fixture = loadFixture(fixturesDir, "openoffice-era.odt");

  it("maps every text:p and text:h to a paragraph (AC-3, AC-5, AC-30)", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.sections[0].blocks).toHaveLength(8);
  });

  it("resolves heading levels, falling back to style:default-outline-level when text:h has no attribute (AC-6, AC-32)", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks[0].headingLevel).toBe(1);
    // "Financials" carries no text:outline-level; its style chain supplies 2.
    expect(blocks[2].headingLevel).toBe(2);
  });

  it("collapses literal XML whitespace like LibreOffice does (AC-8, AC-31)", async () => {
    const parsed = await parseOdt(fixture);
    expect(text(parsed.sections[0].blocks[1] as Paragraph)).toBe(
      "Literal whitespace collapses to single spaces.",
    );
  });

  it("resolves named heading styles to bold/italic/absolute sizes (AC-12, AC-13, AC-14, AC-15, AC-33)", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks[0].runs[0].formatting.fontFamily).toBe("Arial");
    expect(blocks[0].runs[0].formatting.fontSizePt).toBeCloseTo(16.1, 2);
    expect(blocks[0].runs[0].formatting.bold).toBe(true);
    expect(blocks[2].runs[0].formatting.bold).toBe(true);
    expect(blocks[2].runs[0].formatting.italic).toBe(true);
    expect(blocks[2].runs[0].formatting.fontSizePt).toBe(14);
  });

  it("resolves a named character style on a span (AC-12, AC-15)", async () => {
    const parsed = await parseOdt(fixture);
    const runs = (parsed.sections[0].blocks[3] as Paragraph).runs;
    expect(runs.some((r) => r.text === "emphasised" && r.formatting.italic === true)).toBe(true);
  });

  it("maps justify and end alignment (AC-22)", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks[3].formatting?.alignment).toBe("justify");
    expect(blocks[4].formatting?.alignment).toBe("right");
  });

  it("maps inch-based margins and a negative text-indent to a hanging indent (AC-23)", async () => {
    const parsed = await parseOdt(fixture);
    const formatting = (parsed.sections[0].blocks[5] as Paragraph).formatting;
    expect(formatting?.indentStartCm).toBeCloseTo(1.27, 3);
    expect(formatting?.indentEndCm).toBe(0);
    expect(formatting?.indentFirstLineCm).toBeCloseTo(-0.635, 3);
  });

  it("treats any underline line style (dotted) as underlined (AC-16)", async () => {
    const parsed = await parseOdt(fixture);
    const runs = (parsed.sections[0].blocks[6] as Paragraph).runs;
    expect(runs.some((r) => r.text === "underline" && r.formatting.underline === true)).toBe(true);
  });

  it("keeps an empty paragraph as a block with no runs", async () => {
    const parsed = await parseOdt(fixture);
    expect((parsed.sections[0].blocks[7] as Paragraph).runs).toEqual([]);
  });

  it("applies US Letter landscape and 1in margins from the inch-based page layout (AC-25, AC-26, AC-27)", async () => {
    const parsed = await parseOdt(fixture);
    const { pageSize, margins } = parsed.defaultStyles;
    expect(pageSize.widthCm).toBeCloseTo(27.94, 2);
    expect(pageSize.heightCm).toBeCloseTo(21.59, 2);
    expect(pageSize.widthCm).toBeGreaterThan(pageSize.heightCm);
    expect(margins.topCm).toBeCloseTo(2.54, 3);
    expect(margins.leftCm).toBeCloseTo(2.54, 3);
  });

  it("takes the document default font from the default-style (AC-4)", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.defaultStyles.fontFamily).toBe("Times New Roman");
    expect(parsed.defaultStyles.fontSizePt).toBe(12);
  });
});

describe("detectOdf", () => {
  it("detects an ODF package by its mimetype entry, not its name (AC-2)", async () => {
    const odt = loadFixture(fixturesDir, "libreoffice-writer.odt");
    expect(await detectOdf(odt)).toBe("package");
  });

  it("classifies a .docx package (no ODF mimetype entry) as none (AC-2)", async () => {
    const docx = loadFixture(docxFixturesDir, "word-authored.docx");
    expect(await detectOdf(docx)).toBe("none");
  });

  it("classifies flat XML ODF as flat (AC-29)", async () => {
    expect(await detectOdf(new TextEncoder().encode(FLAT_ODF))).toBe("flat");
  });

  it("classifies arbitrary non-ODF bytes as none", async () => {
    expect(await detectOdf(new TextEncoder().encode("plain text"))).toBe("none");
    expect(await detectOdf(new TextEncoder().encode("<html><body/></html>"))).toBe("none");
  });
});

describe("parseOdt — error handling", () => {
  it("rejects a flat XML ODF file with the specific message (AC-29)", async () => {
    const flat = new TextEncoder().encode(FLAT_ODF);
    await expect(parseOdt(flat)).rejects.toThrow(FlatOdfError);
    await expect(parseOdt(flat)).rejects.toThrow(FLAT_ODF_MESSAGE);
  });

  it("rejects a file that is not a ZIP", async () => {
    const notZip = new TextEncoder().encode("not a zip file at all");
    await expect(parseOdt(notZip)).rejects.toThrow(OdtOpenError);
    await expect(parseOdt(notZip)).rejects.toThrow(INVALID_ODT_MESSAGE);
  });

  it("rejects a ZIP without content.xml", async () => {
    // A .docx package is exactly that: a valid ZIP with no content.xml.
    const docx = loadFixture(docxFixturesDir, "word-authored.docx");
    await expect(parseOdt(docx)).rejects.toThrow(OdtOpenError);
    await expect(parseOdt(docx)).rejects.toThrow(INVALID_ODT_MESSAGE);
  });
});
