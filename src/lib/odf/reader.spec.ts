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
import { asParagraph } from "../documentModel";

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
    expect(blocks.map((b) => asParagraph(b).headingLevel)).toEqual([
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

async function makeOdt(contentXml: string, stylesXml?: string): Promise<Uint8Array> {
  const { zip } = await import("../docx/zipWriter");
  const encoder = new TextEncoder();
  const entries = [
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.text"), store: true },
    { name: "content.xml", data: encoder.encode(contentXml) },
  ];
  if (stylesXml !== undefined) {
    entries.push({ name: "styles.xml", data: encoder.encode(stylesXml) });
  }
  return zip(entries);
}

describe("parseOdt — RTL paragraphs (SPEC-014 AC-24, AC-25)", () => {
  // Hand-built to LibreOffice Writer's exact RTL serialisation (matching the
  // other fixtures' provenance note above): LO stores indents PHYSICALLY in
  // fo:margin-left/right per XSL, with style:writing-mode="rl-tb" marking the
  // direction — a "before text" (start) indent of an RTL paragraph lands in
  // fo:margin-right. The reader must swap the physical edges back onto the
  // model's logical start/end.
  const RTL_CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">
 <office:automatic-styles>
  <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard">
   <style:paragraph-properties fo:margin-left="1.27cm" fo:margin-right="2.54cm" fo:text-indent="0.635cm" style:auto-text-indent="false" style:writing-mode="rl-tb"/>
  </style:style>
  <style:style style:name="P2" style:family="paragraph" style:parent-style-name="Standard">
   <style:paragraph-properties style:writing-mode="lr-tb"/>
  </style:style>
 </office:automatic-styles>
 <office:body>
  <office:text>
   <text:p text:style-name="P1">فقرة من اليمين إلى اليسار</text:p>
   <text:p text:style-name="P2">explicit ltr</text:p>
  </office:text>
 </office:body>
</office:document-content>`;

  it("swaps LibreOffice's physical margins onto the logical start/end", async () => {
    const parsed = await parseOdt(await makeOdt(RTL_CONTENT));
    const formatting = (parsed.sections[0].blocks[0] as Paragraph).formatting;
    expect(formatting?.direction).toBe("rtl");
    // Physical right margin (2.54) is the logical START edge for rl-tb;
    // physical left (1.27) is the logical END.
    expect(formatting?.indentStartCm).toBe(2.54);
    expect(formatting?.indentEndCm).toBe(1.27);
    // fo:text-indent is start-relative in both formats — never swapped.
    expect(formatting?.indentFirstLineCm).toBe(0.635);
  });

  it("never stores an explicit ltr — absent means left-to-right", async () => {
    const parsed = await parseOdt(await makeOdt(RTL_CONTENT));
    const formatting = (parsed.sections[0].blocks[1] as Paragraph).formatting;
    expect(formatting?.direction).toBeUndefined();
  });
});

// Hand-built to the markup this project's own writer and LibreOffice emit
// for SPEC-028's pagination properties (same provenance note as the RTL
// fixture above).
function spec028Content(automaticStyles: string, body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">
 <office:automatic-styles>${automaticStyles}</office:automatic-styles>
 <office:body><office:text>${body}</office:text></office:body>
</office:document-content>`;
}

describe("parseOdt — pagination flags (SPEC-028 AC-22..29)", () => {
  it("maps keep-with-next, keep-together and zero widows/orphans onto the model flags", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph">' +
            '<style:paragraph-properties fo:keep-with-next="always" fo:keep-together="always"' +
            ' fo:widows="0" fo:orphans="0"/></style:style>',
          '<text:p text:style-name="P1">flags</text:p>',
        ),
      ),
    );
    const formatting = asParagraph(parsed.sections[0].blocks[0]).formatting;
    expect(formatting?.keepNext).toBe(true);
    expect(formatting?.keepLines).toBe(true);
    expect(formatting?.widowControl).toBe(false);
  });

  it("reads ODF's default values (auto keeps, non-zero widows/orphans) as absent — no model noise", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph">' +
            '<style:paragraph-properties fo:keep-with-next="auto" fo:keep-together="auto"' +
            ' fo:widows="2" fo:orphans="2"/></style:style>',
          '<text:p text:style-name="P1">defaults</text:p>',
        ),
      ),
    );
    const formatting = asParagraph(parsed.sections[0].blocks[0]).formatting;
    expect(formatting?.keepNext).toBeUndefined();
    expect(formatting?.keepLines).toBeUndefined();
    expect(formatting?.widowControl).toBeUndefined();
  });
});

describe("parseOdt — manual page breaks (SPEC-028 AC-32)", () => {
  it("materialises fo:break-before='page' on an AUTOMATIC style as a BreakBlock, dropping the flag", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph">' +
            '<style:paragraph-properties fo:break-before="page"/></style:style>',
          '<text:p>before</text:p><text:p text:style-name="P1">after</text:p>',
        ),
      ),
    );
    const blocks = parsed.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "break", "paragraph"]);
    expect(blocks[1]).toEqual({ type: "break", kind: "page" });
    // The paragraph's own formatting must NOT also carry pageBreakBefore —
    // one channel, canonicalised to the break block.
    expect(asParagraph(blocks[2]).formatting?.pageBreakBefore).toBeUndefined();
  });

  it("keeps fo:break-before='page' from a NAMED style as pure pageBreakBefore formatting", async () => {
    const styles =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
      ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
      ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
      "<office:styles>" +
      '<style:style style:name="PageStart" style:family="paragraph">' +
      '<style:paragraph-properties fo:break-before="page"/></style:style>' +
      "</office:styles></office:document-styles>";
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content("", '<text:p>before</text:p><text:p text:style-name="PageStart">after</text:p>'),
        styles,
      ),
    );
    const blocks = parsed.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    expect(asParagraph(blocks[1]).formatting?.pageBreakBefore).toBe(true);
  });

  it("ignores fo:break-before='column' until columns exist (SPEC-030)", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph">' +
            '<style:paragraph-properties fo:break-before="column"/></style:style>',
          '<text:p>before</text:p><text:p text:style-name="P1">after</text:p>',
        ),
      ),
    );
    const blocks = parsed.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    expect(asParagraph(blocks[1]).formatting?.pageBreakBefore).toBeUndefined();
  });
});

describe("parseOdt — sections via master pages (SPEC-028 AC-33)", () => {
  const SECTION_STYLES =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
    ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
    ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
    "<office:automatic-styles>" +
    '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm"' +
    ' fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm"' +
    ' fo:margin-left="2cm"/></style:page-layout>' +
    '<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="14.8cm"' +
    ' fo:page-height="21cm" fo:margin-top="1cm" fo:margin-right="1cm" fo:margin-bottom="1cm"' +
    ' fo:margin-left="1cm"/></style:page-layout>' +
    "</office:automatic-styles>" +
    "<office:master-styles>" +
    '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
    '<style:master-page style:name="A5Master" style:page-layout-name="pm2"/>' +
    "</office:master-styles></office:document-styles>";

  it("starts a new section at a paragraph whose automatic style names a master page", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph" style:master-page-name="A5Master"/>',
          '<text:p>first</text:p><text:p text:style-name="P1">second</text:p>',
        ),
        SECTION_STYLES,
      ),
    );
    expect(parsed.sections).toHaveLength(2);
    expect(text(asParagraph(parsed.sections[0].blocks[0]))).toBe("first");
    expect(text(asParagraph(parsed.sections[1].blocks[0]))).toBe("second");
    // The new section pins the master page's geometry explicitly; `start`
    // stays absent (a master-page switch is always a page break).
    expect(parsed.sections[1].pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
    expect(parsed.sections[1].margins).toEqual({ topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 });
    expect(parsed.sections[1].start).toBeUndefined();
    // The FIRST master page stays folded into the document defaults.
    expect(parsed.defaultStyles.pageSize).toEqual({ widthCm: 21, heightCm: 29.7 });
    expect(parsed.defaultStyles.margins).toEqual({ topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 });
  });

  it("never opens an empty leading section for a master reference on the very first paragraph", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph" style:master-page-name="A5Master"/>',
          '<text:p text:style-name="P1">only</text:p>',
        ),
        SECTION_STYLES,
      ),
    );
    expect(parsed.sections).toHaveLength(1);
    expect(text(asParagraph(parsed.sections[0].blocks[0]))).toBe("only");
  });

  it("folds the geometry of the master the FIRST paragraph names — not the first-listed one (review #55 bug)", async () => {
    // SECTION_STYLES lists Standard/pm1 (21x29.7) BEFORE A5Master/pm2
    // (14.8x21). The very first body paragraph references A5Master — the
    // master the body actually starts on. The document default and the
    // single section must report A5, never the first-LISTED Standard.
    const parsed = await parseOdt(
      await makeOdt(
        spec028Content(
          '<style:style style:name="P1" style:family="paragraph" style:master-page-name="A5Master"/>',
          '<text:p text:style-name="P1">starts-on-a5</text:p>',
        ),
        SECTION_STYLES,
      ),
    );
    expect(parsed.sections).toHaveLength(1);
    expect(parsed.defaultStyles.pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
    expect(parsed.defaultStyles.margins).toEqual({ topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 });
    // The leading section pins the referenced master's geometry explicitly,
    // the same way later sections that switch masters do (AC-17/18/19).
    expect(parsed.sections[0].pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
    expect(parsed.sections[0].margins).toEqual({ topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 });
  });
});
