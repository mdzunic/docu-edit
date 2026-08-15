import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { deflateRawSync } from "node:zlib";
import { DocxOpenError, LegacyDocError, LEGACY_DOC_MESSAGE, INVALID_DOCX_MESSAGE, parseDocx } from "./reader";
import type { Paragraph } from "../documentModel";
import { asParagraph } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

function text(paragraph: Paragraph): string {
  return paragraph.runs.map((run) => run.text).join("");
}

// These three fixtures are hand-built OPC packages that reproduce known
// structural differences between Word, LibreOffice and Google Docs exports
// (see scripts/build-docx-fixtures.mjs for the rationale and exact markup) —
// this environment has no network access to fetch genuine exports.
describe("parseDocx — word-authored.docx", () => {
  const fixture = loadFixture("word-authored.docx");

  it("maps every w:p to a paragraph (AC-4, AC-36)", async () => {
    const parsed = await parseDocx(fixture);
    expect(parsed.sections[0].blocks).toHaveLength(4);
  });

  it("resolves Heading1/Heading2 pStyle references to heading levels (AC-24, AC-38)", async () => {
    const parsed = await parseDocx(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks[0].headingLevel).toBe(1);
    expect(blocks[1].headingLevel).toBeUndefined();
    expect(blocks[2].headingLevel).toBe(2);
    expect(blocks[3].headingLevel).toBeUndefined();
  });

  it("renders built-in heading styles at their defined size and weight (AC-25)", async () => {
    const parsed = await parseDocx(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks[0].runs[0].formatting.bold).toBe(true);
    expect(blocks[0].runs[0].formatting.fontSizePt).toBe(16);
    expect(blocks[2].runs[0].formatting.bold).toBe(true);
    expect(blocks[2].runs[0].formatting.fontSizePt).toBe(13);
  });

  it("preserves text content, including runs of the heading paragraphs (AC-37)", async () => {
    const parsed = await parseDocx(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(text(blocks[0])).toBe("Quarterly Report");
    expect(text(blocks[2])).toBe("Section Two");
  });

  it("reads bold and italic direct run formatting (AC-9, AC-10)", async () => {
    const parsed = await parseDocx(fixture);
    const runs = (parsed.sections[0].blocks[1] as Paragraph).runs;
    expect(runs.some((r) => r.text === "bold" && r.formatting.bold)).toBe(true);
    expect(runs.some((r) => r.text === "italic" && r.formatting.italic)).toBe(true);
  });

  it("preserves tabs (w:tab) and preserved spaces (xml:space=preserve) (AC-6, AC-7)", async () => {
    const parsed = await parseDocx(fixture);
    const runs = (parsed.sections[0].blocks[1] as Paragraph).runs;
    expect(runs.some((r) => r.text === " Tab\there. ")).toBe(true);
  });

  it("preserves line breaks (w:br) (AC-8)", async () => {
    const parsed = await parseDocx(fixture);
    const runs = (parsed.sections[0].blocks[1] as Paragraph).runs;
    expect(runs.some((r) => r.text === "Line one\nLine two")).toBe(true);
  });

  it("reads paragraph alignment, indents and spacing (AC-17..23)", async () => {
    const parsed = await parseDocx(fixture);
    const formatting = (parsed.sections[0].blocks[3] as Paragraph).formatting;
    expect(formatting?.alignment).toBe("center");
    expect(formatting?.indentStartCm).toBeCloseTo(1.27, 2);
    expect(formatting?.indentEndCm).toBeCloseTo(1.27, 2);
    expect(formatting?.indentFirstLineCm).toBeCloseTo(0.635, 2);
    expect(formatting?.spaceBeforePt).toBe(12);
    expect(formatting?.spaceAfterPt).toBe(12);
    expect(formatting?.lineSpacing).toBe(1.5);
  });

  it("applies A4 portrait page size and standard margins from sectPr (AC-26, AC-28)", async () => {
    const parsed = await parseDocx(fixture);
    expect(parsed.defaultStyles.pageSize.widthCm).toBeCloseTo(21, 1);
    expect(parsed.defaultStyles.pageSize.heightCm).toBeCloseTo(29.7, 1);
    expect(parsed.defaultStyles.margins).toEqual({
      topCm: 2.54,
      rightCm: 2.54,
      bottomCm: 2.54,
      leftCm: 2.54,
    });
  });
});

describe("parseDocx — libreoffice-authored.docx", () => {
  const fixture = loadFixture("libreoffice-authored.docx");

  it("maps every w:p to a paragraph (AC-4, AC-36)", async () => {
    const parsed = await parseDocx(fixture);
    expect(parsed.sections[0].blocks).toHaveLength(5);
  });

  it("resolves heading levels from 'heading N'-named styles (AC-24, AC-38)", async () => {
    const parsed = await parseDocx(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks.map((b) => asParagraph(b).headingLevel)).toEqual([1, undefined, undefined, 2, undefined]);
  });

  it("reads underline and strikethrough direct run formatting (AC-11, AC-12)", async () => {
    const parsed = await parseDocx(fixture);
    const runs = (parsed.sections[0].blocks[1] as Paragraph).runs;
    expect(runs.some((r) => r.text === "underlined" && r.formatting.underline)).toBe(true);
    expect(runs.some((r) => r.text === "struck through" && r.formatting.strikethrough)).toBe(true);
  });

  it("preserves leading and trailing spaces verbatim (AC-6, AC-37)", async () => {
    const parsed = await parseDocx(fixture);
    expect(text(parsed.sections[0].blocks[2] as Paragraph)).toBe("   spaced text   ");
  });

  it("reads right alignment (AC-17)", async () => {
    const parsed = await parseDocx(fixture);
    expect((parsed.sections[0].blocks[4] as Paragraph).formatting?.alignment).toBe("right");
  });

  it("applies landscape orientation and 2cm margins from sectPr (AC-26, AC-27, AC-28)", async () => {
    const parsed = await parseDocx(fixture);
    const { pageSize, margins } = parsed.defaultStyles;
    expect(pageSize.widthCm).toBeGreaterThan(pageSize.heightCm);
    expect(pageSize.widthCm).toBeCloseTo(29.7, 1);
    expect(pageSize.heightCm).toBeCloseTo(21, 1);
    expect(margins.topCm).toBeCloseTo(2, 1);
  });
});

describe("parseDocx — google-docs-exported.docx", () => {
  const fixture = loadFixture("google-docs-exported.docx");

  it("maps every w:p to a paragraph, unaffected by interleaved w:proofErr markers (AC-4, AC-36)", async () => {
    const parsed = await parseDocx(fixture);
    expect(parsed.sections[0].blocks).toHaveLength(6);
  });

  it("resolves three distinct heading levels (AC-24, AC-38)", async () => {
    const parsed = await parseDocx(fixture);
    const blocks = parsed.sections[0].blocks as Paragraph[];
    expect(blocks.map((b) => asParagraph(b).headingLevel)).toEqual([
      1,
      2,
      3,
      undefined,
      undefined,
      undefined,
    ]);
  });

  it("preserves text across a w:proofErr-interrupted run sequence (AC-37)", async () => {
    const parsed = await parseDocx(fixture);
    expect(text(parsed.sections[0].blocks[3] as Paragraph)).toBe("Colored text in Arial 14pt.");
  });

  it("reads font family, font size and text colour (AC-13, AC-14, AC-15)", async () => {
    const parsed = await parseDocx(fixture);
    const run = (parsed.sections[0].blocks[3] as Paragraph).runs[0];
    expect(run.formatting.fontFamily).toBe("Arial");
    expect(run.formatting.fontSizePt).toBe(14);
    expect(run.formatting.color).toBe("#1155cc");
  });

  it("reads highlight (AC-16)", async () => {
    const parsed = await parseDocx(fixture);
    const runs = (parsed.sections[0].blocks[4] as Paragraph).runs;
    expect(runs.some((r) => r.text === "highlighted text" && r.formatting.highlight === "yellow")).toBe(true);
  });

  it("preserves line breaks (AC-8)", async () => {
    const parsed = await parseDocx(fixture);
    expect(text(parsed.sections[0].blocks[5] as Paragraph)).toBe("Line one\nLine two");
  });

  it("applies US Letter page size from sectPr (AC-26)", async () => {
    const parsed = await parseDocx(fixture);
    expect(parsed.defaultStyles.pageSize.widthCm).toBeCloseTo(21.59, 1);
    expect(parsed.defaultStyles.pageSize.heightCm).toBeCloseTo(27.94, 1);
  });
});

describe("parseDocx — special characters", () => {
  it("maps w:noBreakHyphen and w:softHyphen to U+2011 and U+00AD (AC-53)", async () => {
    const documentXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      "<w:p><w:r><w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t><w:softHyphen/><w:t>c</w:t></w:r></w:p>" +
      "</w:body></w:document>";
    const { zip } = await import("../core/zipWriter");
    const bytes = await zip([
      { name: "word/document.xml", data: new TextEncoder().encode(documentXml) },
    ]);
    const parsed = await parseDocx(bytes);
    expect(text(parsed.sections[0].blocks[0] as Paragraph)).toBe("a‑b­c");
  });
});

describe("parseDocx — error handling", () => {
  it("rejects a file that is not a valid ZIP (AC-31)", async () => {
    const notZip = new TextEncoder().encode("not a zip file at all");
    await expect(parseDocx(notZip)).rejects.toThrow(DocxOpenError);
    await expect(parseDocx(notZip)).rejects.toThrow(INVALID_DOCX_MESSAGE);
  });

  it("rejects a ZIP without word/document.xml (AC-32)", async () => {
    const entries = [
      { name: "readme.txt", content: "hello" },
    ];
    const u32 = (v: number) => {
      const b = new Uint8Array(4);
      new DataView(b.buffer).setUint32(0, v, true);
      return b;
    };
    const u16 = (v: number) => {
      const b = new Uint8Array(2);
      new DataView(b.buffer).setUint16(0, v, true);
      return b;
    };
    const concat = (parts: Uint8Array[]) => {
      const total = parts.reduce((s, p) => s + p.length, 0);
      const out = new Uint8Array(total);
      let o = 0;
      for (const p of parts) {
        out.set(p, o);
        o += p.length;
      }
      return out;
    };
    const encoder = new TextEncoder();
    const nameBytes = encoder.encode(entries[0].name);
    const rawBytes = encoder.encode(entries[0].content);
    const dataBytes = deflateRawSync(rawBytes);
    const localHeader = concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(0),
      u32(dataBytes.length),
      u32(rawBytes.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes,
    ]);
    const centralHeader = concat([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(0),
      u32(dataBytes.length),
      u32(rawBytes.length),
      u16(nameBytes.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(0),
      nameBytes,
    ]);
    const localData = concat([localHeader, dataBytes]);
    const eocd = concat([
      u32(0x06054b50),
      u16(0),
      u16(0),
      u16(1),
      u16(1),
      u32(centralHeader.length),
      u32(localData.length),
      u16(0),
    ]);
    const zip = concat([localData, centralHeader, eocd]);

    await expect(parseDocx(zip)).rejects.toThrow(DocxOpenError);
    await expect(parseDocx(zip)).rejects.toThrow(INVALID_DOCX_MESSAGE);
  });

  it("rejects a binary Word 97 (.doc) file (AC-34)", async () => {
    const docMagic = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
    await expect(parseDocx(docMagic)).rejects.toThrow(LegacyDocError);
    await expect(parseDocx(docMagic)).rejects.toThrow(LEGACY_DOC_MESSAGE);
  });
});
