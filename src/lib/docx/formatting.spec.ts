// SPEC-013 — font size, colour, highlight (w:highlight + w:shd fallback)
// and sub/superscript through the .docx reader, writer and validator.
import { describe, it, expect } from "vitest";
import { createDocument, type DocumentModel, type RunFormatting } from "../documentModel";
import { parseRunProperties } from "./properties";
import { parseDocx } from "./reader";
import { buildDocumentXml, writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { parseXml, WORD_NAMESPACE } from "./xml";

function rPr(inner: string): Element {
  return parseXml(`<w:rPr xmlns:w="${WORD_NAMESPACE}">${inner}</w:rPr>`).documentElement;
}

function docWithRun(formatting: RunFormatting, text = "hello"): DocumentModel {
  const model = createDocument("docx-formatting-test");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text, formatting }] }] }];
  return model;
}

describe("parseRunProperties — highlight, shading, vertAlign (SPEC-013)", () => {
  it('maps w:highlight w:val="none" to the explicit false, not an absence', () => {
    expect(parseRunProperties(rPr('<w:highlight w:val="none"/>')).highlight).toBe(false);
  });

  it("keeps a palette name, including white, as the name", () => {
    expect(parseRunProperties(rPr('<w:highlight w:val="yellow"/>')).highlight).toBe("yellow");
    expect(parseRunProperties(rPr('<w:highlight w:val="white"/>')).highlight).toBe("white");
  });

  it("reads a lone w:shd as the highlight, normalised to a palette name when it matches", () => {
    expect(
      parseRunProperties(rPr('<w:shd w:val="clear" w:color="auto" w:fill="FF8800"/>')),
    ).toEqual({ highlight: "#ff8800" });
    expect(
      parseRunProperties(rPr('<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')),
    ).toEqual({ highlight: "yellow" });
  });

  it("keeps both when a run carries w:highlight and w:shd: highlight wins, shading preserved", () => {
    expect(
      parseRunProperties(
        rPr('<w:highlight w:val="yellow"/><w:shd w:val="clear" w:color="auto" w:fill="FF8800"/>'),
      ),
    ).toEqual({ highlight: "yellow", shadingFill: "#ff8800" });
  });

  it('ignores w:shd with fill="auto" and reads w:vertAlign sub/superscript', () => {
    expect(parseRunProperties(rPr('<w:shd w:val="clear" w:fill="auto"/>'))).toEqual({});
    expect(parseRunProperties(rPr('<w:vertAlign w:val="subscript"/>')).vertAlign).toBe("subscript");
    expect(parseRunProperties(rPr('<w:vertAlign w:val="superscript"/>')).vertAlign).toBe(
      "superscript",
    );
    expect(parseRunProperties(rPr('<w:vertAlign w:val="baseline"/>')).vertAlign).toBeUndefined();
  });
});

describe("buildDocumentXml — highlight mapping on save (SPEC-013)", () => {
  it("writes a palette highlight as w:highlight with no shading", () => {
    const xml = buildDocumentXml(docWithRun({ highlight: "yellow" }));
    expect(xml).toContain('<w:highlight w:val="yellow"/>');
    expect(xml).not.toContain("<w:shd");
  });

  it('writes "No colour" as w:highlight w:val="none"', () => {
    const xml = buildDocumentXml(docWithRun({ highlight: false }));
    expect(xml).toContain('<w:highlight w:val="none"/>');
  });

  it("writes an out-of-palette colour as clear shading with NO w:highlight", () => {
    const xml = buildDocumentXml(docWithRun({ highlight: "#ff8800" }));
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FF8800"/>');
    expect(xml).not.toContain("<w:highlight");
  });

  it("re-emits a preserved shadingFill alongside the highlight", () => {
    const xml = buildDocumentXml(docWithRun({ highlight: "yellow", shadingFill: "#ff8800" }));
    expect(xml).toContain('<w:highlight w:val="yellow"/>');
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FF8800"/>');
  });

  it("writes w:vertAlign after w:shd, per the CT_RPr sequence", () => {
    const xml = buildDocumentXml(
      docWithRun({ underline: true, highlight: "#ff8800", vertAlign: "subscript" }),
    );
    const u = xml.indexOf("<w:u ");
    const shd = xml.indexOf("<w:shd");
    const vertAlign = xml.indexOf("<w:vertAlign");
    expect(u).toBeGreaterThan(-1);
    expect(shd).toBeGreaterThan(u);
    expect(vertAlign).toBeGreaterThan(shd);
    expect(xml).toContain('<w:vertAlign w:val="subscript"/>');
  });
});

describe("writeDocx — the SPEC-013 package validates and round-trips", () => {
  const formatted: RunFormatting = {
    fontFamily: "Georgia",
    fontSizePt: 11.5,
    color: "#1155cc",
    highlight: "yellow",
    vertAlign: "superscript",
  };

  it("produces a schema-valid package for every new run property", async () => {
    const model = docWithRun(formatted);
    model.sections[0].blocks[0].runs.push(
      { text: " shaded", formatting: { highlight: "#ff8800" } },
      { text: " off", formatting: { highlight: false } },
      { text: " both", formatting: { highlight: "cyan", shadingFill: "#123456" } },
    );
    const { bytes } = await writeDocx(model);
    expect(await validateDocx(bytes)).toEqual([]);
  });

  it("round-trips font, size, colour, highlight and vertAlign exactly", async () => {
    const { bytes } = await writeDocx(docWithRun(formatted));
    const reopened = await parseDocx(bytes);
    const run = reopened.sections[0].blocks[0].runs[0];
    expect(run.formatting.fontFamily).toBe("Georgia");
    expect(run.formatting.fontSizePt).toBe(11.5);
    expect(run.formatting.color).toBe("#1155cc");
    expect(run.formatting.highlight).toBe("yellow");
    expect(run.formatting.vertAlign).toBe("superscript");
  });

  it("has no font-size drift across a double round trip, at the range ends too (AC)", async () => {
    for (const size of [1, 11.5, 208.5, 409]) {
      const first = await writeDocx(docWithRun({ fontSizePt: size }));
      const opened = await parseDocx(first.bytes);
      expect(opened.sections[0].blocks[0].runs[0].formatting.fontSizePt).toBe(size);
      const model = docWithRun(opened.sections[0].blocks[0].runs[0].formatting);
      const second = await writeDocx(model);
      const reopened = await parseDocx(second.bytes);
      expect(reopened.sections[0].blocks[0].runs[0].formatting.fontSizePt).toBe(size);
    }
  });

  it("round-trips an out-of-palette highlight through the shading fallback (AC)", async () => {
    const { bytes } = await writeDocx(docWithRun({ highlight: "#ff8800" }));
    const reopened = await parseDocx(bytes);
    expect(reopened.sections[0].blocks[0].runs[0].formatting.highlight).toBe("#ff8800");
  });

  it('round-trips "No colour" as an explicit off (AC)', async () => {
    const { bytes } = await writeDocx(docWithRun({ highlight: false }));
    const reopened = await parseDocx(bytes);
    expect(reopened.sections[0].blocks[0].runs[0].formatting.highlight).toBe(false);
  });

  it("keeps w:highlight and w:shd together across a .docx round trip (AC)", async () => {
    const { bytes } = await writeDocx(docWithRun({ highlight: "yellow", shadingFill: "#ff8800" }));
    const reopened = await parseDocx(bytes);
    const formatting = reopened.sections[0].blocks[0].runs[0].formatting;
    expect(formatting.highlight).toBe("yellow");
    expect(formatting.shadingFill).toBe("#ff8800");
  });
});
