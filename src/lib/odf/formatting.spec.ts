// SPEC-013 — font size, colour, highlight (fo:background-color incl. the
// explicit "transparent" off) and sub/superscript (style:text-position)
// through the .odt reader and writer, plus the cross-format run-shading
// drop report.
import { describe, it, expect } from "vitest";
import { asParagraph, createDocument, type DocumentModel, type RunFormatting } from "../documentModel";
import { parseTextProperties } from "./properties";
import { parseOdt } from "./reader";
import { buildContentXml, writeOdt } from "./writer";
import { parseXml } from "../docx/xml";
import { FO_NAMESPACE, STYLE_NAMESPACE } from "./xml";

function textProperties(attrs: string): Element {
  return parseXml(
    `<style:text-properties xmlns:style="${STYLE_NAMESPACE}" xmlns:fo="${FO_NAMESPACE}" ${attrs}/>`,
  ).documentElement;
}

function docWithRun(formatting: RunFormatting, text = "hello"): DocumentModel {
  const model = createDocument("odf-formatting-test");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text, formatting }] }] }];
  return model;
}

describe("parseTextProperties — background and text-position (SPEC-013)", () => {
  const noFonts = new Map<string, string>();

  it("maps a palette background to the OOXML name, including white", () => {
    expect(
      parseTextProperties(textProperties('fo:background-color="#ffff00"'), noFonts).formatting
        .highlight,
    ).toBe("yellow");
    expect(
      parseTextProperties(textProperties('fo:background-color="#ffffff"'), noFonts).formatting
        .highlight,
    ).toBe("white");
  });

  it("keeps an out-of-palette background as a raw hex colour", () => {
    expect(
      parseTextProperties(textProperties('fo:background-color="#FF8800"'), noFonts).formatting
        .highlight,
    ).toBe("#ff8800");
  });

  it('maps "transparent" to the explicit false, not an absence', () => {
    expect(
      parseTextProperties(textProperties('fo:background-color="transparent"'), noFonts).formatting
        .highlight,
    ).toBe(false);
  });

  it("reads style:text-position keywords and signed percentages", () => {
    const read = (value: string) =>
      parseTextProperties(textProperties(`style:text-position="${value}"`), noFonts).formatting
        .vertAlign;
    expect(read("sub 58%")).toBe("subscript");
    expect(read("super 58%")).toBe("superscript");
    expect(read("-33% 58%")).toBe("subscript");
    expect(read("33%")).toBe("superscript");
    expect(read("0% 100%")).toBeUndefined();
  });
});

describe("buildContentXml — background and text-position (SPEC-013)", () => {
  it("writes a palette highlight as its hex background colour", () => {
    expect(buildContentXml(docWithRun({ highlight: "yellow" }))).toContain(
      'fo:background-color="#ffff00"',
    );
  });

  it('writes "No colour" as fo:background-color="transparent"', () => {
    expect(buildContentXml(docWithRun({ highlight: false }))).toContain(
      'fo:background-color="transparent"',
    );
  });

  it("writes an out-of-palette highlight colour directly — .odt needs no fallback", () => {
    expect(buildContentXml(docWithRun({ highlight: "#ff8800" }))).toContain(
      'fo:background-color="#ff8800"',
    );
  });

  it("writes sub/superscript as style:text-position", () => {
    expect(buildContentXml(docWithRun({ vertAlign: "subscript" }))).toContain(
      'style:text-position="sub 58%"',
    );
    expect(buildContentXml(docWithRun({ vertAlign: "superscript" }))).toContain(
      'style:text-position="super 58%"',
    );
  });
});

describe("writeOdt — the SPEC-013 formatting round-trips", () => {
  it("round-trips font, size, colour, highlight and vertAlign exactly", async () => {
    const { bytes } = await writeOdt(
      docWithRun({
        fontFamily: "Georgia",
        fontSizePt: 11.5,
        color: "#1155cc",
        highlight: "yellow",
        vertAlign: "superscript",
      }),
    );
    const reopened = await parseOdt(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.fontFamily).toBe("Georgia");
    expect(run.formatting.fontSizePt).toBe(11.5);
    expect(run.formatting.color).toBe("#1155cc");
    expect(run.formatting.highlight).toBe("yellow");
    expect(run.formatting.vertAlign).toBe("superscript");
  });

  it("has no font-size drift across a double round trip (points in, points out) (AC)", async () => {
    for (const size of [1, 11.5, 208.5, 409]) {
      const first = await writeOdt(docWithRun({ fontSizePt: size }));
      const opened = await parseOdt(first.bytes);
      expect(asParagraph(opened.sections[0].blocks[0]).runs[0].formatting.fontSizePt).toBe(size);
      const second = await writeOdt(docWithRun(asParagraph(opened.sections[0].blocks[0]).runs[0].formatting));
      const reopened = await parseOdt(second.bytes);
      expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting.fontSizePt).toBe(size);
    }
  });

  it("round-trips a custom highlight and the explicit 'No colour' (AC)", async () => {
    const { bytes } = await writeOdt(
      docWithRun({ highlight: "#ff8800" }, "custom"),
    );
    const custom = await parseOdt(bytes);
    expect(asParagraph(custom.sections[0].blocks[0]).runs[0].formatting.highlight).toBe("#ff8800");

    const { bytes: offBytes } = await writeOdt(docWithRun({ highlight: false }, "off"));
    const off = await parseOdt(offBytes);
    expect(asParagraph(off.sections[0].blocks[0]).runs[0].formatting.highlight).toBe(false);
  });

  it('keeps w:highlight w:val="white" alive across .docx -> .odt -> .docx (SPEC-013 decision)', async () => {
    const { bytes } = await writeOdt(docWithRun({ highlight: "white" }));
    const reopened = await parseOdt(bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting.highlight).toBe("white");
  });

  it("reports run shading as dropped on a .docx -> .odt save (SPEC-013)", async () => {
    const model = docWithRun({ highlight: "yellow", shadingFill: "#ff8800" });
    model.format = "docx";
    const { bytes, dropped } = await writeOdt(model);
    const shading = dropped.find((record) => record.kind === "runShading");
    expect(shading).toBeDefined();
    expect(shading!.location).toBe("paragraph 1");
    expect(shading!.reason).toContain("Not preserved");
    // The highlight itself is kept.
    const reopened = await parseOdt(bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting.highlight).toBe("yellow");
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting.shadingFill).toBeUndefined();
  });
});
