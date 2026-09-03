import { describe, expect, it } from "vitest";
import { applyParagraphStyle, caretSelection, updateStyleFromSelection } from "../commands";
import { asParagraph, createDocument, type DocumentModel } from "../documentModel";
import { equalBytes } from "../core/equality";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { buildStylesXml, patchStylesXml, writeDocx } from "./writer";
import { unzip } from "../core/zip";
import { WORD_NAMESPACE } from "./xml";

// SPEC-019 — style definitions in the .docx writer: from-scratch saves
// declare the model's catalogue with each style's own formatting (AC-15,
// AC-11), and same-format saves keep the preserved styles.xml byte-identical
// unless the catalogue diverges from it, in which case only the changed or
// added w:style elements are regenerated.

const decoder = new TextDecoder();

function caretAtStart() {
  return caretSelection({ section: 0, block: 0, offset: 0 });
}

function typedModel(text = "Chapter one"): DocumentModel {
  const model = createDocument("writer-styles");
  model.sections[0].blocks = [{ type: "paragraph", runs: [{ text, formatting: {} }] }];
  return model;
}

/** The model an opened document holds (openDocument.ts's shape), retained
 * source package included. */
function modelFrom(parsed: ParsedDocxDocument): DocumentModel {
  const model: DocumentModel = {
    id: "reopened",
    title: "Reopened",
    format: "docx",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
  if (parsed.styles.length > 0) {
    model.styles = parsed.styles;
  }
  return model;
}

async function partText(bytes: Uint8Array, name: string): Promise<string> {
  return decoder.decode((await unzip(bytes)).get(name)!);
}

describe("buildStylesXml — catalogue definitions (AC-15)", () => {
  it("declares an applied heading with its formatting, name and outline level", () => {
    const model = applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model;
    const xml = buildStylesXml(model);
    expect(xml).toContain('w:styleId="Heading1"');
    expect(xml).toContain('<w:name w:val="Heading 1"/>');
    expect(xml).toContain('<w:basedOn w:val="Normal"/>');
    expect(xml).toContain('<w:outlineLvl w:val="0"/>');
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain('<w:sz w:val="40"/>');
  });

  it("keeps the historical name-only declarations for a model without a catalogue", () => {
    const model = typedModel();
    asParagraph(model.sections[0].blocks[0]).styleId = "Heading1";
    const xml = buildStylesXml(model);
    expect(xml).toContain('<w:name w:val="heading 1"/>');
    expect(xml).not.toContain("<w:outlineLvl");
  });

  it("declares basedOn ancestors of referenced styles", () => {
    const model = typedModel();
    model.styles = [
      { id: "Grand", name: "Grand", runFormatting: { bold: true } },
      { id: "Parent", name: "Parent", basedOn: "Grand" },
    ];
    asParagraph(model.sections[0].blocks[0]).styleId = "Parent";
    const xml = buildStylesXml(model);
    expect(xml).toContain('w:styleId="Parent"');
    expect(xml).toContain('w:styleId="Grand"');
  });
});

describe("writeDocx — applied style round trip (AC-11, AC-15)", () => {
  it("a saved heading is a real style: w:pStyle plus a styles.xml definition", async () => {
    const model = applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model;
    const { bytes } = await writeDocx(model);
    expect(await partText(bytes, "word/document.xml")).toContain('<w:pStyle w:val="Heading1"/>');
    expect(await partText(bytes, "word/styles.xml")).toContain('w:styleId="Heading1"');

    const reopened = await parseDocx(bytes);
    const heading = asParagraph(reopened.sections[0].blocks[0]);
    expect(heading.styleId).toBe("Heading1");
    expect(heading.headingLevel).toBe(1);
    // The reader has always baked the docDefaults font into every run, so
    // it joins the heading formatting on re-open.
    expect(heading.runs[0].formatting).toEqual({
      bold: true,
      fontSizePt: 20,
      fontFamily: "Calibri",
    });
    const definition = reopened.styles.find((style) => style.id === "Heading1");
    expect(definition?.runFormatting).toEqual({ bold: true, fontSizePt: 20 });
    expect(definition?.headingLevel).toBe(1);
  });

  it("style plus layered direct formatting both survive a save (AC-11)", async () => {
    const model = applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model;
    const paragraph = asParagraph(model.sections[0].blocks[0]);
    paragraph.runs[0].formatting = { ...paragraph.runs[0].formatting, italic: true };
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting).toEqual({
      bold: true,
      fontSizePt: 20,
      italic: true,
      fontFamily: "Calibri",
    });
  });

  it("an unedited same-format save keeps styles.xml byte-identical", async () => {
    const first = await writeDocx(applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model);
    const reopenedModel = modelFrom(await parseDocx(first.bytes));
    const second = await writeDocx(reopenedModel);
    const before = (await unzip(first.bytes)).get("word/styles.xml")!;
    const after = (await unzip(second.bytes)).get("word/styles.xml")!;
    expect(equalBytes(before, after)).toBe(true);
  });

  it("an updated style definition reaches the preserved styles.xml (AC-12)", async () => {
    const first = await writeDocx(applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model);
    const reopenedModel = modelFrom(await parseDocx(first.bytes));
    // Make the heading red, then update the style to match.
    const heading = asParagraph(reopenedModel.sections[0].blocks[0]);
    heading.runs[0].formatting = { ...heading.runs[0].formatting, color: "#ff0000" };
    const updated = updateStyleFromSelection(reopenedModel, caretAtStart(), "Heading1").model;

    const { bytes } = await writeDocx(updated);
    const stylesXml = await partText(bytes, "word/styles.xml");
    expect(stylesXml).toContain('w:styleId="Heading1"');
    expect(stylesXml).toContain('<w:color w:val="FF0000"/>');

    const reopened = await parseDocx(bytes);
    // The updated definition captures the paragraph's formatting minus the
    // docDefault-derived font, which is an underlay, not direct formatting.
    expect(reopened.styles.find((style) => style.id === "Heading1")?.runFormatting).toEqual({
      bold: true,
      fontSizePt: 20,
      color: "#ff0000",
    });
  });
});

describe("patchStylesXml", () => {
  const original =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${WORD_NAMESPACE}">` +
    "<w:docDefaults/>" +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
    '<w:name w:val="Normal"/><w:qFormat/></w:style>' +
    "</w:styles>";

  it("returns null when the catalogue matches the part", () => {
    const model = typedModel();
    model.styles = [{ id: "Normal", name: "Normal" }];
    expect(patchStylesXml(model, original)).toBeNull();
  });

  it("returns null for a model without a catalogue", () => {
    expect(patchStylesXml(typedModel(), original)).toBeNull();
  });

  it("rewrites a changed style but keeps its unmodelled attributes and children", () => {
    const model = typedModel();
    model.styles = [{ id: "Normal", name: "Normal", runFormatting: { bold: true } }];
    const patched = patchStylesXml(model, original)!;
    expect(patched).toContain('w:default="1"');
    expect(patched).toContain("<w:qFormat/>");
    expect(patched).toContain("<w:b/>");
    // The preserved child sits before the regenerated rPr (CT_Style order).
    expect(patched.indexOf("<w:qFormat/>")).toBeLessThan(patched.indexOf("<w:b/>"));
  });

  it("appends a definition the part lacks", () => {
    const model = typedModel();
    model.styles = [
      { id: "Normal", name: "Normal" },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", headingLevel: 2 },
    ];
    const patched = patchStylesXml(model, original)!;
    expect(patched).toContain('w:styleId="Heading2"');
    expect(patched).toContain('<w:outlineLvl w:val="1"/>');
    // The untouched Normal definition stays as it was.
    expect(patched).toContain('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>');
  });
});

// SPEC-042 — character styles (w:rStyle), w:next, and table styles the
// editor cannot fully model (unmodeled, preserved verbatim).

describe("writeDocx — character style round trip (SPEC-042 AC-9, AC-13)", () => {
  it("a run's character style survives a save as w:rStyle, correctly layered over the paragraph style", async () => {
    const model = typedModel();
    model.styles = [
      { id: "Emphasis", name: "Emphasis", type: "character", runFormatting: { italic: true, color: "#ff0000" } },
    ];
    const paragraph = asParagraph(model.sections[0].blocks[0]);
    paragraph.runs[0].styleId = "Emphasis";
    paragraph.runs[0].formatting = { ...paragraph.runs[0].formatting, italic: true, color: "#ff0000" };

    const { bytes } = await writeDocx(model);
    expect(await partText(bytes, "word/document.xml")).toContain('<w:rStyle w:val="Emphasis"/>');
    expect(await partText(bytes, "word/styles.xml")).toContain('w:type="character" w:styleId="Emphasis"');

    const reopened = await parseDocx(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.styleId).toBe("Emphasis");
    expect(run.formatting).toMatchObject({ italic: true, color: "#ff0000" });
  });

  it("direct run formatting on top of a character style still wins (OOXML precedence)", async () => {
    const model = typedModel();
    model.styles = [{ id: "Emphasis", name: "Emphasis", type: "character", runFormatting: { bold: true } }];
    const paragraph = asParagraph(model.sections[0].blocks[0]);
    paragraph.runs[0].styleId = "Emphasis";
    // Direct formatting overrides the character style's own bold=true.
    paragraph.runs[0].formatting = { ...paragraph.runs[0].formatting, bold: false };
    const reopened = await parseDocx((await writeDocx(model)).bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting.bold).toBe(false);
  });
});

describe("writeDocx — w:next round trip (SPEC-042 AC-3, AC-4, AC-13)", () => {
  it("a style's 'next' style survives a save", async () => {
    const model = typedModel();
    model.styles = [
      { id: "Normal", name: "Normal" },
      { id: "Title", name: "Title", basedOn: "Normal", next: "Normal" },
    ];
    const { bytes } = await writeDocx(model);
    expect(await partText(bytes, "word/styles.xml")).toContain('<w:next w:val="Normal"/>');
    const reopened = await parseDocx(bytes);
    expect(reopened.styles.find((style) => style.id === "Title")?.next).toBe("Normal");
  });
});

describe("writeDocx — unmodeled styles round trip verbatim (SPEC-042 AC-15..18)", () => {
  // SPEC-076 AC-71 narrows which table styles are unmodeled: the six
  // w:tblStylePr regions it models (firstRow/lastRow/firstCol/lastCol/
  // band1Horz/band1Vert) now come back as NamedStyle.banding and are
  // editable. A region outside those six still keeps the whole style
  // verbatim, which is what SPEC-042 AC-18 guarantees and what this test —
  // rewritten from a firstRow example to an nwCell one — still asserts.
  it("a table-type style with an unmodelled region is preserved and written back unchanged", async () => {
    const source = await writeDocx(typedModel());
    const doc = await parseDocx(source.bytes);
    const model = modelFrom(doc);
    const raw =
      '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'w:type="table" w:styleId="GridTable1"><w:name w:val="Grid Table 1"/>' +
      '<w:tblStylePr w:type="nwCell"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>';
    model.styles = [...(model.styles ?? []), { id: "GridTable1", name: "Grid Table 1", type: "table", unmodeled: true, rawXml: raw }];
    const { bytes } = await writeDocx(model);
    const stylesXml = await partText(bytes, "word/styles.xml");
    expect(stylesXml).toContain("w:tblStylePr");
    expect(stylesXml).toContain('w:type="nwCell"');

    const reopened = await parseDocx(bytes);
    const style = reopened.styles.find((s) => s.id === "GridTable1");
    expect(style?.unmodeled).toBe(true);
    expect(style?.name).toBe("Grid Table 1");
  });

  it("a table style whose conditional formatting is all banding is modelled, not verbatim (SPEC-076 AC-71)", async () => {
    const source = await writeDocx(typedModel());
    const doc = await parseDocx(source.bytes);
    const model = modelFrom(doc);
    const raw =
      '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'w:type="table" w:styleId="BandedOnly"><w:name w:val="Banded Only"/>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>';
    model.styles = [...(model.styles ?? []), { id: "BandedOnly", name: "Banded Only", type: "table", unmodeled: true, rawXml: raw }];
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const style = reopened.styles.find((s) => s.id === "BandedOnly");
    expect(style?.unmodeled).toBeUndefined();
    expect(style?.banding?.firstRow).toEqual({ bold: true });
  });
});
