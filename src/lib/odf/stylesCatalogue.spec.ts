import { describe, expect, it } from "vitest";
import { applyParagraphStyle, caretSelection, updateStyleFromSelection } from "../commands";
import { asParagraph, createDocument, type DocumentModel } from "../documentModel";
import { equalBytes } from "../equality";
import { parseXml } from "../docx/xml";
import { unzip } from "../docx/zip";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { catalogToNamedStyles, parseOdfStyleCatalog } from "./styles";
import { buildStylesXml, patchOdtStylesXml, writeOdt } from "./writer";
import { FO_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TEXT_NAMESPACE } from "./xml";

// SPEC-019 — the ODF side of the style catalogue: named paragraph styles
// read from styles.xml with style:parent-style-name inheritance (AC-7),
// declared on from-scratch saves and patched into preserved styles.xml only
// when the catalogue diverges from it.

const decoder = new TextDecoder();

const XMLNS =
  ` xmlns:office="${OFFICE_NAMESPACE}" xmlns:style="${STYLE_NAMESPACE}"` +
  ` xmlns:fo="${FO_NAMESPACE}" xmlns:text="${TEXT_NAMESPACE}"`;

const EMPTY_CONTENT = `<office:document-content${XMLNS}><office:body><office:text/></office:body></office:document-content>`;

const SAMPLE_STYLES =
  `<office:document-styles${XMLNS}><office:styles>` +
  '<style:default-style style:family="paragraph">' +
  '<style:text-properties fo:font-size="12pt"/></style:default-style>' +
  '<style:style style:name="Standard" style:family="paragraph"/>' +
  '<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph"' +
  ' style:parent-style-name="Standard" style:default-outline-level="1">' +
  '<style:paragraph-properties fo:margin-top="0.4cm"/>' +
  '<style:text-properties fo:font-size="130%" fo:font-weight="bold"/>' +
  "</style:style>" +
  '<style:style style:name="My_20_Quote" style:family="paragraph" style:parent-style-name="Standard">' +
  '<style:text-properties fo:font-style="italic"/>' +
  "</style:style>" +
  "</office:styles></office:document-styles>";

function caretAtStart() {
  return caretSelection({ section: 0, block: 0, offset: 0 });
}

function typedModel(text = "Chapter one"): DocumentModel {
  const model = createDocument("odf-styles");
  model.sections[0].blocks = [{ type: "paragraph", runs: [{ text, formatting: {} }] }];
  return model;
}

function modelFrom(parsed: ParsedOdtDocument): DocumentModel {
  const model: DocumentModel = {
    id: "reopened",
    title: "Reopened",
    format: "odt",
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

describe("catalogToNamedStyles (AC-7)", () => {
  const catalog = parseOdfStyleCatalog(parseXml(EMPTY_CONTENT), parseXml(SAMPLE_STYLES));
  const styles = catalogToNamedStyles(catalog);

  it("converts named paragraph styles with parent links and outline levels", () => {
    const heading = styles.find((style) => style.id === "Heading_20_1");
    expect(heading).toEqual({
      id: "Heading_20_1",
      name: "Heading 1",
      basedOn: "Standard",
      paragraphFormatting: { spaceBeforePt: expect.closeTo(11.34, 1) },
      runFormatting: { bold: true },
      fontSizePercent: 130,
      headingLevel: 1,
    });
  });

  it("decodes an encoded style:name when no display-name is given", () => {
    expect(styles.find((style) => style.id === "My_20_Quote")?.name).toBe("My Quote");
  });

  it("keeps styles.xml order and elides empty formatting", () => {
    expect(styles.map((style) => style.id)).toEqual(["Standard", "Heading_20_1", "My_20_Quote"]);
    const standard = styles.find((style) => style.id === "Standard");
    expect(standard?.paragraphFormatting).toBeUndefined();
    expect(standard?.runFormatting).toBeUndefined();
  });
});

describe("buildStylesXml — named definitions (AC-16, AC-11)", () => {
  it("declares the catalogue's named styles in office:styles", () => {
    const model = typedModel();
    model.styles = [
      { id: "Normal", name: "Normal" },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        runFormatting: { bold: true, fontSizePt: 20 },
        headingLevel: 1,
      },
    ];
    const xml = buildStylesXml(model);
    expect(xml).toContain('style:name="Heading1"');
    expect(xml).toContain('style:display-name="Heading 1"');
    expect(xml).toContain('style:parent-style-name="Normal"');
    expect(xml).toContain('style:default-outline-level="1"');
    expect(xml).toContain('fo:font-weight="bold"');
  });

  it("emits a proportional font size a converted style carries", () => {
    const model = typedModel();
    model.styles = [{ id: "Big", name: "Big", fontSizePercent: 130 }];
    expect(buildStylesXml(model)).toContain('fo:font-size="130%"');
  });
});

describe("writeOdt — applied style round trip (AC-11, AC-16)", () => {
  it("an applied heading saves as text:h and a named style definition", async () => {
    const model = applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model;
    const { bytes } = await writeOdt(model);
    const parts = await unzip(bytes);
    const content = decoder.decode(parts.get("content.xml")!);
    expect(content).toContain('text:outline-level="1"');
    expect(content).toContain("<text:h");
    expect(decoder.decode(parts.get("styles.xml")!)).toContain('style:name="Heading1"');

    const reopened = await parseOdt(bytes);
    const heading = asParagraph(reopened.sections[0].blocks[0]);
    expect(heading.headingLevel).toBe(1);
    expect(heading.runs[0].formatting.bold).toBe(true);
    expect(heading.runs[0].formatting.fontSizePt).toBe(20);
    const definition = reopened.styles.find((style) => style.id === "Heading1");
    expect(definition?.name).toBe("Heading 1");
    expect(definition?.headingLevel).toBe(1);
  });

  it("an unedited same-format save keeps styles.xml byte-identical", async () => {
    const first = await writeOdt(applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model);
    const second = await writeOdt(modelFrom(await parseOdt(first.bytes)));
    const before = (await unzip(first.bytes)).get("styles.xml")!;
    const after = (await unzip(second.bytes)).get("styles.xml")!;
    expect(equalBytes(before, after)).toBe(true);
  });

  it("an updated style definition reaches the preserved styles.xml (AC-12)", async () => {
    const first = await writeOdt(applyParagraphStyle(typedModel(), caretAtStart(), "Heading1").model);
    const reopenedModel = modelFrom(await parseOdt(first.bytes));
    const heading = asParagraph(reopenedModel.sections[0].blocks[0]);
    heading.runs[0].formatting = { ...heading.runs[0].formatting, color: "#ff0000" };
    const updated = updateStyleFromSelection(reopenedModel, caretAtStart(), "Heading1").model;

    const { bytes } = await writeOdt(updated);
    const stylesXml = decoder.decode((await unzip(bytes)).get("styles.xml")!);
    expect(stylesXml).toContain('fo:color="#ff0000"');

    const reopened = await parseOdt(bytes);
    expect(
      reopened.styles.find((style) => style.id === "Heading1")?.runFormatting?.color,
    ).toBe("#ff0000");
  });
});

describe("patchOdtStylesXml", () => {
  it("returns null when the catalogue matches the part, and for no catalogue", () => {
    const catalog = parseOdfStyleCatalog(parseXml(EMPTY_CONTENT), parseXml(SAMPLE_STYLES));
    const model = typedModel();
    model.styles = catalogToNamedStyles(catalog);
    expect(patchOdtStylesXml(model, SAMPLE_STYLES, EMPTY_CONTENT)).toBeNull();
    expect(patchOdtStylesXml(typedModel(), SAMPLE_STYLES, EMPTY_CONTENT)).toBeNull();
  });

  it("rewrites a changed named style and appends a missing one", () => {
    const catalog = parseOdfStyleCatalog(parseXml(EMPTY_CONTENT), parseXml(SAMPLE_STYLES));
    const model = typedModel();
    const styles = catalogToNamedStyles(catalog);
    const quote = styles.find((style) => style.id === "My_20_Quote")!;
    quote.runFormatting = { italic: true, color: "#ff0000" };
    styles.push({ id: "Code", name: "Code", runFormatting: { fontFamily: "Courier New" } });
    model.styles = styles;

    const patched = patchOdtStylesXml(model, SAMPLE_STYLES, EMPTY_CONTENT)!;
    expect(patched).toContain('fo:color="#ff0000"');
    expect(patched).toContain('style:name="Code"');
    // Untouched definitions keep their original serialization.
    expect(patched).toContain('style:display-name="Heading 1"');
    expect(patched).toContain('fo:font-size="130%"');
  });
});

// SPEC-042 — character (family="text") styles, style:next-style-name, and
// styles of a family this app does not model (preserved verbatim).

describe("writeOdt — character style round trip (SPEC-042 AC-9, AC-14)", () => {
  it("a run's character style survives a save as a family=text style, correctly layered over the paragraph", async () => {
    const model = typedModel();
    model.styles = [
      { id: "Emphasis", name: "Emphasis", type: "character", runFormatting: { italic: true, color: "#ff0000" } },
    ];
    const paragraph = asParagraph(model.sections[0].blocks[0]);
    paragraph.runs[0].styleId = "Emphasis";
    paragraph.runs[0].formatting = { ...paragraph.runs[0].formatting, italic: true, color: "#ff0000" };

    const { bytes } = await writeOdt(model);
    const parts = await unzip(bytes);
    expect(decoder.decode(parts.get("content.xml")!)).toContain('text:style-name="Emphasis"');
    expect(decoder.decode(parts.get("styles.xml")!)).toContain('style:name="Emphasis" style:family="text"');

    const reopened = await parseOdt(bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting).toMatchObject({
      italic: true,
      color: "#ff0000",
    });
  });
});

describe("writeOdt — style:next-style-name round trip (SPEC-042 AC-3, AC-4, AC-14)", () => {
  it("a style's 'next' style survives a save", async () => {
    const model = typedModel();
    model.styles = [
      { id: "Standard", name: "Standard" },
      { id: "Title", name: "Title", basedOn: "Standard", next: "Standard" },
    ];
    const { bytes } = await writeOdt(model);
    const stylesXml = decoder.decode((await unzip(bytes)).get("styles.xml")!);
    expect(stylesXml).toContain('style:next-style-name="Standard"');
    const reopened = await parseOdt(bytes);
    expect(reopened.styles.find((style) => style.id === "Title")?.next).toBe("Standard");
  });
});

describe("writeOdt — unmodeled styles round trip verbatim (SPEC-042 AC-15..18)", () => {
  it("a non-paragraph/text-family style read from a source document is preserved and written back unchanged", async () => {
    const model = typedModel();
    const raw = '<style:style style:name="TableGrid" style:family="table">' +
      '<style:table-properties style:width="17cm"/></style:style>';
    model.styles = [{ id: "TableGrid", name: "TableGrid", type: "table", unmodeled: true, rawXml: raw }];
    const { bytes } = await writeOdt(model);
    const stylesXml = decoder.decode((await unzip(bytes)).get("styles.xml")!);
    expect(stylesXml).toContain('style:width="17cm"');

    const reopened = await parseOdt(bytes);
    const style = reopened.styles.find((s) => s.id === "TableGrid");
    expect(style?.unmodeled).toBe(true);
  });
});
