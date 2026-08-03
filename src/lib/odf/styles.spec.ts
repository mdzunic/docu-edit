import { describe, it, expect } from "vitest";
import { parseXml } from "../docx/xml";
import { parseOdfStyleCatalog, resolveParagraphStyle, resolveTextStyle } from "./styles";

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"';

function contentDoc(automaticStyles: string): Document {
  return parseXml(
    `<office:document-content ${NS}><office:automatic-styles>${automaticStyles}</office:automatic-styles></office:document-content>`,
  );
}

function stylesDoc(styles: string): Document {
  return parseXml(`<office:document-styles ${NS}><office:styles>${styles}</office:styles></office:document-styles>`);
}

describe("resolveParagraphStyle", () => {
  it("survives a parent-style-name cycle (AC-13)", () => {
    const catalog = parseOdfStyleCatalog(
      contentDoc(""),
      stylesDoc(`
        <style:style style:name="A" style:family="paragraph" style:parent-style-name="B">
          <style:text-properties fo:font-weight="bold"/>
        </style:style>
        <style:style style:name="B" style:family="paragraph" style:parent-style-name="A">
          <style:text-properties fo:font-style="italic"/>
        </style:style>`),
    );
    const resolved = resolveParagraphStyle(catalog, "A");
    expect(resolved.runFormatting.bold).toBe(true);
    expect(resolved.runFormatting.italic).toBe(true);
  });

  it("lets an automatic style shadow a named style of the same name (AC-11, AC-12)", () => {
    const catalog = parseOdfStyleCatalog(
      contentDoc(`
        <style:style style:name="Quote" style:family="paragraph">
          <style:paragraph-properties fo:text-align="center"/>
        </style:style>`),
      stylesDoc(`
        <style:style style:name="Quote" style:family="paragraph">
          <style:paragraph-properties fo:text-align="justify"/>
        </style:style>`),
    );
    expect(resolveParagraphStyle(catalog, "Quote").paragraphFormatting.alignment).toBe("center");
  });

  it("ignores a percentage font size with no inherited absolute size to resolve against (AC-19)", () => {
    const catalog = parseOdfStyleCatalog(
      contentDoc(""),
      stylesDoc(`
        <style:style style:name="Big" style:family="paragraph">
          <style:text-properties fo:font-size="150%"/>
        </style:style>`),
    );
    expect(resolveParagraphStyle(catalog, "Big").runFormatting.fontSizePt).toBeUndefined();
  });

  it("returns the defaults for an unknown style name", () => {
    const catalog = parseOdfStyleCatalog(contentDoc(""), undefined);
    const resolved = resolveParagraphStyle(catalog, "Missing");
    expect(resolved.runFormatting).toEqual({});
    expect(resolved.paragraphFormatting).toEqual({});
    expect(resolved.headingLevel).toBeUndefined();
  });

  it("cleans a quoted fo:font-family fallback list (AC-18)", () => {
    const catalog = parseOdfStyleCatalog(
      contentDoc(""),
      stylesDoc(`
        <style:style style:name="Mono" style:family="paragraph">
          <style:text-properties fo:font-family="'Courier New', monospace"/>
        </style:style>`),
    );
    expect(resolveParagraphStyle(catalog, "Mono").runFormatting.fontFamily).toBe("Courier New");
  });
});

describe("resolveTextStyle", () => {
  it("resolves a span's percentage font size against the inherited size (AC-19)", () => {
    const catalog = parseOdfStyleCatalog(
      contentDoc(`
        <style:style style:name="T1" style:family="text">
          <style:text-properties fo:font-size="50%"/>
        </style:style>`),
      undefined,
    );
    const resolved = resolveTextStyle(catalog, "T1", { fontSizePt: 24 });
    expect(resolved.fontSizePt).toBe(12);
  });

  it("keeps inherited formatting the style does not override (AC-7)", () => {
    const catalog = parseOdfStyleCatalog(
      contentDoc(`
        <style:style style:name="T1" style:family="text">
          <style:text-properties fo:font-weight="bold"/>
        </style:style>`),
      undefined,
    );
    const resolved = resolveTextStyle(catalog, "T1", { italic: true, fontFamily: "Arial" });
    expect(resolved).toEqual({ italic: true, fontFamily: "Arial", bold: true });
  });

  it("returns the inherited formatting unchanged for an unknown style", () => {
    const catalog = parseOdfStyleCatalog(contentDoc(""), undefined);
    const inherited = { bold: true };
    expect(resolveTextStyle(catalog, "Nope", inherited)).toBe(inherited);
  });
});
