import { describe, it, expect } from "vitest";
import { parseXml } from "../core/xml";
import {
  paragraphAutoStyleSignals,
  parseOdfStyleCatalog,
  resolveParagraphStyle,
  resolveTextStyle,
} from "./styles";

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

describe("parseOdfStyleCatalog — master pages (SPEC-028 AC-33)", () => {
  const FULL_STYLES = parseXml(
    `<office:document-styles ${NS}>` +
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
      "</office:master-styles></office:document-styles>",
  );

  it("catalogues every master page's geometry by name", () => {
    const catalog = parseOdfStyleCatalog(contentDoc(""), FULL_STYLES);
    expect(catalog.masterPageLayouts.get("A5Master")).toEqual({
      pageSize: { widthCm: 14.8, heightCm: 21 },
      margins: { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 },
    });
    // The first master page also stays the document-level default, exactly
    // as before SPEC-028.
    expect(catalog.masterPageLayouts.get("Standard")?.pageSize).toEqual(catalog.pageSize);
    expect(catalog.pageSize).toEqual({ widthCm: 21, heightCm: 29.7 });
    expect(catalog.margins).toEqual({ topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 });
  });

  // SPEC-074 AC-31. A page layout inherits from nothing, so an explicit
  // `none` says exactly what an absent attribute says. (A paragraph is the
  // other way round: there, `fo:border="none"` is how it cancels the border
  // its style would give it — see properties.parseParagraphProperties.)
  function pageLayout(attrs: string) {
    return parseOdfStyleCatalog(
      contentDoc(""),
      parseXml(
        `<office:document-styles ${NS}><office:automatic-styles>` +
          `<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm"` +
          ` fo:page-height="29.7cm" ${attrs}/></style:page-layout>` +
          "</office:automatic-styles><office:master-styles>" +
          '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
          "</office:master-styles></office:document-styles>",
      ),
    ).masterPageLayouts.get("Standard");
  }

  it("reads a page layout with no border as having none", () => {
    expect(pageLayout("")?.pageBorders).toBeUndefined();
  });

  it("reads an explicit fo:border='none' page layout as having none either", () => {
    expect(pageLayout('fo:border="none"')?.pageBorders).toBeUndefined();
    expect(pageLayout('fo:border-top="none" fo:border-bottom="none"')?.pageBorders).toBeUndefined();
  });

  it("keeps the edges a page layout really does draw", () => {
    const borders = pageLayout('fo:border-top="2pt double #c00000" fo:border-bottom="none"')?.pageBorders;
    expect(borders?.top).toEqual({ style: "double", widthPt: 2, color: "#c00000" });
    expect(borders?.bottom).toBeUndefined();
    expect(borders?.measureFrom).toBe("page");
  });
});

describe("paragraphAutoStyleSignals (SPEC-028)", () => {
  const catalog = parseOdfStyleCatalog(
    contentDoc(`
      <style:style style:name="P1" style:family="paragraph" style:master-page-name="A5Master">
        <style:paragraph-properties fo:break-before="page"/>
      </style:style>
      <style:style style:name="P2" style:family="paragraph">
        <style:paragraph-properties fo:break-before="column"/>
      </style:style>`),
    stylesDoc(`
      <style:style style:name="PageStart" style:family="paragraph">
        <style:paragraph-properties fo:break-before="page"/>
      </style:style>
      <style:style style:name="ColStart" style:family="paragraph">
        <style:paragraph-properties fo:break-before="column"/>
      </style:style>`),
  );

  it("reports the break/master signals of an automatic style", () => {
    expect(paragraphAutoStyleSignals(catalog, "P1")).toEqual({
      pageBreakBefore: true,
      columnBreakBefore: false,
      masterPageName: "A5Master",
    });
  });

  it("reports fo:break-before='column' on an automatic style (AC-10/11)", () => {
    expect(paragraphAutoStyleSignals(catalog, "P2")).toEqual({
      pageBreakBefore: false,
      columnBreakBefore: true,
      masterPageName: undefined,
    });
  });

  it("never reports signals from a NAMED style — its break-before stays formatting", () => {
    expect(paragraphAutoStyleSignals(catalog, "PageStart").pageBreakBefore).toBe(false);
    expect(paragraphAutoStyleSignals(catalog, "ColStart").columnBreakBefore).toBe(false);
    expect(paragraphAutoStyleSignals(catalog, "PageStart").masterPageName).toBeUndefined();
    expect(paragraphAutoStyleSignals(catalog, undefined).pageBreakBefore).toBe(false);
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
