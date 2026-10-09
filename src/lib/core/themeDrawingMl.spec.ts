// SPEC-077 AC-29/31: the DrawingML scheme-colour codec — slot spellings in
// both spaces on read, and the source's spelling preserved on write so an
// unedited shape or chart colour round-trips byte-identically.
import { describe, expect, it } from "vitest";
import { readThemePart, schemeClrRefFrom, schemeClrXmlFor, schemeSlotFromVal } from "./themeDrawingMl";
import { DEFAULT_THEME, THEME_COLOR_SLOTS } from "./theme";
import { parseXml } from "./xml";
import { readDocxTheme } from "../docx/theme";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function schemeClrElement(inner: string): Element {
  const doc = parseXml(`<root xmlns:a="${A_NS}">${inner}</root>`);
  return doc.documentElement.firstElementChild!;
}

describe("SPEC-077 AC-29: schemeClr parsing in both slot spaces", () => {
  it("theme-space spellings parse with no spelling carried", () => {
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="dk1"/>'))).toEqual({ theme: "dk1" });
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="accent3"/>'))).toEqual({ theme: "accent3" });
  });

  it("slide-space spellings normalise to the theme-space slot and carry the source spelling", () => {
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="tx1"/>'))).toEqual({ theme: "dk1", spelling: "tx1" });
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="bg1"/>'))).toEqual({ theme: "lt1", spelling: "bg1" });
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="tx2"/>'))).toEqual({ theme: "dk2", spelling: "tx2" });
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="bg2"/>'))).toEqual({ theme: "lt2", spelling: "bg2" });
    expect(schemeSlotFromVal("tx1")).toBe("dk1");
  });

  it("a spelling rides beside the modelled transforms", () => {
    const ref = schemeClrRefFrom(
      schemeClrElement('<a:schemeClr val="bg2"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'),
    );
    expect(ref).toEqual({ theme: "lt2", lumMod: 0.6, lumOff: 0.4, spelling: "bg2" });
  });

  it("phClr stays unmodelled", () => {
    expect(schemeClrRefFrom(schemeClrElement('<a:schemeClr val="phClr"/>'))).toBeUndefined();
  });
});

describe("SPEC-077 AC-31: the source's slot spelling survives to the written XML", () => {
  it("a ref read from a slide-space spelling writes that spelling back", () => {
    expect(schemeClrXmlFor({ theme: "dk1", spelling: "tx1" })).toBe('<a:schemeClr val="tx1"/>');
    expect(schemeClrXmlFor({ theme: "lt1", spelling: "bg1" })).toBe('<a:schemeClr val="bg1"/>');
    expect(schemeClrXmlFor({ theme: "dk2", lumMod: 0.75, spelling: "tx2" })).toBe(
      '<a:schemeClr val="tx2"><a:lumMod val="75000"/></a:schemeClr>',
    );
  });

  it("a product-created ref (no spelling) writes the theme-space name", () => {
    expect(schemeClrXmlFor({ theme: "dk1" })).toBe('<a:schemeClr val="dk1"/>');
    expect(schemeClrXmlFor({ theme: "accent1", lumMod: 0.2, lumOff: 0.8 })).toBe(
      '<a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr>',
    );
  });

  it("a spelling that no longer names the ref's slot falls back to the slot", () => {
    expect(schemeClrXmlFor({ theme: "accent1", spelling: "tx1" })).toBe('<a:schemeClr val="accent1"/>');
  });

  it("read → write → read is a fixed point that keeps the source byte", () => {
    const source = '<a:schemeClr val="tx1"><a:lumMod val="50000"/><a:shade val="25000"/></a:schemeClr>';
    const ref = schemeClrRefFrom(schemeClrElement(source))!;
    expect(schemeClrXmlFor(ref)).toBe(source);
    expect(schemeClrRefFrom(schemeClrElement(schemeClrXmlFor(ref)))).toEqual(ref);
  });
});

// SPEC-117 "Theme": the DrawingML theme-part reader, moved here from
// docx/theme.ts so the spreadsheet product (which may import only core) reads
// an xlsx workbook's theme through the same code as the word processor —
// `a:clrScheme` (srgbClr, or sysClr@lastClr with the windowText/window
// fallback), the major/minor `a:latin` typefaces and the root's name;
// undefined without a clrScheme. readDocxTheme now delegates to it.
describe("SPEC-117 Theme: readThemePart", () => {
  const OFFICE_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A_NS}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

  /** A theme part with the given clrScheme children and fontScheme. */
  function themeXml(slots: string, fonts = "", rootAttrs = ' name="T"'): string {
    return `<a:theme xmlns:a="${A_NS}"${rootAttrs}><a:themeElements><a:clrScheme name="S">${slots}</a:clrScheme>${fonts}</a:themeElements></a:theme>`;
  }

  it("reads an Excel-written Office theme to DEFAULT_THEME's scheme under the part's name", () => {
    expect(readThemePart(OFFICE_THEME)).toEqual({ ...DEFAULT_THEME, name: "Office Theme" });
  });

  it("reads srgbClr values, lowercased", () => {
    const theme = readThemePart(
      themeXml('<a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2>'),
    );
    expect(theme?.colors.accent1).toBe("#156082");
    expect(theme?.colors.accent2).toBe("#e97132");
  });

  it("reads sysClr@lastClr, else windowText → black and window → white", () => {
    const withLast = readThemePart(
      themeXml('<a:dk1><a:sysClr val="windowText" lastClr="1F1F1F"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FAFAFA"/></a:lt1>'),
    );
    expect(withLast?.colors.dk1).toBe("#1f1f1f");
    expect(withLast?.colors.lt1).toBe("#fafafa");
    const fallback = readThemePart(
      themeXml('<a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>'),
    );
    expect(fallback?.colors.dk1).toBe("#000000");
    expect(fallback?.colors.lt1).toBe("#ffffff");
    // A malformed lastClr is ignored in favour of the fallback.
    const malformed = readThemePart(themeXml('<a:lt1><a:sysClr val="window" lastClr="nothex"/></a:lt1>'));
    expect(malformed?.colors.lt1).toBe("#ffffff");
  });

  it("cushions a truncated scheme: a missing or unusable slot reads as black", () => {
    const theme = readThemePart(
      themeXml('<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="12345"/></a:accent2>'),
    );
    expect(theme?.colors.accent1).toBe("#4472c4");
    expect(theme?.colors.accent2).toBe("#000000");
    expect(theme?.colors.hlink).toBe("#000000");
    expect(Object.keys(theme!.colors).toSorted()).toEqual(THEME_COLOR_SLOTS.toSorted());
  });

  it("reads the major and minor latin typefaces, defaulting to Calibri Light / Calibri", () => {
    const slots = '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>';
    const fonts =
      '<a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Aptos Narrow"/></a:minorFont></a:fontScheme>';
    expect(readThemePart(themeXml(slots, fonts))?.fonts).toEqual({ major: "Aptos Display", minor: "Aptos Narrow" });
    expect(readThemePart(themeXml(slots))?.fonts).toEqual({ major: "Calibri Light", minor: "Calibri" });
    const emptyMajor =
      '<a:fontScheme name="F"><a:majorFont><a:latin typeface=""/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Georgia"/></a:minorFont></a:fontScheme>';
    expect(readThemePart(themeXml(slots, emptyMajor))?.fonts).toEqual({ major: "Calibri Light", minor: "Georgia" });
  });

  it("takes the root's name, else Custom", () => {
    const slots = '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>';
    expect(readThemePart(themeXml(slots, "", ' name="Facet"'))?.name).toBe("Facet");
    expect(readThemePart(themeXml(slots, "", ""))?.name).toBe("Custom");
  });

  it("looks elements up by namespace, whatever the prefix", () => {
    const rebound = OFFICE_THEME.replace(/<(\/?)a:/g, "<$1d:").replace(`xmlns:a="${A_NS}"`, `xmlns:d="${A_NS}"`);
    expect(rebound).not.toContain("<a:");
    expect(readThemePart(rebound)).toEqual({ ...DEFAULT_THEME, name: "Office Theme" });
  });

  it("is undefined without a usable clrScheme, for non-theme XML and for malformed XML", () => {
    expect(readThemePart(OFFICE_THEME.replace(/<a:clrScheme[\s\S]*<\/a:clrScheme>/, ""))).toBeUndefined();
    expect(readThemePart(themeXml(""))).toBeUndefined(); // a clrScheme with no slot
    expect(readThemePart(`<a:theme xmlns:a="${A_NS}" name="T"><a:themeElements/></a:theme>`)).toBeUndefined();
    expect(readThemePart(`<a:themeOverride xmlns:a="${A_NS}"><a:clrScheme name="S"/></a:themeOverride>`)).toBeUndefined();
    expect(readThemePart('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>')).toBeUndefined();
    // The right shape in no namespace is not DrawingML.
    const unprefixed = OFFICE_THEME.replace(/(<\/?)a:/g, "$1");
    expect(unprefixed).toContain("<clrScheme");
    expect(readThemePart(unprefixed)).toBeUndefined();
    expect(readThemePart('<a:theme xmlns:a="x"><a:themeElements>')).toBeUndefined();
    expect(readThemePart("not xml at all")).toBeUndefined();
    expect(readThemePart("")).toBeUndefined();
  });

  it("readDocxTheme(xml, partName) is { ...readThemePart(xml), partName }", () => {
    const fonts =
      '<a:fontScheme name="F"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme>';
    for (const xml of [
      OFFICE_THEME,
      themeXml('<a:dk2><a:srgbClr val="123456"/></a:dk2><a:lt1><a:sysClr val="window"/></a:lt1>', fonts),
    ]) {
      const part = readThemePart(xml);
      expect(part).toBeDefined();
      expect(readDocxTheme(xml, "word/theme/theme1.xml")).toEqual({ ...part, partName: "word/theme/theme1.xml" });
    }
    expect(readDocxTheme(themeXml(""), "word/theme/theme1.xml")).toBeUndefined();
    expect(readDocxTheme("<broken", "word/theme/theme1.xml")).toBeUndefined();
  });
});
