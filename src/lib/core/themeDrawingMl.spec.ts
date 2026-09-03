// SPEC-077 AC-29/31: the DrawingML scheme-colour codec — slot spellings in
// both spaces on read, and the source's spelling preserved on write so an
// unedited shape or chart colour round-trips byte-identically.
import { describe, expect, it } from "vitest";
import { schemeClrRefFrom, schemeClrXmlFor, schemeSlotFromVal } from "./themeDrawingMl";
import { parseXml } from "./xml";

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
