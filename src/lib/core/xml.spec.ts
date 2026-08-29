// Markup Compatibility and Extensibility (ECMA-376 Part 3) resolution — the
// `mc:AlternateContent` branch selection every OPC reader needs before it can
// see markup an Office producer wrote behind an `mc:Choice`.
import { describe, it, expect } from "vitest";
import { parseXml, selectAlternateContent } from "./xml";

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const UNDERSTOOD = new Set([WPS]);

/** The single `mc:AlternateContent` in a document that declares the prefixes
 * a real `word/document.xml` would. */
function alternateContent(inner: string, declarations = ""): Element {
  const doc = parseXml(
    `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}"${declarations}>` +
      `<mc:AlternateContent>${inner}</mc:AlternateContent></w:document>`,
  );
  return doc.documentElement.children[0];
}

function names(elements: Element[]): string[] {
  return elements.map((el) => el.localName);
}

describe("selectAlternateContent", () => {
  it("returns the children of a Choice whose Requires is understood", () => {
    const el = alternateContent(
      '<mc:Choice Requires="wps"><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback>',
    );
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["drawing"]);
  });

  it("falls back when the Choice needs a namespace the caller does not model", () => {
    const el = alternateContent(
      '<mc:Choice Requires="wpg"><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback>',
    );
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["pict"]);
  });

  it("requires EVERY prefix in a multi-prefix Requires", () => {
    const el = alternateContent(
      '<mc:Choice Requires="wps wpg"><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback>',
    );
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["pict"]);
    expect(names(selectAlternateContent(el, new Set([WPS, WPG])))).toEqual(["drawing"]);
  });

  it("takes the first understood Choice when several are offered", () => {
    const el = alternateContent(
      '<mc:Choice Requires="wpg"><w:tbl/></mc:Choice>' +
        '<mc:Choice Requires="wps"><w:drawing/></mc:Choice>' +
        "<mc:Fallback><w:pict/></mc:Fallback>",
    );
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["drawing"]);
  });

  it("resolves Requires through the document's own declarations, not by spelling", () => {
    // A producer is free to bind the shape namespace to any prefix; one that
    // binds `x` to it must still be understood, and one that spells `wps` at
    // a different namespace must not.
    const rebound = alternateContent(
      '<mc:Choice Requires="x"><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback>',
      ` xmlns:x="${WPS}"`,
    );
    expect(names(selectAlternateContent(rebound, UNDERSTOOD))).toEqual(["drawing"]);

    const doc = parseXml(
      `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="urn:something-else">` +
        '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice>' +
        "<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:document>",
    );
    expect(names(selectAlternateContent(doc.documentElement.children[0], UNDERSTOOD))).toEqual(["pict"]);
  });

  it("falls back when a Requires prefix is never declared", () => {
    const el = alternateContent(
      '<mc:Choice Requires="nope"><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback>',
    );
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["pict"]);
  });

  it("skips a Choice with no Requires at all (malformed under MCE)", () => {
    const el = alternateContent("<mc:Choice><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback>");
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["pict"]);
  });

  it("returns nothing when no Choice matches and there is no Fallback", () => {
    const el = alternateContent('<mc:Choice Requires="wpg"><w:drawing/></mc:Choice>');
    expect(selectAlternateContent(el, UNDERSTOOD)).toEqual([]);
  });

  it("returns every child of the selected branch, in order", () => {
    const el = alternateContent('<mc:Choice Requires="wps"><w:drawing/><w:t/></mc:Choice>');
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["drawing", "t"]);
  });

  it("ignores non-mc children of the wrapper", () => {
    const el = alternateContent('<w:drawing/><mc:Choice Requires="wps"><w:tbl/></mc:Choice>');
    expect(names(selectAlternateContent(el, UNDERSTOOD))).toEqual(["tbl"]);
  });
});
