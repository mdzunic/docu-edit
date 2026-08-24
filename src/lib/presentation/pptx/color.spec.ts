// Colour resolution + transform maths (SPEC-079). The transforms are pinned
// against known PowerPoint values: accent1 (#4472c4) with lumMod 60% + lumOff
// 40% is PowerPoint's "Blue, Accent 1, Lighter 40%" = #8faadc.
import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import { parseClrMap, parseClrScheme, resolveColorElement, type ColorContext } from "./color";

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

function colorEl(inner: string): Element {
  const doc = parseXml(`<a:wrap ${A}>${inner}</a:wrap>`);
  return doc.documentElement.firstElementChild as Element;
}

const CTX: ColorContext = {
  scheme: { dk1: "#000000", lt1: "#ffffff", accent1: "#4472c4", accent2: "#ed7d31" },
  map: { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" },
};

describe("resolveColorElement", () => {
  it("resolves a:srgbClr to lowercase hex", () => {
    expect(resolveColorElement(colorEl('<a:srgbClr val="4472C4"/>'), CTX)).toBe("#4472c4");
  });

  it("resolves a:sysClr via its lastClr", () => {
    expect(resolveColorElement(colorEl('<a:sysClr val="windowText" lastClr="000000"/>'), CTX)).toBe("#000000");
  });

  it("resolves a:schemeClr through the clrMap into the theme scheme", () => {
    // tx1 → clrMap → dk1 → scheme → #000000.
    expect(resolveColorElement(colorEl('<a:schemeClr val="tx1"/>'), CTX)).toBe("#000000");
    // accent1 maps to itself.
    expect(resolveColorElement(colorEl('<a:schemeClr val="accent1"/>'), CTX)).toBe("#4472c4");
  });

  it("applies lumMod + lumOff per the Reference formula (PowerPoint Lighter 40%)", () => {
    const el = colorEl('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>');
    expect(resolveColorElement(el, CTX)).toBe("#8faadc");
  });

  it("applies shade (c·val)", () => {
    expect(resolveColorElement(colorEl('<a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr>'), CTX)).toBe(
      "#808080",
    );
  });

  it("applies tint (c·val + (1−val))", () => {
    expect(resolveColorElement(colorEl('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>'), CTX)).toBe(
      "#808080",
    );
  });

  it("returns undefined for an unresolvable scheme entry", () => {
    expect(resolveColorElement(colorEl('<a:schemeClr val="accent6"/>'), CTX)).toBeUndefined();
  });
});

describe("parseClrScheme / parseClrMap", () => {
  it("reads the theme's clrScheme entries (srgb and sysClr)", () => {
    const theme = parseXml(
      `<a:theme ${A}><a:themeElements><a:clrScheme name="X">` +
        `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
        `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
        `<a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
        `</a:clrScheme></a:themeElements></a:theme>`,
    );
    const scheme = parseClrScheme(theme.documentElement);
    expect(scheme.dk1).toBe("#000000");
    expect(scheme.lt1).toBe("#ffffff");
    expect(scheme.accent1).toBe("#4472c4");
  });

  it("reads the master's clrMap attributes", () => {
    const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
    const master = parseXml(`<p:sldMaster ${P}><p:clrMap bg1="lt1" tx1="dk1" accent1="accent1"/></p:sldMaster>`);
    const map = parseClrMap(master.documentElement);
    expect(map.tx1).toBe("dk1");
    expect(map.bg1).toBe("lt1");
  });
});
