// SPEC-073 AC-36 one-source: an edit that gives the model a fill/outline/
// shadow value must clear the preserved (unparsed) styling occupying that
// slot, in whichever format flavour the block was read from — otherwise a
// same-format save resurrects the old styling over the user's visible edit.
import { describe, it, expect } from "vitest";
import { clearPreservedStyleSlots } from "./preservedStyle";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

// --- the .docx flavour: a run of a:* property elements -----------------------

const DOCX_FILL = `<a:solidFill xmlns:a="${A_NS}"><a:schemeClr val="accent1"/></a:solidFill>`;
const DOCX_LN = `<a:ln xmlns:a="${A_NS}" w="25400"><a:gradFill><a:gsLst/></a:gradFill></a:ln>`;
const DOCX_GLOW = `<a:glow xmlns:a="${A_NS}" rad="63500"><a:srgbClr val="00FF00"/></a:glow>`;
const DOCX_SHDW = `<a:outerShdw xmlns:a="${A_NS}" blurRad="9"><a:schemeClr val="accent2"/></a:outerShdw>`;

describe("clearPreservedStyleSlots — .docx flavour", () => {
  const all = DOCX_FILL + DOCX_LN + DOCX_GLOW + DOCX_SHDW;

  it("clearing the fill slot removes only the preserved fill element", () => {
    const cleared = clearPreservedStyleSlots(all, { fill: true })!;
    expect(cleared).not.toContain("solidFill");
    expect(cleared).toContain("a:ln");
    expect(cleared).toContain("a:glow");
    expect(cleared).toContain("a:outerShdw");
  });

  it("clearing the outline slot removes only the preserved a:ln", () => {
    const cleared = clearPreservedStyleSlots(all, { outline: true })!;
    expect(cleared).not.toContain("<a:ln");
    expect(cleared).toContain("solidFill");
    expect(cleared).toContain("a:glow");
  });

  it("clearing the shadow slot removes a:outerShdw but keeps other effects (a glow is not the shadow)", () => {
    const cleared = clearPreservedStyleSlots(all, { shadow: true })!;
    expect(cleared).not.toContain("outerShdw");
    expect(cleared).toContain("a:glow");
    expect(cleared).toContain("solidFill");
  });

  it("returns undefined when nothing survives", () => {
    expect(clearPreservedStyleSlots(DOCX_FILL + DOCX_LN, { fill: true, outline: true })).toBeUndefined();
  });

  it("no named slots returns the string unchanged", () => {
    expect(clearPreservedStyleSlots(all, {})).toBe(all);
  });

  it("undefined passes through", () => {
    expect(clearPreservedStyleSlots(undefined, { fill: true })).toBeUndefined();
  });
});

// --- the .odt flavour: one synthetic style:graphic-properties ----------------

const ODF_PROPS =
  `<style:graphic-properties xmlns:style="${STYLE_NS}" xmlns:draw="${DRAW_NS}" xmlns:svg="${SVG_NS}"` +
  ' draw:fill="hatch" draw:fill-hatch-name="Triple" draw:fill-color="#ffffff"' +
  ' draw:marker-start="FancyArrow" draw:marker-start-width="0.4cm"' +
  ' draw:shadow="hidden" draw:shadow-color="#333333"' +
  ' draw:auto-grow-height="true">' +
  `<draw:hatch draw:name="Triple" draw:style="triple" draw:color="#000000"/>` +
  "</style:graphic-properties>";

describe("clearPreservedStyleSlots — .odt flavour", () => {
  it("clearing the fill slot removes the fill attributes AND the definition riding for them", () => {
    const cleared = clearPreservedStyleSlots(ODF_PROPS, { fill: true })!;
    expect(cleared).not.toContain("draw:fill=");
    expect(cleared).not.toContain("fill-hatch-name");
    expect(cleared).not.toContain("fill-color");
    expect(cleared).not.toContain("<draw:hatch");
    // Everything outside the fill slot stays preserved verbatim.
    expect(cleared).toContain("marker-start");
    expect(cleared).toContain("draw:shadow=");
    expect(cleared).toContain("auto-grow-height");
  });

  it("clearing the outline slot removes stroke and marker attributes only", () => {
    const cleared = clearPreservedStyleSlots(ODF_PROPS, { outline: true })!;
    expect(cleared).not.toContain("marker-start");
    expect(cleared).toContain("fill-hatch-name");
    expect(cleared).toContain("draw:shadow=");
  });

  it("clearing the shadow slot removes the draw:shadow-* attributes only", () => {
    const cleared = clearPreservedStyleSlots(ODF_PROPS, { shadow: true })!;
    expect(cleared).not.toContain("draw:shadow");
    expect(cleared).toContain("fill-hatch-name");
    expect(cleared).toContain("marker-start");
  });

  it("returns undefined when every preserved attribute and child is in a cleared slot", () => {
    const onlyFill =
      `<style:graphic-properties xmlns:style="${STYLE_NS}" xmlns:draw="${DRAW_NS}"` +
      ' draw:fill="gradient" draw:fill-gradient-name="g1">' +
      '<draw:gradient draw:name="g1" draw:style="axial"/>' +
      "</style:graphic-properties>";
    expect(clearPreservedStyleSlots(onlyFill, { fill: true })).toBeUndefined();
  });
});
