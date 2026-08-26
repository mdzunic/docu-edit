import { describe, expect, it } from "vitest";
import type { Slide, TextFrame } from "../model";
import { serializePage, spliceAutomaticStyles, StyleRegistry } from "./serialize";

function bodyFrame(paragraphs: TextFrame["paragraphs"]): TextFrame {
  return { id: "b", kind: "text", role: "body", xCm: 2, yCm: 5, widthCm: 14, heightCm: 10, paragraphs };
}

describe("serializePage", () => {
  it("keeps the original draw:page attributes for a dirty page (AC-37)", () => {
    const slide: Slide = {
      id: "s",
      frames: [bodyFrame([{ runs: [{ text: "hi" }] }])],
      preserved: { "@page": ` draw:name="page3" draw:master-page-name="Master" draw:style-name="dp1"` },
    };
    const xml = serializePage(slide, new StyleRegistry(), new Set(), "Default");
    expect(xml).toContain(`draw:master-page-name="Master"`);
    expect(xml).toContain(`draw:style-name="dp1"`);
    expect(xml).toContain(`draw:name="page3"`);
  });

  it("regenerates presentation:notes from the model before preserved anim (SPEC-083, AC-16/38)", () => {
    const slide: Slide = {
      id: "s",
      frames: [],
      notes: ["First note", "Second note"],
      preserved: {
        "@page": ` draw:name="p1"`,
        anim: `<anim:par/>`,
      },
    };
    const xml = serializePage(slide, new StyleRegistry(), new Set(), "Default");
    expect(xml).toContain(`<presentation:notes>`);
    expect(xml).toContain(`presentation:class="notes"`);
    expect(xml).toContain(`<text:p>First note</text:p><text:p>Second note</text:p>`);
    // Notes come before the preserved anim, which comes before </draw:page>.
    expect(xml.indexOf("<presentation:notes>")).toBeLessThan(xml.indexOf("<anim:par/>"));
  });

  it("omits presentation:notes when the slide has none (SPEC-083, AC-25)", () => {
    const slide: Slide = { id: "s", frames: [], preserved: { "@page": ` draw:name="p1"` } };
    const xml = serializePage(slide, new StyleRegistry(), new Set(), "Default");
    expect(xml).not.toContain("presentation:notes");
  });

  it("emits a text run's marks as a generated de-T text style", () => {
    const registry = new StyleRegistry();
    const slide: Slide = { id: "s", frames: [bodyFrame([{ runs: [{ text: "bold", bold: true, color: "#ff0000" }] }])] };
    const xml = serializePage(slide, registry, new Set(), "Default");
    expect(xml).toContain(`<text:span text:style-name="de-T1">bold</text:span>`);
    expect(registry.xml()).toContain(`fo:font-weight="bold"`);
    expect(registry.xml()).toContain(`fo:color="#ff0000"`);
  });

  it("de-duplicates identical run styles onto one name", () => {
    const registry = new StyleRegistry();
    const slide: Slide = {
      id: "s",
      frames: [bodyFrame([{ runs: [{ text: "a", bold: true }] }, { runs: [{ text: "b", bold: true }] }])],
    };
    serializePage(slide, registry, new Set(), "Default");
    expect(registry.xml().match(/style:family="text"/g)?.length).toBe(1);
  });

  it("uniques a duplicated page name against those already used", () => {
    const used = new Set<string>(["page1"]);
    const slide: Slide = { id: "s", frames: [], preserved: { "@page": ` draw:name="page1"` } };
    const xml = serializePage(slide, new StyleRegistry(), used, "Default");
    expect(xml).toContain(`draw:name="page1_2"`);
  });

  it("escapes special characters in run text", () => {
    const slide: Slide = { id: "s", frames: [bodyFrame([{ runs: [{ text: "a & b < c" }] }])] };
    const xml = serializePage(slide, new StyleRegistry(), new Set(), "Default");
    expect(xml).toContain(`a &amp; b &lt; c`);
  });

  it("mints a drawing-page style for a transition and repoints draw:style-name (SPEC-085)", () => {
    const registry = new StyleRegistry();
    const slide: Slide = {
      id: "s",
      frames: [],
      transition: { type: "push", direction: "left", durationMs: 600 },
      preserved: { "@page": ` draw:name="p1" draw:style-name="dp1"` },
    };
    const xml = serializePage(slide, registry, new Set(), "Default");
    expect(xml).toContain(`draw:style-name="de-DP1"`);
    expect(xml).not.toContain(`draw:style-name="dp1"`);
    const styles = registry.xml();
    expect(styles).toContain(`style:family="drawing-page"`);
    expect(styles).toContain(`smil:type="pushWipe" smil:subtype="fromRight"`);
    expect(styles).toContain(`smil:dur="PT0.6S"`);
  });

  it("copies the original drawing-page fill/background into the minted style, dropping the old transition (SPEC-085)", () => {
    const registry = new StyleRegistry();
    const source =
      `<style:style style:name="dp1" style:family="drawing-page">` +
      `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#123456" smil:type="spiralWipe" presentation:transition-type="automatic" presentation:duration="PT5S"/>` +
      `</style:style>`;
    const slide: Slide = {
      id: "s",
      frames: [],
      transition: { type: "fade", durationMs: 600 },
      preserved: { "@page": ` draw:name="p1" draw:style-name="dp1"` },
    };
    serializePage(slide, registry, new Set(), "Default", source);
    const styles = registry.xml();
    // The fill/background is carried over…
    expect(styles).toContain(`draw:fill="solid"`);
    expect(styles).toContain(`draw:fill-color="#123456"`);
    // …the replaced transition's smil attrs are NOT (only the new fade's are).
    expect(styles).not.toContain(`spiralWipe`);
    expect(styles).toContain(`smil:type="fade"`);
  });

  it("mints a fresh drawing-page style even for explicit none, so a replaced transition is dropped", () => {
    const registry = new StyleRegistry();
    const source =
      `<style:style style:name="dp1" style:family="drawing-page">` +
      `<style:drawing-page-properties smil:type="spiralWipe"/></style:style>`;
    const slide: Slide = {
      id: "s",
      frames: [],
      transition: { type: "none" },
      preserved: { "@page": ` draw:name="p1" draw:style-name="dp1"` },
    };
    const xml = serializePage(slide, registry, new Set(), "Default", source);
    expect(xml).toContain(`draw:style-name="de-DP1"`);
    expect(registry.xml()).not.toContain("spiralWipe"); // original transition dropped
  });
});

describe("seedFrom / spliceAutomaticStyles", () => {
  it("bumps generated names past a source that already uses the de- prefix", () => {
    const registry = new StyleRegistry();
    registry.seedFrom(`<style:style style:name="de-T5"/>`);
    const slide: Slide = { id: "s", frames: [bodyFrame([{ runs: [{ text: "x", bold: true }] }])] };
    const xml = serializePage(slide, registry, new Set(), "Default");
    expect(xml).toContain(`de-T6`);
  });

  it("splices styles before the closing office:automatic-styles", () => {
    const content = `<a><office:automatic-styles></office:automatic-styles><office:body/></a>`;
    const spliced = spliceAutomaticStyles(content, `<style:style/>`);
    expect(spliced).toBe(`<a><office:automatic-styles><style:style/></office:automatic-styles><office:body/></a>`);
  });

  it("inserts a fresh automatic-styles block when the source lacks one", () => {
    const content = `<a><office:body/></a>`;
    const spliced = spliceAutomaticStyles(content, `<style:style/>`);
    expect(spliced).toBe(`<a><office:automatic-styles><style:style/></office:automatic-styles><office:body/></a>`);
  });
});
