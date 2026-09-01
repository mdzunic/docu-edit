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

// --- SPEC-093 field declarations + display fold ----------------------------
import { planOdpDeclarations } from "./fields";
import { GOLDEN_ODP_DATE_STYLE, GOLDEN_ODP_DECLS, goldenOdpModel } from "./fieldGolden";

describe("odp field golden (SPEC-093 AC-127/128/134)", () => {
  it("plans the canonical footer/date-time declaration block + date style", () => {
    const plan = planOdpDeclarations(goldenOdpModel());
    expect(plan.declXml).toBe(GOLDEN_ODP_DECLS);
    expect(plan.dateStyleXml).toBe(GOLDEN_ODP_DATE_STYLE);
  });

  it("reuses a source declaration name when the value is unchanged (AC-77)", () => {
    const model = goldenOdpModel();
    // Source prefix already declares the footer under a non-default name.
    const prefix = `<office:presentation><presentation:footer-decl presentation:name="ftr7">Confidential</presentation:footer-decl>`;
    const plan = planOdpDeclarations(model, prefix);
    expect(plan.declXml).toContain(`presentation:name="ftr7">Confidential`);
  });

  it("folds the display flags into the drawing-page style + sets use-*-name (AC-78/79)", () => {
    const model = goldenOdpModel();
    const plan = planOdpDeclarations(model);
    const registry = new StyleRegistry();
    const page = serializePage(model.slides[1], registry, new Set(), "Default", "", undefined, plan.pagePlans.get(model.slides[1]));
    expect(page).toContain(`presentation:use-footer-name="ftr1"`);
    expect(page).toContain(`presentation:use-date-time-name="dtd1"`);
    expect(registry.xml()).toContain(`presentation:display-footer="true"`);
    expect(registry.xml()).toContain(`presentation:display-page-number="true"`);
    expect(registry.xml()).toContain(`presentation:display-date-time="true"`);
  });

  it("never serialises a field placeholder into the draw:page (AC-80)", () => {
    const model = goldenOdpModel();
    const plan = planOdpDeclarations(model);
    const page = serializePage(model.slides[1], new StyleRegistry(), new Set(), "Default", "", undefined, plan.pagePlans.get(model.slides[1]));
    expect(page).not.toContain(`presentation:class="footer"`);
    expect(page).not.toContain(`<text:page-number`);
  });
});

// --- SPEC-094: the de-theme-* parent link -----------------------------------

describe("StyleRegistry — theme references (SPEC-094)", () => {
  it("an untransformed colour ref parents its styles on the scheme styles, beside the literal", () => {
    const registry = new StyleRegistry();
    registry.textStyle({ text: "T", color: { theme: "accent1" } });
    registry.graphicStyle({ fill: { type: "solid", color: { theme: "accent2" } } });
    const xml = registry.xml();
    expect(xml).toContain('style:parent-style-name="de-theme-text-accent1"');
    expect(xml).toContain('fo:color="#4472c4"');
    expect(xml).toContain('style:parent-style-name="de-theme-accent2"');
    expect(xml).toContain('draw:fill-color="#ed7d31"');
  });

  it("a transformed ref writes the literal alone — the stated odp loss", () => {
    const registry = new StyleRegistry();
    registry.graphicStyle({ fill: { type: "solid", color: { theme: "accent1", lumMod: 0.6, lumOff: 0.4 } } });
    const xml = registry.xml();
    expect(xml).not.toContain("parent-style-name");
    expect(xml).toContain('draw:fill-color="#8faadc"');
  });

  it("a theme font ref resolves to its family name — the stated odp loss", () => {
    const registry = new StyleRegistry();
    registry.textStyle({ text: "T", fontFamily: { theme: "major" } });
    expect(registry.xml()).toContain('style:font-name="Calibri Light"');
  });
});

// --- SPEC-096: authored draw:g, draw:modifiers, rotation ---------------------

import { serializeOdpGroupFrame } from "./serialize";
import { parseFrameChildForTest } from "./reader";
import {
  ADJ_GOLDEN as ODP_ADJ_GOLDEN,
  ADJ_GOLDEN_REORDERED as ODP_ADJ_GOLDEN_REORDERED,
  GOLDEN_ADJ_FRAME,
  GOLDEN_GROUP_FRAME,
  GROUP_GOLDEN as ODP_GROUP_GOLDEN,
  GROUP_GOLDEN_AUTO_STYLES,
  GROUP_GOLDEN_REORDERED as ODP_GROUP_GOLDEN_REORDERED,
} from "./groupGolden";
import type { GroupFrame, ShapeFrame, SlideFrame } from "../model";

describe("odp group/adjustment goldens (SPEC-096)", () => {
  it("serialises an authored group exactly as GROUP_GOLDEN (write direction)", () => {
    const { frame, autoStyles } = serializeOdpGroupFrame(structuredClone(GOLDEN_GROUP_FRAME) as SlideFrame);
    expect(frame).toBe(ODP_GROUP_GOLDEN);
    expect(autoStyles).toBe(GROUP_GOLDEN_AUTO_STYLES);
  });

  it("serialises the tuned roundRect exactly as ADJ_GOLDEN, draw:modifiers in our 100000ths", () => {
    const { frame, autoStyles } = serializeOdpGroupFrame(structuredClone(GOLDEN_ADJ_FRAME));
    expect(frame).toBe(ODP_ADJ_GOLDEN);
    expect(autoStyles).toBe("");
  });

  it("reads GROUP_GOLDEN (and reordered) back: children keep their own svg geometry", () => {
    for (const golden of [ODP_GROUP_GOLDEN, ODP_GROUP_GOLDEN_REORDERED]) {
      const frame = parseFrameChildForTest(golden, GROUP_GOLDEN_AUTO_STYLES) as GroupFrame;
      expect(frame.kind).toBe("group");
      expect(frame.xml).toBeTruthy(); // read-side groups stay verbatim until edited
      // A draw:g has no box of its own — bounds are the children's union.
      expect([frame.xCm, frame.yCm, frame.widthCm, frame.heightCm]).toEqual([2, 2, 8, 4]);
      const [rect, ellipse] = frame.children as ShapeFrame[];
      expect([rect.xCm, rect.yCm, rect.widthCm, rect.heightCm]).toEqual([2, 2, 3, 2]);
      expect(rect.fill).toEqual({ type: "solid", color: "#ff0000" });
      expect(ellipse.geometry).toBe("ellipse");
    }
  });

  it("reads ADJ_GOLDEN (and reordered) back to the stored fractions", () => {
    for (const golden of [ODP_ADJ_GOLDEN, ODP_ADJ_GOLDEN_REORDERED]) {
      const frame = parseFrameChildForTest(golden) as ShapeFrame;
      expect(frame.kind).toBe("shape");
      expect(frame.geometry).toBe("roundRect");
      expect(frame.adjustments).toEqual([0.25]);
    }
  });
});

describe("odp draw:transform rotation (SPEC-096)", () => {
  it("reads a lone rotate(r) as clockwise degrees (r is ccw radians)", () => {
    const quarter = (-90 * Math.PI) / 180; // 90° clockwise → rotate(-π/2)
    const xml = `<draw:custom-shape draw:transform="rotate(${quarter.toFixed(6)})" svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"><draw:enhanced-geometry draw:type="round-rectangle"/></draw:custom-shape>`;
    const frame = parseFrameChildForTest(xml) as ShapeFrame;
    expect(frame.rotationDeg).toBe(90);
  });

  it("round-trips the serialiser's own rotate() through the reader", () => {
    const rotated: ShapeFrame = { ...(structuredClone(GOLDEN_ADJ_FRAME)), rotationDeg: 37 };
    const { frame } = serializeOdpGroupFrame(rotated);
    const back = parseFrameChildForTest(frame) as ShapeFrame;
    expect(back.rotationDeg).toBe(37);
  });

  it("does not model a transform carrying anything but a lone rotate()", () => {
    const xml = `<draw:custom-shape draw:transform="rotate(0.5) translate(1cm 2cm)" svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"><draw:enhanced-geometry draw:type="round-rectangle"/></draw:custom-shape>`;
    const frame = parseFrameChildForTest(xml) as ShapeFrame;
    expect(frame.rotationDeg).toBeUndefined();
  });
});
