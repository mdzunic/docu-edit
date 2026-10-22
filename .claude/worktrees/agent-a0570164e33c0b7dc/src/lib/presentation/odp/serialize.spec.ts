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

// --- SPEC-097: slide text depth on the odp wire -----------------------------

import type { ImageFrame } from "../model";

function depthFrame(paragraphs: TextFrame["paragraphs"], extra: Partial<TextFrame> = {}): TextFrame {
  return { id: "t", kind: "text", xCm: 1, yCm: 1, widthCm: 10, heightCm: 4, paragraphs, ...extra };
}

/** Serialises one frame with a fresh registry (frame markup + its styles). */
function writeFrame(frame: SlideFrame): { frame: string; autoStyles: string } {
  return serializeOdpGroupFrame(structuredClone(frame));
}

/** Writes a frame and reads it back through the reader. */
function roundTrip<T extends SlideFrame>(frame: T): T {
  const out = writeFrame(frame);
  return parseFrameChildForTest(out.frame, out.autoStyles) as T;
}

const LABEL = (indentCm: string, hangingCm: string): string =>
  `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">` +
  `<style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="${hangingCm}" fo:margin-left="${indentCm}"/>` +
  `</style:list-level-properties>`;

describe("SPEC-097 paragraph and run properties (odp)", () => {
  it("writes spacing and line height into the paragraph style's property string", () => {
    const registry = new StyleRegistry();
    registry.paragraphStyle({ runs: [], align: "center", lineSpacing: 1.5, spaceBeforePt: 6, spaceAfterPt: 12 });
    registry.paragraphStyle({ runs: [], lineSpacingPt: 18, lineSpacingRule: "exact" });
    registry.paragraphStyle({ runs: [], lineSpacingPt: 18 });
    expect(registry.xml()).toBe(
      `<style:style style:name="de-P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center" fo:margin-top="0.212cm" fo:margin-bottom="0.423cm" fo:line-height="150%"/></style:style>` +
        `<style:style style:name="de-P2" style:family="paragraph"><style:paragraph-properties fo:line-height="18pt"/></style:style>` +
        `<style:style style:name="de-P3" style:family="paragraph"><style:paragraph-properties style:line-height-at-least="18pt"/></style:style>`,
    );
  });

  it("writes the effective indent (1.27cm per level) and the hanging marker box", () => {
    const registry = new StyleRegistry();
    registry.paragraphStyle({ runs: [], bullet: "disc" }); // list item: (0 + 1) × 1.27, hanging 0.635
    registry.paragraphStyle({ runs: [], level: 2 }); // plain text: 2 × 1.27, no hanging
    registry.paragraphStyle({ runs: [], bullet: "number", level: 1, indentCm: 3, hangingCm: 1 });
    expect(registry.xml()).toBe(
      `<style:style style:name="de-P1" style:family="paragraph"><style:paragraph-properties fo:margin-left="1.27cm" fo:text-indent="-0.635cm"/></style:style>` +
        `<style:style style:name="de-P2" style:family="paragraph"><style:paragraph-properties fo:margin-left="2.54cm"/></style:style>` +
        `<style:style style:name="de-P3" style:family="paragraph"><style:paragraph-properties fo:margin-left="3cm" fo:text-indent="-1cm"/></style:style>`,
    );
    // An unindented plain paragraph still mints no style.
    expect(new StyleRegistry().paragraphStyle({ runs: [] })).toBeUndefined();
  });

  it("writes letter spacing into the text style in cm", () => {
    const registry = new StyleRegistry();
    registry.textStyle({ text: "a", letterSpacingPt: 0.5 });
    registry.textStyle({ text: "b", letterSpacingPt: -1.2 });
    expect(registry.xml()).toBe(
      `<style:style style:name="de-T1" style:family="text"><style:text-properties fo:letter-spacing="0.018cm"/></style:style>` +
        `<style:style style:name="de-T2" style:family="text"><style:text-properties fo:letter-spacing="-0.042cm"/></style:style>`,
    );
  });

  it("reads every spacing, indent and letter-spacing field back", () => {
    const paragraphs: TextFrame["paragraphs"] = [
      { runs: [{ text: "multiple", letterSpacingPt: 0.5 }], lineSpacing: 1.5, spaceBeforePt: 6, spaceAfterPt: 12, align: "center" },
      { runs: [{ text: "exact", letterSpacingPt: -1.2 }], lineSpacingPt: 18, lineSpacingRule: "exact" },
      { runs: [{ text: "at least" }], lineSpacingPt: 24.5 },
      { runs: [{ text: "indented" }], bullet: "number", level: 1, indentCm: 3, hangingCm: 1 },
      { runs: [{ text: "plain" }], level: 2 },
    ];
    expect(roundTrip(depthFrame(paragraphs)).paragraphs).toEqual(paragraphs);
  });

  it("reads a plain paragraph's whole-step indent as its level, any other as an explicit indent", () => {
    const back = roundTrip(
      depthFrame([
        { runs: [{ text: "stepped" }], level: 3 },
        { runs: [{ text: "custom" }], indentCm: 2 },
        { runs: [{ text: "hanging" }], level: 1, hangingCm: 0.5 },
      ]),
    );
    expect(back.paragraphs).toEqual([
      { runs: [{ text: "stepped" }], level: 3 },
      { runs: [{ text: "custom" }], indentCm: 2 },
      { runs: [{ text: "hanging" }], level: 1, hangingCm: 0.5 },
    ]);
  });

  it("resolves line height and margins through the paragraph style's parent chain", () => {
    const styles =
      `<style:style style:name="Base" style:family="paragraph"><style:paragraph-properties fo:line-height="120%" fo:margin-bottom="0.3cm"/></style:style>` +
      `<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Base"><style:paragraph-properties fo:margin-top="0.1cm"/></style:style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box><text:p text:style-name="P1">x</text:p></draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "x" }], lineSpacing: 1.2, spaceBeforePt: 2.8, spaceAfterPt: 8.5 });
  });
});

describe("SPEC-097 lists (odp)", () => {
  const NESTED: TextFrame["paragraphs"] = [
    { runs: [{ text: "One" }], bullet: "number", startAt: 3 },
    { runs: [{ text: "One-a" }], bullet: "number", bulletStyle: "lowerAlpha", level: 1 },
    { runs: [{ text: "One-b" }], bullet: "number", bulletStyle: "lowerAlpha", level: 1 },
    { runs: [{ text: "Two" }], bullet: "number", startAt: 3 },
    { runs: [{ text: "Deep" }], bullet: "disc", bulletChar: "★", level: 2 },
  ];

  it("wraps a list run in nested text:list markup over one minted de-L list style", () => {
    const out = writeFrame(depthFrame(NESTED));
    expect(out.frame).toBe(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box>` +
        `<text:list text:style-name="de-L1">` +
        `<text:list-item><text:p text:style-name="de-P1">One</text:p>` +
        `<text:list><text:list-item><text:p text:style-name="de-P2">One-a</text:p></text:list-item>` +
        `<text:list-item><text:p text:style-name="de-P2">One-b</text:p></text:list-item></text:list></text:list-item>` +
        `<text:list-item><text:p text:style-name="de-P1">Two</text:p>` +
        // A skipped level (0 → 2) is bridged by an empty list-item.
        `<text:list><text:list-item><text:list><text:list-item><text:p text:style-name="de-P3">Deep</text:p></text:list-item></text:list></text:list-item></text:list>` +
        `</text:list-item></text:list>` +
        `</draw:text-box></draw:frame>`,
    );
    expect(out.autoStyles).toContain(
      `<text:list-style style:name="de-L1">` +
        `<text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="." text:start-value="3">${LABEL("1.27cm", "-0.635cm")}</text:list-level-style-number>` +
        `<text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=".">${LABEL("2.54cm", "-0.635cm")}</text:list-level-style-number>` +
        `<text:list-level-style-bullet text:level="3" text:bullet-char="★">${LABEL("3.81cm", "-0.635cm")}</text:list-level-style-bullet>` +
        `</text:list-style>`,
    );
  });

  it("reads the nesting depth as the level and the level styles as the list fields", () => {
    expect(roundTrip(depthFrame(NESTED)).paragraphs).toEqual(NESTED);
  });

  it("maps every marker style to its level style and back", () => {
    const paragraphs: TextFrame["paragraphs"] = [
      { runs: [{ text: "circle" }], bullet: "disc", bulletStyle: "circle" },
      { runs: [{ text: "square" }], bullet: "disc", bulletStyle: "square" },
      { runs: [{ text: "A" }], bullet: "number", bulletStyle: "upperAlpha" },
      { runs: [{ text: "i" }], bullet: "number", bulletStyle: "lowerRoman" },
      { runs: [{ text: "I" }], bullet: "number", bulletStyle: "upperRoman" },
    ];
    const out = writeFrame(depthFrame(paragraphs));
    for (const attr of [`text:bullet-char="◦"`, `text:bullet-char="▪"`, `style:num-format="A"`, `style:num-format="i"`, `style:num-format="I"`]) {
      expect(out.autoStyles).toContain(attr);
    }
    // A level's marker changing starts a new list (and list style) each time.
    expect(out.frame.match(/<text:list /g)).toHaveLength(5);
    expect(roundTrip(depthFrame(paragraphs)).paragraphs).toEqual(paragraphs);
  });

  it("de-duplicates list styles by signature and seeds the L counter past a source's names", () => {
    const registry = new StyleRegistry();
    registry.seedFrom(`<text:list-style style:name="de-L4"/>`);
    const list = [{ runs: [{ text: "x" }], bullet: "number" as const }];
    const slide: Slide = { id: "s", frames: [depthFrame(list), { ...depthFrame(list), id: "u" }] };
    const xml = serializePage(slide, registry, new Set(), "Default");
    expect(xml.match(/text:style-name="de-L5"/g)).toHaveLength(2);
    expect(registry.xml().match(/<text:list-style /g)).toHaveLength(1);
  });

  it("keeps a body placeholder's plain discs bare and puts its non-list paragraphs in a list header", () => {
    const out = writeFrame(
      bodyFrame([
        { runs: [{ text: "first" }], bullet: "disc" },
        { runs: [{ text: "second" }], bullet: "disc" },
        { runs: [{ text: "aside" }] },
      ]),
    );
    expect(out.frame).toContain(
      `<draw:text-box><text:p text:style-name="de-P1">first</text:p><text:p text:style-name="de-P1">second</text:p>` +
        `<text:list><text:list-header><text:p>aside</text:p></text:list-header></text:list></draw:text-box>`,
    );
    const back = parseFrameChildForTest(out.frame, out.autoStyles) as TextFrame;
    expect(back.paragraphs.map((p) => p.bullet)).toEqual(["disc", "disc", undefined]);
  });

  it("writes a body placeholder's numbered or nested list as real markup", () => {
    const out = writeFrame(
      bodyFrame([
        { runs: [{ text: "one" }], bullet: "number" },
        { runs: [{ text: "sub" }], bullet: "disc", level: 1 },
      ]),
    );
    expect(out.frame).toContain(`<text:list text:style-name="de-L1"><text:list-item><text:p text:style-name="de-P1">one</text:p><text:list>`);
    const back = parseFrameChildForTest(out.frame, out.autoStyles) as TextFrame;
    expect(back.paragraphs.map((p) => [p.bullet, p.level])).toEqual([
      ["number", undefined],
      ["disc", 1],
    ]);
  });

  it("reads an Impress-style label-alignment list: every item carries the level's start value", () => {
    const styles =
      `<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="." text:start-value="5">` +
      `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">` +
      `<style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-1.27cm" fo:margin-left="1.27cm"/>` +
      `</style:list-level-properties></text:list-level-style-number></text:list-style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box><text:list text:style-name="L1">` +
        `<text:list-item><text:p>a</text:p></text:list-item><text:list-item><text:p>b</text:p></text:list-item>` +
        `</text:list></draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    expect(frame.paragraphs).toEqual([
      { runs: [{ text: "a" }], bullet: "number", startAt: 5, hangingCm: 1.27 },
      { runs: [{ text: "b" }], bullet: "number", startAt: 5, hangingCm: 1.27 },
    ]);
  });

  it("reads a list with no resolvable style as disc, and the older space-before indents", () => {
    const styles =
      `<text:list-style style:name="Lpos"><text:list-level-style-bullet text:level="1" text:bullet-char="●">` +
      `<style:list-level-properties text:space-before="1cm" text:min-label-width="0.5cm"/></text:list-level-style-bullet></text:list-style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box>` +
        `<text:list text:style-name="Missing"><text:list-item><text:p>a</text:p></text:list-item></text:list>` +
        `<text:list text:style-name="Lpos"><text:list-item><text:p>b</text:p></text:list-item></text:list>` +
        `</draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    expect(frame.paragraphs).toEqual([
      { runs: [{ text: "a" }], bullet: "disc" },
      { runs: [{ text: "b" }], bullet: "disc", bulletChar: "●", indentCm: 1.5, hangingCm: 0.5 },
    ]);
  });

  it("writes a picture bullet as a list-level-style-image into Pictures/", () => {
    const out = writeFrame(depthFrame([{ runs: [{ text: "pic" }], bullet: "disc", bulletImageId: "Pictures/bullet.png" }]));
    expect(out.autoStyles).toContain(
      `<text:list-level-style-image text:level="1" xlink:href="Pictures/bullet.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">`,
    );
  });
});

describe("SPEC-097 text-body properties (odp)", () => {
  it("writes the body onto the frame's graphic style, preserved properties after the modelled ones", () => {
    const out = writeFrame(
      depthFrame([{ runs: [{ text: "x" }] }], {
        body: { insetsCm: { left: 0.5, top: 0.25, right: 0.5, bottom: 0.25 }, anchor: "bottom", autofit: "resizeShape", wrap: false, columns: { count: 2, gapCm: 0.5 } },
        // The modelled anchor wins over a preserved one of the same name.
        odpBodyProps: `style:writing-mode="tb-rl" draw:textarea-vertical-align="justify"`,
      }),
    );
    expect(out.frame).toContain(`<draw:frame draw:style-name="de-G1"`);
    expect(out.autoStyles).toBe(
      `<style:style style:name="de-G1" style:family="graphic"><style:graphic-properties` +
        ` fo:padding-left="0.5cm" fo:padding-top="0.25cm" fo:padding-right="0.5cm" fo:padding-bottom="0.25cm"` +
        ` draw:textarea-vertical-align="bottom" draw:auto-grow-height="true" fo:wrap-option="no-wrap" style:writing-mode="tb-rl">` +
        `<style:columns fo:column-count="2" fo:column-gap="0.5cm"/></style:graphic-properties></style:style>`,
    );
  });

  it("reads the body back, kind defaults absent", () => {
    const text = depthFrame([{ runs: [{ text: "x" }] }], {
      body: { insetsCm: { left: 0.5, top: 0.25, right: 0.5, bottom: 0.25 }, anchor: "middle", autofit: "resizeShape", wrap: false, columns: { count: 3, gapCm: 0.4 } },
      odpBodyProps: `style:writing-mode="tb-rl"`,
      fill: { type: "solid", color: "#ff0000" },
    });
    const back = roundTrip(text);
    expect([back.body, back.odpBodyProps, back.fill]).toEqual([text.body, text.odpBodyProps, text.fill]);
    const shape: ShapeFrame = { id: "s", kind: "shape", geometry: "rect", xCm: 1, yCm: 1, widthCm: 4, heightCm: 2, paragraphs: [{ runs: [{ text: "s" }] }], body: { anchor: "top" } };
    expect(roundTrip(shape).body).toEqual({ anchor: "top" });
  });

  it("resolves the body along the draw:style-name chain and drops values equal to the defaults", () => {
    const styles =
      `<style:style style:name="base" style:family="graphic"><style:graphic-properties fo:padding="0.4cm" draw:textarea-vertical-align="top" draw:auto-grow-height="false"/></style:style>` +
      `<style:style style:name="gr1" style:family="graphic" style:parent-style-name="base"><style:graphic-properties fo:padding-left="0.254cm"/></style:style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame draw:style-name="gr1" svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box><text:p>x</text:p></draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    expect(frame.body).toEqual({ insetsCm: { left: 0.254, top: 0.4, right: 0.4, bottom: 0.4 } });
    expect(frame.odpBodyProps).toBeUndefined();
  });

  it("writes no body attribute for a frame without one (existing pages keep their bytes)", () => {
    expect(writeFrame(depthFrame([{ runs: [{ text: "x" }] }])).autoStyles).toBe("");
  });
});

describe("SPEC-097 preservation slots (odp)", () => {
  it("re-emits odpParaProps after the modelled properties — attributes, then child elements — modelled winning a clash", () => {
    const registry = new StyleRegistry();
    registry.paragraphStyle({
      runs: [],
      bullet: "disc",
      odpParaProps: `fo:margin-left="9cm" fo:keep-together="always"<style:tab-stops><style:tab-stop style:position="2cm"/></style:tab-stops>`,
    });
    expect(registry.xml()).toBe(
      `<style:style style:name="de-P1" style:family="paragraph"><style:paragraph-properties fo:margin-left="1.27cm" fo:text-indent="-0.635cm" fo:keep-together="always">` +
        `<style:tab-stops><style:tab-stop style:position="2cm"/></style:tab-stops></style:paragraph-properties></style:style>`,
    );
  });

  it("reads the paragraph's OWN style's unmodelled properties into odpParaProps, in source order", () => {
    const styles =
      `<style:style style:name="Parent" style:family="paragraph"><style:paragraph-properties fo:hyphenate="true"/></style:style>` +
      `<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Parent"><style:paragraph-properties` +
      ` style:writing-mode="tb-rl" fo:text-align="center" fo:margin-left="1cm" fo:text-indent="0.5cm" fo:keep-together="always"/></style:style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box><text:p text:style-name="P1">x</text:p></draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    // The modelled align/indent are fields; a positive (first-line) text-indent
    // is not modelled and is kept; the parent's properties are not copied.
    expect(frame.paragraphs[0]).toEqual({
      runs: [{ text: "x" }],
      align: "center",
      indentCm: 1,
      odpParaProps: `style:writing-mode="tb-rl" fo:text-indent="0.5cm" fo:keep-together="always"`,
    });
    // …and re-emits it after the modelled properties.
    const out = writeFrame(frame);
    expect(out.autoStyles).toContain(
      `<style:paragraph-properties fo:text-align="center" fo:margin-left="1cm" style:writing-mode="tb-rl" fo:text-indent="0.5cm" fo:keep-together="always"/>`,
    );
  });

  it("re-saves a rich page with byte-identical page and style bytes (write → read → write)", () => {
    const frames: SlideFrame[] = [
      bodyFrame([
        { runs: [{ text: "One", letterSpacingPt: 0.5 }], bullet: "number", startAt: 3, lineSpacing: 1.5, spaceBeforePt: 6 },
        { runs: [{ text: "Sub" }], bullet: "disc", bulletStyle: "square", level: 1, odpParaProps: `fo:keep-together="always"` },
        { runs: [{ text: "see " }, { text: "link", linkHref: "https://example.com/", linkTooltip: "Tip" }] },
      ]),
      depthFrame([{ runs: [{ text: "boxed" }] }], {
        body: { anchor: "bottom", insetsCm: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 }, columns: { count: 2, gapCm: 0.3 } },
        odpBodyProps: `style:writing-mode="tb-rl"`,
        linkHref: "https://example.org/",
      }),
    ];
    const first = new StyleRegistry();
    const page1 = serializePage({ id: "s", frames, preserved: { "@page": ` draw:name="p1"` } }, first, new Set(), "Default");
    const inner = page1.replace(/^<draw:page[^>]*>/, "").replace(/<\/draw:page>$/, "");
    const read: SlideFrame[] = [];
    for (const element of inner.match(/<draw:(?:a|frame)\b[\s\S]*?<\/draw:(?:a|frame)>(?=<draw:|$)/g) ?? []) {
      read.push(parseFrameChildForTest(element, first.xml())!);
    }
    expect(read).toHaveLength(2);
    const second = new StyleRegistry();
    const page2 = serializePage({ id: "s", frames: read, preserved: { "@page": ` draw:name="p1"` } }, second, new Set(), "Default");
    expect(page2).toBe(page1);
    expect(second.xml()).toBe(first.xml());
  });
});

describe("SPEC-097 links (odp)", () => {
  it("wraps consecutive runs sharing a link in one text:a", () => {
    const out = writeFrame(
      depthFrame([
        {
          runs: [
            { text: "see " },
            { text: "here", linkHref: "https://a.example/", linkTooltip: "Tip" },
            { text: "!", linkHref: "https://a.example/", linkTooltip: "Tip", bold: true },
            { text: "b", linkHref: "https://b.example/?q=1&r=2" },
          ],
        },
      ]),
    );
    expect(out.frame).toContain(
      `<text:p>see <text:a xlink:href="https://a.example/" xlink:type="simple" office:title="Tip">here<text:span text:style-name="de-T1">!</text:span></text:a>` +
        `<text:a xlink:href="https://b.example/?q=1&amp;r=2" xlink:type="simple">b</text:a></text:p>`,
    );
  });

  it("reads a text:a back as run links (spans inside keep their style) and strips a dangerous href", () => {
    const runs: TextFrame["paragraphs"][number]["runs"] = [
      { text: "see " },
      { text: "here", linkHref: "https://a.example/", linkTooltip: "Tip" },
      { text: "!", bold: true, linkHref: "https://a.example/", linkTooltip: "Tip" },
    ];
    expect(roundTrip(depthFrame([{ runs }])).paragraphs[0].runs).toEqual(runs);
    expect(writeFrame(depthFrame([{ runs: [{ text: "x", linkHref: "javascript:alert(1)" }] }])).frame).not.toContain("text:a");
    const read = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p><text:a xlink:href="javascript:alert(1)">x</text:a></text:p></draw:text-box></draw:frame>`,
    ) as TextFrame;
    expect(read.paragraphs[0].runs).toEqual([{ text: "x" }]);
  });

  it("wraps a linked text, image or shape frame in draw:a and reads it back", () => {
    const image: ImageFrame = { id: "i", kind: "image", xCm: 1, yCm: 1, widthCm: 4, heightCm: 3, mediaId: "Pictures/a.png", linkHref: "https://img.example/" };
    expect(writeFrame(image).frame).toBe(
      `<draw:a xlink:href="https://img.example/" xlink:type="simple"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="3cm">` +
        `<draw:image xlink:href="Pictures/a.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></draw:a>`,
    );
    const shape: ShapeFrame = { id: "s", kind: "shape", geometry: "ellipse", xCm: 1, yCm: 1, widthCm: 4, heightCm: 2, linkHref: "https://shape.example/", linkTooltip: "Shape" };
    expect(writeFrame(shape).frame).toMatch(/^<draw:a xlink:href="https:\/\/shape.example\/" xlink:type="simple" office:title="Shape"><draw:ellipse /);
    const back = roundTrip(shape);
    expect([back.kind, back.linkHref, back.linkTooltip]).toEqual(["shape", "https://shape.example/", "Shape"]);
    const text = roundTrip(depthFrame([{ runs: [{ text: "t" }] }], { linkHref: "https://text.example/" }));
    expect(text.linkHref).toBe("https://text.example/");
  });

  it("never wraps a non-linkable frame, and keeps a draw:a around one verbatim on read", () => {
    const group: SlideFrame = {
      id: "g",
      kind: "group",
      xCm: 1,
      yCm: 1,
      widthCm: 4,
      heightCm: 2,
      children: [{ id: "r", kind: "shape", geometry: "rect", xCm: 1, yCm: 1, widthCm: 4, heightCm: 2 }],
      linkHref: "https://group.example/",
    };
    expect(writeFrame(group).frame).not.toContain("draw:a");
    const read = parseFrameChildForTest(
      `<draw:a xlink:href="https://g.example/" xlink:type="simple"><draw:g><draw:rect svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"/></draw:g></draw:a>`,
    );
    expect(read?.kind).toBe("foreign");
  });
});

// --- SPEC-104 hidden pages ---------------------------------------------------
import { GOLDEN_ODP_HIDDEN_PAGE_ATTR, GOLDEN_ODP_HIDDEN_STYLE, ODP_HIDDEN_VISIBILITY_PROP } from "./visibilityGolden";

describe("odp visibility golden (SPEC-104)", () => {
  it("a hidden from-scratch page mints exactly the golden drawing-page style and reference", () => {
    const registry = new StyleRegistry();
    const xml = serializePage({ id: "s", frames: [], hidden: true }, registry, new Set(), "Default");
    expect(xml).toBe(`<draw:page draw:name="page" draw:master-page-name="Default"${GOLDEN_ODP_HIDDEN_PAGE_ATTR}></draw:page>`);
    expect(registry.xml()).toContain(GOLDEN_ODP_HIDDEN_STYLE);
  });

  it("strips an inherited visibility from the copied base and states the page's own", () => {
    const source =
      `<style:style style:name="dp1" style:family="drawing-page">` +
      `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#112233" presentation:visibility="visible"/></style:style>`;
    const registry = new StyleRegistry();
    const xml = serializePage(
      { id: "s", frames: [], hidden: true, preserved: { "@page": ` draw:name="p1" draw:style-name="dp1"` } },
      registry,
      new Set(),
      "Default",
      source,
    );
    expect(xml).toContain(`draw:style-name="de-DP1"`);
    expect(registry.xml()).toContain(
      `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#112233" ${ODP_HIDDEN_VISIBILITY_PROP}/>`,
    );
  });

  it("a shown page whose style states nothing about visibility keeps its style", () => {
    const source = `<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="none"/></style:style>`;
    const registry = new StyleRegistry();
    const xml = serializePage(
      { id: "s", frames: [], preserved: { "@page": ` draw:name="p1" draw:style-name="dp1"` } },
      registry,
      new Set(),
      "Default",
      source,
    );
    expect(xml).toBe(`<draw:page draw:name="p1" draw:style-name="dp1"></draw:page>`);
  });
});
