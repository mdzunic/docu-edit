import { describe, expect, it } from "vitest";
import type { Relationship } from "../../core/xml";
import type { Slide, TextFrame } from "../model";
import { serializeSlideXml } from "./serialize";

const NO_RELS = new Map<string, Relationship>();

function textSlide(frames: Slide["frames"], preserved?: Record<string, string>): Slide {
  const slide: Slide = { id: "s", frames };
  if (preserved) slide.preserved = preserved;
  return slide;
}

function titleFrame(paragraphs: TextFrame["paragraphs"]): TextFrame {
  return {
    id: "t",
    kind: "text",
    role: "title",
    phType: "title",
    xCm: 2,
    yCm: 1,
    widthCm: 20,
    heightCm: 3,
    paragraphs,
  };
}

describe("serializeSlideXml — text frame", () => {
  it("emits a placeholder p:sp with p:ph, xfrm, and explicit run properties", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: "Hello", bold: true, sizePt: 24, color: "#ff0000" }] }])]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<p:ph type="title"/>`);
    expect(xml).toContain(`<a:off x="720000" y="360000"/>`);
    expect(xml).toContain(`sz="2400"`);
    expect(xml).toContain(` b="1"`);
    expect(xml).toContain(`<a:srgbClr val="FF0000"/>`);
    expect(xml).toContain(`<a:t>Hello</a:t>`);
  });

  it("writes the 18pt / #000000 defaults explicitly so effective values survive", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: "plain" }] }])]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`sz="1800"`);
    expect(xml).toContain(`<a:srgbClr val="000000"/>`);
  });

  it("emits a:buNone for an unbulleted paragraph and a:buChar for a disc", () => {
    const slide = textSlide([
      titleFrame([
        { runs: [{ text: "a" }] },
        { runs: [{ text: "b" }], bullet: "disc", align: "center", level: 1 },
      ]),
    ]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<a:buNone/>`);
    expect(xml).toContain(`<a:pPr algn="ctr" lvl="1"><a:buChar char="•"/></a:pPr>`);
  });

  it("emits an empty paragraph as a:p with a:endParaRPr", () => {
    const slide = textSlide([titleFrame([{ runs: [] }])]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<a:endParaRPr lang="en-US"/>`);
  });

  it("escapes XML special characters and non-ASCII in run text", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: `a & b < c > "d" — ünïcode` }] }])]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<a:t>a &amp; b &lt; c &gt; "d" — ünïcode</a:t>`);
  });
});

describe("serializeSlideXml — preserved XML placement", () => {
  it("re-emits preserved bg/clrMapOvr/transition/timing/extLst in schema position and root attrs", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: "x" }] }])], {
      "@root": ` xmlns:a="A" xmlns:r="R" xmlns:p="P" showMasterSp="0"`,
      "p:bg": `<p:bg><p:bgPr/></p:bg>`,
      "p:clrMapOvr": `<p:clrMapOvr><a:overrideClrMapping/></p:clrMapOvr>`,
      "p:transition": `<p:transition/>`,
      "p:timing": `<p:timing/>`,
      "p:extLst": `<p:extLst/>`,
    });
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<p:sld xmlns:a="A" xmlns:r="R" xmlns:p="P" showMasterSp="0">`);
    // p:bg opens the cSld, before the spTree.
    expect(xml.indexOf("<p:bg>")).toBeLessThan(xml.indexOf("<p:spTree>"));
    // clrMapOvr → transition → timing → extLst, all after the cSld closes.
    const order = ["</p:cSld>", "<p:clrMapOvr>", "<p:transition/>", "<p:timing/>", "<p:extLst/>", "</p:sld>"];
    let last = -1;
    for (const token of order) {
      const at = xml.indexOf(token);
      expect(at, token).toBeGreaterThan(last);
      last = at;
    }
  });

  it("emits a modelled transition in schema position (SPEC-085 AC-11)", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: "x" }] }])]);
    slide.transition = { type: "push", direction: "left", durationMs: 1000 };
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<p:transition spd="slow"><p:push dir="l"/></p:transition>`);
    expect(xml.indexOf("<p:clrMapOvr>")).toBeLessThan(xml.indexOf("<p:transition"));
  });

  it("prefers the modelled transition over a preserved one (never both)", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: "x" }] }])], {
      "p:transition": `<p:transition><p:morph/></p:transition>`,
    });
    slide.transition = { type: "fade", durationMs: 600 };
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<p:transition spd="med"><p:fade/></p:transition>`);
    expect(xml).not.toContain("p:morph");
  });
});

describe("serializeSlideXml — foreign, group, image", () => {
  it("re-emits group and foreign frames verbatim, in z-order", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        { id: "g", kind: "group", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, children: [], xml: "<p:grpSp/>" },
        { id: "f", kind: "foreign", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, xml: "<p:graphicFrame/>", label: "Table" },
      ],
    };
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml.indexOf("<p:grpSp/>")).toBeLessThan(xml.indexOf("<p:graphicFrame/>"));
  });

  it("resolves an image frame's r:embed through the carried rels", () => {
    const rels = new Map<string, Relationship>([
      ["rId2", { id: "rId2", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: "../media/image1.png" }],
    ]);
    const slide: Slide = {
      id: "s",
      frames: [{ id: "i", kind: "image", xCm: 1, yCm: 1, widthCm: 2, heightCm: 2, mediaId: "ppt/media/image1.png" }],
    };
    const xml = serializeSlideXml(slide, rels, "ppt/slides/slide1.xml");
    expect(xml).toContain(`<a:blip r:embed="rId2"/>`);
  });
});

// --- SPEC-093 deck field placeholders --------------------------------------
import { canonicalFieldSp, resolveSlideFields } from "./serialize";
import { parseFieldPlaceholderForTest } from "./reader";
import {
  GOLDEN_DATE_ENTRY,
  GOLDEN_DATE_RF,
  GOLDEN_DATE_SP,
  GOLDEN_FOOTER_ENTRY,
  GOLDEN_FOOTER_RF,
  GOLDEN_FOOTER_SP,
  GOLDEN_FOOTER_SP_REORDERED,
  GOLDEN_SLIDENUM_ENTRY,
  GOLDEN_SLIDENUM_RF,
  GOLDEN_SLIDENUM_SP,
} from "./fieldGolden";
import type { PresentationModel } from "../model";

describe("pptx field golden (SPEC-093 AC-127..129)", () => {
  it("serialises each canonical field p:sp exactly (write direction)", () => {
    expect(canonicalFieldSp(GOLDEN_FOOTER_RF, 5)).toBe(GOLDEN_FOOTER_SP);
    expect(canonicalFieldSp(GOLDEN_SLIDENUM_RF, 6)).toBe(GOLDEN_SLIDENUM_SP);
    expect(canonicalFieldSp(GOLDEN_DATE_RF, 7)).toBe(GOLDEN_DATE_SP);
  });

  it("reads each canonical field p:sp back to its entry (read direction)", () => {
    expect(parseFieldPlaceholderForTest(GOLDEN_FOOTER_SP, "footer")).toEqual(GOLDEN_FOOTER_ENTRY);
    expect(parseFieldPlaceholderForTest(GOLDEN_SLIDENUM_SP, "slideNumber")).toEqual(GOLDEN_SLIDENUM_ENTRY);
    expect(parseFieldPlaceholderForTest(GOLDEN_DATE_SP, "date")).toEqual(GOLDEN_DATE_ENTRY);
  });

  it("accepts an attribute-reordered PowerPoint-flavoured variant", () => {
    expect(parseFieldPlaceholderForTest(GOLDEN_FOOTER_SP_REORDERED, "footer")).toEqual(GOLDEN_FOOTER_ENTRY);
  });
});

describe("pptx field serialisation (SPEC-093 AC-66/67/68/69)", () => {
  const deck = (slide: Slide): PresentationModel => ({
    kind: "presentation",
    id: "d",
    title: "D",
    format: "pptx",
    slides: [slide],
    slideWidthCm: 25.4,
    slideHeightCm: 19.05,
  });

  it("rewrites only the cached a:t inside a preserved field p:sp", () => {
    const preserved = GOLDEN_SLIDENUM_SP; // cached "3"
    const slide: Slide = {
      id: "s",
      frames: [],
      fields: [{ kind: "slideNumber", xCm: 2, yCm: 18, widthCm: 8, heightCm: 1, align: "right" }],
      preserved: { "p:ph:sldNum": preserved },
    };
    const model = deck(slide);
    const fields = resolveSlideFields(model, 0, new Date(2026, 4, 14));
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml", fields);
    // The preserved p:sp is re-emitted with only its a:t refreshed to the live number "1".
    expect(xml).toContain(`type="slidenum"><a:t>1</a:t></a:fld>`);
    expect(xml).toContain(`name="Slide Number Placeholder"`);
  });

  it("emits a plain a:r for a footer and a:fld for a number", () => {
    const slide: Slide = {
      id: "s",
      frames: [],
      fields: [
        { kind: "footer", xCm: 2, yCm: 18, widthCm: 8, heightCm: 1, align: "center", text: "Confidential" },
        { kind: "slideNumber", xCm: 2, yCm: 18, widthCm: 8, heightCm: 1, align: "right" },
      ],
    };
    const model = deck(slide);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml", resolveSlideFields(model, 0, new Date()));
    expect(xml).toContain(`<a:r><a:rPr lang="en-US"/><a:t>Confidential</a:t></a:r>`);
    expect(xml).toContain(`type="slidenum">`);
  });
});

// --- SPEC-094: theme reference emission (AC-21) -----------------------------

describe("serializeSlideXml — theme references (SPEC-094)", () => {
  const DECK_THEME = {
    name: "Office",
    colors: {
      dk1: "#000000", lt1: "#ffffff", dk2: "#44546a", lt2: "#e7e6e6",
      accent1: "#4472c4", accent2: "#ed7d31", accent3: "#a5a5a5", accent4: "#ffc000",
      accent5: "#5b9bd5", accent6: "#70ad47", hlink: "#0563c1", folHlink: "#954f72",
    },
    fonts: { major: "Calibri Light", minor: "Calibri" },
    colorMap: { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2", accent1: "accent1" },
  } as const;

  it("writes a run colour ref as a:schemeClr through the clrMap, in canonical transform order", () => {
    const slide = textSlide([
      titleFrame([{ runs: [{ text: "T", color: { theme: "dk1", lumMod: 0.6, lumOff: 0.4 } }] }]),
    ]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml", [], structuredClone(DECK_THEME));
    expect(xml).toContain('<a:schemeClr val="tx1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>');
  });

  it("writes a theme font ref as +mn-lt / +mj-lt", () => {
    const slide = textSlide([
      titleFrame([{ runs: [{ text: "T", fontFamily: { theme: "minor" } }, { text: "U", fontFamily: { theme: "major" } }] }]),
    ]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml", [], structuredClone(DECK_THEME));
    expect(xml).toContain('<a:latin typeface="+mn-lt"/>');
    expect(xml).toContain('<a:latin typeface="+mj-lt"/>');
  });

  it("still writes a literal as a:srgbClr beside a ref (AC-19)", () => {
    const slide = textSlide([
      titleFrame([{ runs: [{ text: "T", color: { theme: "accent1" } }, { text: "U", color: "#ff0000" }] }]),
    ]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml", [], structuredClone(DECK_THEME));
    expect(xml).toContain('<a:schemeClr val="accent1"/>');
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
  });

  it("writes refs through a parseable p:clrMapOvr override and literals under an unparseable one (AC-22/23)", () => {
    const ref = { runs: [{ text: "T", color: { theme: "accent1" } }] };
    const withOverride = textSlide([titleFrame([ref])], {
      "p:clrMapOvr": '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" accent1="accent1" accent2="accent1"/></p:clrMapOvr>',
    });
    const overrideXml = serializeSlideXml(withOverride, NO_RELS, "ppt/slides/slide1.xml", [], structuredClone(DECK_THEME));
    expect(overrideXml).toContain('<a:schemeClr val="accent1"/>');
    const optedOut = textSlide([titleFrame([ref])], {
      "p:clrMapOvr": "<p:clrMapOvr><p:garbage/></p:clrMapOvr>",
    });
    const literalXml = serializeSlideXml(optedOut, NO_RELS, "ppt/slides/slide1.xml", [], structuredClone(DECK_THEME));
    expect(literalXml).not.toContain("a:schemeClr");
    expect(literalXml).toContain('<a:srgbClr val="4472C4"/>'); // the resolved accent1
  });

  it("with no deck theme, a (stray) ref writes its Office-resolved literal", () => {
    const slide = textSlide([titleFrame([{ runs: [{ text: "T", color: { theme: "accent1" } }] }])]);
    const xml = serializeSlideXml(slide, NO_RELS, "ppt/slides/slide1.xml");
    expect(xml).not.toContain("a:schemeClr");
    expect(xml).toContain('<a:srgbClr val="4472C4"/>');
  });
});
