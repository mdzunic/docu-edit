// SPEC-097 pptx wire mapping: every a:pPr / a:bodyPr / a:rPr@spc /
// a:hlinkClick form the reader models, what it keeps in the preservation
// slots, the bytes the writer emits (attribute and ECMA-376 child order), and
// the read → write → read → write stability of both slots.
import { describe, expect, it } from "vitest";
import type { Relationship } from "../../core/xml";
import type { ImageFrame, ShapeFrame, Slide, SlideParagraph, TextFrame } from "../model";
import { parseShapeTreeChildForTest } from "./reader";
import { serializeSlideXml } from "./serialize";
import { bodyPrXml, collectRelNeeds, parseSlot, pPrXml, type RelLookup } from "./textDepthWire";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"`;
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE1 = "ppt/slides/slide1.xml";

const RELS = new Map<string, Relationship>([
  ["rId2", { id: "rId2", type: IMAGE_REL, target: "../media/image1.png" }],
  ["rId3", { id: "rId3", type: HYPERLINK_REL, target: "https://example.com/a", targetMode: "External" }],
  ["rId4", { id: "rId4", type: HYPERLINK_REL, target: "javascript:alert(1)", targetMode: "External" }],
  ["rId5", { id: "rId5", type: SLIDE_REL, target: "slide3.xml" }],
]);
const SLIDE_NUMBERS = new Map([
  ["ppt/slides/slide1.xml", 1],
  ["ppt/slides/slide2.xml", 2],
  ["ppt/slides/slide3.xml", 3],
]);
const CONTEXT = { rels: RELS, slidePartPath: SLIDE1, mediaIds: ["ppt/media/image1.png"], slideNumberOf: SLIDE_NUMBERS };
const LOOKUP: RelLookup = {
  rels: RELS,
  slidePartPath: SLIDE1,
  slideParts: ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml", "ppt/slides/slide3.xml"],
};

/** A text box `p:sp` with the given txBody content, read through the reader. */
function readTextBox(txBodyInner: string, opts: { bodyPr?: string; cNvPrInner?: string; ph?: string } = {}): TextFrame {
  const cNvPr = opts.cNvPrInner ? `<p:cNvPr id="2" name="T">${opts.cNvPrInner}</p:cNvPr>` : `<p:cNvPr id="2" name="T"/>`;
  const xml =
    `<p:sp ${NS}><p:nvSpPr>${cNvPr}<p:cNvSpPr/><p:nvPr>${opts.ph ?? ""}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3600000" cy="1800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody>${opts.bodyPr ?? "<a:bodyPr/>"}<a:lstStyle/>${txBodyInner}</p:txBody></p:sp>`;
  const frame = parseShapeTreeChildForTest(xml, CONTEXT);
  if (!frame || frame.kind !== "text") throw new Error("expected a text frame");
  return frame;
}

/** The first paragraph of a text box whose single `a:p` carries `pPr`. */
function readPara(pPr: string, opts: { ph?: string } = {}): SlideParagraph {
  return readTextBox(`<a:p>${pPr}<a:r><a:t>x</a:t></a:r></a:p>`, opts).paragraphs[0];
}

/** The paragraph fields beyond the runs. */
function fields(para: SlideParagraph): Omit<SlideParagraph, "runs"> {
  const { runs: _runs, ...rest } = para;
  return rest;
}

function slideOf(frames: Slide["frames"]): Slide {
  return { id: "s", frames };
}

function textBox(paragraphs: SlideParagraph[], extra: Partial<TextFrame> = {}): TextFrame {
  return { id: "t", kind: "text", xCm: 0, yCm: 0, widthCm: 10, heightCm: 5, paragraphs, ...extra };
}

describe("a:pPr read — indents", () => {
  it("drops a list item's marL/indent when they equal the effective default (EMU precision)", () => {
    expect(fields(readPara(`<a:pPr marL="457200" indent="-228600"><a:buChar char="•"/></a:pPr>`))).toEqual({ bullet: "disc" });
    expect(fields(readPara(`<a:pPr marL="914400" indent="-228600" lvl="1"><a:buChar char="•"/></a:pPr>`))).toEqual({
      bullet: "disc",
      level: 1,
    });
  });

  it("models a non-default list indent and hanging in cm (r3)", () => {
    expect(fields(readPara(`<a:pPr marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr>`))).toEqual({
      bullet: "disc",
      indentCm: 0.953,
      hangingCm: 0.953,
    });
  });

  it("keeps an explicit marL/indent on a plain level-0 paragraph (the writer emits it only when explicit)", () => {
    expect(fields(readPara(`<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>`))).toEqual({ indentCm: 0, hangingCm: 0 });
  });

  it("drops a plain indented level's default marL and zero indent", () => {
    expect(fields(readPara(`<a:pPr marL="457200" indent="0" lvl="1"><a:buNone/></a:pPr>`))).toEqual({ level: 1 });
  });
});

describe("a:pPr read — spacing", () => {
  it("reads a:lnSpc spcPct as a multiple and spcPts as an exact point value", () => {
    expect(fields(readPara(`<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:buNone/></a:pPr>`))).toEqual({ lineSpacing: 1.5 });
    expect(fields(readPara(`<a:pPr><a:lnSpc><a:spcPct val="150%"/></a:lnSpc><a:buNone/></a:pPr>`))).toEqual({ lineSpacing: 1.5 });
    expect(fields(readPara(`<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:buNone/></a:pPr>`))).toEqual({
      lineSpacingPt: 24,
      lineSpacingRule: "exact",
    });
  });

  it("reads a:spcBef/a:spcAft spcPts as points", () => {
    expect(
      fields(readPara(`<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="1200"/></a:spcAft><a:buNone/></a:pPr>`)),
    ).toEqual({ spaceBeforePt: 6, spaceAfterPt: 12 });
  });

  it("keeps an unmodelled spacing form (spcPct in spcBef) and an explicit default (100%, 0pt) in pptxPPr", () => {
    const para = readPara(
      `<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="20000"/></a:spcBef>` +
        `<a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:pPr>`,
    );
    expect(fields(para)).toEqual({
      pptxPPr:
        `<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="20000"/></a:spcBef>` +
        `<a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>`,
    });
  });
});

describe("a:pPr read — bullets", () => {
  it("reads buNone as no bullet, even in a body placeholder", () => {
    expect(fields(readPara(`<a:pPr><a:buNone/></a:pPr>`, { ph: `<p:ph type="body" idx="1"/>` }))).toEqual({});
  });

  it("keeps the body-role default disc when no bullet element is present", () => {
    expect(readPara(`<a:pPr/>`, { ph: `<p:ph type="body" idx="1"/>` }).bullet).toBe("disc");
  });

  it("maps buChar •/○/▪ to disc/circle/square and any other grapheme to bulletChar", () => {
    expect(fields(readPara(`<a:pPr><a:buChar char="•"/></a:pPr>`))).toEqual({ bullet: "disc" });
    expect(fields(readPara(`<a:pPr><a:buChar char="○"/></a:pPr>`))).toEqual({ bullet: "disc", bulletStyle: "circle" });
    expect(fields(readPara(`<a:pPr><a:buChar char="▪"/></a:pPr>`))).toEqual({ bullet: "disc", bulletStyle: "square" });
    expect(fields(readPara(`<a:pPr><a:buChar char="➢"/></a:pPr>`))).toEqual({ bullet: "disc", bulletChar: "➢" });
  });

  it("keeps a recognised glyph the writer would not reproduce as its exact bulletChar beside the style", () => {
    expect(fields(readPara(`<a:pPr><a:buChar char="◦"/></a:pPr>`))).toEqual({ bullet: "disc", bulletStyle: "circle", bulletChar: "◦" });
    expect(fields(readPara(`<a:pPr><a:buChar char="·"/></a:pPr>`))).toEqual({ bullet: "disc", bulletChar: "·" });
  });

  it("preserves a multi-grapheme buChar verbatim while still modelling a bullet", () => {
    expect(fields(readPara(`<a:pPr><a:buChar char="-&gt;"/></a:pPr>`))).toEqual({
      bullet: "disc",
      pptxPPr: `<a:pPr><a:buChar char="-&gt;"/></a:pPr>`,
    });
  });

  it("maps every modelled buAutoNum type and startAt (only when ≠ 1)", () => {
    const style = (type: string): Omit<SlideParagraph, "runs"> => fields(readPara(`<a:pPr><a:buAutoNum type="${type}"/></a:pPr>`));
    expect(style("arabicPeriod")).toEqual({ bullet: "number" });
    expect(style("alphaLcPeriod")).toEqual({ bullet: "number", bulletStyle: "lowerAlpha" });
    expect(style("alphaUcPeriod")).toEqual({ bullet: "number", bulletStyle: "upperAlpha" });
    expect(style("romanLcPeriod")).toEqual({ bullet: "number", bulletStyle: "lowerRoman" });
    expect(style("romanUcPeriod")).toEqual({ bullet: "number", bulletStyle: "upperRoman" });
    expect(fields(readPara(`<a:pPr><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>`))).toEqual({ bullet: "number", startAt: 3 });
    expect(fields(readPara(`<a:pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr>`))).toEqual({ bullet: "number" });
  });

  it("reads an unrecognised buAutoNum scheme as number (style absent) and keeps the element", () => {
    expect(fields(readPara(`<a:pPr><a:buAutoNum type="arabicParenR" startAt="2"/></a:pPr>`))).toEqual({
      bullet: "number",
      startAt: 2,
      pptxPPr: `<a:pPr><a:buAutoNum type="arabicParenR" startAt="2"/></a:pPr>`,
    });
  });

  it("reads a buBlip picture bullet to its media part", () => {
    expect(fields(readPara(`<a:pPr><a:buBlip><a:blip r:embed="rId2"/></a:buBlip></a:pPr>`))).toEqual({
      bullet: "disc",
      bulletImageId: "ppt/media/image1.png",
    });
  });

  it("keeps a buBlip that does not resolve verbatim", () => {
    const para = readPara(`<a:pPr><a:buBlip><a:blip r:embed="rId99"/></a:buBlip></a:pPr>`);
    expect(para.bullet).toBe("disc");
    expect(para.bulletImageId).toBeUndefined();
    // The platform serialiser declares the relationships namespace on the
    // fragment itself (jsdom under a generated prefix); either way the
    // embed survives and re-reads to the same slot.
    expect(para.pptxPPr).toMatch(/^<a:pPr><a:buBlip><a:blip xmlns:(\w+)="[^"]+" \1:embed="rId99"\/><\/a:buBlip><\/a:pPr>$/);
    const xml = serializeSlideXml(slideOf([textBox([para])]), RELS, SLIDE1);
    const sp = /<p:sp>[\s\S]*<\/p:sp>/.exec(xml)![0].replace("<p:sp>", `<p:sp ${NS}>`);
    expect((parseShapeTreeChildForTest(sp, CONTEXT) as TextFrame).paragraphs[0].pptxPPr).toBe(para.pptxPPr);
  });
});

describe("a:pPr read — the preservation slot", () => {
  it("keeps unmodelled attributes and children (schema-slot order), never the modelled ones", () => {
    const para = readPara(
      `<a:pPr marL="342900" indent="-342900" algn="ctr" rtl="0" fontAlgn="base">` +
        `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>` +
        `<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buSzPct val="80000"/><a:buFont typeface="Arial"/>` +
        `<a:buChar char="•"/><a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst><a:defRPr sz="1800"/></a:pPr>`,
    );
    expect(fields(para)).toEqual({
      align: "center",
      bullet: "disc",
      indentCm: 0.953,
      hangingCm: 0.953,
      lineSpacing: 0.9,
      pptxPPr:
        `<a:pPr rtl="0" fontAlgn="base"><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buSzPct val="80000"/>` +
        `<a:buFont typeface="Arial"/><a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst><a:defRPr sz="1800"/></a:pPr>`,
    });
  });

  it("leaves pptxPPr absent when everything is modelled", () => {
    expect(readPara(`<a:pPr lvl="1" algn="r"><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buAutoNum type="alphaLcPeriod"/></a:pPr>`).pptxPPr).toBeUndefined();
  });

  it("parseSlot splits a slot back into attributes and top-level children", () => {
    const parsed = parseSlot(`<a:pPr rtl="0"><a:buFont typeface="A"/><a:tabLst><a:tab pos="1"/></a:tabLst></a:pPr>`);
    expect(parsed.attrs).toEqual([["rtl", ` rtl="0"`]]);
    expect(parsed.children).toEqual([
      { local: "buFont", xml: `<a:buFont typeface="A"/>` },
      { local: "tabLst", xml: `<a:tabLst><a:tab pos="1"/></a:tabLst>` },
    ]);
    expect(parseSlot(`<a:bodyPr vert="vert270"/>`)).toEqual({ attrs: [["vert", ` vert="vert270"`]], children: [] });
    expect(parseSlot(undefined)).toEqual({ attrs: [], children: [] });
  });
});

describe("a:pPr write — attribute and child order", () => {
  it("writes marL/indent for a list item, then lvl and algn", () => {
    expect(pPrXml({ runs: [], bullet: "disc", level: 2, align: "right" }, LOOKUP)).toBe(
      `<a:pPr marL="1371600" indent="-228600" lvl="2" algn="r"><a:buChar char="•"/></a:pPr>`,
    );
  });

  it("keeps a plain level-0 paragraph free of marL unless an indent is explicit", () => {
    expect(pPrXml({ runs: [] }, LOOKUP)).toBe(`<a:pPr><a:buNone/></a:pPr>`);
    expect(pPrXml({ runs: [], indentCm: 0, hangingCm: 0 }, LOOKUP)).toBe(`<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>`);
    expect(pPrXml({ runs: [], level: 1 }, LOOKUP)).toBe(`<a:pPr marL="457200" indent="0" lvl="1"><a:buNone/></a:pPr>`);
  });

  it("writes each modelled bullet form", () => {
    const bullet = (para: Partial<SlideParagraph>): string => pPrXml({ runs: [], ...para }, LOOKUP).replace(/^<a:pPr[^>]*>|<\/a:pPr>$/g, "");
    expect(bullet({ bullet: "disc", bulletStyle: "circle" })).toBe(`<a:buChar char="○"/>`);
    expect(bullet({ bullet: "disc", bulletStyle: "square" })).toBe(`<a:buChar char="▪"/>`);
    expect(bullet({ bullet: "disc", bulletStyle: "circle", bulletChar: "➢" })).toBe(`<a:buChar char="➢"/>`);
    expect(bullet({ bullet: "number" })).toBe(`<a:buAutoNum type="arabicPeriod"/>`);
    expect(bullet({ bullet: "number", bulletStyle: "upperRoman", startAt: 4 })).toBe(`<a:buAutoNum type="romanUcPeriod" startAt="4"/>`);
    expect(bullet({ bullet: "number", bulletStyle: "lowerAlpha", startAt: 1 })).toBe(`<a:buAutoNum type="alphaLcPeriod"/>`);
    expect(bullet({ bullet: "disc", bulletImageId: "ppt/media/image1.png" })).toBe(`<a:buBlip><a:blip r:embed="rId2"/></a:buBlip>`);
  });

  it("writes spacing in CT_TextParagraphProperties order ahead of the bullet", () => {
    expect(
      pPrXml({ runs: [], bullet: "number", lineSpacing: 1.5, spaceBeforePt: 6, spaceAfterPt: 3 }, LOOKUP),
    ).toBe(
      `<a:pPr marL="457200" indent="-228600"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>` +
        `<a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="300"/></a:spcAft>` +
        `<a:buAutoNum type="arabicPeriod"/></a:pPr>`,
    );
    expect(pPrXml({ runs: [], lineSpacing: 1.5, lineSpacingPt: 24, lineSpacingRule: "exact" }, LOOKUP)).toBe(
      `<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:buNone/></a:pPr>`,
    );
  });

  it("re-emits each preserved child at its schema slot, a preserved spacing/bullet replacing the modelled one", () => {
    const para: SlideParagraph = {
      runs: [],
      bullet: "number",
      spaceAfterPt: 6,
      pptxPPr:
        `<a:pPr rtl="0"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/>` +
        `<a:buAutoNum type="arabicParenR"/><a:extLst><a:ext uri="{X}"/></a:extLst></a:pPr>`,
    };
    expect(pPrXml(para, LOOKUP)).toBe(
      `<a:pPr marL="457200" indent="-228600" rtl="0"><a:spcBef><a:spcPct val="20000"/></a:spcBef>` +
        `<a:spcAft><a:spcPts val="600"/></a:spcAft><a:buFont typeface="Arial"/><a:buAutoNum type="arabicParenR"/>` +
        `<a:extLst><a:ext uri="{X}"/></a:extLst></a:pPr>`,
    );
  });

  it("never emits a preserved attribute beside the modelled one of the same name", () => {
    expect(pPrXml({ runs: [], bullet: "disc", pptxPPr: `<a:pPr marL="1" lvl="3" rtl="0"/>` }, LOOKUP)).toBe(
      `<a:pPr marL="457200" indent="-228600" lvl="3" rtl="0"><a:buChar char="•"/></a:pPr>`,
    );
  });
});

describe("a:rPr — letter spacing and run links", () => {
  const run = (rPr: string): TextFrame["paragraphs"][number]["runs"][number] =>
    readTextBox(`<a:p><a:r>${rPr}<a:t>go</a:t></a:r></a:p>`).paragraphs[0].runs[0];

  it("reads spc (100ths of a point) into letterSpacingPt on the 0.1 grid, dropping 0", () => {
    expect(run(`<a:rPr lang="en-US" spc="50"/>`).letterSpacingPt).toBe(0.5);
    expect(run(`<a:rPr lang="en-US" spc="-120"/>`).letterSpacingPt).toBe(-1.2);
    expect(run(`<a:rPr lang="en-US" spc="5000"/>`).letterSpacingPt).toBe(10);
    expect(run(`<a:rPr lang="en-US" spc="0"/>`).letterSpacingPt).toBeUndefined();
  });

  it("reads an external hyperlink (sanitised) and its tooltip", () => {
    const linked = run(`<a:rPr lang="en-US"><a:hlinkClick r:id="rId3" tooltip="Go there"/></a:rPr>`);
    expect(linked.linkHref).toBe("https://example.com/a");
    expect(linked.linkTooltip).toBe("Go there");
    expect(linked.linkSlide).toBeUndefined();
  });

  it("drops a dangerous href", () => {
    const linked = run(`<a:rPr lang="en-US"><a:hlinkClick r:id="rId4"/></a:rPr>`);
    expect(linked.linkHref).toBeUndefined();
  });

  it("reads a slide jump to the target's 1-based deck position", () => {
    const linked = run(`<a:rPr lang="en-US"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:hlinkClick r:id="rId5" action="ppaction://hlinksldjump"/></a:rPr>`);
    expect(linked.linkSlide).toBe(3);
    expect(linked.linkHref).toBeUndefined();
  });

  it("ignores actions outside the modelled pair (custom shows, next-slide jumps)", () => {
    expect(run(`<a:rPr lang="en-US"><a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/></a:rPr>`).linkSlide).toBeUndefined();
    expect(run(`<a:rPr lang="en-US"><a:hlinkClick r:id="rId5"/></a:rPr>`).linkSlide).toBeUndefined();
  });

  it("writes spc and the a:hlinkClick after a:latin (CT_TextCharacterProperties order)", () => {
    const xml = serializeSlideXml(
      slideOf([textBox([{ runs: [{ text: "go", letterSpacingPt: 0.5, fontFamily: "Arial", linkHref: "https://example.com/a", linkTooltip: "Go" }] }])]),
      RELS,
      SLIDE1,
      [],
      undefined,
      LOOKUP.slideParts,
    );
    expect(xml).toContain(
      `<a:rPr lang="en-US" sz="1800" spc="50"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>` +
        `<a:hlinkClick xmlns:r="${R_NS}" r:id="rId3" tooltip="Go"/></a:rPr>`,
    );
  });

  it("writes a slide-jump run link against the target part's slide relationship", () => {
    const xml = serializeSlideXml(slideOf([textBox([{ runs: [{ text: "go", linkSlide: 3 }] }])]), RELS, SLIDE1, [], undefined, LOOKUP.slideParts);
    expect(xml).toContain(`<a:hlinkClick xmlns:r="${R_NS}" r:id="rId5" action="ppaction://hlinksldjump"/>`);
  });
});

describe("frame links — p:cNvPr/a:hlinkClick", () => {
  it("reads a text frame's and a shape's own link", () => {
    expect(readTextBox(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`, { cNvPrInner: `<a:hlinkClick r:id="rId3"/>` }).linkHref).toBe("https://example.com/a");
    const shape = parseShapeTreeChildForTest(
      `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="3" name="S"><a:hlinkClick r:id="rId5" action="ppaction://hlinksldjump" tooltip="Results"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:sp>`,
      CONTEXT,
    ) as ShapeFrame;
    expect([shape.kind, shape.linkSlide, shape.linkTooltip]).toEqual(["shape", 3, "Results"]);
  });

  it("reads an image frame's link from p:nvPicPr/p:cNvPr", () => {
    const pic = parseShapeTreeChildForTest(
      `<p:pic ${NS}><p:nvPicPr><p:cNvPr id="4" name="P"><a:hlinkClick r:id="rId3"/></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rId2"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>`,
      CONTEXT,
    ) as ImageFrame;
    expect([pic.kind, pic.linkHref]).toEqual(["image", "https://example.com/a"]);
  });

  it("writes the link as p:cNvPr's first child on text, shape and image frames", () => {
    const frames: Slide["frames"] = [
      textBox([{ runs: [{ text: "t" }] }], { linkHref: "https://example.com/a" }),
      { id: "s", kind: "shape", geometry: "ellipse", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, linkSlide: 3, linkTooltip: "R" },
      { id: "i", kind: "image", mediaId: "ppt/media/image1.png", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, linkHref: "https://example.com/a" },
    ];
    const xml = serializeSlideXml(slideOf(frames), RELS, SLIDE1, [], undefined, LOOKUP.slideParts);
    expect(xml).toContain(`<p:cNvPr id="2" name="TextBox"><a:hlinkClick xmlns:r="${R_NS}" r:id="rId3"/></p:cNvPr>`);
    expect(xml).toContain(`<p:cNvPr id="3" name="Shape 3"><a:hlinkClick xmlns:r="${R_NS}" r:id="rId5" action="ppaction://hlinksldjump" tooltip="R"/></p:cNvPr>`);
    expect(xml).toContain(`<p:cNvPr id="4" name="Picture 4"><a:hlinkClick xmlns:r="${R_NS}" r:id="rId3"/></p:cNvPr>`);
  });

  it("collects relationship needs only from frames the serialiser rebuilds", () => {
    const needs = collectRelNeeds([
      textBox([{ runs: [{ text: "a", linkHref: "https://x.test/" }, { text: "b", linkSlide: 2 }], bullet: "disc", bulletImageId: "ppt/media/b.png" }], {
        linkHref: "https://x.test/",
      }),
      {
        id: "g",
        kind: "group",
        xCm: 0,
        yCm: 0,
        widthCm: 1,
        heightCm: 1,
        xml: "<p:grpSp/>",
        children: [textBox([{ runs: [{ text: "c", linkHref: "https://verbatim.test/" }] }])],
      },
    ]);
    expect(needs).toEqual({ images: ["ppt/media/b.png"], hrefs: ["https://x.test/"], slides: [2] });
  });
});

describe("a:bodyPr read", () => {
  const body = (bodyPr: string, kind: "text" | "shape" = "text"): Pick<TextFrame, "body" | "pptxBodyPr"> => {
    if (kind === "text") {
      const frame = readTextBox(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`, { bodyPr });
      return { body: frame.body, pptxBodyPr: frame.pptxBodyPr };
    }
    const shape = parseShapeTreeChildForTest(
      `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="3" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${bodyPr}<a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>`,
      CONTEXT,
    ) as ShapeFrame;
    return { body: shape.body, pptxBodyPr: shape.pptxBodyPr };
  };

  it("reads insets (all four, missing ones defaulted) when any differs", () => {
    expect(body(`<a:bodyPr lIns="180000" tIns="45720"/>`)).toEqual({
      body: { insetsCm: { left: 0.5, top: 0.127, right: 0.254, bottom: 0.127 } },
      pptxBodyPr: undefined,
    });
  });

  it("reads anchor, columns, wrap and the autofit child", () => {
    expect(body(`<a:bodyPr anchor="ctr" numCol="2" spcCol="360000" wrap="none"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>`)).toEqual({
      body: { anchor: "middle", columns: { count: 2, gapCm: 1 }, wrap: false, autofit: "shrink", fontScalePct: 62.5, lineScalePct: 80 },
      pptxBodyPr: undefined,
    });
    expect(body(`<a:bodyPr><a:normAutofit/></a:bodyPr>`)).toEqual({ body: { autofit: "shrink" }, pptxBodyPr: undefined });
    expect(body(`<a:bodyPr anchor="b"><a:spAutoFit/></a:bodyPr>`)).toEqual({ body: { anchor: "bottom", autofit: "resizeShape" }, pptxBodyPr: undefined });
  });

  it("stores a shape's anchor only when it is not middle", () => {
    expect(body(`<a:bodyPr anchor="t"/>`, "shape")).toEqual({ body: { anchor: "top" }, pptxBodyPr: undefined });
    expect(body(`<a:bodyPr anchor="ctr"/>`, "shape")).toEqual({ body: undefined, pptxBodyPr: `<a:bodyPr anchor="ctr"/>` });
  });

  it("keeps explicit defaults and the out-of-subset rest verbatim, modelling what it can", () => {
    expect(
      body(
        `<a:bodyPr vert="vert270" wrap="square" lIns="91440" rtlCol="0" anchor="ctr">` +
          `<a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp><a:noAutofit/><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d></a:bodyPr>`,
      ),
    ).toEqual({
      body: { anchor: "middle" },
      pptxBodyPr:
        `<a:bodyPr vert="vert270" wrap="square" lIns="91440" rtlCol="0"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>` +
        `<a:noAutofit/><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d></a:bodyPr>`,
    });
  });
});

describe("a:bodyPr write", () => {
  it("stays the bare element when the frame has no body properties", () => {
    expect(bodyPrXml(undefined, undefined)).toBe(`<a:bodyPr/>`);
  });

  it("writes insets, columns, anchor and wrap, then the one autofit child", () => {
    expect(
      bodyPrXml(
        {
          insetsCm: { left: 0.5, top: 0.3, right: 0.5, bottom: 0.3 },
          columns: { count: 2, gapCm: 1 },
          anchor: "middle",
          wrap: false,
          autofit: "shrink",
          fontScalePct: 62.5,
          lineScalePct: 90,
        },
        undefined,
      ),
    ).toBe(
      `<a:bodyPr lIns="180000" tIns="108000" rIns="180000" bIns="108000" numCol="2" spcCol="360000" anchor="ctr" wrap="none">` +
        `<a:normAutofit fontScale="62500" lnSpcReduction="10000"/></a:bodyPr>`,
    );
    expect(bodyPrXml({ autofit: "shrink" }, undefined)).toBe(`<a:bodyPr><a:normAutofit/></a:bodyPr>`);
    expect(bodyPrXml({ autofit: "resizeShape" }, undefined)).toBe(`<a:bodyPr><a:spAutoFit/></a:bodyPr>`);
    expect(bodyPrXml({ anchor: "bottom" }, undefined)).toBe(`<a:bodyPr anchor="b"/>`);
  });

  it("re-emits the preserved warp before the autofit child and scene3d/extLst after it", () => {
    expect(
      bodyPrXml(
        { autofit: "resizeShape", insetsCm: { left: 0, top: 0, right: 0, bottom: 0 } },
        `<a:bodyPr vert="vert270" lIns="91440"><a:prstTxWarp prst="textWave1"/><a:noAutofit/><a:scene3d/><a:extLst/></a:bodyPr>`,
      ),
    ).toBe(
      `<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" vert="vert270"><a:prstTxWarp prst="textWave1"/><a:spAutoFit/><a:scene3d/><a:extLst/></a:bodyPr>`,
    );
    expect(bodyPrXml(undefined, `<a:bodyPr vert="vert270"><a:noAutofit/></a:bodyPr>`)).toBe(`<a:bodyPr vert="vert270"><a:noAutofit/></a:bodyPr>`);
  });

  it("writes a shape's body even when it carries no text", () => {
    const xml = serializeSlideXml(
      slideOf([{ id: "s", kind: "shape", geometry: "rect", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, body: { anchor: "top" } }]),
      RELS,
      SLIDE1,
    );
    expect(xml).toContain(`<p:txBody><a:bodyPr anchor="t"/><a:lstStyle/><a:p/></p:txBody>`);
  });
});

describe("byte stability — read → write → read → write", () => {
  it("reproduces both slots and the slide part exactly", () => {
    const source = readTextBox(
      `<a:p><a:pPr marL="342900" indent="-342900" rtl="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>` +
        `<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buFont typeface="Wingdings" pitchFamily="2" charset="2"/><a:buChar char="§"/>` +
        `<a:tabLst/><a:extLst><a:ext uri="{35155182-B16C-46BC-9424-99874614C6A1}"><a14:x xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" v="1"/></a:ext></a:extLst></a:pPr>` +
        `<a:r><a:rPr lang="en-US" spc="30"/><a:t>one</a:t></a:r></a:p>` +
        `<a:p><a:pPr lvl="1"><a:buAutoNum type="thaiAlphaPeriod" startAt="4"/></a:pPr><a:r><a:t>two</a:t></a:r></a:p>`,
      { bodyPr: `<a:bodyPr vert="vert270" wrap="square" anchor="b" numCol="2" spcCol="91440"><a:normAutofit fontScale="85000"/><a:sp3d/></a:bodyPr>` },
    );
    const write = (frame: TextFrame): string => serializeSlideXml(slideOf([frame]), RELS, SLIDE1, [], undefined, LOOKUP.slideParts);
    const first = write(source);
    const spXml = /<p:sp>[\s\S]*<\/p:sp>/.exec(first)![0].replace("<p:sp>", `<p:sp ${NS}>`);
    const reread = parseShapeTreeChildForTest(spXml, CONTEXT) as TextFrame;
    expect(reread.paragraphs).toEqual(source.paragraphs);
    expect(reread.body).toEqual(source.body);
    expect(reread.pptxBodyPr).toBe(source.pptxBodyPr);
    expect(write(reread)).toBe(first);
    expect(source.paragraphs[0].bulletChar).toBe("§");
    expect(source.paragraphs[1].pptxPPr).toBe(`<a:pPr><a:buAutoNum type="thaiAlphaPeriod" startAt="4"/></a:pPr>`);
  });
});
