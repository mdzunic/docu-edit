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
