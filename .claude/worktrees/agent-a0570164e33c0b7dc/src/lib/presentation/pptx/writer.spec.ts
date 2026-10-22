// The pptx deck-level write-back (SPEC-090): p:sldSz + firstSlideNum splices
// into presentation.xml, p:notesSz left verbatim, and docProps/core.xml rebuilt
// from the deck metadata.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { cmToEmu } from "../../core/units";
import { buildPptxFixtureBytes } from "../fixtures/build";
import { readPresentationBytes } from "../open";
import { applyDeckProperties, applyDeckSize, setFirstSlideNumber } from "../deckSetup";
import { writePptx } from "./writer";
import type { PresentationModel } from "../model";

const decoder = new TextDecoder("utf-8");
const MODIFIED = new Date("2026-03-01T12:00:00Z");

async function openDeck(): Promise<PresentationModel> {
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
}

async function writtenParts(model: PresentationModel): Promise<Map<string, Uint8Array>> {
  const { bytes } = await writePptx(model, { modified: MODIFIED });
  return unzip(bytes);
}

function partText(parts: Map<string, Uint8Array>, name: string): string {
  return decoder.decode(parts.get(name)!);
}

describe("pptx p:sldSz splice (SPEC-090 AC-15)", () => {
  it("rewrites cx/cy to the new size in whole EMU, notesSz untouched", async () => {
    const model = await openDeck();
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).toContain(`<p:sldSz cx="${cmToEmu(25.4)}" cy="${cmToEmu(19.05)}"/>`);
    // notesSz is sized independently — left byte-verbatim.
    expect(pres).toContain(`<p:notesSz cx="6858000" cy="9144000"/>`);
  });

  it("leaves presentation.xml's sldSz verbatim when the deck is unchanged", async () => {
    const model = await openDeck();
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).toContain(`<p:sldSz cx="12192000" cy="6858000"/>`);
  });
});

describe("pptx firstSlideNum (SPEC-090 AC-15)", () => {
  it("sets firstSlideNum on the p:presentation start tag when present", async () => {
    const model = await openDeck();
    setFirstSlideNumber(model, 7);
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).toMatch(/<p:presentation\b[^>]*\bfirstSlideNum="7"[^>]*>/);
  });

  it("removes firstSlideNum when the origin returns to 1", async () => {
    const model = await openDeck();
    setFirstSlideNumber(model, 7);
    setFirstSlideNumber(model, 1);
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).not.toContain("firstSlideNum");
  });
});

describe("pptx docProps/core.xml (SPEC-090 AC-20)", () => {
  it("rebuilds core.xml from the deck metadata on a deck change", async () => {
    const model = await openDeck();
    applyDeckProperties(model, { ...model.properties, author: "Grace Hopper", subject: "Q3" });
    const core = partText(await writtenParts(model), "docProps/core.xml");
    expect(core).toContain("Grace Hopper");
    expect(core).toContain("Q3");
  });

  it("leaves core.xml verbatim when the deck is unchanged", async () => {
    const model = await openDeck();
    const parts = await writtenParts(model);
    const source = await unzip(await buildPptxFixtureBytes());
    expect(partText(parts, "docProps/core.xml")).toBe(decoder.decode(source.get("docProps/core.xml")!));
  });
});

// --- SPEC-097: picture bullets, hyperlink and slide-jump relationships -------
import { readPresentationBytes as reopenBytes } from "../open";
import { addSlide, moveSlide } from "../slides";
import { markSlideEdited } from "../text";
import type { Slide, TextFrame } from "../model";

const PNG_2X2 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGMw7jhj3HGGAUIBACiuBh3+iAU4AAAAAElFTkSuQmCC";
const GIF_1X1 = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

function linkedTextFrame(id: string, paragraphs: TextFrame["paragraphs"], extra: Partial<TextFrame> = {}): TextFrame {
  return { id, kind: "text", xCm: 1, yCm: 1, widthCm: 10, heightCm: 4, paragraphs, ...extra };
}

/** A from-scratch three-slide deck: slide 1 carries a picture bullet, a run
 * link to slide 3, an external run link and a frame link. */
function scratchDeck(): PresentationModel {
  const slide1: Slide = {
    id: "s1",
    frames: [
      linkedTextFrame(
        "t1",
        [
          { runs: [{ text: "Pictured" }], bullet: "disc", bulletImageId: "ppt/media/image1.png" },
          {
            runs: [
              { text: "see " },
              { text: "results", linkSlide: 3 },
              { text: " and " },
              { text: "detail", linkHref: "https://example.com/detail?a=1&b=2", linkTooltip: "More" },
            ],
          },
        ],
        { linkHref: "https://example.com/" },
      ),
    ],
  };
  return {
    kind: "presentation",
    id: "d",
    title: "Links",
    format: "pptx",
    slideWidthCm: 25.4,
    slideHeightCm: 14.288,
    media: [{ mediaId: "ppt/media/image1.png", contentType: "image/png", data: PNG_2X2 }],
    slides: [slide1, { id: "s2", frames: [] }, { id: "s3", frames: [] }],
  };
}

describe("pptx SPEC-097 relationships — from scratch", () => {
  it("saves a picture-bullet deck with its media part, content-type Default row and image relationship", async () => {
    const parts = await writtenParts(scratchDeck());
    expect(parts.has("ppt/media/image1.png")).toBe(true);
    expect(partText(parts, "[Content_Types].xml")).toContain(`<Default Extension="png" ContentType="image/png"/>`);
    const rels = partText(parts, "ppt/slides/_rels/slide1.xml.rels");
    expect(rels).toContain(`<Relationship Id="rId2" Type="${IMAGE_REL_TYPE}" Target="../media/image1.png"/>`);
    expect(partText(parts, "ppt/slides/slide1.xml")).toContain(`<a:buBlip><a:blip r:embed="rId2"/></a:buBlip>`);
  });

  it("mints External hyperlink relationships and a slide relationship relative to the slide part", async () => {
    const parts = await writtenParts(scratchDeck());
    const rels = partText(parts, "ppt/slides/_rels/slide1.xml.rels");
    // Minted after the image, in first-use order per kind: hrefs, then slides.
    expect(rels).toContain(`<Relationship Id="rId3" Type="${HYPERLINK_REL_TYPE}" Target="https://example.com/" TargetMode="External"/>`);
    expect(rels).toContain(
      `<Relationship Id="rId4" Type="${HYPERLINK_REL_TYPE}" Target="https://example.com/detail?a=1&amp;b=2" TargetMode="External"/>`,
    );
    expect(rels).toContain(`<Relationship Id="rId5" Type="${SLIDE_REL_TYPE}" Target="slide3.xml"/>`);
    const slide = partText(parts, "ppt/slides/slide1.xml");
    expect(slide).toContain(`<p:cNvPr id="2" name="TextBox"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3"/></p:cNvPr>`);
    expect(slide).toContain(`r:id="rId5" action="ppaction://hlinksldjump"/>`);
    expect(slide).toContain(`r:id="rId4" tooltip="More"/>`);
  });

  it("reads every link and the picture bullet back", async () => {
    const { bytes } = await writePptx(scratchDeck(), { modified: MODIFIED });
    const reopened = await reopenBytes(bytes, "pptx", "Links");
    const frame = reopened.slides[0].frames[0] as TextFrame;
    expect(frame.linkHref).toBe("https://example.com/");
    expect(frame.paragraphs[0]).toMatchObject({ bullet: "disc", bulletImageId: "ppt/media/image1.png" });
    expect(frame.paragraphs[1].runs.map((run) => [run.text, run.linkSlide, run.linkHref, run.linkTooltip])).toEqual([
      ["see ", undefined, undefined, undefined],
      ["results", 3, undefined, undefined],
      [" and ", undefined, undefined, undefined],
      ["detail", undefined, "https://example.com/detail?a=1&b=2", "More"],
    ]);
  });
});

describe("pptx SPEC-097 relationships — retained package", () => {
  it("adds a picture bullet's part, Default row and relationship on a text edit, other parts verbatim", async () => {
    const model = await openDeck();
    const source = await unzip(await buildPptxFixtureBytes());
    model.media = [...(model.media ?? []), { mediaId: "ppt/media/image97.gif", contentType: "image/gif", data: GIF_1X1 }];
    const frame = model.slides[0].frames.find((f): f is TextFrame => f.kind === "text")!;
    frame.paragraphs[0] = { ...frame.paragraphs[0], bullet: "disc", bulletImageId: "ppt/media/image97.gif" };
    markSlideEdited(model.slides[0], "pptx");
    const parts = await writtenParts(model);
    const slidePart = model.slides[0].partName!;
    expect(parts.has("ppt/media/image97.gif")).toBe(true);
    expect(partText(parts, "[Content_Types].xml")).toContain(`<Default Extension="gif" ContentType="image/gif"/>`);
    const rels = partText(parts, slidePart.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels"));
    const rel = /<Relationship Id="(rId\d+)" Type="[^"]*\/image" Target="\.\.\/media\/image97\.gif"\/>/.exec(rels);
    expect(rel).not.toBeNull();
    expect(partText(parts, slidePart)).toContain(`<a:buBlip><a:blip r:embed="${rel![1]}"/></a:buBlip>`);
    // Slides the edit did not touch stay byte-verbatim.
    for (const slide of model.slides.slice(1)) {
      expect(partText(parts, slide.partName!)).toBe(decoder.decode(source.get(slide.partName!)!));
    }
  });

  it("reuses an existing hyperlink relationship and never deletes an unreferenced one", async () => {
    const model = await openDeck();
    const frame = model.slides[0].frames.find((f): f is TextFrame => f.kind === "text")!;
    frame.linkHref = "https://example.com/";
    markSlideEdited(model.slides[0], "pptx");
    const once = await writtenParts(model);
    const relsPath = model.slides[0].partName!.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels");
    const firstRels = partText(once, relsPath);
    expect(firstRels.match(/TargetMode="External"/g)).toHaveLength(1);
    // Reopen, drop the link, edit again: the now-unreferenced relationship stays.
    const reopened = await reopenBytes((await writePptx(model, { modified: MODIFIED })).bytes, "pptx", "Deck");
    const again = reopened.slides[0].frames.find((f): f is TextFrame => f.kind === "text")!;
    expect(again.linkHref).toBe("https://example.com/");
    delete again.linkHref;
    markSlideEdited(reopened.slides[0], "pptx");
    expect(partText(await writtenParts(reopened), relsPath)).toBe(firstRels);
  });

  it("resolves a slide jump to the part its target is saved under after a structural change", async () => {
    const model = await openDeck();
    const count = model.slides.length;
    // A new slide appended at the end, then moved to the front: it is saved
    // under a fresh part name, and slide 1 (now the old first slide, at
    // index 1) links to it.
    const added = addSlide(model, count - 1);
    moveSlide(model, added, 0);
    const linking = model.slides[1];
    const frame = linking.frames.find((f): f is TextFrame => f.kind === "text")!;
    frame.paragraphs[0].runs = [{ ...frame.paragraphs[0].runs[0], linkSlide: 1 }];
    markSlideEdited(linking, "pptx");
    const parts = await writtenParts(model);
    const pres = partText(parts, "ppt/presentation.xml");
    const presRels = partText(parts, "ppt/_rels/presentation.xml.rels");
    const firstRelId = /<p:sldId id="\d+" r:id="(rId\d+)"\/>/.exec(pres)![1];
    const firstTarget = new RegExp(`<Relationship Id="${firstRelId}" Type="[^"]*/slide" Target="([^"]+)"/>`).exec(presRels)![1];
    const linkingPart = linking.partName!;
    const slideRels = partText(parts, linkingPart.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels"));
    const jump = new RegExp(`<Relationship Id="(rId\\d+)" Type="${SLIDE_REL_TYPE}" Target="([^"]+)"/>`).exec(slideRels);
    expect(jump).not.toBeNull();
    expect(`ppt/slides/${jump![2]}`).toBe(`ppt/${firstTarget}`);
    expect(partText(parts, linkingPart)).toContain(`r:id="${jump![1]}" action="ppaction://hlinksldjump"/>`);
    // And it reads back as a jump to slide 1.
    const reopened = await reopenBytes((await writePptx(model, { modified: MODIFIED })).bytes, "pptx", "Deck");
    const back = reopened.slides[1].frames.find((f): f is TextFrame => f.kind === "text")!;
    expect(back.paragraphs[0].runs[0].linkSlide).toBe(1);
  });
});
