// SPEC-095: the pptx media wire, in SPEC-086's golden shape — golden write
// equality, a "" (no-media) case, golden → model read-back, a
// PowerPoint-flavoured tolerant variant, and preserve-don't-parse for a
// non-media p:pic.
import { describe, it, expect } from "vitest";
import type { Relationship } from "../../core/xml";
import type { MediaFrame, Slide, SlideFrame } from "../model";
import { serializeSlideXml } from "./serialize";
import { parseMediaPicForTest } from "./reader";
import {
  GOLDEN_AUDIO_RELS,
  GOLDEN_MEDIA_AUDIO,
  GOLDEN_MEDIA_PARTS,
  GOLDEN_MEDIA_VIDEO,
  GOLDEN_VIDEO_RELS,
  MEDIA_AUDIO_PIC_GOLDEN,
  MEDIA_VIDEO_PIC_GOLDEN,
  MEDIA_VIDEO_PIC_VARIANT,
  NON_MEDIA_PIC,
} from "./mediaGolden";

function mediaSlide(frame: Omit<MediaFrame, "id">): Slide {
  return { id: "s", frames: [{ id: "m1", ...frame }] };
}

function relsMap(rels: Relationship[]): Map<string, Relationship> {
  return new Map(rels.map((rel) => [rel.id, rel]));
}

function modelled(frame: SlideFrame): Omit<MediaFrame, "id"> {
  expect(frame.kind).toBe("media");
  const media = frame as MediaFrame;
  const { id: _id, sourceXml: _sourceXml, ...rest } = media;
  return rest;
}

describe("pptx media serialiser (SPEC-095)", () => {
  it("serialises the video golden model to the golden p:pic", () => {
    const xml = serializeSlideXml(mediaSlide(GOLDEN_MEDIA_VIDEO), relsMap(GOLDEN_VIDEO_RELS), "ppt/slides/slide1.xml");
    expect(xml).toContain(MEDIA_VIDEO_PIC_GOLDEN);
  });

  it("serialises the audio golden model to the golden p:pic", () => {
    const xml = serializeSlideXml(mediaSlide(GOLDEN_MEDIA_AUDIO), relsMap(GOLDEN_AUDIO_RELS), "ppt/slides/slide1.xml");
    expect(xml).toContain(MEDIA_AUDIO_PIC_GOLDEN);
  });

  it("re-emits an untouched clip's verbatim sourceXml, not a rebuild", () => {
    const frame: MediaFrame = { id: "m1", ...GOLDEN_MEDIA_VIDEO, sourceXml: "<p:pic>VERBATIM</p:pic>" };
    const xml = serializeSlideXml({ id: "s", frames: [frame] }, relsMap(GOLDEN_VIDEO_RELS), "ppt/slides/slide1.xml");
    expect(xml).toContain("<p:pic>VERBATIM</p:pic>");
    expect(xml).not.toContain("Video 2");
  });
});

describe("pptx media reader (SPEC-095)", () => {
  it("reads the video golden back to the model", () => {
    const frame = parseMediaPicForTest(MEDIA_VIDEO_PIC_GOLDEN, GOLDEN_VIDEO_RELS, GOLDEN_MEDIA_PARTS);
    expect(modelled(frame)).toEqual(GOLDEN_MEDIA_VIDEO);
  });

  it("reads the audio golden back to the model", () => {
    const frame = parseMediaPicForTest(MEDIA_AUDIO_PIC_GOLDEN, GOLDEN_AUDIO_RELS, GOLDEN_MEDIA_PARTS);
    expect(modelled(frame)).toEqual(GOLDEN_MEDIA_AUDIO);
  });

  it("keeps the verbatim source on a read clip for byte-exact re-emit", () => {
    const frame = parseMediaPicForTest(MEDIA_VIDEO_PIC_GOLDEN, GOLDEN_VIDEO_RELS, GOLDEN_MEDIA_PARTS) as MediaFrame;
    expect(frame.sourceXml).toContain("<p:pic");
    expect(frame.sourceXml).toContain("p14:media");
  });

  it("tolerates a PowerPoint-flavoured variant (p14:media only, reordered attrs)", () => {
    const frame = parseMediaPicForTest(MEDIA_VIDEO_PIC_VARIANT, GOLDEN_VIDEO_RELS, GOLDEN_MEDIA_PARTS) as MediaFrame;
    expect(frame.kind).toBe("media");
    expect(frame.media).toBe("video");
    expect(frame.mediaId).toBe("ppt/media/media1.mp4");
    expect(frame.posterMediaId).toBe("ppt/media/image1.png");
    expect(frame.contentType).toBe("video/mp4");
  });

  it("models a linked clip (rel target has no part) as linkTarget without mediaId", () => {
    const linkedPic =
      `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Video 2"/><p:cNvPicPr/><p:nvPr>` +
      `<a:videoFile r:link="rId9"/></p:nvPr></p:nvPicPr>` +
      `<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="720000" cy="720000"/></a:xfrm></p:spPr></p:pic>`;
    const rels: Relationship[] = [
      { id: "rId9", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video", target: "../media/remote.mp4" },
    ];
    const frame = parseMediaPicForTest(linkedPic, rels, []) as MediaFrame;
    expect(frame.kind).toBe("media");
    expect(frame.linkTarget).toBe("../media/remote.mp4");
    expect(frame.mediaId).toBeUndefined();
  });

  it("preserves a non-media p:pic as the SPEC-079 Media ForeignFrame", () => {
    const frame = parseMediaPicForTest(NON_MEDIA_PIC, [], []);
    expect(frame.kind).toBe("foreign");
    if (frame.kind === "foreign") {
      expect(frame.label).toBe("Media");
      expect(frame.xml).toContain("<p:pic");
    }
  });
});
