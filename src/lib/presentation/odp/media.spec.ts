// SPEC-095: the odp media wire (draw:plugin), in SPEC-086's golden shape —
// golden write equality, a "" (no-media) case, golden → model read-back, an
// Impress-flavoured tolerant variant, and preserve-don't-parse for a non-media
// draw:frame.
import { describe, it, expect } from "vitest";
import type { MediaFrame, SlideFrame } from "../model";
import { serializeOdpGroupFrame } from "./serialize";
import { parseMediaFrameForTest } from "./reader";
import {
  MEDIA_PLUGIN_AUDIO_GOLDEN,
  MEDIA_PLUGIN_GOLDEN,
  MEDIA_PLUGIN_VARIANT,
  NON_MEDIA_FRAME,
  ODP_GOLDEN_MEDIA_AUDIO,
  ODP_GOLDEN_MEDIA_VIDEO,
} from "./mediaGolden";

const PARTS: Record<string, Uint8Array> = {
  "media/media1.mp4": new Uint8Array([0, 1, 2]),
  "media/media2.wav": new Uint8Array([3, 4, 5]),
  "Pictures/image1.png": new Uint8Array([6, 7, 8]),
};

function frameFrom(model: Omit<MediaFrame, "id">): SlideFrame {
  return { id: "m1", ...model };
}

function modelled(frame: SlideFrame): Omit<MediaFrame, "id"> {
  expect(frame.kind).toBe("media");
  const media = frame as MediaFrame;
  const { id: _id, sourceXml: _sourceXml, ...rest } = media;
  return rest;
}

describe("odp media serialiser (SPEC-095)", () => {
  it("serialises the video golden model to the golden draw:frame", () => {
    expect(serializeOdpGroupFrame(frameFrom(ODP_GOLDEN_MEDIA_VIDEO)).frame).toBe(MEDIA_PLUGIN_GOLDEN);
  });

  it("serialises the audio golden model (no poster, no flags)", () => {
    expect(serializeOdpGroupFrame(frameFrom(ODP_GOLDEN_MEDIA_AUDIO)).frame).toBe(MEDIA_PLUGIN_AUDIO_GOLDEN);
  });

  it("re-emits an untouched clip's verbatim sourceXml, not a rebuild", () => {
    const frame: MediaFrame = { id: "m1", ...ODP_GOLDEN_MEDIA_VIDEO, sourceXml: "<draw:frame>VERBATIM</draw:frame>" };
    expect(serializeOdpGroupFrame(frame).frame).toBe("<draw:frame>VERBATIM</draw:frame>");
  });
});

describe("odp media reader (SPEC-095)", () => {
  it("reads the video golden back to the model", () => {
    expect(modelled(parseMediaFrameForTest(MEDIA_PLUGIN_GOLDEN, PARTS))).toEqual(ODP_GOLDEN_MEDIA_VIDEO);
  });

  it("reads the audio golden back to the model", () => {
    expect(modelled(parseMediaFrameForTest(MEDIA_PLUGIN_AUDIO_GOLDEN, PARTS))).toEqual(ODP_GOLDEN_MEDIA_AUDIO);
  });

  it("keeps the verbatim source for byte-exact re-emit", () => {
    const frame = parseMediaFrameForTest(MEDIA_PLUGIN_GOLDEN, PARTS) as MediaFrame;
    expect(frame.sourceXml).toContain("<draw:plugin");
  });

  it("tolerates a media-typed draw:object with VolumeDB (Impress-flavoured)", () => {
    const frame = parseMediaFrameForTest(MEDIA_PLUGIN_VARIANT, PARTS) as MediaFrame;
    expect(frame.kind).toBe("media");
    expect(frame.media).toBe("video");
    expect(frame.mediaId).toBe("media/media1.mp4");
    expect(frame.loop).toBe(true);
    // -6 dB ≈ 50% linear.
    expect(frame.volume).toBe(50);
  });

  it("models an href outside the package as a linked clip", () => {
    const linked =
      `<draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="4cm">` +
      `<draw:plugin xlink:href="../far/remote.mp4" draw:mime-type="video/mp4"/></draw:frame>`;
    const frame = parseMediaFrameForTest(linked, {}) as MediaFrame;
    expect(frame.kind).toBe("media");
    expect(frame.linkTarget).toBe("../far/remote.mp4");
    expect(frame.mediaId).toBeUndefined();
  });

  it("preserves a non-media draw:frame as the SPEC-079 Media ForeignFrame", () => {
    const frame = parseMediaFrameForTest(NON_MEDIA_FRAME, {});
    expect(frame.kind).toBe("foreign");
    if (frame.kind === "foreign") {
      expect(frame.label).toBe("Media");
    }
  });
});
