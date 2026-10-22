// SPEC-095: the format-neutral media helpers, unit-tested without a format or
// the DOM.
import { describe, it, expect } from "vitest";
import {
  mediaBasename,
  mediaElementVolume,
  mediaIsEmbedded,
  mediaKindForContentType,
} from "./mediaFrames";
import type { MediaFrame } from "./model";

function frame(over: Partial<MediaFrame>): MediaFrame {
  return { id: "f", kind: "media", media: "video", xCm: 0, yCm: 0, widthCm: 16, heightCm: 9, ...over };
}

describe("mediaKindForContentType (SPEC-095)", () => {
  it("maps audio/* to audio and everything else to video", () => {
    expect(mediaKindForContentType("audio/mpeg")).toBe("audio");
    expect(mediaKindForContentType("audio/wav")).toBe("audio");
    expect(mediaKindForContentType("video/mp4")).toBe("video");
    expect(mediaKindForContentType("application/octet-stream")).toBe("video");
    expect(mediaKindForContentType(undefined)).toBe("video");
  });
});

describe("mediaBasename (SPEC-095)", () => {
  it("takes the last path segment of the target or embedded id", () => {
    expect(mediaBasename(frame({ mediaId: "ppt/media/media1.mp4" }))).toBe("media1.mp4");
    expect(mediaBasename(frame({ linkTarget: "../clips/intro.mp4" }))).toBe("intro.mp4");
    expect(mediaBasename(frame({ linkTarget: "C:\\Media\\song.mp3" }))).toBe("song.mp3");
    expect(mediaBasename(frame({}))).toBe("");
  });

  it("prefers the link target over the embedded id", () => {
    expect(mediaBasename(frame({ mediaId: "ppt/media/media1.mp4", linkTarget: "remote.mp4" }))).toBe("remote.mp4");
  });
});

describe("mediaElementVolume (SPEC-095)", () => {
  it("scales 0–100 to 0.0–1.0, absent = full, and clamps", () => {
    expect(mediaElementVolume(undefined)).toBe(1);
    expect(mediaElementVolume(100)).toBe(1);
    expect(mediaElementVolume(50)).toBe(0.5);
    expect(mediaElementVolume(0)).toBe(0);
    expect(mediaElementVolume(-20)).toBe(0);
    expect(mediaElementVolume(250)).toBe(1);
  });
});

describe("mediaIsEmbedded (SPEC-095)", () => {
  it("is true only for an embedded, non-oversize clip", () => {
    expect(mediaIsEmbedded(frame({ mediaId: "ppt/media/media1.mp4" }))).toBe(true);
    expect(mediaIsEmbedded(frame({ linkTarget: "remote.mp4" }))).toBe(false);
    expect(mediaIsEmbedded(frame({ mediaId: "ppt/media/media1.mp4", oversize: true }))).toBe(false);
    expect(mediaIsEmbedded(frame({}))).toBe(false);
  });
});
