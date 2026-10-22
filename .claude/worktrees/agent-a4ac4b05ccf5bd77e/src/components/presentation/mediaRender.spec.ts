// SPEC-095: the pure stage-media classifier — the degrade ladder (link →
// oversize → unsupported → poster/plate) with the probe injected.
import { describe, it, expect } from "vitest";
import {
  MEDIA_LINK_CAPTION,
  MEDIA_OVERSIZE_CAPTION,
  MEDIA_UNSUPPORTED_CAPTION,
  mediaRenderDescriptor,
} from "./mediaRender";
import type { MediaFrame } from "../../lib/presentation/model";

function frame(over: Partial<MediaFrame>): MediaFrame {
  return { id: "f", kind: "media", media: "video", xCm: 0, yCm: 0, widthCm: 16, heightCm: 9, ...over };
}

const poster = (id: string): string => `data:image/png;base64,${id}`;
const always = (): boolean => true;
const never = (): boolean => false;

describe("mediaRenderDescriptor (SPEC-095)", () => {
  it("shows the link placeholder for a linked clip (AC-32)", () => {
    const d = mediaRenderDescriptor(frame({ linkTarget: "../far/remote.mp4" }), poster, always);
    expect(d.mode).toBe("link");
    expect(d.caption).toBe(MEDIA_LINK_CAPTION);
    expect(d.basename).toBe("remote.mp4");
    expect(d.playable).toBe(false);
  });

  it("shows the oversize chip for an oversize clip (AC-11)", () => {
    const d = mediaRenderDescriptor(frame({ mediaId: "ppt/media/big.mp4", oversize: true }), poster, always);
    expect(d.mode).toBe("oversize");
    expect(d.caption).toBe(MEDIA_OVERSIZE_CAPTION);
    expect(d.playable).toBe(false);
  });

  it("shows unsupported when the browser can't play it, keeping the poster (AC-33)", () => {
    const d = mediaRenderDescriptor(frame({ mediaId: "m.mp4", posterMediaId: "p", contentType: "video/mp4" }), poster, never);
    expect(d.mode).toBe("unsupported");
    expect(d.caption).toBe(MEDIA_UNSUPPORTED_CAPTION);
    expect(d.posterSrc).toBe("data:image/png;base64,p");
    expect(d.playable).toBe(false);
  });

  it("shows unsupported when a runtime error fired even if canPlay was optimistic (AC-33)", () => {
    const d = mediaRenderDescriptor(frame({ mediaId: "m.mp4", contentType: "video/mp4" }), poster, always, true);
    expect(d.mode).toBe("unsupported");
  });

  it("shows the poster for a playable clip with a poster (AC-18)", () => {
    const d = mediaRenderDescriptor(frame({ mediaId: "m.mp4", posterMediaId: "p", contentType: "video/mp4" }), poster, always);
    expect(d.mode).toBe("poster");
    expect(d.playable).toBe(true);
  });

  it("shows the neutral plate for a playable clip with no poster (AC-18)", () => {
    const d = mediaRenderDescriptor(frame({ mediaId: "m.mp4", contentType: "video/mp4" }), poster, always);
    expect(d.mode).toBe("plate");
    expect(d.playable).toBe(true);
    expect(d.noun).toBe("Video");
  });
});
