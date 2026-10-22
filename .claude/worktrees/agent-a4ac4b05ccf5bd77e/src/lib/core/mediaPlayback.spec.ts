// SPEC-095: the injected-probe playability decision. The probe stands in for
// an `HTMLMediaElement`'s own `canPlayType`, so these tests are DOM-free.
import { describe, it, expect } from "vitest";
import { canPlayMedia } from "./mediaPlayback";

describe("canPlayMedia (SPEC-095)", () => {
  it('is true when the probe answers "probably" or "maybe"', () => {
    expect(canPlayMedia("video/mp4", () => "probably")).toBe(true);
    expect(canPlayMedia("audio/wav", () => "maybe")).toBe(true);
  });

  it('is false when the probe answers "" (the browser cannot play it)', () => {
    expect(canPlayMedia("video/ogg", () => "")).toBe(false);
  });

  it("is false for an absent or empty content type without probing", () => {
    let probed = false;
    const probe = (): string => {
      probed = true;
      return "probably";
    };
    expect(canPlayMedia(undefined, probe)).toBe(false);
    expect(canPlayMedia("", probe)).toBe(false);
    expect(probed).toBe(false);
  });

  it("passes the content type through to the probe verbatim", () => {
    const seen: string[] = [];
    canPlayMedia('video/mp4; codecs="avc1.42E01E"', (type) => {
      seen.push(type);
      return "maybe";
    });
    expect(seen).toEqual(['video/mp4; codecs="avc1.42E01E"']);
  });
});
