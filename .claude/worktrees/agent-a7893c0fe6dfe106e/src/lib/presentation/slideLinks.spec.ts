import { describe, expect, it } from "vitest";
import type { ImageFrame, TextFrame } from "./model";
import { cleanLinkTarget, effectiveRunLink, isLinkableKind, setFrameLink, setRunLink } from "./slideLinks";

const text = (): TextFrame => ({ id: "t", kind: "text", xCm: 0, yCm: 0, widthCm: 5, heightCm: 2, paragraphs: [{ runs: [{ text: "click here" }] }] });

describe("slide links (SPEC-097)", () => {
  it("only text/image/shape frames are linkable", () => {
    expect(["text", "image", "shape"].every((k) => isLinkableKind(k as TextFrame["kind"]))).toBe(true);
    expect(isLinkableKind("group")).toBe(false);
    expect(isLinkableKind("foreign")).toBe(false);
    expect(isLinkableKind("table")).toBe(false);
  });

  it("cleanLinkTarget sanitises hrefs and bounds slide numbers", () => {
    expect(cleanLinkTarget({ href: "  https://example.com  ", tooltip: " " }, 3)).toEqual({ href: "https://example.com" });
    expect(cleanLinkTarget({ href: "javascript:alert(1)" }, 3)).toBeNull();
    expect(cleanLinkTarget({ slide: 3, tooltip: "Results" }, 3)).toEqual({ slide: 3, tooltip: "Results" });
    expect(cleanLinkTarget({ slide: 4 }, 3)).toBeNull();
  });

  it("setRunLink links exactly the range and unlinks with null", () => {
    const frame = text();
    setRunLink(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 5 } }, { slide: 3 });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "click", linkSlide: 3 }, { text: " here" }]);
    setRunLink(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 10 } }, null);
    expect(frame.paragraphs[0].runs).toEqual([{ text: "click here" }]);
  });

  it("setFrameLink sets one target and a run link wins over its frame's", () => {
    const frame = text();
    setFrameLink(frame, { href: "https://a.test", tooltip: "A" });
    expect(frame).toMatchObject({ linkHref: "https://a.test", linkTooltip: "A" });
    setFrameLink(frame, { slide: 2 });
    expect(frame.linkHref).toBeUndefined();
    expect(frame.linkSlide).toBe(2);
    expect(effectiveRunLink({ text: "x" }, frame)).toEqual({ slide: 2, tooltip: undefined });
    expect(effectiveRunLink({ text: "x", linkHref: "https://b.test" }, frame)).toEqual({ href: "https://b.test", tooltip: undefined });
    const image: ImageFrame = { id: "i", kind: "image", mediaId: "m", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 };
    expect(setFrameLink(image, null)).toBe(true);
  });
});
