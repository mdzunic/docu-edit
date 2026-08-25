// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { cssColorToHex, normalizeTextFrameDom, renderEditableParagraphs } from "./slideTextDom";
import type { SlideParagraph } from "../../lib/presentation/model";

function host(): HTMLElement {
  return document.createElement("div");
}

describe("render → normalize round trip", () => {
  it("recovers paragraphs, run marks and splits", () => {
    const el = host();
    const paragraphs: SlideParagraph[] = [
      { runs: [{ text: "Hello " }, { text: "world", bold: true }] },
      { runs: [{ text: "line two", italic: true, color: "#ff0000" }] },
    ];
    renderEditableParagraphs(el, paragraphs);
    const back = normalizeTextFrameDom(el);
    expect(back).toEqual([
      { runs: [{ text: "Hello " }, { text: "world", bold: true }] },
      { runs: [{ text: "line two", italic: true, color: "#ff0000" }] },
    ]);
  });

  it("renders an empty paragraph with a <br> and reads it back as empty", () => {
    const el = host();
    renderEditableParagraphs(el, [{ runs: [] }]);
    expect(el.querySelector("br")).not.toBeNull();
    expect(normalizeTextFrameDom(el)).toEqual([{ runs: [] }]);
  });

  it("treats a <br> inside a block as a paragraph split", () => {
    const el = host();
    el.innerHTML = `<div>one<br>two</div>`;
    expect(normalizeTextFrameDom(el)).toEqual([{ runs: [{ text: "one" }] }, { runs: [{ text: "two" }] }]);
  });

  it("reads bare text (no block wrapper) as one paragraph", () => {
    const el = host();
    el.innerHTML = `plain text`;
    expect(normalizeTextFrameDom(el)).toEqual([{ runs: [{ text: "plain text" }] }]);
  });

  it("coalesces adjacent spans with identical marks", () => {
    const el = host();
    el.innerHTML = `<div><span style="font-weight:700">a</span><span style="font-weight:700">b</span></div>`;
    expect(normalizeTextFrameDom(el)).toEqual([{ runs: [{ text: "ab", bold: true }] }]);
  });
});

describe("cssColorToHex", () => {
  it("parses #rrggbb, #rgb and rgb()", () => {
    expect(cssColorToHex("#FF0000")).toBe("#ff0000");
    expect(cssColorToHex("#f00")).toBe("#ff0000");
    expect(cssColorToHex("rgb(255, 0, 0)")).toBe("#ff0000");
    expect(cssColorToHex("")).toBeUndefined();
  });
});
