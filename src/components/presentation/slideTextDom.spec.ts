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

  it("recovers paragraph-level props (align/bullet/level) per block", () => {
    const el = host();
    const paragraphs: SlideParagraph[] = [
      { runs: [{ text: "a" }], bullet: "disc" },
      { runs: [{ text: "b" }], bullet: "disc", level: 2 },
      { runs: [{ text: "c" }], align: "center" },
    ];
    renderEditableParagraphs(el, paragraphs);
    expect(normalizeTextFrameDom(el)).toEqual(paragraphs);
  });
});

// The blocking regression: paragraph props must stay glued to their paragraph
// when a keystroke changes the paragraph count. Pressing Enter mid-body makes
// the browser insert a fresh block (no data-p marker); the surviving bulleted
// lines must keep their bullets, and the new line inherits the bullet from the
// paragraph above it — not have props shifted by array index.
describe("paragraph props survive a paragraph-count change", () => {
  it("keeps trailing bullets and inherits onto a browser-inserted line (AC-11/42/45)", () => {
    const el = host();
    renderEditableParagraphs(el, [
      { runs: [{ text: "A" }], bullet: "disc" },
      { runs: [{ text: "B" }], bullet: "disc" },
      { runs: [{ text: "C" }], bullet: "disc" },
    ]);
    // Simulate "Enter at end of A, type word": the browser inserts a plain <div>
    // (no data-p) between A and B; A/B/C blocks keep their attributes.
    const inserted = el.ownerDocument.createElement("div");
    inserted.textContent = "word";
    el.children[0].after(inserted);
    expect(normalizeTextFrameDom(el)).toEqual([
      { runs: [{ text: "A" }], bullet: "disc" },
      { runs: [{ text: "word" }], bullet: "disc" }, // inherits the bullet above
      { runs: [{ text: "B" }], bullet: "disc" },
      { runs: [{ text: "C" }], bullet: "disc" }, // NOT shifted/dropped
    ]);
  });

  it("does not clobber an explicit toggle-off following a bulleted line", () => {
    const el = host();
    renderEditableParagraphs(el, [
      { runs: [{ text: "A" }], bullet: "disc" },
      { runs: [{ text: "B" }] }, // bullet toggled off — explicit no-bullet
    ]);
    expect(normalizeTextFrameDom(el)).toEqual([{ runs: [{ text: "A" }], bullet: "disc" }, { runs: [{ text: "B" }] }]);
  });

  it("preserves per-line indent levels when a line is inserted (no index shift)", () => {
    const el = host();
    renderEditableParagraphs(el, [
      { runs: [{ text: "A" }], bullet: "disc" },
      { runs: [{ text: "B" }], bullet: "disc", level: 1 },
      { runs: [{ text: "C" }], bullet: "disc", level: 2 },
    ]);
    const inserted = el.ownerDocument.createElement("div");
    inserted.textContent = "x";
    el.children[0].after(inserted);
    expect(normalizeTextFrameDom(el)).toEqual([
      { runs: [{ text: "A" }], bullet: "disc" },
      { runs: [{ text: "x" }], bullet: "disc" },
      { runs: [{ text: "B" }], bullet: "disc", level: 1 },
      { runs: [{ text: "C" }], bullet: "disc", level: 2 },
    ]);
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
