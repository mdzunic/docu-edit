import { describe, expect, it } from "vitest";
import type { SlideParagraph, TextFrame } from "./model";
import {
  applyParagraphFormat,
  applyRunFormat,
  markSlideEdited,
  normalizeRange,
  selectionFormatState,
  setIndentLevel,
  toggleBullet,
  wholeFrameRange,
} from "./text";

function frameOf(paragraphs: SlideParagraph[]): TextFrame {
  return { id: "f", kind: "text", xCm: 0, yCm: 0, widthCm: 10, heightCm: 5, paragraphs };
}

describe("applyRunFormat", () => {
  it("splits a single run at both selection edges and marks only the middle", () => {
    const frame = frameOf([{ runs: [{ text: "Hello world" }] }]);
    applyRunFormat(frame, { start: { para: 0, offset: 6 }, end: { para: 0, offset: 11 } }, { bold: true });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "Hello " }, { text: "world", bold: true }]);
  });

  it("coalesces adjacent runs that end up identical", () => {
    const frame = frameOf([{ runs: [{ text: "ab", bold: true }, { text: "cd" }] }]);
    // Un-bold the whole paragraph → one plain run.
    applyRunFormat(frame, wholeFrameRange(frame), { bold: false });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "abcd" }]);
  });

  it("applies across paragraph boundaries", () => {
    const frame = frameOf([{ runs: [{ text: "one" }] }, { runs: [{ text: "two" }] }]);
    applyRunFormat(frame, { start: { para: 0, offset: 1 }, end: { para: 1, offset: 2 } }, { italic: true });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "o" }, { text: "ne", italic: true }]);
    expect(frame.paragraphs[1].runs).toEqual([{ text: "tw", italic: true }, { text: "o" }]);
  });

  it("sets font family, size and colour on the selected span", () => {
    const frame = frameOf([{ runs: [{ text: "styled" }] }]);
    applyRunFormat(frame, wholeFrameRange(frame), { fontFamily: "Arial", sizePt: 24, color: "#ff0000" });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "styled", fontFamily: "Arial", sizePt: 24, color: "#ff0000" }]);
  });

  it("leaves a collapsed range untouched", () => {
    const frame = frameOf([{ runs: [{ text: "abc" }] }]);
    const changed = applyRunFormat(frame, { start: { para: 0, offset: 1 }, end: { para: 0, offset: 1 } }, { bold: true });
    expect(changed).toBe(false);
    expect(frame.paragraphs[0].runs).toEqual([{ text: "abc" }]);
  });
});

describe("applyParagraphFormat / bullets / indent", () => {
  it("sets and clears alignment (left drops the field)", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }]);
    applyParagraphFormat(frame, { startPara: 0, endPara: 0 }, { align: "center" });
    expect(frame.paragraphs[0].align).toBe("center");
    applyParagraphFormat(frame, { startPara: 0, endPara: 0 }, { align: "left" });
    expect(frame.paragraphs[0].align).toBeUndefined();
  });

  it("toggles disc bullets on, then off", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }, { runs: [{ text: "b" }] }]);
    toggleBullet(frame, { startPara: 0, endPara: 1 });
    expect(frame.paragraphs.map((p) => p.bullet)).toEqual(["disc", "disc"]);
    toggleBullet(frame, { startPara: 0, endPara: 1 });
    expect(frame.paragraphs.map((p) => p.bullet)).toEqual([undefined, undefined]);
  });

  it("turns bullets on for all when the selection is mixed", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "disc" }, { runs: [{ text: "b" }] }]);
    toggleBullet(frame, { startPara: 0, endPara: 1 });
    expect(frame.paragraphs.map((p) => p.bullet)).toEqual(["disc", "disc"]);
  });

  it("clamps indent to 0..8", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }]);
    setIndentLevel(frame, { startPara: 0, endPara: 0 }, 10);
    expect(frame.paragraphs[0].level).toBe(8);
    setIndentLevel(frame, { startPara: 0, endPara: 0 }, -20);
    expect(frame.paragraphs[0].level).toBeUndefined();
  });
});

describe("selectionFormatState", () => {
  it("reports uniform marks and mixed where they differ", () => {
    const frame = frameOf([{ runs: [{ text: "ab", bold: true }, { text: "cd" }], align: "center" }]);
    const state = selectionFormatState(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 4 } });
    expect(state.bold).toBe("mixed");
    expect(state.align).toBe("center");
    expect(state.italic).toBe(false);
  });

  it("reads the character left of a collapsed caret", () => {
    const frame = frameOf([{ runs: [{ text: "ab", bold: true }] }]);
    const state = selectionFormatState(frame, { start: { para: 0, offset: 2 }, end: { para: 0, offset: 2 } });
    expect(state.bold).toBe(true);
  });
});

describe("normalizeRange", () => {
  it("orders reversed endpoints and clamps out-of-range offsets", () => {
    const frame = frameOf([{ runs: [{ text: "abc" }] }]);
    const norm = normalizeRange(frame, { start: { para: 0, offset: 9 }, end: { para: 0, offset: 1 } });
    expect(norm).toEqual({ start: { para: 0, offset: 1 }, end: { para: 0, offset: 3 } });
  });
});

describe("markSlideEdited", () => {
  it("sets dirty and keeps sourceXml for pptx", () => {
    const slide = { dirty: undefined as true | undefined, sourceXml: "<p/>" };
    markSlideEdited(slide, "pptx");
    expect(slide.dirty).toBe(true);
    expect(slide.sourceXml).toBe("<p/>");
  });

  it("drops sourceXml for odp", () => {
    const slide = { dirty: undefined as true | undefined, sourceXml: "<page/>" };
    markSlideEdited(slide, "odp");
    expect(slide.dirty).toBe(true);
    expect(slide.sourceXml).toBeUndefined();
  });
});
