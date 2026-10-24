import { describe, expect, it } from "vitest";
import type { SlideParagraph, TextFrame } from "./model";
import {
  applyParagraphFormat,
  applyRunFormat,
  demote,
  lineSpacingKey,
  markSlideEdited,
  normalizeRange,
  paragraphMarkerStyle,
  parseLineSpacingKey,
  promote,
  replaceParagraphText,
  selectionFormatState,
  setIndentLevel,
  setLetterSpacing,
  setLineSpacing,
  setListStyle,
  setParagraphSpacing,
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

// --- SPEC-097: lists, spacing, letter spacing, preservation clearing --------

describe("SPEC-097 widened ParagraphFormatPatch", () => {
  const all = { startPara: 0, endPara: 0 };

  it("sets every new paragraph field and null clears each back to absent", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "number" }]);
    applyParagraphFormat(frame, all, {
      startAt: 4,
      indentCm: 2,
      hangingCm: 0.5,
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 12,
    });
    expect(frame.paragraphs[0]).toEqual({
      runs: [{ text: "a" }],
      bullet: "number",
      startAt: 4,
      indentCm: 2,
      hangingCm: 0.5,
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 12,
    });
    applyParagraphFormat(frame, all, {
      startAt: null,
      indentCm: null,
      hangingCm: null,
      lineSpacing: null,
      spaceBeforePt: null,
      spaceAfterPt: null,
    });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], bullet: "number" });
  });

  it("clamps to the documented ranges and drops values equal to the default", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "number" }]);
    applyParagraphFormat(frame, all, { startAt: 99999, lineSpacing: 9, spaceBeforePt: 900, lineSpacingPt: 1 });
    expect(frame.paragraphs[0].startAt).toBe(32767);
    expect(frame.paragraphs[0].lineSpacing).toBe(5);
    expect(frame.paragraphs[0].spaceBeforePt).toBe(500);
    expect(frame.paragraphs[0].lineSpacingPt).toBe(4);
    applyParagraphFormat(frame, all, { startAt: 1, lineSpacing: 1, spaceBeforePt: 0 });
    expect(frame.paragraphs[0].startAt).toBeUndefined();
    expect(frame.paragraphs[0].lineSpacing).toBeUndefined();
    expect(frame.paragraphs[0].spaceBeforePt).toBeUndefined();
  });

  it("clears the preserved a:pPr / ODF props when a shadowed property is edited, and only then", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "disc", pptxPPr: "<a:pPr><a:buFont/></a:pPr>", odpParaProps: "x" }]);
    applyParagraphFormat(frame, all, { align: "center" });
    applyParagraphFormat(frame, all, { indentCm: 3 });
    setIndentLevel(frame, all, 1);
    expect(frame.paragraphs[0].pptxPPr).toBe("<a:pPr><a:buFont/></a:pPr>");
    expect(frame.paragraphs[0].odpParaProps).toBe("x");
    setLineSpacing(frame, all, { mode: "multiple", value: 1.5 });
    expect(frame.paragraphs[0].pptxPPr).toBeUndefined();
    expect(frame.paragraphs[0].odpParaProps).toBeUndefined();
    frame.paragraphs[0].pptxPPr = "<a:pPr/>";
    setListStyle(frame, all, "square");
    expect(frame.paragraphs[0].pptxPPr).toBeUndefined();
  });
});

describe("toggleBullet's style argument", () => {
  it("toggles numbering on and off with a number style, leaving the default call disc", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }, { runs: [{ text: "b" }] }]);
    toggleBullet(frame, { startPara: 0, endPara: 1 }, "decimal");
    expect(frame.paragraphs.map((p) => p.bullet)).toEqual(["number", "number"]);
    expect(frame.paragraphs.map((p) => p.bulletStyle)).toEqual([undefined, undefined]);
    toggleBullet(frame, { startPara: 0, endPara: 1 }, "decimal");
    expect(frame.paragraphs.map((p) => p.bullet)).toEqual([undefined, undefined]);
  });

  it("switches a bulleted list to numbered (not off) and stores a non-default style", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "disc", bulletChar: "→" }]);
    toggleBullet(frame, { startPara: 0, endPara: 0 }, "upperRoman");
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], bullet: "number", bulletStyle: "upperRoman" });
  });

  it("turning a list off drops every marker field with it", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "number", bulletStyle: "lowerAlpha", startAt: 3 }]);
    toggleBullet(frame, { startPara: 0, endPara: 0 }, "decimal");
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }] });
  });
});

describe("setListStyle", () => {
  const all = { startPara: 0, endPara: 0 };

  it("a marker style decides the list type (markerType agrees with bullet)", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }]);
    setListStyle(frame, all, "lowerRoman");
    expect(frame.paragraphs[0].bullet).toBe("number");
    expect(paragraphMarkerStyle(frame.paragraphs[0])).toBe("lowerRoman");
    setListStyle(frame, all, "circle");
    expect(frame.paragraphs[0].bullet).toBe("disc");
    expect(frame.paragraphs[0].bulletStyle).toBe("circle");
    setListStyle(frame, all, "disc");
    expect(frame.paragraphs[0].bulletStyle).toBeUndefined();
  });

  it("a custom bullet char makes a disc list and wins over the style; a picture wins over both", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "number" }]);
    setListStyle(frame, all, { bulletChar: "★" });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], bullet: "disc", bulletChar: "★" });
    setListStyle(frame, all, { bulletImageId: "ppt/media/image3.png" });
    expect(frame.paragraphs[0].bulletImageId).toBe("ppt/media/image3.png");
    setListStyle(frame, all, "decimal");
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], bullet: "number" });
  });

  it("paragraphMarkerStyle ignores a style that disagrees with bullet", () => {
    expect(paragraphMarkerStyle({ runs: [], bullet: "number", bulletStyle: "square" })).toBe("decimal");
    expect(paragraphMarkerStyle({ runs: [], bullet: "disc" })).toBe("disc");
    expect(paragraphMarkerStyle({ runs: [] })).toBeUndefined();
  });
});

describe("spacing commands", () => {
  const all = { startPara: 0, endPara: 0 };

  it("setLineSpacing: a point value wins and a multiple drops it", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }]);
    setLineSpacing(frame, all, { mode: "exact", pt: 24 });
    expect(frame.paragraphs[0]).toMatchObject({ lineSpacingPt: 24, lineSpacingRule: "exact" });
    expect(lineSpacingKey(frame.paragraphs[0])).toBe("exact:24");
    setLineSpacing(frame, all, { mode: "atLeast", pt: 18 });
    expect(frame.paragraphs[0].lineSpacingRule).toBeUndefined();
    expect(lineSpacingKey(frame.paragraphs[0])).toBe("atLeast:18");
    setLineSpacing(frame, all, { mode: "multiple", value: 1.5 });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], lineSpacing: 1.5 });
    expect(parseLineSpacingKey("multiple:1.5")).toEqual({ mode: "multiple", value: 1.5 });
    expect(parseLineSpacingKey("exact:24")).toEqual({ mode: "exact", pt: 24 });
  });

  it("setParagraphSpacing sets before/after independently", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }]);
    setParagraphSpacing(frame, all, { spaceAfterPt: 6 });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], spaceAfterPt: 6 });
    setParagraphSpacing(frame, all, { spaceBeforePt: 3, spaceAfterPt: 0 });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "a" }], spaceBeforePt: 3 });
  });

  it("setLetterSpacing marks only the range, clamps to ±10 in 0.1 steps, 0 clears and runs coalesce", () => {
    const frame = frameOf([{ runs: [{ text: "abcd" }] }]);
    setLetterSpacing(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 2 } }, 0.54);
    expect(frame.paragraphs[0].runs).toEqual([{ text: "ab", letterSpacingPt: 0.5 }, { text: "cd" }]);
    setLetterSpacing(frame, { start: { para: 0, offset: 2 }, end: { para: 0, offset: 4 } }, 0.5);
    expect(frame.paragraphs[0].runs).toEqual([{ text: "abcd", letterSpacingPt: 0.5 }]);
    setLetterSpacing(frame, wholeFrameRange(frame), 40);
    expect(frame.paragraphs[0].runs[0].letterSpacingPt).toBe(10);
    setLetterSpacing(frame, wholeFrameRange(frame), 0);
    expect(frame.paragraphs[0].runs).toEqual([{ text: "abcd" }]);
  });

  it("run links coalesce per key and a tooltip never outlives its link", () => {
    const frame = frameOf([{ runs: [{ text: "abcd" }] }]);
    applyRunFormat(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 2 } }, { linkSlide: 3, linkTooltip: "Go" });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "ab", linkSlide: 3, linkTooltip: "Go" }, { text: "cd" }]);
    applyRunFormat(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 2 } }, { linkHref: "https://x.test" });
    expect(frame.paragraphs[0].runs[0]).toEqual({ text: "ab", linkHref: "https://x.test", linkTooltip: "Go" });
    applyRunFormat(frame, wholeFrameRange(frame), { linkHref: null });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "abcd" }]);
  });
});

describe("demote / promote", () => {
  it("demotes a numbered item to the next level's default style (3. → a.) and back", () => {
    const frame = frameOf([{ runs: [{ text: "c" }], bullet: "number" }]);
    demote(frame, { startPara: 0, endPara: 0 });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "c" }], bullet: "number", level: 1, bulletStyle: "lowerAlpha" });
    demote(frame, { startPara: 0, endPara: 0 });
    expect(frame.paragraphs[0].bulletStyle).toBe("lowerRoman");
    promote(frame, { startPara: 0, endPara: 0 });
    promote(frame, { startPara: 0, endPara: 0 });
    expect(frame.paragraphs[0]).toEqual({ runs: [{ text: "c" }], bullet: "number" });
  });

  it("keeps an explicitly chosen non-cycle style and shifts an explicit indent", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], bullet: "number", bulletStyle: "upperRoman", indentCm: 1 }]);
    demote(frame, { startPara: 0, endPara: 0 });
    expect(frame.paragraphs[0]).toMatchObject({ level: 1, bulletStyle: "upperRoman", indentCm: 2.27 });
  });

  it("clamps 0..8", () => {
    const frame = frameOf([{ runs: [{ text: "a" }], level: 8 }, { runs: [{ text: "b" }] }]);
    demote(frame, { startPara: 0, endPara: 1 });
    expect(frame.paragraphs.map((p) => p.level)).toEqual([8, 1]);
    promote(frame, { startPara: 1, endPara: 1 });
    promote(frame, { startPara: 1, endPara: 1 });
    expect(frame.paragraphs[1].level).toBeUndefined();
  });
});

describe("SelectionFormatState SPEC-097 fields", () => {
  it("folds list, spacing and letter spacing, reporting mixed where they differ", () => {
    const frame = frameOf([
      { runs: [{ text: "ab", letterSpacingPt: 1 }], bullet: "number", startAt: 3, lineSpacing: 1.5, spaceAfterPt: 6 },
      { runs: [{ text: "cd" }], bullet: "number", startAt: 3, lineSpacing: 1.5, spaceAfterPt: 6, spaceBeforePt: 2 },
    ]);
    const state = selectionFormatState(frame, { start: { para: 0, offset: 0 }, end: { para: 1, offset: 2 } });
    expect(state.bulletStyle).toBe("decimal");
    expect(state.numbered).toBe(true);
    expect(state.startAt).toBe(3);
    expect(state.lineSpacing).toBe("multiple:1.5");
    expect(state.spaceAfterPt).toBe(6);
    expect(state.spaceBeforePt).toBe("mixed");
    expect(state.letterSpacingPt).toBe("mixed");
  });

  it("reports none for non-list paragraphs and mixed list styles", () => {
    const frame = frameOf([{ runs: [{ text: "a" }] }, { runs: [{ text: "b" }], bullet: "disc" }]);
    expect(selectionFormatState(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 1 } }).bulletStyle).toBe("none");
    expect(selectionFormatState(frame, wholeFrameRange(frame)).bulletStyle).toBe("mixed");
  });
});

describe("replaceParagraphText (SPEC-098)", () => {
  it("takes the marks of the run containing the start and leaves the paragraph's own fields alone", () => {
    const para: SlideParagraph = { runs: [{ text: "say " }, { text: "hello", bold: true, lang: "de-DE" }, { text: " now" }], align: "center" };
    replaceParagraphText(para, 4, 9, "hi");
    expect(para).toEqual({ runs: [{ text: "say " }, { text: "hi", bold: true, lang: "de-DE" }, { text: " now" }], align: "center" });
  });

  it("spans several runs, re-coalesces identical neighbours and can delete outright", () => {
    const para: SlideParagraph = { runs: [{ text: "ab" }, { text: "cd", italic: true }, { text: "ef" }] };
    replaceParagraphText(para, 1, 5, "X");
    expect(para.runs).toEqual([{ text: "aXf" }]);
    replaceParagraphText(para, 0, 3, "");
    expect(para.runs).toEqual([]);
  });

  it("clamps its range to the paragraph", () => {
    const para: SlideParagraph = { runs: [{ text: "abc" }] };
    replaceParagraphText(para, 2, 99, "Z");
    expect(para.runs).toEqual([{ text: "abZ" }]);
  });
});

describe("run language and noProof (SPEC-098)", () => {
  it("sets and clears lang and noProof through RunFormatPatch", () => {
    const frame = frameOf([{ runs: [{ text: "Guten Tag" }] }]);
    applyRunFormat(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 5 } }, { lang: "de-DE", noProof: true });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "Guten", lang: "de-DE", noProof: true }, { text: " Tag" }]);
    applyRunFormat(frame, wholeFrameRange(frame), { clearLang: true, noProof: false });
    expect(frame.paragraphs[0].runs).toEqual([{ text: "Guten Tag" }]);
  });

  it("keeps runs that differ only in language (or noProof) apart", () => {
    const frame = frameOf([{ runs: [{ text: "colour", lang: "en-GB" }, { text: " color", lang: "en-US" }] }]);
    applyRunFormat(frame, wholeFrameRange(frame), { bold: true });
    expect(frame.paragraphs[0].runs).toEqual([
      { text: "colour", bold: true, lang: "en-GB" },
      { text: " color", bold: true, lang: "en-US" },
    ]);
  });

  it("reports lang and noProof in the selection state, mixed across runs", () => {
    const frame = frameOf([{ runs: [{ text: "ab", lang: "de-DE" }, { text: "cd", noProof: true }] }]);
    const first = selectionFormatState(frame, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 2 } });
    expect([first.lang, first.noProof]).toEqual(["de-DE", false]);
    const all = selectionFormatState(frame, wholeFrameRange(frame));
    expect([all.lang, all.noProof]).toEqual(["mixed", "mixed"]);
    const plain = selectionFormatState(frame, { start: { para: 0, offset: 2 }, end: { para: 0, offset: 4 } });
    expect([plain.lang, plain.noProof]).toEqual(["", true]);
  });
});
