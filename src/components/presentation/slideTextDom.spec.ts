// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  cssColorToHex,
  EDIT_LINK_CLASS,
  marksEqual,
  normalizeTextFrameDom,
  placeCaretAtPoint,
  renderEditableParagraphs,
} from "./slideTextDom";
import { cmPx } from "./slideRender";
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

// jsdom has no layout, so the hit-test API is absent; stub it on the document
// as a plain property (index-signature cast so it can be deleted afterwards).
function stubHitTest(resolve: (x: number, y: number) => Range): void {
  (document as unknown as Record<string, unknown>).caretRangeFromPoint = resolve;
}
function clearHitTest(): void {
  delete (document as unknown as Record<string, unknown>).caretRangeFromPoint;
}

describe("placeCaretAtPoint (AC-1: caret at the click point)", () => {
  it("places the caret at the hit-tested point when it lands inside the host", () => {
    const el = host();
    document.body.appendChild(el);
    renderEditableParagraphs(el, [{ runs: [{ text: "Quarterly Review" }] }]);
    const textNode = el.querySelector("span")!.firstChild!;
    stubHitTest(() => {
      const range = document.createRange();
      range.setStart(textNode, 0); // the point-before-'Q' the AC exercises
      return range;
    });
    try {
      expect(placeCaretAtPoint(el, 5, 5)).toBe(true);
      const sel = document.getSelection()!;
      expect(sel.focusNode).toBe(textNode);
      expect(sel.focusOffset).toBe(0);
    } finally {
      clearHitTest();
      document.body.removeChild(el);
    }
  });

  it("returns false when the point resolves outside the host (caller falls back)", () => {
    const el = host();
    document.body.appendChild(el);
    renderEditableParagraphs(el, [{ runs: [{ text: "hi" }] }]);
    const outside = document.createElement("div");
    outside.textContent = "elsewhere";
    document.body.appendChild(outside);
    stubHitTest(() => {
      const range = document.createRange();
      range.setStart(outside.firstChild!, 0);
      return range;
    });
    try {
      expect(placeCaretAtPoint(el, 5, 5)).toBe(false);
    } finally {
      clearHitTest();
      document.body.removeChild(el);
      document.body.removeChild(outside);
    }
  });

  it("returns false when neither hit-test API exists (jsdom)", () => {
    const el = host();
    renderEditableParagraphs(el, [{ runs: [{ text: "hi" }] }]);
    expect(placeCaretAtPoint(el, 1, 1)).toBe(false);
  });

  // SPEC-097 insets pad the body around the host: a double-click in that
  // margin, just left of the first character, resolves to the body — retry
  // at the nearest point inside the host's box rather than falling back to
  // the end (SPEC-081 AC-1's "XXXQuarterly Review").
  it("retries a point in the body's inset margin at the nearest point inside the host", () => {
    const body = document.createElement("div");
    const el = host();
    body.appendChild(el);
    document.body.appendChild(body);
    renderEditableParagraphs(el, [{ runs: [{ text: "Quarterly Review" }] }]);
    const textNode = el.querySelector("span")!.firstChild!;
    el.getBoundingClientRect = () => new DOMRect(10, 20, 100, 30); // left 10, top 20, right 110, bottom 50
    const probes: Array<[number, number]> = [];
    stubHitTest((x: number, y: number) => {
      probes.push([x, y]);
      const range = document.createRange();
      if (x < 10) range.setStart(body, 0); // the inset padding: the body, not the host
      else range.setStart(textNode, 0);
      return range;
    });
    try {
      expect(placeCaretAtPoint(el, 3, 35)).toBe(true);
      expect(probes).toEqual([
        [3, 35],
        [11, 35],
      ]);
      const sel = document.getSelection()!;
      expect(sel.focusNode).toBe(textNode);
      expect(sel.focusOffset).toBe(0);
    } finally {
      clearHitTest();
      document.body.removeChild(body);
    }
  });

  it("does not retry a point already inside the host's box that resolves elsewhere", () => {
    const el = host();
    document.body.appendChild(el);
    renderEditableParagraphs(el, [{ runs: [{ text: "hi" }] }]);
    el.getBoundingClientRect = () => new DOMRect(0, 0, 100, 30);
    let probes = 0;
    stubHitTest(() => {
      probes++;
      const range = document.createRange();
      range.setStart(document.body, 0);
      return range;
    });
    try {
      expect(placeCaretAtPoint(el, 5, 5)).toBe(false);
      expect(probes).toBe(1);
    } finally {
      clearHitTest();
      document.body.removeChild(el);
    }
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

// SPEC-097: every new paragraph field rides its block as data-*, every new
// run mark its span, so an editing session never drops list, spacing,
// preservation or link state (AC-1..10, 16, 43).
describe("SPEC-097 fields survive the contenteditable round trip", () => {
  const rich: SlideParagraph[] = [
    {
      runs: [{ text: "Seeded" }],
      bullet: "number",
      level: 1,
      bulletStyle: "lowerRoman",
      startAt: 7,
      indentCm: 2.5,
      hangingCm: 0.8,
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 12.5,
      pptxPPr: '{"attrs":" fontAlgn=\\"ctr\\"","before":"","mid":"<a:buFont typeface=\\"Wingdings\\"/>","after":""}',
    },
    { runs: [{ text: "Starred" }], bullet: "disc", bulletChar: "★", lineSpacingPt: 24, lineSpacingRule: "exact" },
    { runs: [{ text: "Pictured" }], bullet: "disc", bulletImageId: "ppt/media/image2.png", lineSpacingPt: 18 },
    { runs: [{ text: "Odp" }], odpParaProps: 'fo:text-indent="0cm" style:punctuation-wrap="hanging"' },
  ];

  it("recovers every paragraph field per block", () => {
    const el = host();
    renderEditableParagraphs(el, rich);
    expect(normalizeTextFrameDom(el)).toEqual(rich);
  });

  it("recovers letter spacing and both link kinds per run", () => {
    const el = host();
    const paragraphs: SlideParagraph[] = [
      {
        runs: [
          { text: "wide", letterSpacingPt: 0.5 },
          { text: " tight", letterSpacingPt: -1.2, bold: true },
          { text: " site", linkHref: "https://example.com/", linkTooltip: "Example" },
          { text: " slide", linkSlide: 3 },
        ],
      },
    ];
    renderEditableParagraphs(el, paragraphs);
    expect(normalizeTextFrameDom(el)).toEqual(paragraphs);
  });

  it("marks a linked span by class and data-*, never an inline underline", () => {
    const el = host();
    renderEditableParagraphs(el, [{ runs: [{ text: "go", linkSlide: 2, linkTooltip: "Two" }] }]);
    const span = el.querySelector("span") as HTMLElement;
    expect(span.classList.contains(EDIT_LINK_CLASS)).toBe(true);
    expect(span.dataset.linkSlide).toBe("2");
    expect(span.dataset.linkTooltip).toBe("Two");
    expect(span.style.textDecoration).toBe("");
  });

  it("coalesces only spans whose letter spacing and links agree (AC-10)", () => {
    const el = host();
    el.innerHTML =
      '<div class="ce-p" data-p="1"><span style="letter-spacing: 2px">a</span><span style="letter-spacing: 2px">b</span>' +
      '<span>c</span><span data-link-slide="2">d</span><span data-link-slide="2">e</span><span data-link-slide="3">f</span></div>';
    expect(normalizeTextFrameDom(el)).toEqual([
      {
        runs: [
          { text: "ab", letterSpacingPt: 1.5 },
          { text: "c" },
          { text: "de", linkSlide: 2 },
          { text: "f", linkSlide: 3 },
        ],
      },
    ]);
    expect(marksEqual({ text: "x", linkHref: "https://a.example/" }, { text: "y", linkHref: "https://b.example/" })).toBe(false);
    expect(marksEqual({ text: "x", linkSlide: 2, linkTooltip: "t" }, { text: "y", linkSlide: 2 })).toBe(false);
    expect(marksEqual({ text: "x", letterSpacingPt: 0 }, { text: "y" })).toBe(true);
  });

  it("a browser-created line inherits list/spacing/preservation props, but not startAt", () => {
    const el = host();
    renderEditableParagraphs(el, [rich[0]]);
    const inserted = el.ownerDocument.createElement("div");
    inserted.textContent = "next";
    el.appendChild(inserted);
    const [first, second] = normalizeTextFrameDom(el);
    expect(first).toEqual(rich[0]);
    const { startAt: _seed, ...carried } = rich[0];
    expect(second).toEqual({ ...carried, runs: [{ text: "next" }] });
  });

  it("a block the browser CLONED on Enter (same data-pi) keeps the props but drops the start-at seed", () => {
    const el = host();
    renderEditableParagraphs(el, [rich[0], { runs: [{ text: "after" }] }]);
    // Chrome's Enter split clones the block, attributes and all.
    const clone = el.children[0].cloneNode(false) as HTMLElement;
    clone.textContent = "split";
    el.children[0].after(clone);
    const result = normalizeTextFrameDom(el);
    expect(result).toHaveLength(3);
    expect(result[0].startAt).toBe(7);
    expect(result[1].startAt).toBeUndefined();
    expect(result[1].bulletStyle).toBe("lowerRoman");
    expect(result[1].pptxPPr).toBe(rich[0].pptxPPr);
    expect(result[2]).toEqual({ runs: [{ text: "after" }] });
  });

  it("a paragraph's source id (a preserved-animation target) round-trips but stays on the first line of a split", () => {
    const el = host();
    const target: SlideParagraph = { runs: [{ text: "Target" }], odpParaProps: 'fo:text-indent="0cm"', sourceWireId: "id4" };
    renderEditableParagraphs(el, [target]);
    expect(normalizeTextFrameDom(el)).toEqual([target]);
    // Chrome's Enter split clones the block, attributes and all.
    const clone = el.children[0].cloneNode(false) as HTMLElement;
    clone.textContent = "split";
    el.children[0].after(clone);
    // …and a block the browser minted itself inherits from the line above.
    const inserted = el.ownerDocument.createElement("div");
    inserted.textContent = "new";
    el.appendChild(inserted);
    const [first, split, created] = normalizeTextFrameDom(el);
    expect(first.sourceWireId).toBe("id4");
    expect(split.sourceWireId).toBeUndefined();
    expect(split.odpParaProps).toBe(target.odpParaProps);
    expect(created.sourceWireId).toBeUndefined();
  });

  it("a line's pptx source paragraph round-trips and carries onto every line split off it", () => {
    const el = host();
    const line: SlideParagraph = { runs: [{ text: "Target" }], pptxSourceParagraph: 0 };
    const next: SlideParagraph = { runs: [{ text: "Next" }], pptxSourceParagraph: 2 };
    renderEditableParagraphs(el, [line, next]);
    expect(normalizeTextFrameDom(el)).toEqual([line, next]);
    // Chrome's Enter split clones the block; a soft break stays inside it…
    const clone = el.children[0].cloneNode(false) as HTMLElement;
    clone.append("split", el.ownerDocument.createElement("br"), "soft");
    el.children[0].after(clone);
    // …and a block the browser minted itself inherits from the line above.
    const inserted = el.ownerDocument.createElement("div");
    inserted.textContent = "new";
    el.children[1].after(inserted);
    const lines = normalizeTextFrameDom(el);
    expect(lines.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Target", "split", "soft", "new", "Next"]);
    // All of it came from source a:p 0 — a preserved by-paragraph build
    // still reveals it together.
    expect(lines.map((p) => p.pptxSourceParagraph)).toEqual([0, 0, 0, 0, 2]);
  });

  it("renders the paragraph's indent and spacing on its block (WYSIWYG), with no hanging first line for a list item", () => {
    const el = host();
    renderEditableParagraphs(el, [
      { runs: [{ text: "item" }], bullet: "disc", spaceAfterPt: 6, lineSpacing: 2 },
      { runs: [{ text: "plain" }], hangingCm: 0.5, indentCm: 1 },
    ]);
    const [item, plain] = Array.from(el.children) as HTMLElement[];
    expect(item.style.marginLeft).toBe(cmPx(1.27));
    expect(item.style.textIndent).toBe("");
    expect(item.style.marginBottom).toBe("calc(8px * var(--fit-scale, 1))");
    expect(item.style.lineHeight).toBe("calc(2.4 * var(--line-scale, 1))");
    expect(plain.style.marginLeft).toBe(cmPx(1));
    expect(plain.style.textIndent).toBe(cmPx(-0.5));
  });
});

describe("SPEC-098: run language and noProof survive the contenteditable round trip", () => {
  it("carries lang on the span and noProof as data-no-proof, and keeps runs differing only in them apart", () => {
    const el = host();
    const paragraphs: SlideParagraph[] = [
      { runs: [{ text: "Guten Tag", lang: "de-DE" }, { text: " hello" }, { text: " lorem", noProof: true }, { text: " colour", lang: "en-gb" }] },
    ];
    renderEditableParagraphs(el, paragraphs);
    const spans = el.querySelectorAll("span");
    expect(spans[0].lang).toBe("de-DE");
    expect(spans[2].dataset.noProof).toBe("1");
    expect(normalizeTextFrameDom(el)).toEqual(paragraphs);
    expect(marksEqual({ text: "a", lang: "de-DE" }, { text: "b" })).toBe(false);
    expect(marksEqual({ text: "a", noProof: true }, { text: "b" })).toBe(false);
  });
});
