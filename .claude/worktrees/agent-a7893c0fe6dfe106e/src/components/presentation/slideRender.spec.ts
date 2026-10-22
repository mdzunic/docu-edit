// SPEC-097 — the stage's text rendering: derived list markers (never stored),
// the model-derived indent (the old private 0.95cm/level is gone — pinned
// here as a deliberate correction to the writers' 1.27cm), paragraph line and
// space-before/after pitch, run letter spacing and the fit-scale hooks, and
// the link anchors (a run's own link, else its frame's).
import { afterEach, describe, expect, it, vi } from "vitest";
import { mount } from "@vue/test-utils";
import type { SlideFrame, SlideParagraph, TextFrame } from "../../lib/presentation/model";
import { cmToPx } from "../../lib/core/units";
import SlideFrameView from "./SlideFrameView.vue";
import {
  cmPx,
  followEditorLink,
  frameLinkRender,
  linkRender,
  markParagraphList,
  markedParagraphs,
  markerStyle,
  paragraphLineHeight,
  paragraphStyle,
  runLineHeight,
  runRenders,
  runStyle,
  SLIDE_LINK_FOLLOW_EVENT,
  slideLinkElement,
  slideLinkTargetOf,
  tableRenderRows,
} from "./slideRender";

const num = (text: string, level = 0, extra: Partial<SlideParagraph> = {}): SlideParagraph => ({
  runs: [{ text }],
  bullet: "number",
  ...(level ? { level } : {}),
  ...extra,
});
const disc = (text: string, level = 0, extra: Partial<SlideParagraph> = {}): SlideParagraph => ({
  runs: [{ text }],
  bullet: "disc",
  ...(level ? { level } : {}),
  ...extra,
});
const plain = (text: string, extra: Partial<SlideParagraph> = {}): SlideParagraph => ({ runs: [{ text }], ...extra });

function textFrame(paragraphs: SlideParagraph[], extra: Partial<TextFrame> = {}): TextFrame {
  return { id: "t1", kind: "text", xCm: 1, yCm: 1, widthCm: 10, heightCm: 5, paragraphs, ...extra };
}

const markers = (paragraphs: SlideParagraph[]): string[] => markParagraphList(paragraphs).map((entry) => entry.marker);

describe("markedParagraphs — numbering is derived by adjacency (AC-30)", () => {
  it("numbers a contiguous run 1., 2., 3. and bullets with the style glyph", () => {
    expect(markers([num("a"), num("b"), num("c")])).toEqual(["1.", "2.", "3."]);
    expect(markers([disc("a"), disc("b", 0, { bulletStyle: "square" })])).toEqual(["•", "▪"]);
  });

  it("a non-list paragraph between two items restarts the count", () => {
    expect(markers([num("a"), num("b"), plain("break"), num("c")])).toEqual(["1.", "2.", "", "1."]);
  });

  it("a shallower item between two deeper items restarts the deeper level", () => {
    const paragraphs = [
      num("one"),
      num("one-a", 1, { bulletStyle: "lowerAlpha" }),
      num("one-b", 1, { bulletStyle: "lowerAlpha" }),
      num("two"),
      num("two-a", 1, { bulletStyle: "lowerAlpha" }),
    ];
    expect(markers(paragraphs)).toEqual(["1.", "a.", "b.", "2.", "a."]);
  });

  it("a deeper run does not restart its parent level", () => {
    expect(markers([num("1"), num("1a", 1), num("1b", 1), num("2")])).toEqual(["1.", "1.", "2.", "2."]);
  });

  it("startAt seeds a run and later items continue it", () => {
    expect(markers([num("a", 0, { startAt: 4 }), num("b"), num("c")])).toEqual(["4.", "5.", "6."]);
    expect(markers([num("a", 0, { bulletStyle: "upperRoman", startAt: 3 }), num("b", 0, { bulletStyle: "upperRoman" })])).toEqual([
      "III.",
      "IV.",
    ]);
  });

  it("a different number style at the same level restarts that level", () => {
    expect(markers([num("a"), num("b"), num("c", 0, { bulletStyle: "lowerRoman" })])).toEqual(["1.", "2.", "i."]);
  });

  it("a custom bullet character wins over the style; a picture bullet wins over both", () => {
    const marked = markParagraphList([
      disc("star", 0, { bulletChar: "★" }),
      disc("pic", 0, { bulletChar: "★", bulletImageId: "ppt/media/image3.png" }),
    ]);
    expect(marked[0].marker).toBe("★");
    expect(marked[1].marker).toBe("");
    expect(marked[1].markerImageId).toBe("ppt/media/image3.png");
  });

  it("reads shape text through the same derivation; other kinds carry none", () => {
    const shape: SlideFrame = {
      id: "s1",
      kind: "shape",
      geometry: "rect",
      xCm: 0,
      yCm: 0,
      widthCm: 4,
      heightCm: 2,
      paragraphs: [num("x"), num("y")],
    };
    expect(markedParagraphs(shape).map((e) => e.marker)).toEqual(["1.", "2."]);
    const image: SlideFrame = { id: "i1", kind: "image", mediaId: "m", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 };
    expect(markedParagraphs(image)).toEqual([]);
  });

  it("table cells number through the same helper", () => {
    const table: SlideFrame = {
      id: "t",
      kind: "table",
      xCm: 0,
      yCm: 0,
      widthCm: 4,
      heightCm: 2,
      columnWidthsCm: [4],
      rows: [{ cells: [{ paragraphs: [num("a"), num("b")] }] }],
    };
    expect(tableRenderRows(table)[0][0].paragraphs.map((e) => e.marker)).toEqual(["1.", "2."]);
  });
});

describe("paragraphStyle — indent derived from the model (AC-32/33)", () => {
  it("pins the deliberate correction: a list item's per-level indent is 1.27cm, not the old 0.95cm", () => {
    // level 0 list item: text edge (0 + 1) × 1.27, marker box 0.635 hanging.
    expect(paragraphStyle(disc("a")).marginLeft).toBe(cmPx(1.27));
    expect(paragraphStyle(disc("a")).textIndent).toBe(cmPx(-0.635));
    // level 2: (2 + 1) × 1.27 — every level step widens from 0.95cm to 1.27cm.
    const level2 = parseFloat(paragraphStyle(disc("a", 2)).marginLeft);
    const level1 = parseFloat(paragraphStyle(disc("a", 1)).marginLeft);
    expect(level2 - level1).toBeCloseTo(cmToPx(1.27), 6);
    expect(level2 - level1).not.toBeCloseTo(cmToPx(0.95), 2);
  });

  it("a plain paragraph stays flush left at level 0 and indents level × 1.27cm without a hanging marker", () => {
    expect(paragraphStyle(plain("t")).marginLeft).toBe(cmPx(0));
    expect(paragraphStyle(plain("t")).textIndent).toBeUndefined();
    expect(paragraphStyle(plain("t", { level: 2 })).marginLeft).toBe(cmPx(2 * 1.27));
  });

  it("the paragraph's own indentCm/hangingCm win", () => {
    const style = paragraphStyle(disc("a", 3, { indentCm: 2, hangingCm: 1 }));
    expect(style.marginLeft).toBe(cmPx(2));
    expect(style.textIndent).toBe(cmPx(-1));
  });

  it("the marker box is hangingCm wide (at least), sized and coloured like the first run", () => {
    const [entry] = markParagraphList([{ runs: [{ text: "x", sizePt: 30, color: "#112233" }], bullet: "disc", hangingCm: 1 }]);
    const style = markerStyle(entry);
    expect(style.minWidth).toBe(cmPx(1));
    expect(style.display).toBe("inline-block");
    expect(style.textIndent).toBe("0");
    expect(style.fontSize).toBe("calc(40px * var(--fit-scale, 1))");
    expect(style.color).toBe("#112233");
  });

  it("alignment: the paragraph's own, else the frame's default (shape text centres)", () => {
    expect(paragraphStyle(plain("t")).textAlign).toBe("left");
    expect(paragraphStyle(plain("t"), "center").textAlign).toBe("center");
    expect(paragraphStyle(plain("t", { align: "right" }), "center").textAlign).toBe("right");
  });
});

describe("paragraph and run spacing (AC-6..10)", () => {
  it("single spacing is SPEC-079's 1.2, scaled by the body's line scale", () => {
    expect(paragraphLineHeight(plain("t"))).toBe("calc(1.2 * var(--line-scale, 1))");
    expect(paragraphLineHeight(plain("t", { lineSpacing: 1.5 }))).toBe("calc(1.8 * var(--line-scale, 1))");
  });

  it("exact point spacing is that pitch in px (24pt → 32px), on the runs too", () => {
    const para = plain("t", { lineSpacingPt: 24, lineSpacingRule: "exact" });
    expect(paragraphLineHeight(para)).toBe("calc(32px * var(--fit-scale, 1) * var(--line-scale, 1))");
    // Exact: the runs inherit the paragraph's pitch.
    expect(runLineHeight(para)).toBeUndefined();
    expect(runRenders(para, null)[0].style.lineHeight).toBeUndefined();
  });

  it("'at least' point spacing puts the pitch on the paragraph strut and gives runs their natural pitch", () => {
    const para = plain("t", { lineSpacingPt: 24 });
    expect(paragraphLineHeight(para)).toBe("calc(32px * var(--fit-scale, 1) * var(--line-scale, 1))");
    expect(runLineHeight(para)).toBe("calc(1.2 * var(--line-scale, 1))");
    expect(runRenders(para, null)[0].style.lineHeight).toBe("calc(1.2 * var(--line-scale, 1))");
  });

  it("space before/after render as pt → px margins (6pt → 8px), absent = none", () => {
    const style = paragraphStyle(plain("t", { spaceBeforePt: 6, spaceAfterPt: 12 }));
    expect(style.marginTop).toBe("calc(8px * var(--fit-scale, 1))");
    expect(style.marginBottom).toBe("calc(16px * var(--fit-scale, 1))");
    expect(paragraphStyle(plain("t")).marginTop).toBeUndefined();
  });

  it("font size goes through --fit-scale; letter spacing renders pt → px", () => {
    expect(runStyle({ text: "a" }).fontSize).toBe("calc(24px * var(--fit-scale, 1))");
    expect(runStyle({ text: "a", sizePt: 12 }).fontSize).toBe("calc(16px * var(--fit-scale, 1))");
    expect(runStyle({ text: "a", letterSpacingPt: 1.5 }).letterSpacing).toBe("calc(2px * var(--fit-scale, 1))");
    expect(runStyle({ text: "a", letterSpacingPt: -0.75 }).letterSpacing).toBe("calc(-1px * var(--fit-scale, 1))");
    expect(runStyle({ text: "a" }).letterSpacing).toBeUndefined();
  });
});

describe("links (AC-47)", () => {
  it("a run's own link wins; else its frame's; underlined in the run's own colour", () => {
    const frame = textFrame([], { linkSlide: 2 });
    const para: SlideParagraph = {
      runs: [
        { text: "own", color: "#aa0000", linkHref: "https://example.com/", linkTooltip: "Ex" },
        { text: "framed" },
      ],
    };
    const [own, framed] = runRenders(para, frame);
    expect(own.link).toEqual({ href: "https://example.com/", external: "https://example.com/", slide: 0, tooltip: "Ex" });
    expect(own.style.textDecoration).toBe("underline");
    expect(own.style.color).toBe("#aa0000");
    expect(framed.link).toEqual({ href: "#slide-2", external: "", slide: 2, tooltip: undefined });
    expect(runRenders({ runs: [{ text: "x" }] }, null)[0].link).toBeNull();
  });

  it("re-sanitises an external href at render; a dangerous one renders no link", () => {
    expect(linkRender({ href: "javascript:alert(1)" })).toBeNull();
    expect(linkRender({ slide: 0 })).toBeNull();
  });

  it("only text/image/shape frames render a frame link", () => {
    const group = { id: "g", kind: "group", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, children: [], linkSlide: 2 } as SlideFrame;
    expect(frameLinkRender(group)).toBeNull();
    expect(frameLinkRender(textFrame([], { linkHref: "https://a.example/" }))?.external).toBe("https://a.example/");
  });

  it("reads a rendered link element's target back, and follows an in-deck one with a bubbling event", () => {
    const el = document.createElement("a");
    el.setAttribute("data-testid", "slide-link");
    el.setAttribute("data-link-slide", "3");
    const inner = document.createElement("span");
    el.appendChild(inner);
    document.body.appendChild(el);
    expect(slideLinkElement(inner)).toBe(el);
    expect(slideLinkTargetOf(el)).toEqual({ slide: 3 });
    const seen: unknown[] = [];
    const listener = (event: Event) => seen.push((event as CustomEvent).detail);
    document.addEventListener(SLIDE_LINK_FOLLOW_EVENT, listener);
    expect(followEditorLink(el)).toBe(true);
    document.removeEventListener(SLIDE_LINK_FOLLOW_EVENT, listener);
    expect(seen).toEqual([{ slide: 3 }]);
    el.remove();
  });
});

const resolvePic = (id: string): string => (id === "pic" ? "data:image/png;base64,AA==" : "");
const mountFrame = (frame: SlideFrame, props: Record<string, unknown> = {}) =>
  mount(SlideFrameView, { props: { frame, resolveMedia: resolvePic, ...props }, attachTo: document.body });

describe("SlideFrameView text body (AC-11..15, 30, 47/48)", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("renders paragraphs inside a .text-body on a frame root carrying data-frame-id", () => {
    const wrapper = mountFrame(textFrame([num("a", 0, { startAt: 3 }), num("b")]));
    const root = wrapper.get('[data-testid="slide-frame"]');
    expect(root.attributes("data-frame-id")).toBe("t1");
    const body = wrapper.get(".text-body");
    expect(body.findAll('[data-testid="slide-paragraph"]')).toHaveLength(2);
    expect(wrapper.findAll(".bullet-marker").map((m) => m.text())).toEqual(["3.", "4."]);
    // The stored defaults: top anchor, the pptx insets, unscaled.
    const style = (body.element as HTMLElement).style;
    expect(style.getPropertyValue("justify-content")).toBe("flex-start");
    expect(style.getPropertyValue("--fit-scale")).toBe("1");
    expect(style.getPropertyValue("padding")).toBe(`${cmPx(0.127)} ${cmPx(0.254)} ${cmPx(0.127)} ${cmPx(0.254)}`);
    wrapper.unmount();
  });

  it("a thumbnail (not top-level) carries no data-frame-id", () => {
    const wrapper = mountFrame(textFrame([plain("x")]), { topLevel: false });
    expect(wrapper.find("[data-frame-id]").exists()).toBe(false);
    wrapper.unmount();
  });

  it("starts a shrink body from its STORED scale and never writes the model when nothing lays out", async () => {
    const frame = textFrame([plain("Big title")], { body: { autofit: "shrink", fontScalePct: 62.5, lineScalePct: 90 } });
    const before = JSON.stringify(frame);
    const wrapper = mountFrame(frame);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const style = (wrapper.get(".text-body").element as HTMLElement).style;
    expect(style.getPropertyValue("--fit-scale")).toBe("0.625");
    expect(style.getPropertyValue("--line-scale")).toBe("0.9");
    expect(JSON.stringify(frame)).toBe(before);
    wrapper.unmount();
  });

  it("re-renders a shrink body at its own measured scale, locally (the model is untouched)", async () => {
    // A synthetic layout: content is 4cm tall per 100% of font scale; the
    // frame offers 5cm − 0.254cm of insets, so the ladder settles on 85%.
    vi.spyOn(HTMLElement.prototype, "offsetHeight", "get").mockImplementation(function (this: HTMLElement) {
      const scale = Number(this.style.getPropertyValue("--fit-scale") || "1");
      return this.classList.contains("text-body") ? cmToPx(6 * scale) : 0;
    });
    const frame = textFrame([plain("Overflowing title")], { body: { autofit: "shrink" } });
    const before = JSON.stringify(frame);
    const wrapper = mountFrame(frame);
    await new Promise((resolve) => setTimeout(resolve, 0));
    await wrapper.vm.$nextTick();
    const style = (wrapper.get(".text-body").element as HTMLElement).style;
    // 6 × 0.775 = 4.65 ≤ 4.746, the first fitting step.
    expect(style.getPropertyValue("--fit-scale")).toBe("0.775");
    expect(JSON.stringify(frame)).toBe(before);
    wrapper.unmount();
  });

  it("anchors a shape's text middle by default and centres paragraphs with no alignment", () => {
    const shape: SlideFrame = {
      id: "s1",
      kind: "shape",
      geometry: "rect",
      xCm: 0,
      yCm: 0,
      widthCm: 4,
      heightCm: 2,
      paragraphs: [disc("x"), plain("y", { align: "left" })],
    };
    const wrapper = mountFrame(shape);
    const body = wrapper.get(".shape-text.text-body").element as HTMLElement;
    expect(body.style.getPropertyValue("justify-content")).toBe("center");
    const paras = wrapper.findAll(".paragraph").map((p) => (p.element as HTMLElement).style.textAlign);
    expect(paras).toEqual(["center", "left"]);
    expect(wrapper.findAll(".bullet-marker").map((m) => m.text())).toEqual(["•"]);
    wrapper.unmount();
  });

  it("renders a picture bullet as an <img> through the media resolver, falling back to • when unresolved", () => {
    const wrapper = mountFrame(textFrame([disc("a", 0, { bulletImageId: "pic" }), disc("b", 0, { bulletImageId: "gone" })]));
    const markersEls = wrapper.findAll(".bullet-marker");
    expect(markersEls[0].get("img.bullet-image").attributes("src")).toBe("data:image/png;base64,AA==");
    expect(markersEls[1].find("img").exists()).toBe(false);
    expect(markersEls[1].text()).toBe("•");
    wrapper.unmount();
  });

  it("renders a linked run as an inert <a data-testid=slide-link> on the editor stage; Cmd/Ctrl+click follows", async () => {
    const open = vi.spyOn(window, "open").mockReturnValue(null);
    const frame = textFrame([
      { runs: [{ text: "see " }, { text: "site", linkHref: "https://example.com/", linkTooltip: "Example" }, { text: " deck", linkSlide: 3 }] },
    ]);
    const wrapper = mountFrame(frame);
    const links = wrapper.findAll('[data-testid="slide-link"]');
    expect(links).toHaveLength(2);
    expect(links[0].element.tagName).toBe("A");
    expect(links[0].attributes("href")).toBe("https://example.com/");
    expect(links[0].attributes("title")).toBe("Example");
    expect(links[0].attributes("tabindex")).toBe("-1");
    expect(links[0].attributes("draggable")).toBe("false");
    expect((links[0].element as HTMLElement).style.textDecoration).toBe("underline");
    expect(links[1].attributes("href")).toBe("#slide-3");

    // A plain click is inert (default prevented, nothing followed).
    const plainClick = new MouseEvent("click", { bubbles: true, cancelable: true });
    links[0].element.dispatchEvent(plainClick);
    expect(plainClick.defaultPrevented).toBe(true);
    expect(open).not.toHaveBeenCalled();

    // Cmd+click opens an external link in a new tab, no opener.
    links[0].element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, metaKey: true }));
    expect(open).toHaveBeenCalledWith("https://example.com/", "_blank", "noopener,noreferrer");

    // Ctrl+click on an in-deck link dispatches the bubbling follow event.
    const followed: unknown[] = [];
    const listener = (event: Event) => followed.push((event as CustomEvent).detail);
    document.addEventListener(SLIDE_LINK_FOLLOW_EVENT, listener);
    links[1].element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, ctrlKey: true }));
    document.removeEventListener(SLIDE_LINK_FOLLOW_EVENT, listener);
    expect(followed).toEqual([{ slide: 3 }]);
    wrapper.unmount();
  });

  it("a linked image frame gets a whole-frame link overlay; playback links are focusable and never self-navigate", () => {
    const open = vi.spyOn(window, "open").mockReturnValue(null);
    const image: SlideFrame = { id: "i1", kind: "image", mediaId: "pic", xCm: 0, yCm: 0, widthCm: 2, heightCm: 2, linkSlide: 4 };
    const wrapper = mountFrame(image, { topLevel: false, playbackMode: true });
    const overlay = wrapper.get('a.frame-link[data-testid="slide-link"]');
    expect(overlay.attributes("data-link-slide")).toBe("4");
    expect(overlay.attributes("tabindex")).toBeUndefined();
    const click = new MouseEvent("click", { bubbles: true, cancelable: true, metaKey: true });
    overlay.element.dispatchEvent(click);
    expect(click.defaultPrevented).toBe(true);
    expect(open).not.toHaveBeenCalled();
    wrapper.unmount();
  });
});
