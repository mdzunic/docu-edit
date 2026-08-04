import { describe, it, expect, afterEach } from "vitest";
import {
  caretAtClientPoint,
  caretClientRect,
  caretNearClientPoint,
} from "./caretGeometry";

// jsdom has no layout: every rect is zero and caret-from-point hit testing
// does not exist. These tests stub exactly those seams — element rects and
// document.caretRangeFromPoint — to drive the geometry logic.

interface RectSpec {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

function stubRect(element: Element, rect: RectSpec): void {
  element.getBoundingClientRect = () =>
    new DOMRect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
}

/** Two stacked paragraphs ("Hello world", "Second paragraph") with stubbed
 * layout: the content area spans x 100–500, paragraph bands y 100–120 and
 * y 120–140. */
function buildStubbedContent(): HTMLElement {
  const root = document.createElement("div");
  const texts = ["Hello world", "Second paragraph"];
  texts.forEach((text, blockIndex) => {
    const p = document.createElement("p");
    p.setAttribute("data-section", "0");
    p.setAttribute("data-block", String(blockIndex));
    const span = document.createElement("span");
    span.textContent = text;
    p.appendChild(span);
    root.appendChild(p);
  });
  document.body.appendChild(root);
  stubRect(root, { top: 100, bottom: 800, left: 100, right: 500 });
  stubRect(root.children[0], { top: 100, bottom: 120, left: 100, right: 500 });
  stubRect(root.children[1], { top: 120, bottom: 140, left: 100, right: 500 });
  return root;
}

// jsdom's document has no own caretRangeFromPoint; tests install one and
// remove it again (through a loose cast — lib.dom types it non-optional).
const caretRangeDocument = document as Document & {
  caretRangeFromPoint: (x: number, y: number) => Range | null;
};

afterEach(() => {
  delete (document as unknown as Record<string, unknown>).caretRangeFromPoint;
  document.body.innerHTML = "";
});

describe("caretNearClientPoint", () => {
  it("uses the browser's caret hit testing when available", () => {
    const root = buildStubbedContent();
    const textNode = root.querySelectorAll("span")[0].firstChild!;
    caretRangeDocument.caretRangeFromPoint = () => {
      const range = document.createRange();
      range.setStart(textNode, 4);
      return range;
    };
    expect(caretNearClientPoint(root, 150, 110)).toEqual({
      section: 0,
      block: 0,
      offset: 4,
    });
  });

  it("clamps the hit-test point into the content area and paragraph band", () => {
    const root = buildStubbedContent();
    const seen: Array<[number, number]> = [];
    caretRangeDocument.caretRangeFromPoint = (x, y) => {
      seen.push([x, y]);
      return null;
    };
    // Far left of the page, above the first paragraph.
    caretNearClientPoint(root, -300, 20);
    expect(seen[0][0]).toBeGreaterThanOrEqual(100);
    expect(seen[0][1]).toBeGreaterThanOrEqual(100);
    // Far right, below everything: clamped into the last paragraph's band.
    caretNearClientPoint(root, 900, 700);
    expect(seen[1][0]).toBeLessThanOrEqual(500);
    expect(seen[1][1]).toBeLessThanOrEqual(140);
  });

  it("falls back to the nearest paragraph's start above the content (AC-1)", () => {
    const root = buildStubbedContent();
    expect(caretNearClientPoint(root, 110, 10)).toEqual({
      section: 0,
      block: 0,
      offset: 0,
    });
  });

  it("falls back to the nearest paragraph's end below the content (AC-1)", () => {
    const root = buildStubbedContent();
    expect(caretNearClientPoint(root, 110, 700)).toEqual({
      section: 0,
      block: 1,
      offset: "Second paragraph".length,
    });
  });

  it("clicking beside a line resolves to that line's paragraph (AC-1)", () => {
    const root = buildStubbedContent();
    // Left page margin beside the second paragraph: its start.
    expect(caretNearClientPoint(root, 20, 130)).toEqual({
      section: 0,
      block: 1,
      offset: 0,
    });
    // Right page margin beside the first paragraph: its end.
    expect(caretNearClientPoint(root, 640, 110)).toEqual({
      section: 0,
      block: 0,
      offset: "Hello world".length,
    });
  });

  it("returns null when the canvas has no paragraphs", () => {
    const root = document.createElement("div");
    document.body.appendChild(root);
    expect(caretNearClientPoint(root, 10, 10)).toBeNull();
  });
});

describe("caretAtClientPoint", () => {
  it("returns the hit-tested caret inside the root", () => {
    const root = buildStubbedContent();
    const textNode = root.querySelectorAll("span")[1].firstChild!;
    caretRangeDocument.caretRangeFromPoint = () => {
      const range = document.createRange();
      range.setStart(textNode, 3);
      return range;
    };
    expect(caretAtClientPoint(root, 150, 130)).toEqual({
      section: 0,
      block: 1,
      offset: 3,
    });
  });

  it("returns null when the point resolves outside the root", () => {
    const root = buildStubbedContent();
    const outside = document.createElement("p");
    outside.textContent = "elsewhere";
    document.body.appendChild(outside);
    caretRangeDocument.caretRangeFromPoint = () => {
      const range = document.createRange();
      range.setStart(outside.firstChild!, 0);
      return range;
    };
    expect(caretAtClientPoint(root, 150, 130)).toBeNull();
  });

  it("returns null when hit testing is unavailable (jsdom)", () => {
    const root = buildStubbedContent();
    expect(caretAtClientPoint(root, 150, 130)).toBeNull();
  });
});

describe("caretClientRect", () => {
  it("returns null when nothing is measurable (jsdom's zero rects)", () => {
    const root = document.createElement("div");
    const p = document.createElement("p");
    p.setAttribute("data-section", "0");
    p.setAttribute("data-block", "0");
    p.appendChild(document.createElement("br"));
    root.appendChild(p);
    document.body.appendChild(root);
    expect(caretClientRect(root, { section: 0, block: 0, offset: 0 })).toBeNull();
  });

  it("falls back to the containing element's rect when the range has none", () => {
    const root = buildStubbedContent();
    const span = root.querySelectorAll("span")[0];
    stubRect(span, { top: 101, bottom: 119, left: 100, right: 180 });
    const rect = caretClientRect(root, { section: 0, block: 0, offset: 4 });
    expect(rect?.top).toBe(101);
    expect(rect?.left).toBe(100);
  });

  it("returns null for a paragraph that is not rendered", () => {
    const root = buildStubbedContent();
    expect(caretClientRect(root, { section: 0, block: 9, offset: 0 })).toBeNull();
  });
});
