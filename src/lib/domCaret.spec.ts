import { describe, it, expect } from "vitest";
import {
  caretFromDomPoint,
  domPointForCaret,
  extendDomSelection,
  modelRangeFromSelection,
  setDomCaret,
  setDomSelection,
} from "./domCaret";

/** Builds the paragraph DOM shape DocumentView renders: paragraphs with
 * data-section/data-block, runs as spans, empty paragraphs as a lone <br>. */
function buildContent(paragraphs: string[][]): HTMLElement {
  const root = document.createElement("div");
  paragraphs.forEach((runs, blockIndex) => {
    const p = document.createElement("p");
    p.setAttribute("data-section", "0");
    p.setAttribute("data-block", String(blockIndex));
    if (runs.length === 0) {
      p.appendChild(document.createElement("br"));
    } else {
      for (const text of runs) {
        const span = document.createElement("span");
        span.textContent = text;
        p.appendChild(span);
      }
    }
    root.appendChild(p);
  });
  document.body.appendChild(root);
  return root;
}

describe("caretFromDomPoint", () => {
  it("maps a point inside the first run", () => {
    const root = buildContent([["Hello ", "world"]]);
    const textNode = root.querySelectorAll("span")[0].firstChild!;
    expect(caretFromDomPoint(root, textNode, 3)).toEqual({
      section: 0,
      block: 0,
      offset: 3,
    });
    root.remove();
  });

  it("accumulates offsets across earlier runs", () => {
    const root = buildContent([["Hello ", "world"]]);
    const textNode = root.querySelectorAll("span")[1].firstChild!;
    expect(caretFromDomPoint(root, textNode, 2)).toEqual({
      section: 0,
      block: 0,
      offset: 8,
    });
    root.remove();
  });

  it("addresses the right paragraph", () => {
    const root = buildContent([["first"], ["second"]]);
    const textNode = root.querySelectorAll("span")[1].firstChild!;
    expect(caretFromDomPoint(root, textNode, 0)).toEqual({
      section: 0,
      block: 1,
      offset: 0,
    });
    root.remove();
  });

  it("returns null for a point outside the root", () => {
    const root = buildContent([["inside"]]);
    const outside = document.createElement("p");
    outside.setAttribute("data-section", "0");
    outside.setAttribute("data-block", "0");
    outside.textContent = "outside";
    document.body.appendChild(outside);
    expect(caretFromDomPoint(root, outside.firstChild!, 1)).toBeNull();
    outside.remove();
    root.remove();
  });
});

describe("setDomCaret / modelRangeFromSelection round-trip", () => {
  it("places the caret at a model offset and reads it back", () => {
    const root = buildContent([["Hello ", "world"]]);
    setDomCaret(root, { section: 0, block: 0, offset: 8 });
    const range = modelRangeFromSelection(root, window.getSelection());
    expect(range).toEqual({
      start: { section: 0, block: 0, offset: 8 },
      end: { section: 0, block: 0, offset: 8 },
    });
    root.remove();
  });

  it("places the caret in an empty paragraph", () => {
    const root = buildContent([["before"], []]);
    setDomCaret(root, { section: 0, block: 1, offset: 0 });
    const selection = window.getSelection()!;
    expect(selection.rangeCount).toBe(1);
    const paragraph = root.querySelector('[data-block="1"]')!;
    expect(paragraph.contains(selection.getRangeAt(0).startContainer)).toBe(true);
    root.remove();
  });

  it("clamps an offset past the end of the paragraph to its end", () => {
    const root = buildContent([["abc"]]);
    setDomCaret(root, { section: 0, block: 0, offset: 99 });
    const range = modelRangeFromSelection(root, window.getSelection());
    expect(range?.end).toEqual({ section: 0, block: 0, offset: 3 });
    root.remove();
  });

  it("does nothing for a paragraph that is not rendered", () => {
    const root = buildContent([["abc"]]);
    window.getSelection()?.removeAllRanges();
    setDomCaret(root, { section: 0, block: 5, offset: 0 });
    expect(window.getSelection()?.rangeCount ?? 0).toBe(0);
    root.remove();
  });
});

describe("domPointForCaret", () => {
  it("resolves an offset spanning runs to the containing text node", () => {
    const root = buildContent([["Hello ", "world"]]);
    const point = domPointForCaret(root, { section: 0, block: 0, offset: 8 });
    expect(point?.node).toBe(root.querySelectorAll("span")[1].firstChild);
    expect(point?.offset).toBe(2);
    root.remove();
  });

  it("clamps an overshooting offset to the last text node's end", () => {
    const root = buildContent([["abc"]]);
    const point = domPointForCaret(root, { section: 0, block: 0, offset: 99 });
    expect(point?.node).toBe(root.querySelector("span")!.firstChild);
    expect(point?.offset).toBe(3);
    root.remove();
  });

  it("resolves an empty paragraph to the paragraph element itself", () => {
    const root = buildContent([[]]);
    const point = domPointForCaret(root, { section: 0, block: 0, offset: 0 });
    expect(point?.node).toBe(root.querySelector('[data-block="0"]'));
    expect(point?.offset).toBe(0);
    root.remove();
  });

  it("returns null for a paragraph that is not rendered", () => {
    const root = buildContent([["abc"]]);
    expect(domPointForCaret(root, { section: 0, block: 7, offset: 0 })).toBeNull();
    root.remove();
  });
});

describe("setDomSelection / extendDomSelection", () => {
  it("spans two model positions across paragraphs", () => {
    const root = buildContent([["first"], ["second"]]);
    setDomSelection(
      root,
      { section: 0, block: 0, offset: 2 },
      { section: 0, block: 1, offset: 4 },
    );
    const range = modelRangeFromSelection(root, window.getSelection());
    expect(range).toEqual({
      start: { section: 0, block: 0, offset: 2 },
      end: { section: 0, block: 1, offset: 4 },
    });
    root.remove();
  });

  it("keeps the anchor fixed so the focus end stays extendable", () => {
    const root = buildContent([["abcdef"]]);
    setDomSelection(
      root,
      { section: 0, block: 0, offset: 1 },
      { section: 0, block: 0, offset: 3 },
    );
    const selection = window.getSelection()!;
    expect(selection.anchorOffset).toBe(1);
    expect(selection.focusOffset).toBe(3);
    root.remove();
  });

  it("extendDomSelection moves only the focus end", () => {
    const root = buildContent([["first"], ["second"]]);
    setDomCaret(root, { section: 0, block: 0, offset: 2 });
    extendDomSelection(root, { section: 0, block: 1, offset: 3 });
    const range = modelRangeFromSelection(root, window.getSelection());
    expect(range).toEqual({
      start: { section: 0, block: 0, offset: 2 },
      end: { section: 0, block: 1, offset: 3 },
    });
    expect(window.getSelection()!.anchorOffset).toBe(2);
    root.remove();
  });

  it("extendDomSelection with no existing selection places the caret", () => {
    const root = buildContent([["abc"]]);
    window.getSelection()?.removeAllRanges();
    extendDomSelection(root, { section: 0, block: 0, offset: 1 });
    const range = modelRangeFromSelection(root, window.getSelection());
    expect(range?.end).toEqual({ section: 0, block: 0, offset: 1 });
    root.remove();
  });
});
