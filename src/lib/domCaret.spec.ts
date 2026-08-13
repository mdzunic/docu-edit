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

  // SPEC-051: a shape's text frame (and any other data-canvas-ignore subtree
  // — header/footer bands, note areas) renders its paragraphs with synthetic
  // section-0 addressing. Without the ignore marker they would shadow the
  // body's own blocks: the LAST DOM match on a data-slice-start tie is a
  // shape paragraph inside the frame, so the body caret at {0,0} would land
  // inside the shape. The fix asserts the real body block wins.
  it("does not let a synthetic block in a data-canvas-ignore subtree shadow a body block", () => {
    const root = document.createElement("div");
    const body = document.createElement("p");
    body.setAttribute("data-section", "0");
    body.setAttribute("data-block", "0");
    body.textContent = "X";
    const frame = document.createElement("div");
    frame.setAttribute("data-canvas-ignore", "1");
    const shadow = document.createElement("p");
    shadow.setAttribute("data-section", "0");
    shadow.setAttribute("data-block", "0");
    shadow.textContent = "shape";
    frame.appendChild(shadow);
    root.append(body, frame);
    document.body.appendChild(root);
    // Both carry data-slice-start 0 and the shadow sits later in the DOM;
    // without the ignore marker the last-match rule would pick the shadow.
    const point = domPointForCaret(root, { section: 0, block: 0, offset: 0 });
    expect(point?.node).toBe(body.firstChild);
    expect(point?.offset).toBe(0);
    setDomCaret(root, { section: 0, block: 0, offset: 0 });
    const range = modelRangeFromSelection(root, window.getSelection());
    expect(range).toEqual({
      start: { section: 0, block: 0, offset: 0 },
      end: { section: 0, block: 0, offset: 0 },
    });
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

// --- caret atoms (SPEC-031) --------------------------------------------------
// A note reference mark renders its computed number (any length) but stands
// for a one-character model placeholder, declared via data-caret-atom.

/** One paragraph: "ab" + a mark rendering "12" for model text "1" + "cd". */
function buildAtomContent(): HTMLElement {
  const root = document.createElement("div");
  const p = document.createElement("p");
  p.setAttribute("data-section", "0");
  p.setAttribute("data-block", "0");
  const before = document.createElement("span");
  before.textContent = "ab";
  const atom = document.createElement("span");
  atom.setAttribute("data-caret-atom", "1");
  atom.textContent = "12";
  const after = document.createElement("span");
  after.textContent = "cd";
  p.append(before, atom, after);
  root.appendChild(p);
  document.body.appendChild(root);
  return root;
}

describe("caret atoms (SPEC-031)", () => {
  it("counts an atom as one model character in caretFromDomPoint", () => {
    const root = buildAtomContent();
    const afterText = root.querySelectorAll("span")[2].firstChild!;
    // DOM point 1 char into "cd": rendered offset is 2+2+1, model 2+1+1.
    expect(caretFromDomPoint(root, afterText, 1)).toEqual({ section: 0, block: 0, offset: 4 });
    root.remove();
  });

  it("snaps a point inside the atom's rendered label to just after it", () => {
    const root = buildAtomContent();
    const atomText = root.querySelectorAll("span")[1].firstChild!;
    expect(caretFromDomPoint(root, atomText, 1)).toEqual({ section: 0, block: 0, offset: 3 });
    root.remove();
  });

  it("resolves model offsets across the atom in domPointForCaret", () => {
    const root = buildAtomContent();
    const spans = root.querySelectorAll("span");
    // Offset 2: end of "ab" (before the atom).
    expect(domPointForCaret(root, { section: 0, block: 0, offset: 2 })).toEqual({
      node: spans[0].firstChild,
      offset: 2,
    });
    // Offset 3: just after the atom — the start of "cd".
    expect(domPointForCaret(root, { section: 0, block: 0, offset: 3 })).toEqual({
      node: spans[2].firstChild,
      offset: 0,
    });
    // Offset 5: end of "cd".
    expect(domPointForCaret(root, { section: 0, block: 0, offset: 5 })).toEqual({
      node: spans[2].firstChild,
      offset: 2,
    });
    root.remove();
  });
});

describe("canvasModelText and renderedOffsetsToModel (SPEC-031)", () => {
  it("substitutes atoms' model text in canvasModelText", async () => {
    const { canvasModelText } = await import("./domCaret");
    const root = buildAtomContent();
    expect(canvasModelText(root)).toBe("ab1cd");
    root.remove();
  });

  it("maps rendered line offsets back to model space", async () => {
    const { renderedOffsetsToModel } = await import("./domCaret");
    const root = buildAtomContent();
    const p = root.querySelector("p")!;
    // Rendered text "ab12cd": offset 4 (after the label) is model 3; offset 3
    // (inside the label) snaps to the atom's start at model 2.
    expect(renderedOffsetsToModel(p, [0, 2, 4, 6])).toEqual([0, 2, 3, 5]);
    expect(renderedOffsetsToModel(p, [3])).toEqual([2]);
    root.remove();
  });
});
