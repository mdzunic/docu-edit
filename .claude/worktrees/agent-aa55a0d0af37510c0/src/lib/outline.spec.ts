import { describe, expect, it } from "vitest";
import type { Block, DocumentModel, Paragraph, Section } from "./documentModel";
import { blockText } from "./textEditing";
import {
  activeHeadingAnchor,
  buildOutline,
  canMoveHeading,
  headingContentLength,
  hiddenBlockAddresses,
  moveHeading,
  visibleOutlineRows,
  type OutlineNode,
} from "./outline";

function para(text: string, headingLevel?: number): Paragraph {
  const paragraph: Paragraph = { type: "paragraph", runs: [{ text, formatting: {} }] };
  if (headingLevel !== undefined) {
    paragraph.headingLevel = headingLevel;
  }
  return paragraph;
}

function model(blocks: Block[], extra: Partial<Section> = {}): DocumentModel {
  return {
    id: "test",
    title: "Test",
    sections: [{ blocks, ...extra }],
    // Only the fields outline.ts touches matter here.
  } as unknown as DocumentModel;
}

function anchorOfHeading(doc: DocumentModel, text: string): string {
  const match = buildOutline(doc)
    .flatMap(function flatten(node): OutlineNode[] {
      return [node, ...node.children.flatMap(flatten)];
    })
    .find((node) => node.text === text);
  if (!match) {
    throw new Error(`no heading "${text}"`);
  }
  return match.anchor;
}

describe("buildOutline", () => {
  it("nests headings by level", () => {
    const doc = model([
      para("A", 1),
      para("body"),
      para("A.1", 2),
      para("A.1.1", 3),
      para("A.2", 2),
      para("B", 1),
    ]);
    const outline = buildOutline(doc);
    expect(outline.map((node) => node.text)).toEqual(["A", "B"]);
    expect(outline[0].children.map((node) => node.text)).toEqual(["A.1", "A.2"]);
    expect(outline[0].children[0].children.map((node) => node.text)).toEqual(["A.1.1"]);
    expect(outline[1].children).toEqual([]);
  });

  it("nests a skipped level under the nearest shallower heading", () => {
    // An H1 directly followed by an H3 (no H2) still nests the H3 under it.
    const doc = model([para("Top", 1), para("Deep", 3)]);
    const outline = buildOutline(doc);
    expect(outline.map((node) => node.text)).toEqual(["Top"]);
    expect(outline[0].children.map((node) => node.text)).toEqual(["Deep"]);
  });

  it("ignores non-heading paragraphs", () => {
    expect(buildOutline(model([para("just body"), para("more")]))).toEqual([]);
  });
});

describe("visibleOutlineRows", () => {
  const doc = model([para("A", 1), para("A.1", 2), para("A.1.1", 3), para("B", 1)]);
  const outline = buildOutline(doc);

  it("returns every row with tree depth when nothing is collapsed", () => {
    const rows = visibleOutlineRows(outline, new Set());
    expect(rows.map((row) => [row.node.text, row.depth])).toEqual([
      ["A", 0],
      ["A.1", 1],
      ["A.1.1", 2],
      ["B", 0],
    ]);
    expect(rows[0].hasChildren).toBe(true);
    expect(rows[3].hasChildren).toBe(false);
  });

  it("omits the descendants of a collapsed node", () => {
    const collapsed = new Set([anchorOfHeading(doc, "A")]);
    const rows = visibleOutlineRows(outline, collapsed);
    expect(rows.map((row) => row.node.text)).toEqual(["A", "B"]);
    expect(rows[0].collapsed).toBe(true);
  });

  it("only marks a node collapsed when it actually has children", () => {
    const collapsed = new Set([anchorOfHeading(doc, "B")]);
    const rows = visibleOutlineRows(outline, collapsed);
    expect(rows.find((row) => row.node.text === "B")?.collapsed).toBe(false);
  });
});

describe("activeHeadingAnchor", () => {
  const doc = model([para("Intro", 1), para("body"), para("Details", 2), para("more")]);

  it("is null when the caret is before the first heading", () => {
    // No heading precedes a caret in a document whose first block is body text.
    const bodyFirst = model([para("body"), para("Later", 1)]);
    expect(activeHeadingAnchor(bodyFirst, { section: 0, block: 0, offset: 0 })).toBeNull();
  });

  it("is null when there is no caret", () => {
    expect(activeHeadingAnchor(doc, null)).toBeNull();
  });

  it("resolves the heading the caret sits under", () => {
    expect(activeHeadingAnchor(doc, { section: 0, block: 1, offset: 2 })).toBe(
      anchorOfHeading(doc, "Intro"),
    );
    expect(activeHeadingAnchor(doc, { section: 0, block: 3, offset: 0 })).toBe(
      anchorOfHeading(doc, "Details"),
    );
  });

  it("treats a caret on the heading line as inside that heading", () => {
    expect(activeHeadingAnchor(doc, { section: 0, block: 2, offset: 1 })).toBe(
      anchorOfHeading(doc, "Details"),
    );
  });
});

describe("headingContentLength", () => {
  const doc = model([para("A", 1), para("a-body"), para("A.1", 2), para("B", 1)]);

  it("counts the content blocks a heading owns", () => {
    expect(headingContentLength(doc, 0, 0)).toBe(2); // a-body + A.1
    expect(headingContentLength(doc, 0, 2)).toBe(0); // A.1 owns nothing before B
  });

  it("is 0 for a non-heading block", () => {
    expect(headingContentLength(doc, 0, 1)).toBe(0);
  });
});

describe("hiddenBlockAddresses", () => {
  const doc = model([
    para("A", 1), // 0
    para("a-body"), // 1
    para("A.1", 2), // 2
    para("a1-body"), // 3
    para("B", 1), // 4
    para("b-body"), // 5
  ]);

  it("hides a heading's content but not the heading itself", () => {
    const hidden = hiddenBlockAddresses(doc, new Set(["0:2"]));
    expect([...hidden]).toEqual(["0:3"]);
  });

  it("hides the whole subtree under a collapsed top-level heading", () => {
    const hidden = hiddenBlockAddresses(doc, new Set(["0:0"]));
    // Everything from just after A up to (not including) B.
    expect([...hidden].toSorted()).toEqual(["0:1", "0:2", "0:3"]);
  });

  it("ignores a stale address that no longer points at a heading", () => {
    expect(hiddenBlockAddresses(doc, new Set(["0:1"])).size).toBe(0);
    expect(hiddenBlockAddresses(doc, new Set(["9:9"])).size).toBe(0);
  });

  it("returns an empty set when nothing is collapsed", () => {
    expect(hiddenBlockAddresses(doc, new Set()).size).toBe(0);
  });
});

describe("moveHeading", () => {
  function siblings(): DocumentModel {
    return model([
      para("A", 1), // 0
      para("a-body"), // 1
      para("B", 1), // 2
      para("b-body"), // 3
      para("C", 1), // 4
      para("c-body"), // 5
    ]);
  }

  it("moves a heading and its content among siblings as one result", () => {
    const doc = siblings();
    const result = moveHeading(doc, anchorOfHeading(doc, "C"), anchorOfHeading(doc, "A"), "before");
    expect(result).not.toBeNull();
    const texts = result!.model.sections[0].blocks.map((block) => blockText(block));
    expect(texts).toEqual(["C", "c-body", "A", "a-body", "B", "b-body"]);
    // The caret lands at the moved heading's new start.
    expect(result!.selection.focus).toEqual({ section: 0, block: 0, offset: 0 });
  });

  it("moves a heading with its whole subtree", () => {
    const doc = model([
      para("A", 1), // 0
      para("A.1", 2), // 1
      para("a1-body"), // 2
      para("B", 1), // 3
    ]);
    const result = moveHeading(doc, anchorOfHeading(doc, "A"), anchorOfHeading(doc, "B"), "after");
    const texts = result!.model.sections[0].blocks.map((block) => blockText(block));
    expect(texts).toEqual(["B", "A", "A.1", "a1-body"]);
  });

  it("does not mutate the input model", () => {
    const doc = siblings();
    const before = doc.sections[0].blocks.map((block) => blockText(block));
    moveHeading(doc, anchorOfHeading(doc, "C"), anchorOfHeading(doc, "A"), "before");
    expect(doc.sections[0].blocks.map((block) => blockText(block))).toEqual(before);
  });

  it("rejects dropping onto a heading of a different level", () => {
    const doc = model([para("A", 1), para("A.1", 2), para("B", 1)]);
    expect(
      moveHeading(doc, anchorOfHeading(doc, "A.1"), anchorOfHeading(doc, "B"), "before"),
    ).toBeNull();
    expect(canMoveHeading(doc, anchorOfHeading(doc, "A.1"), anchorOfHeading(doc, "B"), "before")).toBe(
      false,
    );
  });

  it("rejects dropping onto a non-sibling of the same level", () => {
    // Two H2s under different H1 parents are the same level but not siblings.
    const doc = model([para("A", 1), para("A.1", 2), para("B", 1), para("B.1", 2)]);
    expect(
      moveHeading(doc, anchorOfHeading(doc, "A.1"), anchorOfHeading(doc, "B.1"), "before"),
    ).toBeNull();
  });

  it("rejects dropping a heading onto itself", () => {
    const doc = siblings();
    expect(moveHeading(doc, anchorOfHeading(doc, "A"), anchorOfHeading(doc, "A"), "before")).toBeNull();
  });

  it("rejects a no-op reorder to the same place", () => {
    const doc = siblings();
    // Dropping A "before" B when A already immediately precedes B changes nothing.
    expect(moveHeading(doc, anchorOfHeading(doc, "A"), anchorOfHeading(doc, "B"), "before")).toBeNull();
  });

  it("re-anchors preserved opaque block XML across the move", () => {
    const doc = model(
      [para("A", 1), para("a-body"), para("B", 1), para("b-body")],
      { opaqueBlocks: [{ blockIndex: 2, xml: "<opaque/>" }] },
    );
    // The opaque node sits before block 2 (heading "B"); after moving B before
    // A it must still sit before B — now at index 0.
    const result = moveHeading(doc, anchorOfHeading(doc, "B"), anchorOfHeading(doc, "A"), "before");
    expect(result!.model.sections[0].opaqueBlocks).toEqual([{ blockIndex: 0, xml: "<opaque/>" }]);
  });
});
