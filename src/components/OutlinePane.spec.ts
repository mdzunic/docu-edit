import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import OutlinePane from "./OutlinePane.vue";
import { buildOutline, OUTLINE_EMPTY_TEXT, type OutlineNode } from "../lib/outline";
import type { Block, DocumentModel, Paragraph } from "../lib/documentModel";

function para(text: string, headingLevel?: number): Paragraph {
  const p: Paragraph = { type: "paragraph", runs: [{ text, formatting: {} }] };
  if (headingLevel !== undefined) {
    p.headingLevel = headingLevel;
  }
  return p;
}

function outlineFor(blocks: Block[]): OutlineNode[] {
  const model = {
    id: "t",
    title: "t",
    sections: [{ blocks }],
  } as unknown as DocumentModel;
  return buildOutline(model);
}

function anchorFor(outline: OutlineNode[], text: string): string {
  const found = outline
    .flatMap(function flat(node): OutlineNode[] {
      return [node, ...node.children.flatMap(flat)];
    })
    .find((node) => node.text === text);
  if (!found) {
    throw new Error(`no ${text}`);
  }
  return found.anchor;
}

const SAMPLE = [para("A", 1), para("body"), para("A.1", 2), para("B", 1)];

function mountPane(props = {}) {
  return mount(OutlinePane, {
    props: {
      outline: outlineFor(SAMPLE),
      collapsed: new Set(),
      activeAnchor: undefined,
      ...props,
    },
  });
}

describe("OutlinePane", () => {
  it("is a navigation landmark containing a tree (AC-20)", () => {
    const wrapper = mountPane();
    const nav = wrapper.get('[data-testid="outline-pane"]');
    expect(nav.element.tagName).toBe("NAV");
    const tree = wrapper.get('[data-testid="outline-tree"]');
    expect(tree.attributes("role")).toBe("tree");
    expect(tree.attributes("tabindex")).toBe("0");
  });

  it("lists Heading 1-6 as treeitems indented by depth (AC-1, AC-2)", () => {
    const wrapper = mountPane();
    const items = wrapper.findAll('[role="treeitem"]');
    expect(wrapper.findAll(".outline-label").map((l) => l.text())).toEqual(["A", "A.1", "B"]);
    // A.1 is one level deeper than A/B.
    expect(items[0].attributes("aria-level")).toBe("1");
    expect(items[1].attributes("aria-level")).toBe("2");
    expect(items[2].attributes("aria-level")).toBe("1");
  });

  it("shows the empty-state message when there are no headings (AC-12)", () => {
    const wrapper = mountPane({ outline: outlineFor([para("just body")]) });
    expect(wrapper.get('[data-testid="outline-empty"]').text()).toBe(OUTLINE_EMPTY_TEXT);
    expect(wrapper.find('[data-testid="outline-tree"]').exists()).toBe(false);
  });

  it("highlights the heading containing the caret (AC-3)", () => {
    const outline = outlineFor(SAMPLE);
    const wrapper = mountPane({ outline, activeAnchor: anchorFor(outline, "A.1") });
    const active = wrapper.get(`[data-testid="outline-item-${anchorFor(outline, "A.1")}"]`);
    expect(active.attributes("aria-current")).toBe("true");
    expect(active.classes()).toContain("is-active");
  });

  it("navigates on click (AC-5, AC-6)", async () => {
    const outline = outlineFor(SAMPLE);
    const wrapper = mountPane({ outline });
    await wrapper.get(`[data-testid="outline-item-${anchorFor(outline, "B")}"]`).trigger("click");
    expect(wrapper.emitted("navigate")).toEqual([[anchorFor(outline, "B")]]);
  });

  it("toggles a node with children via its chevron (AC-7, AC-8)", async () => {
    const outline = outlineFor(SAMPLE);
    const wrapper = mountPane({ outline });
    const aAnchor = anchorFor(outline, "A");
    await wrapper.get(`[data-testid="outline-twisty-${aAnchor}"]`).trigger("click");
    expect(wrapper.emitted("toggle")).toEqual([[aAnchor]]);
    // A leaf has no chevron.
    expect(wrapper.find(`[data-testid="outline-twisty-${anchorFor(outline, "B")}"]`).exists()).toBe(
      false,
    );
  });

  it("hides the descendants of a collapsed node (AC-8)", () => {
    const outline = outlineFor(SAMPLE);
    const wrapper = mountPane({ outline, collapsed: new Set([anchorFor(outline, "A")]) });
    expect(wrapper.findAll(".outline-label").map((l) => l.text())).toEqual(["A", "B"]);
  });

  it("moves the selection with arrow keys and jumps on Enter (AC-18, AC-19)", async () => {
    const outline = outlineFor(SAMPLE);
    const wrapper = mountPane({ outline });
    const tree = wrapper.get('[data-testid="outline-tree"]');
    await tree.trigger("keydown", { key: "ArrowDown" });
    await tree.trigger("keydown", { key: "ArrowDown" });
    // First row is pre-selected; two ArrowDowns land on the third row ("B").
    expect(tree.attributes("aria-activedescendant")).toBe(`outline-item-${anchorFor(outline, "B")}`);
    await tree.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("navigate")).toEqual([[anchorFor(outline, "B")]]);
  });

  it("emits a move when a heading is dragged before a sibling (AC-9)", async () => {
    // Two H1 siblings so the move is valid.
    const outline = outlineFor([para("A", 1), para("B", 1)]);
    const wrapper = mountPane({ outline });
    const aAnchor = anchorFor(outline, "A");
    const bAnchor = anchorFor(outline, "B");
    const target = wrapper.get(`[data-testid="outline-item-${bAnchor}"]`);
    target.element.getBoundingClientRect = () =>
      ({ top: 0, height: 20, left: 0, right: 0, bottom: 20, width: 0, x: 0, y: 0, toJSON() {} }) as DOMRect;
    await wrapper
      .get(`[data-testid="outline-item-${aAnchor}"]`)
      .trigger("dragstart", { dataTransfer: { setData() {} } });
    // clientY near the top third => "before".
    await target.trigger("dragover", { clientY: 2, dataTransfer: {} });
    await target.trigger("drop", { clientY: 2, dataTransfer: {} });
    expect(wrapper.emitted("move")).toEqual([[{ source: aAnchor, target: bAnchor, position: "before" }]]);
  });

  it("rejects a drop onto a heading of a different level (AC-11)", async () => {
    const outline = outlineFor(SAMPLE); // A(1), A.1(2), B(1)
    const wrapper = mountPane({ outline });
    const source = anchorFor(outline, "A"); // level 1
    const targetAnchor = anchorFor(outline, "A.1"); // level 2
    const target = wrapper.get(`[data-testid="outline-item-${targetAnchor}"]`);
    target.element.getBoundingClientRect = () =>
      ({ top: 0, height: 20, left: 0, right: 0, bottom: 20, width: 0, x: 0, y: 0, toJSON() {} }) as DOMRect;
    await wrapper
      .get(`[data-testid="outline-item-${source}"]`)
      .trigger("dragstart", { dataTransfer: { setData() {} } });
    await target.trigger("dragover", { clientY: 2, dataTransfer: {} });
    await target.trigger("drop", { clientY: 2, dataTransfer: {} });
    expect(wrapper.emitted("move")).toBeUndefined();
  });
});
