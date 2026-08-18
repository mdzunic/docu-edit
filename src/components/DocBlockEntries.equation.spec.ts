// SPEC-056 — an equation run renders through the paint path as a laid-out,
// selectable DOM equation (not the "1" placeholder), as a caret atom.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import DocBlockEntries from "./DocBlockEntries.vue";
import type { RenderEntry } from "../lib/pageRender";
import { parseLinear } from "../lib/math/linearParse";

function paragraphEntry(): RenderEntry {
  return {
    key: "p0",
    memo: [],
    kind: "paragraph",
    tag: "p",
    runs: [
      {
        key: "r0",
        href: "",
        styleObj: {},
        text: "",
        segments: null,
        atomText: "1",
        equation: { display: "inline", body: parseLinear("\\frac{a}{b}") },
        lang: "en-US",
      },
    ],
  } as RenderEntry;
}

describe("DocBlockEntries equation rendering (SPEC-056)", () => {
  it("renders the equation and not the placeholder character", () => {
    const w = mount(DocBlockEntries, { props: { entries: [paragraphEntry()] } });
    expect(w.findComponent({ name: "EquationView" }).exists()).toBe(true);
    expect(w.find(".math-frac-bar").exists()).toBe(true);
    expect(w.get(".doc-equation").text()).not.toBe("1");
  });

  it("marks the equation as a contenteditable=false caret atom of length one", () => {
    const w = mount(DocBlockEntries, { props: { entries: [paragraphEntry()] } });
    const atom = w.get(".doc-equation");
    expect(atom.attributes("contenteditable")).toBe("false");
    expect(atom.attributes("data-caret-atom")).toBe("1");
  });
});
