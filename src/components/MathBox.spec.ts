// SPEC-056 — a stretchy delimiter scales by CSS transform up to 3x and beyond
// that is assembled from Unicode bracket pieces U+239B–U+23AD (AC-44).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import MathBox from "./MathBox.vue";
import { layoutRow } from "../lib/math/render";
import type { MathNode } from "../lib/documentModel";

function matrixDelim(rows: number): MathNode {
  const cells = Array.from({ length: rows }, () => [[]]);
  return { kind: "delim", open: "(", close: ")", body: [{ kind: "matrix", rows: cells }] };
}

describe("MathBox stretchy delimiters (AC-44)", () => {
  it("uses a CSS transform for a delimiter up to 3x", () => {
    const wrapper = mount(MathBox, { props: { box: layoutRow([matrixDelim(3)]) } });
    expect(wrapper.find(".math-delim-piece").exists()).toBe(false);
    const open = wrapper.find(".math-delim-open");
    expect(open.attributes("style")).toContain("scaleY(3)");
  });

  it("assembles a taller delimiter from Unicode bracket pieces", () => {
    const wrapper = mount(MathBox, { props: { box: layoutRow([matrixDelim(6)]) } });
    // Both the open and close delimiters become a stacked group of pieces.
    const groups = wrapper.findAll(".math-delim-pieces");
    expect(groups.length).toBe(2);
    const openPieces = groups[0].findAll(".math-delim-piece");
    expect(openPieces.length).toBeGreaterThan(2);
    // Top hook of the left parenthesis: U+239B; bottom hook: U+239D.
    expect(openPieces[0].text()).toBe("⎛");
    expect(openPieces[openPieces.length - 1].text()).toBe("⎝");
  });
});
