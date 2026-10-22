import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import EquationView from "./EquationView.vue";
import type { EquationSpec } from "../lib/documentModel";
import { parseLinear } from "../lib/math/linearParse";

function eq(src: string, display: "inline" | "block" = "inline"): EquationSpec {
  return { display, body: parseLinear(src) };
}

describe("EquationView (SPEC-056 rendering)", () => {
  it("renders identifiers italic and numbers upright (AC-8/9)", () => {
    const w = mount(EquationView, { props: { equation: eq("x2") } });
    const glyphs = w.findAll(".math-glyph");
    expect(glyphs[0].attributes("style")).toContain("italic");
    expect(glyphs[1].attributes("style")).toContain("normal");
  });

  it("renders a fraction with a bar element (AC-42) as selectable DOM text (AC-29)", () => {
    const w = mount(EquationView, { props: { equation: eq("\\frac{a}{b}") } });
    expect(w.find(".math-frac-bar").exists()).toBe(true);
    // The numerator/denominator glyphs are real text nodes, not an image.
    expect(w.get(".math-frac-num").text()).toBe("a");
    expect(w.get(".math-frac-den").text()).toBe("b");
    expect(w.find("img").exists()).toBe(false);
  });

  it("draws \\binom without a bar (noBar)", () => {
    const w = mount(EquationView, { props: { equation: eq("\\binom{n}{k}") } });
    expect(w.find(".math-frac-nobar").exists()).toBe(true);
  });

  it("exposes the alt text to assistive technology (AC-21) and defaults it to the linear form (AC-19)", () => {
    const w = mount(EquationView, { props: { equation: eq("\\frac{a}{b}") } });
    expect(w.get('[role="math"]').attributes("aria-label")).toBe("\\frac{a}{b}");
  });

  it("uses an explicit altText when set (AC-20)", () => {
    const w = mount(EquationView, {
      props: { equation: { display: "inline", body: parseLinear("x"), altText: "the unknown x" } },
    });
    expect(w.get('[role="math"]').attributes("aria-label")).toBe("the unknown x");
  });

  it("shows the read-only box with the fixed message for an unsupported construct (AC-24)", () => {
    const w = mount(EquationView, {
      props: {
        equation: { display: "inline", body: [{ kind: "opaqueMath", source: "<x/>", linear: "f(x)" }] },
      },
    });
    const box = w.get('[data-testid="equation-unsupported"]');
    expect(box.text()).toContain("f(x)");
    expect(box.text()).toContain("We can't edit this equation yet — it's kept exactly as it was");
  });

  it("centres a display equation on its own line (AC-2)", () => {
    const w = mount(EquationView, { props: { equation: eq("x", "block") } });
    expect(w.get('[role="math"]').classes()).toContain("equation-block");
  });

  it("renders Greek glyphs from the SPEC-026 catalogue as text (AC-28)", () => {
    const w = mount(EquationView, { props: { equation: eq("\\alpha") } });
    expect(w.get(".math-glyph").text()).toBe("α");
  });
});
