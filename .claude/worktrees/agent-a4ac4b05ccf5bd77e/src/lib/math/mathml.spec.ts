// SPEC-056 — Presentation-MathML serialiser/parser reversibility. For every
// construct in the closed set, `parseMathML(serializeMathML(eq))` must
// reproduce `eq`'s model exactly (structural equality ignoring `id`s).
import { describe, expect, it } from "vitest";
import { MATHML_NAMESPACE, parseMathML, serializeMathML } from "./mathml";
import { parseLinear } from "./linearParse";
import { equalStructural, firstStructuralDifference } from "../equality";
import { canonicalizeXml, parseXml } from "../core/xml";
import type { EquationSpec, MathNode } from "../documentModel";

function toElement(xml: string): Element {
  return parseXml(xml).documentElement;
}

function roundTrip(latex: string, display: "inline" | "block"): EquationSpec {
  const body = parseLinear(latex);
  const eq: EquationSpec = { display, body };
  const xml = serializeMathML(eq);
  const back = parseMathML(toElement(xml));
  const diff = firstStructuralDifference(eq, back);
  expect(diff, diff ? `differs at ${diff}\n${xml}` : "").toBeNull();
  return back;
}

describe("MathML round trip", () => {
  const cases: Array<[string, string, "inline" | "block"]> = [
    ["italic symbol", "x", "inline"],
    ["number", "12", "inline"],
    ["addition", "x+1", "inline"],
    ["fraction", "\\frac{a}{b}", "inline"],
    ["binomial (no bar)", "\\binom{n}{k}", "inline"],
    ["superscript", "x^2", "inline"],
    ["subscript", "x_i", "inline"],
    ["sub and super", "x_i^2", "inline"],
    ["square root", "\\sqrt{x}", "inline"],
    ["nth root", "\\sqrt[3]{x}", "inline"],
    ["n-ary sum", "\\sum_{i=1}^{n} i", "inline"],
    ["n-ary integral", "\\int_a^b", "inline"],
    ["delimiters", "\\left( a \\right)", "inline"],
    ["matrix", "\\matrix{a & b \\\\ c & d}", "inline"],
    ["accent", "\\hat{x}", "inline"],
    ["function", "\\sin", "inline"],
    ["text run", "\\text{if }", "inline"],
  ];

  for (const [label, latex, display] of cases) {
    it(`preserves a ${label}`, () => {
      roundTrip(latex, display);
    });
  }

  it("preserves display mode", () => {
    const back = roundTrip("\\frac{a}{b}", "block");
    expect(back.display).toBe("block");
  });

  it("carries linethickness=0 for a binomial (bar:false)", () => {
    const xml = serializeMathML({ display: "inline", body: parseLinear("\\binom{n}{k}") });
    expect(xml).toContain(`linethickness="0"`);
    const back = parseMathML(toElement(xml));
    expect((back.body[0] as { bar: boolean }).bar).toBe(false);
  });

  it("preserves an unrecognised construct verbatim as opaqueMath", () => {
    const opaque: MathNode = {
      kind: "opaqueMath",
      source: `<merror xmlns="${MATHML_NAMESPACE}"><mtext>x</mtext></merror>`,
      linear: "x",
    };
    const eq: EquationSpec = { display: "inline", body: [opaque] };
    const back = parseMathML(toElement(serializeMathML(eq)));
    expect(back.body[0].kind).toBe("opaqueMath");
    // opaqueMath.source is compared with canonicalizeXml via equalStructural.
    expect(equalStructural(eq.body, back.body)).toBe(true);
    // Sanity: the two sources are XML-equivalent even if not byte-identical.
    const backSource = (back.body[0] as { source: string }).source;
    expect(canonicalizeXml(backSource)).toBe(canonicalizeXml(opaque.source));
  });
});
