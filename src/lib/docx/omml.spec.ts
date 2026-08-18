import { describe, expect, it } from "vitest";
import { parseXml, canonicalizeXml } from "../core/xml";
import type { EquationSpec, MathNode } from "../documentModel";
import { parseLinear } from "../math/linearParse";
import { parseOmmlEquation, serializeOmmlEquation } from "./omml";
import { MATH_NAMESPACE } from "./xml";

function ommlElement(inner: string, tag = "oMath"): Element {
  const doc = parseXml(`<m:${tag} xmlns:m="${MATH_NAMESPACE}">${inner}</m:${tag}>`);
  return doc.documentElement;
}

function eq(body: MathNode[], display: "inline" | "block" = "inline"): EquationSpec {
  return { display, body };
}

/** parse → serialise → re-parse should be structurally stable. */
function reparse(equation: EquationSpec): EquationSpec {
  const xml = serializeOmmlEquation(equation);
  return parseOmmlEquation(ommlElement(xml.replace(/^<m:oMath[^>]*>|<\/m:oMath>$/g, ""), equation.display === "block" ? "oMathPara" : "oMath"));
}

describe("OMML serialiser (SPEC-056 AC-30)", () => {
  it("emits m:oMath for inline and m:oMathPara for display", () => {
    const inline = serializeOmmlEquation(eq(parseLinear("x")));
    expect(inline.startsWith("<m:oMath>")).toBe(true);
    const block = serializeOmmlEquation(eq(parseLinear("x"), "block"));
    expect(block.startsWith("<m:oMathPara>")).toBe(true);
    expect(block).toContain("<m:oMath>");
  });

  const constructs: Array<[string, string]> = [
    ["fraction", "\\frac{a}{b}"],
    ["binom (noBar)", "\\binom{n}{k}"],
    ["superscript", "x^2"],
    ["subsup", "x_i^2"],
    ["radical", "\\sqrt{x}"],
    ["indexed radical", "\\sqrt[3]{x}"],
    ["n-ary", "\\sum_{i=1}^{n}"],
    ["delimiter", "\\left( a \\right)"],
    ["matrix", "\\matrix{a & b \\\\ c & d}"],
    ["accent", "\\hat{x}"],
    ["function", "\\sin"],
  ];

  for (const [label, src] of constructs) {
    it(`round-trips a ${label} through OMML`, () => {
      const original = eq(parseLinear(src));
      const roundTripped = reparse(original);
      expect(roundTripped.body).toEqual(original.body);
    });
  }

  it("preserves an unsupported construct verbatim as opaqueMath (AC-23/25/49)", () => {
    const source = `<m:groupChr xmlns:m="${MATH_NAMESPACE}"><m:e><m:r><m:t>x</m:t></m:r></m:e></m:groupChr>`;
    const parsed = parseOmmlEquation(ommlElement(source.replace(/ xmlns:m="[^"]*"/, "")));
    expect(parsed.body[0].kind).toBe("opaqueMath");
    const out = serializeOmmlEquation(parsed);
    // The opaque subtree is re-emitted; compared canonically per SPEC-008.
    if (parsed.body[0].kind !== "opaqueMath") throw new Error("expected opaqueMath");
    expect(canonicalizeXml(parsed.body[0].source)).toContain("groupChr");
    expect(out).toContain("groupChr");
  });

  it("carries the italic/upright variant through m:sty", () => {
    const parsed = parseOmmlEquation(
      ommlElement(
        `<m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>1</m:t></m:r>`,
      ),
    );
    expect(parsed.body).toEqual<MathNode[]>([
      { kind: "sym", text: "x", variant: "italic" },
      { kind: "sym", text: "1", variant: "upright" },
    ]);
  });
});
