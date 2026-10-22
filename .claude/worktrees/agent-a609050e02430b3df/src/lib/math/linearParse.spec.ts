import { describe, expect, it } from "vitest";
import type { MathNode } from "../documentModel";
import { parseLinear } from "./linearParse";
import { linearForm } from "./linearForm";

describe("parseLinear (SPEC-056 linear input)", () => {
  it("parses \\frac{a}{b} as a fraction with a bar", () => {
    expect(parseLinear("\\frac{a}{b}")).toEqual<MathNode[]>([
      {
        kind: "frac",
        bar: true,
        num: [{ kind: "sym", text: "a", variant: "italic" }],
        den: [{ kind: "sym", text: "b", variant: "italic" }],
      },
    ]);
  });

  it("parses \\binom{n}{k} as a fraction without a bar", () => {
    const [node] = parseLinear("\\binom{n}{k}");
    expect(node).toMatchObject({ kind: "frac", bar: false });
  });

  it("parses x^2, x_i and x_i^2 as scripts", () => {
    expect(parseLinear("x^2")).toEqual<MathNode[]>([
      {
        kind: "script",
        base: [{ kind: "sym", text: "x", variant: "italic" }],
        sup: [{ kind: "sym", text: "2", variant: "upright" }],
      },
    ]);
    const [sub] = parseLinear("x_i");
    expect(sub).toMatchObject({ kind: "script" });
    const [both] = parseLinear("x_i^2");
    expect(both).toMatchObject({ kind: "script", sub: expect.anything(), sup: expect.anything() });
  });

  it("parses \\sqrt{x} and \\sqrt[3]{x}", () => {
    expect(parseLinear("\\sqrt{x}")).toEqual<MathNode[]>([
      { kind: "radical", radicand: [{ kind: "sym", text: "x", variant: "italic" }] },
    ]);
    const [rooted] = parseLinear("\\sqrt[3]{x}");
    expect(rooted).toMatchObject({
      kind: "radical",
      index: [{ kind: "sym", text: "3", variant: "upright" }],
    });
  });

  it("parses \\sum_{i=1}^{n} as an n-ary with limits", () => {
    const [node] = parseLinear("\\sum_{i=1}^{n}");
    expect(node).toMatchObject({ kind: "nary", op: "∑" });
    if (node.kind !== "nary") throw new Error("expected nary");
    expect(linearForm(node.lower!)).toBe("i=1");
    expect(linearForm(node.upper!)).toBe("n");
  });

  it("parses \\left( … \\right) as a delimiter", () => {
    const [node] = parseLinear("\\left( a \\right)");
    expect(node).toMatchObject({ kind: "delim", open: "(", close: ")" });
  });

  it("parses \\matrix{a & b \\\\ c & d} into rows of cells", () => {
    const [node] = parseLinear("\\matrix{a & b \\\\ c & d}");
    expect(node.kind).toBe("matrix");
    if (node.kind !== "matrix") throw new Error("expected matrix");
    expect(node.rows.length).toBe(2);
    expect(node.rows[0].length).toBe(2);
    expect(linearForm(node.rows[1][1])).toBe("d");
  });

  it("parses \\hat{x} as an accent", () => {
    const [node] = parseLinear("\\hat{x}");
    expect(node).toMatchObject({ kind: "accent" });
  });

  it("parses \\sin as an upright function name", () => {
    expect(parseLinear("\\sin")).toEqual<MathNode[]>([{ kind: "func", name: "sin", body: [] }]);
  });

  it("parses \\text{if } x preserving spaces", () => {
    const nodes = parseLinear("\\text{if } x > 0");
    expect(nodes[0]).toEqual({ kind: "text", text: "if " });
  });

  it("resolves \\alpha against the SPEC-026 Greek catalogue (no second table)", () => {
    expect(parseLinear("\\alpha")).toEqual<MathNode[]>([
      { kind: "sym", text: "α", variant: "italic" },
    ]);
    expect(parseLinear("\\Gamma")).toEqual<MathNode[]>([
      { kind: "sym", text: "Γ", variant: "italic" },
    ]);
  });

  it("renders identifiers italic and numbers/operators upright (variant rule)", () => {
    expect(parseLinear("x")).toEqual([{ kind: "sym", text: "x", variant: "italic" }]);
    expect(parseLinear("12")).toEqual([{ kind: "sym", text: "12", variant: "upright" }]);
    expect(parseLinear("+")).toEqual([{ kind: "sym", text: "+", variant: "upright" }]);
  });

  it("inserts input outside the table as literal sym nodes rather than erroring", () => {
    const nodes = parseLinear("\\foo");
    expect(nodes.every((n) => n.kind === "sym")).toBe(true);
    expect(linearForm(nodes)).toBe("\\foo");
  });

  it("never hangs and drops stray closers", () => {
    expect(() => parseLinear("}}}")).not.toThrow();
    expect(() => parseLinear("\\frac{a}")).not.toThrow();
  });
});

describe("linearForm (regeneration for editing and altText)", () => {
  const cases = [
    "\\frac{a}{b}",
    "x^2",
    "\\sqrt[3]{x}",
    "\\hat{x}",
    "\\sin",
    "\\alpha",
  ];
  for (const input of cases) {
    it(`round-trips ${input} through parse → linearForm`, () => {
      expect(linearForm(parseLinear(input))).toBe(input);
    });
  }

  it("regenerates the quadratic formula legibly", () => {
    const src = "x = \\frac{-b}{2a}";
    expect(linearForm(parseLinear(src))).toContain("\\frac{-b}{2a}");
  });
});
