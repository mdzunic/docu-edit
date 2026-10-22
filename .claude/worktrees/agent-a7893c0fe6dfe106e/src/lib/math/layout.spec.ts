import { describe, expect, it } from "vitest";
import type { MathNode } from "../documentModel";
import {
  classify,
  clampFontSize,
  eighteenthsToEm,
  fractionScale,
  MIN_FONT_SIZE_PT,
  RULE_CLEARANCE_EM,
  RULE_THICKNESS_EM,
  scriptScale,
  spaceEighteenths,
  spacingForRow,
} from "./layout";

const sym = (text: string, variant: "italic" | "upright" = "upright"): MathNode => ({
  kind: "sym",
  text,
  variant,
});

describe("SPEC-056 spacing table (eighteenths of an em)", () => {
  it("ordinary → ordinary is 0", () => {
    expect(spaceEighteenths(sym("a", "italic"), sym("b", "italic"))).toBe(0);
  });

  it("ordinary ↔ binary operator is 4/18", () => {
    expect(spaceEighteenths(sym("a", "italic"), sym("+"))).toBe(4);
    // binary→ordinary measured through a row so the sign rule doesn't fire
    expect(spacingForRow([sym("a", "italic"), sym("+"), sym("b", "italic")])).toEqual([0, 4, 4]);
  });

  it("ordinary ↔ relation is 5/18", () => {
    expect(spaceEighteenths(sym("a", "italic"), sym("="))).toBe(5);
    expect(spaceEighteenths(sym("="), sym("b", "italic"))).toBe(5);
  });

  it("after a comma or semicolon is 3/18", () => {
    expect(spaceEighteenths(sym(","), sym("b", "italic"))).toBe(3);
  });

  it("a func name to its argument is 3/18", () => {
    expect(spaceEighteenths({ kind: "func", name: "sin", body: [] }, sym("x", "italic"))).toBe(3);
  });

  it("either side of an n-ary body is 3/18", () => {
    const nary: MathNode = { kind: "nary", op: "∑", body: [] };
    expect(spaceEighteenths(sym("a", "italic"), nary)).toBe(3);
    expect(spaceEighteenths(nary, sym("a", "italic"))).toBe(3);
  });
});

describe("SPEC-056 AC-11 sign rule", () => {
  it("a binary operator with nothing to its left takes no space", () => {
    // −1 : the leading minus is a sign
    expect(spacingForRow([sym("−"), sym("1")])).toEqual([0, 0]);
  });

  it("a binary operator after another operator is a sign (e.g. =−1)", () => {
    const row = [sym("a", "italic"), sym("="), sym("−"), sym("1")];
    // a  =(5)  −(sign,0)  1(0)
    expect(spacingForRow(row)).toEqual([0, 5, 0, 0]);
  });

  it("a real binary operator between operands keeps its 4/18 on both sides", () => {
    expect(spacingForRow([sym("1"), sym("+"), sym("1")])).toEqual([0, 4, 4]);
  });
});

describe("SPEC-056 sizes", () => {
  it("script is 71% of base, script-of-script 50% and no smaller (AC-39)", () => {
    expect(scriptScale(0)).toBe(1);
    expect(scriptScale(1)).toBe(0.71);
    expect(scriptScale(2)).toBe(0.5);
    expect(scriptScale(3)).toBe(0.5);
  });

  it("a fraction inside a fraction is 71% of its parent (AC-40)", () => {
    expect(fractionScale(1)).toBe(0.71);
  });

  it("nothing renders below 5 pt (AC-41)", () => {
    expect(MIN_FONT_SIZE_PT).toBe(5);
    expect(clampFontSize(3)).toBe(5);
    expect(clampFontSize(12)).toBe(12);
  });
});

describe("SPEC-056 rule metrics (AC-42)", () => {
  it("bar thickness is 1/18 em and clearance is 1/6 em", () => {
    expect(RULE_THICKNESS_EM).toBeCloseTo(1 / 18);
    expect(RULE_CLEARANCE_EM).toBeCloseTo(1 / 6);
  });

  it("converts eighteenths to em strings", () => {
    expect(eighteenthsToEm(0)).toBe("0");
    expect(eighteenthsToEm(3)).toBe("0.1667em");
    expect(eighteenthsToEm(18)).toBe("1em");
  });
});

describe("classify", () => {
  it("labels ordinary, binary, relation and punctuation", () => {
    expect(classify(sym("x", "italic"))).toBe("ord");
    expect(classify(sym("+"))).toBe("bin");
    expect(classify(sym("="))).toBe("rel");
    expect(classify(sym(","))).toBe("punct");
  });
});
