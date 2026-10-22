import { describe, expect, it } from "vitest";
import { parseLinear } from "./linearParse";
import { layoutRow, type RenderBox } from "./render";

function row(box: RenderBox) {
  if (box.type !== "row") throw new Error("expected a row");
  return box.items;
}

describe("layoutRow (SPEC-056 render tree)", () => {
  it("embeds the 4/18 em operator spacing for a + b (AC-10)", () => {
    const items = row(layoutRow(parseLinear("a+b")));
    expect(items.map((i) => i.spaceBefore)).toEqual(["0", "0.2222em", "0.2222em"]);
  });

  it("gives a leading sign no space (AC-11)", () => {
    const items = row(layoutRow(parseLinear("-1")));
    expect(items[0].spaceBefore).toBe("0");
    expect(items[1].spaceBefore).toBe("0");
  });

  it("uses the 5/18 em relation spacing for x = y", () => {
    const items = row(layoutRow(parseLinear("x=y")));
    expect(items[1].spaceBefore).toBe("0.2778em");
  });

  it("carries the italic variant on identifiers and upright on numbers (AC-8/9)", () => {
    const items = row(layoutRow(parseLinear("x2")));
    const x = items[0].box;
    const two = items[1].box;
    expect(x.type === "glyph" && x.variant).toBe("italic");
    expect(two.type === "glyph" && two.variant).toBe("upright");
  });

  it("lays a fraction out as num/den boxes with the bar flag", () => {
    const [box] = row(layoutRow(parseLinear("\\frac{a}{b}"))).map((i) => i.box);
    expect(box).toMatchObject({ type: "frac", bar: true });
  });

  it("marks a 3-row matrix delimiter as needing 3× stretch (AC-44)", () => {
    const [box] = row(
      layoutRow(parseLinear("\\left( \\matrix{a & b \\\\ c & d \\\\ e & f} \\right)")),
    ).map((i) => i.box);
    expect(box).toMatchObject({ type: "delim", stretch: 3 });
  });

  it("renders an unsupported construct as its linear form (AC-24)", () => {
    const box = layoutRow([{ kind: "opaqueMath", source: "<x/>", linear: "??" }]);
    expect(row(box)[0].box).toEqual({ type: "opaque", linear: "??" });
  });

  it("shrinks each script level and floors a script-of-a-script (AC-39/41)", () => {
    const script1 = row(layoutRow(parseLinear("x^{y^{z^{w}}}")))[0].box;
    if (script1.type !== "script" || !script1.sup) throw new Error("expected a script");
    const script2 = row(script1.sup)[0].box;
    if (script2.type !== "script" || !script2.sup) throw new Error("expected a nested script");
    const script3 = row(script2.sup)[0].box;
    if (script3.type !== "script") throw new Error("expected a doubly-nested script");
    // 71% at depth 1, then 0.5/0.71 to reach the 50% floor at depth 2, then no
    // further shrink at depth 3 (floor), each with the 5 pt clamp (AC-41).
    expect(script1.stackScale).toBe("max(5pt, 0.71em)");
    expect(script2.stackScale).toBe("max(5pt, 0.7042em)");
    expect(script3.stackScale).toBeUndefined();
  });

  it("leaves a top-level fraction full size but shrinks a nested one (AC-40)", () => {
    const outer = row(layoutRow(parseLinear("\\frac{\\frac{a}{b}}{c}")))[0].box;
    if (outer.type !== "frac") throw new Error("expected a fraction");
    expect(outer.fontScale).toBeUndefined();
    const inner = row(outer.num)[0].box;
    if (inner.type !== "frac") throw new Error("expected a nested fraction");
    expect(inner.fontScale).toBe("max(5pt, 0.71em)");
  });
});
