// SPEC-122 "Point mode": canPointAt after "=", "(", ",", each binary
// operator and a unary sign (trailing spaces ignored) and nowhere else;
// referenceSpans' offsets in the whole editor text and colours by first
// appearance of each distinct reference text, modulo 8, for cells, ranges,
// whole columns and rows and sheet-qualified references; placePointReference
// inserting at the caret then replacing its own span; cycleAbsolute's
// A1 → $A$1 → A$1 → $A1 → A1 cycle on cells, ranges (both ends together),
// whole columns and rows, keeping the sheet prefix as written.
import { describe, expect, it } from "vitest";
import { canPointAt, cycleAbsolute, placePointReference, referenceSpans } from "./pointMode";

/** canPointAt with the caret at the end of the text. */
const atEnd = (text: string): boolean => canPointAt(text, text.length);

describe("canPointAt", () => {
  it.each(["=", "=SUM(", "=SUM(A1,", "=A1+", "=A1-", "=A1*", "=A1/", "=A1^", '="a"&', "=A1=", "=A1<", "=A1>", "=A1<=", "=A1>=", "=A1<>", "=-", "=+", "=(", "=(A1,", "=SUM( ", "=A1 + ", "=IF(A1>0,\n"])(
    "allows a reference after %j",
    (text) => {
      expect(atEnd(text)).toBe(true);
    },
  );

  it.each(["A1+", "1+", "=A1", "=SUM", "=SUM(A1", "=A1%", "=A1:", "=1", '="a', '="a"', "=TRUE", "=SUM(A1)", "={1,", "=#N/A", "=Rate", "=A1 "])(
    "refuses after %j",
    (text) => {
      expect(atEnd(text)).toBe(false);
    },
  );

  it("looks only at the text before the caret", () => {
    expect(canPointAt("=A1+B2", 4)).toBe(true);
    expect(canPointAt("=A1+B2", 3)).toBe(false);
    expect(canPointAt("=SUM(B1)", 5)).toBe(true);
    expect(canPointAt("=A1", 1)).toBe(true);
    expect(canPointAt("=A1", 0)).toBe(false);
    expect(canPointAt("=A1", 9)).toBe(false);
  });
});

describe("referenceSpans", () => {
  it("returns nothing for text that is not a formula", () => {
    expect(referenceSpans("A1+B1")).toEqual([]);
    expect(referenceSpans("")).toEqual([]);
  });

  it("finds every kind of reference with its offsets in the editor text", () => {
    const text = "=SUM(A1:B2,'My Sheet'!C3,Sheet2!A1:B3,D:D,4:5,$E$6)+x";
    const spans = referenceSpans(text);
    expect(spans.map((span) => text.slice(span.start, span.end))).toEqual(["A1:B2", "'My Sheet'!C3", "Sheet2!A1:B3", "D:D", "4:5", "$E$6"]);
    expect(spans[0]!.ref).toEqual({
      start: { row: 0, col: 0, rowAbs: false, colAbs: false },
      end: { row: 1, col: 1, rowAbs: false, colAbs: false },
    });
    expect(spans[1]!.ref.sheet).toBe("My Sheet");
    expect(spans[2]!.ref.sheet).toBe("Sheet2");
    expect(spans[3]!.ref.start).toEqual({ col: 3, rowAbs: false, colAbs: false });
    expect(spans[4]!.ref.start).toEqual({ row: 3, rowAbs: false, colAbs: false });
    expect(spans.map((span) => span.colorIndex)).toEqual([0, 1, 2, 3, 4, 5]);
  });

  it("colours by first appearance of each distinct reference text, case-insensitively", () => {
    const spans = referenceSpans("=A1+B2+a1+$A$1+B2");
    expect(spans.map((span) => span.colorIndex)).toEqual([0, 1, 0, 2, 1]);
  });

  it("cycles the eight colours", () => {
    const refs = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A1"];
    const spans = referenceSpans(`=${refs.join("+")}`);
    expect(spans.map((span) => span.colorIndex)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 0]);
  });

  it("does not count names, functions or strings", () => {
    const spans = referenceSpans('=SUM(Rate,"A1",LOG10(B1))');
    expect(spans.map((span) => [span.start, span.end])).toEqual([[21, 23]]);
  });
});

describe("placePointReference", () => {
  it("inserts at the caret, then replaces its own span", () => {
    const first = placePointReference("=SUM(", { start: 5, end: 5 }, "B1");
    expect(first).toEqual({ text: "=SUM(B1", span: { start: 5, end: 7 } });
    const second = placePointReference(first.text, first.span, "B1:B2");
    expect(second).toEqual({ text: "=SUM(B1:B2", span: { start: 5, end: 10 } });
    const third = placePointReference(`${second.text})`, second.span, "'My Sheet'!B2");
    expect(third).toEqual({ text: "=SUM('My Sheet'!B2)", span: { start: 5, end: 18 } });
  });

  it("inserts in the middle of the text without touching the rest", () => {
    expect(placePointReference("=A1+*2", { start: 4, end: 4 }, "C3")).toEqual({ text: "=A1+C3*2", span: { start: 4, end: 6 } });
  });

  it("inserts at the caret given, or at the end, when there is no span yet", () => {
    expect(placePointReference("=1+", null, "A1")).toEqual({ text: "=1+A1", span: { start: 3, end: 5 } });
    expect(placePointReference("=+2", null, "A1", 1)).toEqual({ text: "=A1+2", span: { start: 1, end: 3 } });
  });
});

describe("cycleAbsolute", () => {
  it("cycles a cell A1 → $A$1 → A$1 → $A1 → A1", () => {
    const seen: string[] = [];
    let state = { text: "=A1", caret: 3 };
    for (let i = 0; i < 4; i += 1) {
      state = cycleAbsolute(state.text, state.caret);
      seen.push(state.text);
    }
    expect(seen).toEqual(["=$A$1", "=A$1", "=$A1", "=A1"]);
    expect(state.caret).toBe(3);
  });

  it("finds the reference the caret is in, after or before, and puts the caret after it", () => {
    expect(cycleAbsolute("=A1+B22*2", 5)).toEqual({ text: "=A1+$B$22*2", caret: 9 });
    expect(cycleAbsolute("=A1+B22*2", 7)).toEqual({ text: "=A1+$B$22*2", caret: 9 });
    expect(cycleAbsolute("=A1+B22*2", 3)).toEqual({ text: "=$A$1+B22*2", caret: 5 });
    expect(cycleAbsolute("=A1+B22*2", 4)).toEqual({ text: "=A1+$B$22*2", caret: 9 });
    expect(cycleAbsolute("=SUM(A1)", 5)).toEqual({ text: "=SUM($A$1)", caret: 9 });
  });

  it("cycles a range's ends together, from its first end's state", () => {
    const steps = ["=A1:B2", "=$A$1:$B$2", "=A$1:B$2", "=$A1:$B2", "=A1:B2"];
    for (let i = 0; i < steps.length - 1; i += 1) {
      expect(cycleAbsolute(steps[i]!, steps[i]!.length).text).toBe(steps[i + 1]);
    }
    expect(cycleAbsolute("=$A1:B$2", 8).text).toBe("=A1:B2");
  });

  it("toggles the one marker of a whole column or row", () => {
    expect(cycleAbsolute("=SUM(A:A)", 8).text).toBe("=SUM($A:$A)");
    expect(cycleAbsolute("=SUM($A:$A)", 10).text).toBe("=SUM(A:A)");
    expect(cycleAbsolute("=SUM(3:4)", 8).text).toBe("=SUM($3:$4)");
    expect(cycleAbsolute("=SUM($3:$4)", 10).text).toBe("=SUM(3:4)");
  });

  it("keeps the sheet prefix as written", () => {
    expect(cycleAbsolute("='My Sheet'!B2", 14)).toEqual({ text: "='My Sheet'!$B$2", caret: 16 });
    expect(cycleAbsolute("=sheet2!a1:b3", 13)).toEqual({ text: "=sheet2!$A$1:$B$3", caret: 17 });
  });

  it("leaves text without a reference at the caret unchanged", () => {
    expect(cycleAbsolute("=A1+1", 5)).toEqual({ text: "=A1+1", caret: 5 });
    expect(cycleAbsolute("=SUM(Rate)", 9)).toEqual({ text: "=SUM(Rate)", caret: 9 });
    expect(cycleAbsolute("A1", 2)).toEqual({ text: "A1", caret: 2 });
  });
});
