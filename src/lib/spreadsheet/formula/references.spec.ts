// SPEC-119 "References": rewriteReferences is token-level — only the ref
// tokens it rewrites change, every other token and all trivia survive
// byte-for-byte, an untouched formula comes back as the identical string —
// and on it shiftFormula (relative parts move, absolute stay, whole columns
// ignore dRow and whole rows dCol, any part leaving the grid at any of its
// four edges makes the reference #REF! keeping its sheet prefix),
// renameSheetInFormula (case-insensitive, 3-D endpoints, sheet-scoped names
// and Sheet!#REF!, re-quoted per the rule) and relativeFormulaKey (equal for
// shifted copies, different otherwise); findArrayAnchor is memoised per
// `sheet.rows` array.
import { describe, expect, it } from "vitest";
import type { Cell, Row, Sheet } from "../model";
import { findArrayAnchor, relativeFormulaKey, renameSheetInFormula, rewriteReferences, shiftFormula, type ReferenceNode } from "./references";

describe("rewriteReferences", () => {
  it("replaces only the ref tokens the rewrite returns, verbatim elsewhere — whitespace, case and spelling kept", () => {
    const text = "  sum( a1 ,\r\n\t$B$2 ,'My Sheet'!c3 )  &\"A1\"";
    const out = rewriteReferences(text, (ref) => (ref.start.col === 0 ? { ...ref, start: { ...ref.start, row: 9 } } : undefined));
    expect(out).toBe("  sum( A10 ,\r\n\t$B$2 ,'My Sheet'!c3 )  &\"A1\"");
  });

  it("turns null into the token's sheet prefix (as written) + #REF!", () => {
    expect(rewriteReferences("A1+Sheet2!B2+'My Sheet'!C3:D4+Jan:Mar!E5", () => null)).toBe("#REF!+Sheet2!#REF!+'My Sheet'!#REF!+Jan:Mar!#REF!");
  });

  it("returns the identical string when nothing is rewritten, and never calls back for non-ref tokens", () => {
    const text = "SUM(A1, Rate, Table1[Col], [1]Sheet1!A1, #REF!A1)";
    const seen: ReferenceNode[] = [];
    expect(rewriteReferences(text, (ref) => {
      seen.push(ref);
      return undefined;
    })).toBe(text);
    expect(seen).toHaveLength(1);
  });
});

describe("shiftFormula", () => {
  it.each([
    ["A1", 1, 1, "B2"],
    ["$A1", 1, 1, "$A2"],
    ["A$1", 1, 1, "B$1"],
    ["$A$1", 5, 5, "$A$1"],
    ["A1:B2", 2, 0, "A3:B4"],
    ["$A$1:B2", 2, 1, "$A$1:C4"],
    ["SUM(A:A)", 7, 1, "SUM(B:B)"],
    ["SUM($A:$A)", 7, 1, "SUM($A:$A)"],
    ["SUM(1:1)", 1, 7, "SUM(2:2)"],
    ["SUM($1:$1)", 1, 7, "SUM($1:$1)"],
    ["SUM(Jan:Mar!B2)", 1, 0, "SUM(Jan:Mar!B3)"],
    ["'My Sheet'!A1*B1", 0, 2, "'My Sheet'!C1*D1"],
    ["B2*$F$1", 1, 0, "B3*$F$1"],
    ["A1", -1, 0, "#REF!"],
    ["A1", 0, -1, "#REF!"],
    ["A1048576", 1, 0, "#REF!"],
    ["XFD1", 0, 1, "#REF!"],
    ["SUM(A1:A1048576)", 1, 0, "SUM(#REF!)"],
    ["Sheet2!A1+1", -1, 0, "Sheet2!#REF!+1"],
    ["'My Sheet'!A1:B2", 0, -1, "'My Sheet'!#REF!"],
    ["A1048575*2", 2, 0, "#REF!*2"],
    ["SUM(A:A)", -1, 0, "SUM(A:A)"],
    ["SUM(1:1)", 0, -1, "SUM(1:1)"],
  ])("%s shifted by (%i, %i) → %s", (text, dRow, dCol, expected) => {
    expect(shiftFormula(text, dRow, dCol)).toBe(expected);
  });

  it("leaves every non-reference token and all trivia untouched", () => {
    expect(shiftFormula("  SUM( A1 ,\n B1 ) & \"C1\" & Rate & Table1[A1] & [1]Sheet1!A1 & #REF!A1 ", 1, 0)).toBe(
      "  SUM( A2 ,\n B2 ) & \"C1\" & Rate & Table1[A1] & [1]Sheet1!A1 & #REF!A1 ",
    );
  });

  it("keeps an unmoved reference's spelling, and returns the same string for a zero shift", () => {
    const text = "$a$1+b2";
    expect(shiftFormula(text, 0, 0)).toBe(text);
    expect(shiftFormula(text, 1, 0)).toBe("$a$1+B3");
  });

  it("normalises a range whose corners cross, each axis keeping its own $", () => {
    expect(shiftFormula("SUM(A1:$A$5)", 10, 0)).toBe("SUM(A$5:$A11)");
    expect(shiftFormula("SUM(A1:$C$1)", 0, 5)).toBe("SUM($C1:F$1)");
  });
});

describe("renameSheetInFormula", () => {
  it("renames case-insensitively, adding and removing quotes per the rule", () => {
    expect(renameSheetInFormula("SUM(sheet1!A1:B2)+Sheet1!C3", "SHEET1", "My Data")).toBe("SUM('My Data'!A1:B2)+'My Data'!C3");
    expect(renameSheetInFormula("'My Data'!A1*2", "my data", "Data")).toBe("Data!A1*2");
    expect(renameSheetInFormula("Data!A1", "Data", "O'Brien")).toBe("'O''Brien'!A1");
  });

  it("renames either 3-D endpoint, re-quoting the prefix whole", () => {
    expect(renameSheetInFormula("SUM(Jan:Mar!B2)", "Mar", "March 2026")).toBe("SUM('Jan:March 2026'!B2)");
    expect(renameSheetInFormula("SUM('Q1 2026:Q4 2026'!C3)", "q1 2026", "First")).toBe("SUM('First:Q4 2026'!C3)");
    expect(renameSheetInFormula("SUM('First:Q4 2026'!C3)", "Q4 2026", "Last")).toBe("SUM(First:Last!C3)");
    // A cell-shaped name is always quoted.
    expect(renameSheetInFormula("SUM(First:Last!C3)", "Last", "Q4")).toBe("SUM('First:Q4'!C3)");
  });

  it("renames sheet-scoped names and Sheet!#REF! too; external references and other sheets are left alone", () => {
    expect(renameSheetInFormula("Jan!Rate+Jan!#REF!+[1]Jan!A1+Feb!A1", "Jan", "January")).toBe("January!Rate+January!#REF!+[1]Jan!A1+Feb!A1");
  });

  it("returns the identical string when the formula names no such sheet, whitespace intact", () => {
    const text = " Feb!A1 +  'Other Sheet'!B2 ";
    expect(renameSheetInFormula(text, "Jan", "January")).toBe(text);
    expect(renameSheetInFormula(" Jan!A1 +  B2 ", "Jan", "January")).toBe(" January!A1 +  B2 ");
  });
});

describe("relativeFormulaKey", () => {
  it("writes each reference in R1C1 form relative to the host, trivia dropped, sheet prefixes as written", () => {
    expect(relativeFormulaKey("B3 + $C$5", { row: 0, col: 0 })).toBe("R[2]C[1]+R5C3");
    expect(relativeFormulaKey("SUM(A:A, $2:3)", { row: 0, col: 0 })).toBe("SUM(C[0]:C[0],R2:R[2])");
    expect(relativeFormulaKey("'My Sheet'!A$1:$B2", { row: 4, col: 4 })).toBe("'My Sheet'!R1C[-4]:R[-3]C2");
  });

  it("is equal for shifted copies and different otherwise", () => {
    const master = relativeFormulaKey("B2*$F$1", { row: 1, col: 2 });
    expect(relativeFormulaKey(shiftFormula("B2*$F$1", 5, 3), { row: 6, col: 5 })).toBe(master);
    expect(relativeFormulaKey("B9*$F$1", { row: 8, col: 2 })).toBe(master);
    expect(relativeFormulaKey("B9*$F$1", { row: 7, col: 2 })).not.toBe(master);
    expect(relativeFormulaKey("B2*$F$2", { row: 1, col: 2 })).not.toBe(master);
    expect(relativeFormulaKey("B2+$F$1", { row: 1, col: 2 })).not.toBe(master);
    expect(relativeFormulaKey("Sheet2!B2*$F$1", { row: 1, col: 2 })).not.toBe(master);
  });
});

describe("findArrayAnchor", () => {
  const anchor: Cell = { col: 1, value: 6, formula: "A2:A4*2", arrayRange: { top: 1, left: 1, bottom: 3, right: 2 } };
  const rows: Row[] = [
    { index: 1, cells: [anchor] },
    { index: 2, cells: [{ col: 2, value: 4 }] },
  ];
  const sheet: Sheet = { conditionalFormats: [], id: "s", name: "S", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };

  it("finds the anchor whose arrayRange holds a position — its own and every member's, stored or not", () => {
    expect(findArrayAnchor(sheet, 1, 1)).toBe(anchor);
    expect(findArrayAnchor(sheet, 2, 2)).toBe(anchor);
    expect(findArrayAnchor(sheet, 3, 1)).toBe(anchor);
    expect(findArrayAnchor(sheet, 4, 1)).toBeUndefined();
    expect(findArrayAnchor(sheet, 1, 0)).toBeUndefined();
  });

  it("is memoised per rows array: a new rows array (an edit, copy-on-write) is looked at afresh", () => {
    expect(findArrayAnchor(sheet, 2, 2)).toBe(anchor);
    const edited: Sheet = { ...sheet, rows: [{ index: 1, cells: [{ col: 1, value: 6 }] }, rows[1]!] };
    expect(findArrayAnchor(edited, 2, 2)).toBeUndefined();
    expect(findArrayAnchor(sheet, 2, 2)).toBe(anchor);
  });
});
