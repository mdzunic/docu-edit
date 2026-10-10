// SPEC-124 "Formula adjustment": every example of the table, plus a
// randomised check that an insert followed by the delete of the same band
// restores every formula the insert did not turn into #REF!.
import { describe, expect, it } from "vitest";
import { colToLetters } from "../address";
import { MAX_COLS, MAX_ROWS } from "../model";
import type { StructuralEdit } from "../structure";
import { adjustFormula, adjustFormulaReferences } from "./adjust";

const SHEETS: Record<string, string> = { s1: "Sheet1", s2: "Sheet2", s3: "Sheet3", my: "My Sheet" };
const nameOf = (id: string): string => SHEETS[id] ?? "";

function edit(kind: StructuralEdit["kind"], at: number, count: number, sheetId = "s1"): StructuralEdit {
  return { kind, sheetId, at, count };
}

describe("adjustFormula — the table (host Sheet1, edit on Sheet1 unless noted)", () => {
  const rows: { formula: string; edit: StructuralEdit; host?: string; result: string }[] = [
    { formula: "A5+1", edit: edit("insertRows", 2, 2), result: "A7+1" },
    { formula: "$A$5*2", edit: edit("insertRows", 2, 2), result: "$A$7*2" },
    { formula: "SUM(A1:A10)", edit: edit("insertRows", 4, 1), result: "SUM(A1:A11)" },
    { formula: "SUM(A5:A10)", edit: edit("insertRows", 4, 1), result: "SUM(A6:A11)" },
    { formula: "SUM(A1:A10)", edit: edit("insertRows", 10, 1), result: "SUM(A1:A10)" },
    { formula: "A5 +  B5", edit: edit("insertRows", 0, 1), result: "A6 +  B6" },
    { formula: "A5", edit: edit("deleteRows", 4, 1), result: "#REF!" },
    { formula: "SUM(A1:A10)", edit: edit("deleteRows", 2, 3), result: "SUM(A1:A7)" },
    { formula: "SUM(A4:A10)", edit: edit("deleteRows", 2, 3), result: "SUM(A3:A7)" },
    { formula: "SUM(A1:A4)", edit: edit("deleteRows", 2, 3), result: "SUM(A1:A2)" },
    { formula: "SUM(A3:A5)", edit: edit("deleteRows", 2, 3), result: "SUM(#REF!)" },
    { formula: "5:10", edit: edit("insertRows", 0, 1), result: "6:11" },
    { formula: "A:A", edit: edit("insertRows", 0, 1), result: "A:A" },
    { formula: "SUM(B:D)", edit: edit("deleteCols", 2, 1), result: "SUM(B:C)" },
    { formula: "A1048576", edit: edit("insertRows", 0, 1), result: "#REF!" },
    { formula: "A1:A1048576", edit: edit("insertRows", 0, 1), result: "A2:A1048576" },
    { formula: "Sheet1!C1", host: "s2", edit: edit("insertCols", 1, 1), result: "Sheet1!D1" },
    { formula: "'my sheet'!A5", host: "s2", edit: edit("insertRows", 0, 1, "my"), result: "'my sheet'!A6" },
    { formula: "Sheet2!A5", host: "s1", edit: edit("deleteRows", 4, 1, "s2"), result: "Sheet2!#REF!" },
    { formula: "INDIRECT(\"A5\")", edit: edit("insertRows", 0, 1), result: "INDIRECT(\"A5\")" },
  ];
  for (const row of rows) {
    it(`${row.formula} — ${row.edit.kind} a=${row.edit.at} n=${row.edit.count} → ${row.result}`, () => {
      expect(adjustFormula(row.formula, row.host ?? "s1", row.edit, nameOf)).toBe(row.result);
    });
  }

  it("leaves 3-D and external references alone, under any edit", () => {
    for (const kind of ["insertRows", "deleteRows", "insertCols", "deleteCols"] as const) {
      expect(adjustFormula("Sheet1:Sheet3!A5", "s1", edit(kind, 0, 3), nameOf)).toBe("Sheet1:Sheet3!A5");
      expect(adjustFormula("[1]Sheet1!A5", "s1", edit(kind, 0, 3), nameOf)).toBe("[1]Sheet1!A5");
    }
  });

  it("returns the identical string for an unaffected formula", () => {
    const text = "SUM( A1:A3 ) + Sheet2!B9";
    expect(adjustFormula(text, "s1", edit("insertRows", 10, 2), nameOf)).toBe(text);
    expect(adjustFormula(text, "s3", edit("insertRows", 0, 2), nameOf)).toBe(text);
  });
});

describe("adjustFormula — what survives byte-for-byte", () => {
  it("keeps trivia, $ markers, the written sheet prefix and every other token", () => {
    expect(adjustFormula("  SUM( $A$5 ,\n'Sheet1'!B$2:$B9 ) & \"A5\" & Rate & Table1[A5]", "s2", edit("insertRows", 0, 1), nameOf)).toBe(
      "  SUM( $A$5 ,\n'Sheet1'!B$3:$B10 ) & \"A5\" & Rate & Table1[A5]",
    );
    expect(adjustFormula("sheet1!a5", "s2", edit("insertRows", 0, 1), nameOf)).toBe("sheet1!A6");
  });

  it("moves absolute and relative parts alike and follows each axis on its own", () => {
    expect(adjustFormula("$B$2+B2+$B2+B$2", "s1", edit("insertCols", 0, 2), nameOf)).toBe("$D$2+D2+$D2+D$2");
    expect(adjustFormula("$B$2+B2", "s1", edit("insertRows", 5, 2), nameOf)).toBe("$B$2+B2");
  });

  it("names the edited sheet case-insensitively, and never another sheet's references", () => {
    expect(adjustFormula("SHEET1!A5+A5", "s2", edit("insertRows", 0, 1), nameOf)).toBe("SHEET1!A6+A5");
    expect(adjustFormula("Sheet2!A5+A5", "s1", edit("insertRows", 0, 1), nameOf)).toBe("Sheet2!A5+A6");
  });

  it("shrinks a straddled range and keeps a reversed range's ends", () => {
    expect(adjustFormula("SUM(A10:A1)", "s1", edit("deleteRows", 2, 3), nameOf)).toBe("SUM(A7:A1)");
    expect(adjustFormula("SUM(B1:F1)", "s1", edit("deleteCols", 0, 3), nameOf)).toBe("SUM(A1:C1)");
  });

  it("lets whole columns ignore row edits and whole rows column edits", () => {
    expect(adjustFormula("SUM(C:C)+SUM(3:4)", "s1", edit("insertRows", 0, 5), nameOf)).toBe("SUM(C:C)+SUM(8:9)");
    expect(adjustFormula("SUM(C:C)+SUM(3:4)", "s1", edit("insertCols", 0, 1), nameOf)).toBe("SUM(D:D)+SUM(3:4)");
    expect(adjustFormula("SUM(C:C)", "s1", edit("deleteCols", 2, 1), nameOf)).toBe("SUM(#REF!)");
  });

  it("takes the sheet test from the caller for retained text", () => {
    expect(adjustFormulaReferences("Old!A5+A5", edit("insertRows", 0, 1), (sheet) => sheet === "Old")).toBe("Old!A6+A5");
  });
});

/** A small deterministic generator (mulberry32). */
function random(seed: number): () => number {
  let state = seed;
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let t = Math.imul(state ^ (state >>> 15), 1 | state);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe("adjustFormula — insert then delete of the same band", () => {
  it("restores every formula the insert did not turn into #REF! (randomised)", () => {
    const next = random(124);
    const int = (n: number): number => Math.floor(next() * n);
    const dollar = (): string => (next() < 0.3 ? "$" : "");
    const cellText = (maxRow: number, maxCol: number): string => `${dollar()}${colToLetters(int(maxCol))}${dollar()}${int(maxRow) + 1}`;
    const reference = (): string => {
      const roll = next();
      const prefix = next() < 0.3 ? (next() < 0.5 ? "Sheet1!" : "'Sheet1'!") : "";
      if (roll < 0.35) {
        // Anywhere on the grid: an insert may push it off the end.
        return prefix + cellText(MAX_ROWS, MAX_COLS);
      }
      if (roll < 0.7) {
        // Ranges well inside the grid: inserts never clamp them.
        const a = cellText(1000, 200);
        const b = cellText(1000, 200);
        return `${prefix}${a}:${b}`;
      }
      if (roll < 0.85) {
        const a = colToLetters(int(200));
        const b = colToLetters(int(200));
        return `${prefix}${a}:${b}`;
      }
      return `${prefix}${int(1000) + 1}:${int(1000) + 1}`;
    };
    for (let k = 0; k < 400; k += 1) {
      const parts = Array.from({ length: 1 + int(4) }, reference);
      const formula = `SUM(${parts.join(",")})+Sheet2!A${int(50) + 1}`;
      const rows = next() < 0.5;
      const at = int(rows ? 1200 : 220);
      const count = 1 + int(50);
      const inserted = adjustFormula(formula, "s1", edit(rows ? "insertRows" : "insertCols", at, count), nameOf);
      const restored = adjustFormula(inserted, "s1", edit(rows ? "deleteRows" : "deleteCols", at, count), nameOf);
      if (inserted.includes("#REF!")) {
        continue;
      }
      expect(restored, `${formula} through ${rows ? "rows" : "cols"} a=${at} n=${count}: ${inserted}`).toBe(formula);
    }
  });
});
