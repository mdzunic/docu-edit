// SPEC-116 "Addressing": the A1 helpers every later spreadsheet spec
// addresses cells through — bijective column letters over the whole A..XFD
// grid (round-tripped for every column), A1 references with optional `$`
// markers that return null and never throw outside A1..XFD1048576, ranges and
// sqref lists normalised to top-left:bottom-right, and the sparse-sheet
// lookups: findRow/getCell are binary searches (proven by counting index
// reads on a 100,000-row sheet) and usedRange is the stored cells' bounding
// box.
import { describe, expect, it } from "vitest";
import {
  colToLetters,
  findRow,
  formatA1,
  formatRange,
  getCell,
  lettersToCol,
  parseA1,
  parseRange,
  parseSqref,
  usedRange,
} from "./address";
import { MAX_COLS, MAX_ROWS, type Cell, type Row, type Sheet } from "./model";

/** A grid sheet over the given rows. */
function sheetOf(rows: Row[]): Sheet {
  return { conditionalFormats: [], validations: [], id: "sheet-1", name: "Sheet1", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

/** A row storing a value at each listed column (value = row * 1000 + col, so
 * a lookup that lands on the wrong cell shows). */
function rowOf(index: number, cols: number[]): Row {
  return { index, cells: cols.map((col): Cell => ({ col, value: index * 1000 + col })) };
}

/** `items` behind a Proxy counting reads of numeric indices — how the
 * binary-search tests see how many entries one lookup touched. */
function countingArray<T>(items: T[]): { array: T[]; reads: () => number; reset: () => void } {
  let count = 0;
  const array = new Proxy(items, {
    get(target, property, receiver) {
      if (typeof property === "string" && /^\d+$/.test(property)) {
        count += 1;
      }
      return Reflect.get(target, property, receiver);
    },
  });
  return {
    array,
    reads: () => count,
    reset: () => {
      count = 0;
    },
  };
}

describe("colToLetters", () => {
  it("names columns in bijective base 26, A..XFD", () => {
    expect(colToLetters(0)).toBe("A");
    expect(colToLetters(25)).toBe("Z");
    expect(colToLetters(26)).toBe("AA");
    expect(colToLetters(701)).toBe("ZZ");
    expect(colToLetters(702)).toBe("AAA");
    expect(colToLetters(16_383)).toBe("XFD");
    expect(colToLetters(MAX_COLS - 1)).toBe("XFD");
  });

  it("throws a RangeError for an index that is not a non-negative integer", () => {
    expect(() => colToLetters(-1)).toThrow(RangeError);
    expect(() => colToLetters(1.5)).toThrow(RangeError);
    expect(() => colToLetters(Number.NaN)).toThrow(RangeError);
  });
});

describe("lettersToCol", () => {
  it("reads column letters, case-insensitively", () => {
    expect(lettersToCol("A")).toBe(0);
    expect(lettersToCol("Z")).toBe(25);
    expect(lettersToCol("AA")).toBe(26);
    expect(lettersToCol("ZZ")).toBe(701);
    expect(lettersToCol("AAA")).toBe(702);
    expect(lettersToCol("XFD")).toBe(16_383);
    expect(lettersToCol("xfd")).toBe(16_383);
    expect(lettersToCol("xFd")).toBe(16_383);
  });

  it.each(["", "XFE", "ZZZ", "AAAA", "A1", "$A", " A", "A ", "1", "-", "É"])(
    "returns null for %j — not a column inside the grid",
    (text) => {
      expect(lettersToCol(text)).toBeNull();
    },
  );

  it("inverts colToLetters for every column 0..MAX_COLS-1, upper-case out and either case in", () => {
    const failures: string[] = [];
    for (let col = 0; col < MAX_COLS; col += 1) {
      const letters = colToLetters(col);
      if (!/^[A-Z]{1,3}$/.test(letters) || lettersToCol(letters) !== col || lettersToCol(letters.toLowerCase()) !== col) {
        failures.push(`${col} → ${letters}`);
      }
    }
    expect(failures).toEqual([]);
  });
});

describe("formatA1", () => {
  it("writes a 0-based cell as A1 text", () => {
    expect(formatA1({ row: 2, col: 1 })).toBe("B3");
    expect(formatA1({ row: 0, col: 0 })).toBe("A1");
    expect(formatA1({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })).toBe("XFD1048576");
  });
});

describe("parseA1", () => {
  it("parses a relative reference to its 0-based cell", () => {
    expect(parseA1("B3")).toEqual({ row: 2, col: 1, rowAbs: false, colAbs: false });
    expect(parseA1("A1")).toEqual({ row: 0, col: 0, rowAbs: false, colAbs: false });
  });

  it("accepts optional $ markers and reports which parts were absolute", () => {
    expect(parseA1("$B$3")).toEqual({ row: 2, col: 1, rowAbs: true, colAbs: true });
    expect(parseA1("B$3")).toEqual({ row: 2, col: 1, rowAbs: true, colAbs: false });
    expect(parseA1("$B3")).toEqual({ row: 2, col: 1, rowAbs: false, colAbs: true });
  });

  it("accepts lower-case letters", () => {
    expect(parseA1("b3")).toEqual({ row: 2, col: 1, rowAbs: false, colAbs: false });
    expect(parseA1("$xfd$1048576")).toEqual({ row: 1_048_575, col: 16_383, rowAbs: true, colAbs: true });
  });

  it("reaches the grid's last cell, XFD1048576", () => {
    expect(parseA1("XFD1048576")).toEqual({ row: MAX_ROWS - 1, col: MAX_COLS - 1, rowAbs: false, colAbs: false });
  });

  it("inverts formatA1", () => {
    for (const ref of [
      { row: 0, col: 0 },
      { row: 2, col: 1 },
      { row: 99, col: 26 },
      { row: 1_048_575, col: 16_383 },
    ]) {
      expect(parseA1(formatA1(ref))).toEqual({ ...ref, rowAbs: false, colAbs: false });
    }
  });

  it.each([
    "A0",
    "A1048577",
    "XFD1048577",
    "XFE1",
    "AAAA1",
    "A99999999",
    "",
    " A1",
    "A1 ",
    "A1\n",
    "A 1",
    "1A",
    "A01",
    "A-1",
    "A1.5",
    "A",
    "1",
    "$",
    "$$A1",
    "A$$1",
    "A1$",
    "B3:C4",
    "R1C1",
    "hello",
  ])("returns null, never throwing, for %j", (text) => {
    expect(() => parseA1(text)).not.toThrow();
    expect(parseA1(text)).toBeNull();
  });
});

describe("formatRange", () => {
  it("writes a single cell as one reference", () => {
    expect(formatRange({ top: 2, left: 1, bottom: 2, right: 1 })).toBe("B3");
  });

  it("writes anything larger as top-left:bottom-right", () => {
    expect(formatRange({ top: 0, left: 0, bottom: 2, right: 2 })).toBe("A1:C3");
    expect(formatRange({ top: 0, left: 0, bottom: 0, right: 2 })).toBe("A1:C1");
    expect(formatRange({ top: 1, left: 1, bottom: 4, right: 1 })).toBe("B2:B5");
  });
});

describe("parseRange", () => {
  const A1_C3 = { top: 0, left: 0, bottom: 2, right: 2 };

  it("parses a two-corner range", () => {
    expect(parseRange("A1:C3")).toEqual(A1_C3);
    expect(parseRange("$A$1:$C$3")).toEqual(A1_C3);
    expect(parseRange("a1:c3")).toEqual(A1_C3);
  });

  it("normalises reversed corners to top-left:bottom-right", () => {
    expect(parseRange("C3:A1")).toEqual(A1_C3);
    expect(parseRange("A3:C1")).toEqual(A1_C3);
    expect(parseRange("C1:A3")).toEqual(A1_C3);
  });

  it("parses a single cell as a one-cell range", () => {
    expect(parseRange("B3")).toEqual({ top: 2, left: 1, bottom: 2, right: 1 });
  });

  it.each(["", "A1:", ":A1", "A1:B2:C3", "A1-B2", "A0:B2", "A1:XFE1", "hello", "A1 :B2"])("returns null for %j", (text) => {
    expect(parseRange(text)).toBeNull();
  });

  it("inverts formatRange", () => {
    for (const range of [A1_C3, { top: 2, left: 1, bottom: 2, right: 1 }, { top: 5, left: 3, bottom: 1_048_575, right: 16_383 }]) {
      expect(parseRange(formatRange(range))).toEqual(range);
    }
  });
});

describe("parseSqref", () => {
  it("parses a space-separated list, normalising each range", () => {
    expect(parseSqref("A1:B2 D4  F6:E5")).toEqual([
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 3, left: 3, bottom: 3, right: 3 },
      { top: 4, left: 4, bottom: 5, right: 5 },
    ]);
  });

  it("drops invalid entries and keeps the rest in order", () => {
    expect(parseSqref(" A1  bogus B2:A1 XFE1 A0 C3 ")).toEqual([
      { top: 0, left: 0, bottom: 0, right: 0 },
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 2, left: 2, bottom: 2, right: 2 },
    ]);
  });

  it("is empty for an empty or blank list", () => {
    expect(parseSqref("")).toEqual([]);
    expect(parseSqref("   ")).toEqual([]);
  });
});

describe("findRow and getCell", () => {
  // Stored: A1, C1 · B4 · A5, F5, J5 · D11 — gaps between rows and cells.
  const sparse = sheetOf([rowOf(0, [0, 2]), rowOf(3, [1]), rowOf(4, [0, 5, 9]), rowOf(10, [3])]);

  it("finds a stored row by index, and nothing between rows or past the end", () => {
    expect(findRow(sparse, 0)).toBe(sparse.rows[0]);
    expect(findRow(sparse, 3)).toBe(sparse.rows[1]);
    expect(findRow(sparse, 4)).toBe(sparse.rows[2]);
    expect(findRow(sparse, 10)).toBe(sparse.rows[3]);
    expect(findRow(sparse, 2)).toBeUndefined();
    expect(findRow(sparse, 5)).toBeUndefined();
    expect(findRow(sparse, 11)).toBeUndefined();
    expect(findRow(sparse, MAX_ROWS - 1)).toBeUndefined();
  });

  it("finds a stored cell, and nothing between cells, past a row's last cell, in a missing row or past the end", () => {
    expect(getCell(sparse, 4, 5)).toBe(sparse.rows[2]!.cells[1]);
    expect(getCell(sparse, 0, 0)?.value).toBe(0);
    expect(getCell(sparse, 10, 3)?.value).toBe(10_003);
    expect(getCell(sparse, 4, 4)).toBeUndefined();
    expect(getCell(sparse, 4, 10)).toBeUndefined();
    expect(getCell(sparse, 2, 0)).toBeUndefined();
    expect(getCell(sparse, 11, 0)).toBeUndefined();
  });

  it("answers exactly the stored cells over the whole neighbourhood", () => {
    const stored = new Set(sparse.rows.flatMap((row) => row.cells.map((cell) => `${row.index}:${cell.col}`)));
    for (let row = 0; row <= 12; row += 1) {
      for (let col = 0; col <= 11; col += 1) {
        const cell = getCell(sparse, row, col);
        expect(cell !== undefined, `${row}:${col}`).toBe(stored.has(`${row}:${col}`));
        if (cell) {
          expect(cell.value).toBe(row * 1000 + col);
        }
      }
    }
  });

  it("finds nothing in an empty sheet", () => {
    const empty = sheetOf([]);
    expect(findRow(empty, 0)).toBeUndefined();
    expect(getCell(empty, 0, 0)).toBeUndefined();
  });

  it("is a binary search over the rows: a lookup on a 100,000-row sheet reads a few dozen row entries", () => {
    // Every third row stored, one cell each — a linear scan would read
    // tens of thousands of entries; log2(100,000) is about 17.
    const counted = countingArray(Array.from({ length: 100_000 }, (_, i) => rowOf(i * 3, [i % 7])));
    const sheet = sheetOf(counted.array);
    const probes: { row: number; col: number; hit: boolean }[] = [
      { row: 0, col: 0, hit: true },
      { row: 50_000 * 3, col: 50_000 % 7, hit: true },
      { row: 99_999 * 3, col: 99_999 % 7, hit: true },
      { row: 50_000 * 3 + 1, col: 0, hit: false },
      { row: 50_000 * 3, col: (50_000 % 7) + 1, hit: false },
      { row: 400_000, col: 0, hit: false },
    ];
    for (const { row, col, hit } of probes) {
      counted.reset();
      const cell = getCell(sheet, row, col);
      expect(cell !== undefined, `${row}:${col}`).toBe(hit);
      if (cell) {
        expect(cell.value).toBe(row * 1000 + col);
      }
      expect(counted.reads(), `${row}:${col}`).toBeGreaterThan(0);
      expect(counted.reads(), `${row}:${col}`).toBeLessThanOrEqual(40);
    }
    counted.reset();
    expect(findRow(sheet, 77_777 * 3)?.index).toBe(77_777 * 3);
    expect(counted.reads()).toBeLessThanOrEqual(40);
  });

  it("is a binary search over a row's cells: a lookup in a row filling all 16,384 columns reads a few dozen cells", () => {
    const counted = countingArray(Array.from({ length: MAX_COLS }, (_, col): Cell => ({ col, value: col })));
    const sheet = sheetOf([{ index: 0, cells: counted.array }]);
    for (const col of [0, 12_345, MAX_COLS - 1]) {
      counted.reset();
      expect(getCell(sheet, 0, col)?.value).toBe(col);
      expect(counted.reads()).toBeGreaterThan(0);
      expect(counted.reads()).toBeLessThanOrEqual(40);
    }
  });
});

describe("usedRange", () => {
  it("is null for a sheet that stores nothing", () => {
    expect(usedRange(sheetOf([]))).toBeNull();
  });

  it("is the one cell of a single-cell sheet", () => {
    expect(usedRange(sheetOf([rowOf(4, [2])]))).toEqual({ top: 4, left: 2, bottom: 4, right: 2 });
  });

  it("is the bounding box of ragged rows", () => {
    // Rows start and end at different columns; the widest extent on each
    // side comes from a middle row, not the first or the last.
    const ragged = sheetOf([rowOf(1, [5, 6]), rowOf(3, [4]), rowOf(4, [0, 9, 20]), rowOf(6, [3, 7]), rowOf(9, [8])]);
    expect(usedRange(ragged)).toEqual({ top: 1, left: 0, bottom: 9, right: 20 });
  });

  it("spans the grid's far corner without visiting empty space", () => {
    const corner = sheetOf([rowOf(0, [0]), rowOf(MAX_ROWS - 1, [MAX_COLS - 1])]);
    expect(usedRange(corner)).toEqual({ top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 });
  });
});
