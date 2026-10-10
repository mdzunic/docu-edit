// SPEC-120 "Recalculation", "Cycles", "Stale cells", "Time slicing",
// "Opening a workbook": changedCellsBetween over copy-on-write edits; an edit
// pass evaluates the changed cells' transitive dependents and the volatile
// cells, in topological order, seeded in sheet order then row-major; cycles
// are 0 after an edit or F9 (their cached values at open; uncomputable when
// the workbook iterates) and reported by their first cell in sheet order;
// stale cells and forceFullCalcOnLoad; copy-on-write writes that leave
// untouched rows (and an unchanged value) alone; array results distributed
// with broadcasting, a changed value dropping what was bound to the old one
// (SPEC-122's drop rules); a pass driven in 0 ms steps equal to the synchronous
// one; cancel-and-inherit; MAX_EVAL_DEPTH re-queueing of a long dynamic
// chain; dynamic dependencies replaced on re-evaluation. Random draws come
// from a seeded mulberry32 and the clock is injected — nothing reads the
// wall clock.
import { describe, expect, it } from "vitest";
import { getCell, parseA1 } from "../address";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import type { Cell, CellValue, Row, SpreadsheetModel } from "../model";
import { cellKey, graphFor } from "./dependencies";
import { MAX_EVAL_DEPTH } from "./evaluate";
import {
  CALCULATING_STATUS_DELAY_MS,
  changedCellsBetween,
  needsFullCalcOnOpen,
  RECALC_SLICE_MS,
  recalculate,
  startRecalc,
  uncomputableReasonOf,
} from "./recalc";

const NOW = (): Date => new Date(2026, 9, 8, 9, 30);

/** A seeded generator (mulberry32), so random draws are reproducible. */
function mulberry32(seed: number): () => number {
  let state = seed;
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let t = Math.imul(state ^ (state >>> 15), 1 | state);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

function keyOf(a1: string, sheet = "Sheet1"): string {
  const at = parseA1(a1)!;
  return cellKey(`sheet-${sheet}`, at.row, at.col);
}

function valueOf(model: SpreadsheetModel, a1: string, sheet = "Sheet1"): CellValue {
  const at = parseA1(a1)!;
  return getCell(model.sheets.find((candidate) => candidate.name === sheet)!, at.row, at.col)?.value ?? null;
}

/** Edits one cell copy-on-write (a new Cell, a new Row, a new rows array)
 * and returns the keys changedCellsBetween finds. */
function edit(model: SpreadsheetModel, a1: string, spec: Partial<Cell>, sheetName = "Sheet1"): string[] {
  const index = model.sheets.findIndex((sheet) => sheet.name === sheetName);
  const sheet = model.sheets[index]!;
  const at = parseA1(a1)!;
  const before = sheet.rows;
  const rows = before.slice();
  const position = rows.findIndex((row) => row.index === at.row);
  const row: Row = position >= 0 ? rows[position]! : { index: at.row, cells: [] };
  const cells = row.cells.filter((cell) => cell.col !== at.col);
  cells.push({ col: at.col, value: null, ...spec });
  const next: Row = { ...row, cells: cells.toSorted((a, b) => a.col - b.col) };
  if (position >= 0) {
    rows[position] = next;
  } else {
    rows.push(next);
    rows.sort((a, b) => a.index - b.index);
  }
  model.sheets = model.sheets.map((candidate, i) => (i === index ? { ...candidate, rows } : candidate));
  return changedCellsBetween(sheet.id, before, rows);
}

/** Three independent RAND() cells on two sheets, Sheet2 first. */
function randomBook(): SpreadsheetModel {
  return calcWorkbook({ Sheet2: { A1: "=RAND()" }, Sheet1: { B2: "=RAND()", A3: "=RAND()" } });
}

/** A computable chain read by an add-in call, and an unrelated unknown
 * function. */
function staleBook(): SpreadsheetModel {
  return calcWorkbook({
    Sheet1: {
      A1: 1,
      B1: { formula: "A1*2", value: 2 },
      C1: { formula: "_xll.CUSTOM(B1)", value: 42 },
      D1: { formula: "ZZZ.UNKNOWN(5)", value: 7 },
    },
  });
}

describe("the constants", () => {
  it("slices of RECALC_SLICE_MS = 12 and the Calculating… delay CALCULATING_STATUS_DELAY_MS = 250", () => {
    expect(RECALC_SLICE_MS).toBe(12);
    expect(CALCULATING_STATUS_DELAY_MS).toBe(250);
  });
});

describe("changedCellsBetween (the identity diff)", () => {
  it("finds exactly the cells a copy-on-write edit touched, row-major — untouched rows and cells by identity", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: 2, A2: 3, A5: 4 } });
    expect(edit(model, "B1", { value: 20 })).toEqual([keyOf("B1")]);
    expect(edit(model, "C3", { value: "new" })).toEqual([keyOf("C3")]);
    const sheet = model.sheets[0]!;
    const without = sheet.rows.filter((row) => row.index !== 4);
    expect(changedCellsBetween(sheet.id, sheet.rows, without)).toEqual([keyOf("A5")]);
    expect(changedCellsBetween(sheet.id, sheet.rows, sheet.rows.slice())).toEqual([]);
  });
});

describe("edit passes", () => {
  it("evaluate the changed cells' transitive dependents — and nothing else — in topological order", () => {
    const model = calcWorkbook({
      Sheet1: { A1: 1, A2: { formula: "A1*2", value: 2 }, A3: { formula: "A2+1", value: 3 }, B1: { formula: "5+5", value: 99 } },
    });
    const result = recalculate(model, edit(model, "A1", { value: 10 }), { now: NOW });
    expect(valueOf(model, "A2")).toBe(20);
    expect(valueOf(model, "A3")).toBe(21);
    expect(valueOf(model, "B1")).toBe(99);
    expect(result.evaluated).toBe(2);
    expect(result.valueChanged).toEqual([keyOf("A2"), keyOf("A3")]);
    expect(result.changedSheetIds).toEqual(["sheet-Sheet1"]);
  });

  it("orders a chain written backwards by its dependencies, not its position", () => {
    const model = calcWorkbook({ Sheet1: { A1: "=A2+1", A2: "=A3+1", A3: "=1" } });
    recalculate(model, [], { full: true, reason: "manual" });
    expect([valueOf(model, "A1"), valueOf(model, "A2"), valueOf(model, "A3")]).toEqual([3, 2, 1]);
  });

  it("draws random numbers in sheet order, then row-major, among independent cells — reproducibly", () => {
    const build = randomBook;
    const first = build();
    recalculate(first, [], { full: true, reason: "manual", random: mulberry32(7) });
    const draws = mulberry32(7);
    const expected = [draws(), draws(), draws()];
    // Sheet2 is first in the workbook, then Sheet1's B2 (row 2) before A3 (row 3).
    expect([valueOf(first, "A1", "Sheet2"), valueOf(first, "B2"), valueOf(first, "A3")]).toEqual(expected);
    const second = build();
    recalculate(second, [], { full: true, reason: "manual", random: mulberry32(7) });
    expect(second.sheets.map((sheet) => sheet.rows)).toEqual(first.sheets.map((sheet) => sheet.rows));
  });

  it("seeds every volatile cell and its dependents, even with nothing changed", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "TODAY()", value: 1 }, A2: { formula: "A1+1", value: 2 }, B1: { formula: "1+1", value: 99 } } });
    const result = recalculate(model, [], { reason: "open", now: NOW });
    expect(valueOf(model, "A1")).toBe(46303);
    expect(valueOf(model, "A2")).toBe(46304);
    expect(valueOf(model, "B1")).toBe(99);
    expect(result.evaluated).toBe(2);
  });
});

describe("cycles", () => {
  it("are 0 after an edit or F9 pass, their dependents compute from that 0, and each is reported by its first cell", () => {
    const model = calcWorkbook({
      Sheet2: { B2: { formula: "B2+1", value: 9 } },
      Sheet1: { A1: { formula: "B1+1", value: 7 }, B1: { formula: "A1+1", value: 8 }, C1: { formula: "A1+5", value: 1 } },
    });
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect([valueOf(model, "A1"), valueOf(model, "B1"), valueOf(model, "C1"), valueOf(model, "B2", "Sheet2")]).toEqual([0, 0, 5, 0]);
    expect(result.circular).toEqual([keyOf("B2", "Sheet2"), keyOf("A1")]);
  });

  it("keep their cached values in an open pass (reported all the same)", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "B1+1", value: 7 }, B1: { formula: "A1+1", value: 8 } } });
    const result = recalculate(model, [], { full: true, reason: "open" });
    expect([valueOf(model, "A1"), valueOf(model, "B1")]).toEqual([7, 8]);
    expect(result.circular).toEqual([keyOf("A1")]);
  });

  it("are uncomputable (\"iterative\") in an iterating workbook: cached values kept, no circular report", () => {
    const model = calcWorkbook(
      { Sheet1: { A1: 1, B1: { formula: "C1+A1", value: 7 }, C1: { formula: "B1/2", value: 3.5 } } },
      { calc: { iterate: true } },
    );
    const result = recalculate(model, edit(model, "A1", { value: 2 }), { now: NOW });
    expect([valueOf(model, "B1"), valueOf(model, "C1")]).toEqual([7, 3.5]);
    expect(result.circular).toEqual([]);
    expect(result.stale).toEqual([keyOf("B1"), keyOf("C1")]);
    expect(uncomputableReasonOf(model, keyOf("B1"))).toEqual({ kind: "iterative" });
  });

  it("closes a dynamic cycle (OFFSET onto itself) found in progress", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "OFFSET(B1,0,-1)+1", value: 5 }, B1: 1 } });
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(valueOf(model, "A1")).toBe(0);
    expect(result.circular).toEqual([keyOf("A1")]);
  });
});

describe("stale cells", () => {
  const book = staleBook;

  it("an uncomputable dependent of a changed value keeps its cached value, goes stale and sets forceFullCalcOnLoad", () => {
    const model = book();
    const result = recalculate(model, edit(model, "A1", { value: 5 }), { now: NOW });
    expect(valueOf(model, "B1")).toBe(10);
    expect(valueOf(model, "C1")).toBe(42);
    expect(result.stale).toEqual([keyOf("C1")]);
    expect(model.forceFullCalcOnLoad).toBe(true);
    expect(uncomputableReasonOf(model, keyOf("C1"))).toEqual({ kind: "function", name: "_XLL.CUSTOM" });
  });

  it("is not stale when nothing it reads changed value — so F9 over unchanged values leaves none", () => {
    const model = book();
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.stale).toEqual([]);
    expect(model.forceFullCalcOnLoad).toBeUndefined();
  });

  it("is every uncomputable cell after an open-time full pass (the file asked for recalculation)", () => {
    const model = book();
    expect(recalculate(model, [], { full: true, reason: "open" }).stale).toEqual([keyOf("C1"), keyOf("D1")]);
  });

  it("covers every cell of a data-table range after an edit or manual pass, never at open", () => {
    const xml = `<f t="dataTable" ref="L3:L4" dt2D="0" dtr="0" r1="F2"/>`;
    const build = () => calcWorkbook({ Sheet1: { F2: 1, L3: { dataTableXml: xml, value: 10 }, L4: 20 } });
    const manual = build();
    expect(recalculate(manual, [], { full: true, reason: "manual" }).stale).toEqual([keyOf("L3"), keyOf("L4")]);
    expect(uncomputableReasonOf(manual, keyOf("L4"))).toEqual({ kind: "dataTable" });
    expect(recalculate(build(), [], { full: true, reason: "open" }).stale).toEqual([]);
  });
});

describe("writes", () => {
  it("are copy-on-write: one new Row per touched row, untouched rows and sheets keep their identity", () => {
    const model = calcWorkbook({
      Sheet1: { A1: 1, A2: { formula: "A1+1", value: 2 }, A3: 5, A4: { formula: "A3+1", value: 6 } },
      Sheet2: { A1: 3 },
    });
    const [first, second, third, fourth] = model.sheets[0]!.rows;
    const other = model.sheets[1];
    recalculate(model, edit(model, "A1", { value: 10 }), { now: NOW });
    const rows = model.sheets[0]!.rows;
    expect(rows[1]).not.toBe(second);
    expect(rows[1]!.cells[0]).toMatchObject({ formula: "A1+1", value: 11 });
    expect(rows[2]).toBe(third);
    expect(rows[3]).toBe(fourth);
    expect(rows[0]).not.toBe(first); // the edit's own row
    expect(model.sheets[1]).toBe(other);
  });

  it("drop what was bound to a changed value (SPEC-122's drop rules) and keep it with an unchanged one", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: { formula: '"site "&A1', value: "site 1" }, C1: { formula: "A1*2", value: 2 }, D1: { formula: "LEN(B1)", value: 6 } } });
    // A string result whose text:p carries a link, value metadata and runs on a number, a kept paragraph.
    const [, b1, c1, d1] = model.sheets[0]!.rows[0]!.cells;
    Object.assign(b1!, { preservedText: `<text:p><text:a xlink:href="https://example.com/">site 1</text:a></text:p>` });
    Object.assign(c1!, { preservedAttrs: ` vm="1" ph="1"`, sstIndex: 3, runs: [{ text: "2" }] });
    Object.assign(d1!, { preservedText: "<text:p>6</text:p>" });
    recalculate(model, edit(model, "A1", { value: 2 }), { now: NOW });
    const row = model.sheets[0]!.rows[0]!;
    expect(row.cells[1]).toEqual({ col: 1, formula: '"site "&A1', value: "site 2" });
    expect(row.cells[2]).toEqual({ col: 2, formula: "A1*2", value: 4, preservedAttrs: ` ph="1"` });
    expect(row.cells[3]).toBe(d1);
  });

  it("write nothing for an unchanged value — a full pass that confirms every cache replaces nothing", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: { formula: "A1+1", value: 2 } } });
    const sheets = model.sheets;
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(model.sheets).toBe(sheets);
    expect(result.valueChanged).toEqual([]);
    expect(result.evaluated).toBe(1);
  });

  it("distribute an array result over its range with broadcasting — every member written, row 3 #N/A — before its readers", () => {
    const model = calcWorkbook({
      Sheet1: { A1: 1, A2: 2, B1: { formula: "A1:A2*{10,100}", array: "B1:C3" }, D1: "=SUM(C1:C2)", E1: "=COUNTA(B3:C3)" },
    });
    recalculate(model, [], { full: true, reason: "manual" });
    const values = ["B1", "C1", "B2", "C2", "B3", "C3"].map((a1) => valueOf(model, a1));
    expect(values).toEqual([10, 100, 20, 200, { error: "#N/A" }, { error: "#N/A" }]);
    expect(valueOf(model, "D1")).toBe(300);
    expect(valueOf(model, "E1")).toBe(2);
  });
});

/** A workbook big enough to need many slices: a running sum, RAND and a
 * cross-sheet total. */
function slicedBook(): SpreadsheetModel {
  const cells: Record<string, CellSpec> = { A1: 1 };
  for (let row = 2; row <= 400; row += 1) {
    cells[`A${row}`] = `=A${row - 1}+1`;
    cells[`B${row}`] = `=SUM($A$1:A${row})`;
    if (row % 50 === 0) {
      cells[`C${row}`] = "=RAND()";
    }
  }
  return calcWorkbook({ Sheet1: cells, Totals: { A1: "=SUM(Sheet1!B:B)" } });
}

describe("time slicing", () => {
  it("a pass driven in 0 ms steps does at least one unit per step and ends equal to the synchronous recalculate", () => {
    const sync = slicedBook();
    const expected = recalculate(sync, [], { full: true, reason: "manual", random: mulberry32(3) });
    const sliced = slicedBook();
    const pass = startRecalc(sliced, [], { full: true, reason: "manual", random: mulberry32(3) });
    let steps = 0;
    while (!pass.step(0)) {
      steps += 1;
    }
    expect(steps).toBeGreaterThan(10);
    expect(pass.result).toEqual(expected);
    expect(sliced.sheets.map((sheet) => sheet.rows)).toEqual(sync.sheets.map((sheet) => sheet.rows));
  });

  it("a cancelled pass hands back the cells it had not evaluated; a pass inheriting them finishes the job", () => {
    const reference = slicedBook();
    recalculate(reference, [], { full: true, reason: "manual", random: () => 0.25 });
    const model = slicedBook();
    const pass = startRecalc(model, [], { full: true, reason: "manual", random: () => 0.25 });
    for (let i = 0; i < 40; i += 1) {
      pass.step(0);
    }
    const left = pass.cancel();
    expect(left.length).toBeGreaterThan(0);
    expect(pass.step(0)).toBe(true);
    recalculate(model, left, { reason: "manual", random: () => 0.25 });
    expect(model.sheets.map((sheet) => sheet.rows)).toEqual(reference.sheets.map((sheet) => sheet.rows));
  });
});

describe("dynamic dependencies", () => {
  it("a long INDIRECT chain re-queues past MAX_EVAL_DEPTH instead of overflowing, and finishes", () => {
    const length = MAX_EVAL_DEPTH * 4;
    const cells: Record<string, CellSpec> = { [`A${length}`]: 1 };
    for (let row = 1; row < length; row += 1) {
      cells[`A${row}`] = `=INDIRECT("A${row + 1}")+1`;
    }
    const model = calcWorkbook({ Sheet1: cells });
    recalculate(model, [], { full: true, reason: "manual" });
    expect(valueOf(model, "A1")).toBe(length);
    expect(graphFor(model).precedentsOf(keyOf("A1"))).toContainEqual({ sheetId: "sheet-Sheet1", top: 1, left: 0, bottom: 1, right: 0 });
  });

  it("are recorded at evaluation and replaced on the next one", () => {
    const model = calcWorkbook({ Sheet1: { A1: '=INDIRECT(B1)', B1: "C1", C1: 1, C2: 2 } });
    recalculate(model, [], { full: true, reason: "manual" });
    expect(valueOf(model, "A1")).toBe(1);
    expect(graphFor(model).dependentsOf("sheet-Sheet1", 0, 2)).toContain(keyOf("A1"));
    recalculate(model, edit(model, "B1", { value: "C2" }), { now: NOW });
    expect(valueOf(model, "A1")).toBe(2);
    expect(graphFor(model).dependentsOf("sheet-Sheet1", 0, 2)).not.toContain(keyOf("A1"));
    expect(graphFor(model).dependentsOf("sheet-Sheet1", 1, 2)).toContain(keyOf("A1"));
    recalculate(model, edit(model, "C2", { value: 20 }), { now: NOW });
    expect(valueOf(model, "A1")).toBe(20);
  });
});

describe("the open-time policy's test", () => {
  it("needs a full pass when the file asks (fullCalcOnLoad) or a formula has no cached value — not otherwise", () => {
    expect(needsFullCalcOnOpen(calcWorkbook({ Sheet1: { A1: { formula: "1+1", value: 2 } } }))).toBe(false);
    expect(needsFullCalcOnOpen(calcWorkbook({ Sheet1: { A1: "=1+1" } }))).toBe(true);
    expect(needsFullCalcOnOpen(calcWorkbook({ Sheet1: { A1: 1 } }, { calc: { fullCalcOnLoad: true } }))).toBe(true);
  });

  it("an open pass never sets a sheet dirty", () => {
    const model = calcWorkbook({ Sheet1: { A1: "=1+1", B1: { formula: "TODAY()", value: 3 } } });
    recalculate(model, [], { full: true, reason: "open", now: NOW });
    expect(valueOf(model, "A1")).toBe(2);
    expect(model.sheets.some((sheet) => "dirty" in sheet)).toBe(false);
  });
});
