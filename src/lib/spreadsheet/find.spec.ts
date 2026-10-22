// SPEC-134's find.ts over models built here (no file I/O): a cell's search
// text per kind in both Look-in modes, every match option, scope and order
// (multi-range selections and wrap included), the chunked walk under an
// injected clock, the replace plan with every skip reason and refusal,
// literal replacement, the Go To grammar's acceptances and rejections, Go To
// Special per option and type filter, rectangle coalescing and the range
// limit.
import { describe, expect, it } from "vitest";
import { isQueryError } from "../core/findQuery";
import type { CellRef, GridRange } from "./address";
import {
  breaksFormulaMessage,
  cellMatches,
  coalesceCells,
  compileCellQuery,
  createSheetSearch,
  FIND_ALL_LIMIT,
  FIND_CHECK_INTERVAL,
  FIND_SLICE_BUDGET_MS,
  findAllCapMessage,
  findAreas,
  GO_TO_HIDDEN_SHEET_MESSAGE,
  GO_TO_INVALID_MESSAGE,
  GO_TO_NOT_A_RANGE_MESSAGE,
  GO_TO_SPECIAL_MAX_RANGES,
  GO_TO_SPECIAL_NONE_MESSAGE,
  GO_TO_SPECIAL_TOO_MANY_MESSAGE,
  hitIsHidden,
  planReplace,
  replaceReport,
  resolveGoToTarget,
  searchText,
  selectSpecial,
  type CompiledCellQuery,
  type FindHit,
  type SheetFindOptions,
  type SpecialTypes,
} from "./find";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type Row, type Sheet, type SheetTable, type SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import type { FormatContext } from "./numberFormat/render";
import { singleCell, type SheetSelection } from "./selection";
import { createBlankWorkbook } from "./sheets";

const ctx: FormatContext = { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" };
const OPTIONS: SheetFindOptions = { matchCase: false, entireCell: false, regex: false, lookIn: "values", within: "sheet" };
const ALL_TYPES: SpecialTypes = { numbers: true, text: true, logicals: true, errors: true };
/** styles[1]: a currency numFmt; [2]: unlocked; [3]: quote-prefixed. */
const CURRENCY = 1;
const UNLOCKED = 2;
const QUOTED = 3;

/** "B3" → { row: 2, col: 1 }. */
function at(a1: string): CellRef {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(a1)!;
  let col = 0;
  for (const letter of letters!) {
    col = col * 26 + letter.charCodeAt(0) - 64;
  }
  return { row: Number(digits) - 1, col: col - 1 };
}

/** Sparse rows from A1 → a value or a cell's fields. */
function rowsOf(cells: Record<string, CellValue | Partial<Cell>>): Row[] {
  const rows = new Map<number, Cell[]>();
  for (const [ref, spec] of Object.entries(cells)) {
    const { row, col } = at(ref);
    const fields = spec !== null && typeof spec === "object" && !("error" in spec) ? spec : { value: spec as CellValue };
    const cell: Cell = { col, value: null, ...fields };
    rows.set(row, [...(rows.get(row) ?? []), cell]);
  }
  return [...rows.entries()].toSorted(([a], [b]) => a - b).map(([index, list]) => ({ index, cells: list.toSorted((a, b) => a.col - b.col) }));
}

function book(sheets: Partial<Sheet>[]): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  model.sheets = sheets.map((sheet, index) => ({ ...base, id: `s${index + 1}`, name: `Sheet${index + 1}`, rows: [], ...sheet }));
  model.styles = [...model.styles, { numberFormat: '"$"#,##0.00' }, { protection: { locked: false } }, { quotePrefix: true }];
  return model;
}

function query(text: string, options: Partial<SheetFindOptions> = {}): CompiledCellQuery {
  const compiled = compileCellQuery(text, { ...OPTIONS, ...options });
  if (compiled === null || isQueryError(compiled)) {
    throw new Error(`"${text}" did not compile`);
  }
  return compiled;
}

const range = (a1: string, b1 = a1): GridRange => {
  const a = at(a1);
  const b = at(b1);
  return { top: a.row, left: a.col, bottom: b.row, right: b.col };
};
const select = (active: string, ...ranges: GridRange[]): SheetSelection => ({ active: at(active), ranges, activeRangeIndex: 0 });
const hit = (sheetId: string, a1: string): FindHit => ({ sheetId, ...at(a1) });

/** Every hit of a whole walk, as "s1!B3" strings. */
function find(model: SpreadsheetModel, text: string, options: Partial<SheetFindOptions> = {}, selection: SheetSelection = singleCell({ row: 0, col: 0 }), activeSheetId = model.sheets[0]!.id): string[] {
  const all = { ...OPTIONS, ...options };
  const search = createSheetSearch(model, query(text, all), all, activeSheetId, selection, ctx);
  search.step(() => false);
  expect(search.done).toBe(true);
  return search.hits.map((h) => `${h.sheetId}!${String.fromCharCode(65 + h.col)}${h.row + 1}`);
}

const ifNorth = (n: number, value: number): Partial<Cell> => ({ formula: `IF(A${n}="North",1,0)`, value });

/** The fixture's shape (spec134-find.xlsx), in memory. */
function salesBook(): SpreadsheetModel {
  const sales = rowsOf({
    A1: "Region",
    B1: "Amount",
    C1: "Flag",
    A2: "North",
    B2: { value: 1200, styleId: CURRENCY },
    C2: ifNorth(2, 1),
    A3: "South",
    B3: { value: 800, styleId: CURRENCY },
    C3: ifNorth(3, 0),
    A4: "Northwest",
    B4: { value: 450, styleId: CURRENCY },
    C4: ifNorth(4, 0),
    B5: { value: 300, styleId: CURRENCY },
    C5: { formula: "B5/0", value: { error: "#DIV/0!" } },
    A6: "north",
    B6: { value: 50, styleId: CURRENCY },
    C6: ifNorth(6, 0),
    B10: { formula: "SUM(B2:B6)", value: 2800 },
  });
  const model = book([
    { id: "sales", name: "Sales", rows: sales },
    { id: "locked", name: "Locked", rows: rowsOf({ A1: "North office", A2: "South office" }), protection: { sheet: true } },
    { id: "archive", name: "Archive", visibility: "hidden", rows: rowsOf({ A1: "North" }) },
  ]);
  model.names = [{ name: "Totals", formula: "Sales!$B$10" }];
  return model;
}

describe("searchText — per cell kind, both Look-in modes", () => {
  const cell = (model: SpreadsheetModel, a1: string, sheet = 0): { sheet: Sheet; cell: Cell; row: number } => {
    const s = model.sheets[sheet]!;
    const { row, col } = at(a1);
    return { sheet: s, row, cell: s.rows.find((r) => r.index === row)!.cells.find((c) => c.col === col)! };
  };
  const text = (model: SpreadsheetModel, a1: string, lookIn: "values" | "formulas"): string | null => {
    const found = cell(model, a1);
    return searchText(model, found.sheet, found.cell, lookIn, ctx, found.row);
  };

  it("a constant: its displayed text in Values, its edit text in Formulas", () => {
    const model = book([
      {
        rows: rowsOf({
          A1: { value: 1200, styleId: CURRENCY },
          A2: "North",
          A3: true,
          A4: { error: "#N/A" },
          A5: { value: "Plain", runs: [{ text: "Pl" }, { text: "ain" }] },
          A6: { value: "007", styleId: QUOTED },
          A7: 0.5,
        }),
      },
    ]);
    expect(text(model, "A1", "values")).toBe("$1,200.00");
    expect(text(model, "A1", "formulas")).toBe("1200");
    expect([text(model, "A2", "values"), text(model, "A2", "formulas")]).toEqual(["North", "North"]);
    expect([text(model, "A3", "values"), text(model, "A3", "formulas")]).toEqual(["TRUE", "TRUE"]);
    expect([text(model, "A4", "values"), text(model, "A4", "formulas")]).toEqual(["#N/A", "#N/A"]);
    // Rich text searches its plain value.
    expect(text(model, "A5", "values")).toBe("Plain");
    // A quote-prefixed constant's edit text keeps its "'".
    expect([text(model, "A6", "values"), text(model, "A6", "formulas")]).toEqual(["007", "'007"]);
    expect(text(model, "A7", "values")).toBe("0.5");
  });

  it("a formula cell: its displayed result in Values, \"=\" + formula in Formulas", () => {
    const model = book([{ rows: rowsOf({ A1: { formula: 'IF(B1="North",1,0)', value: 1 }, A2: { formula: "B5/0", value: { error: "#DIV/0!" } }, A3: { formula: "NOW()", value: null } }) }]);
    expect(text(model, "A1", "values")).toBe("1");
    expect(text(model, "A1", "formulas")).toBe('=IF(B1="North",1,0)');
    expect(text(model, "A2", "values")).toBe("#DIV/0!");
    // An uncomputed formula shows nothing but is still a formula to search.
    expect(text(model, "A3", "values")).toBe("");
    expect(text(model, "A3", "formulas")).toBe("=NOW()");
  });

  it("an array: the anchor searched in both modes, a member in Values only", () => {
    const model = book([{ rows: rowsOf({ A1: { formula: "ROW(A1:A3)", value: 1, arrayRange: range("A1", "A3") }, A2: 2, A3: 3 }) }]);
    expect(text(model, "A1", "formulas")).toBe("=ROW(A1:A3)");
    expect(text(model, "A2", "values")).toBe("2");
    expect(text(model, "A2", "formulas")).toBeNull();
  });

  it("never searches a cell storing no value, nor a merge-covered cell", () => {
    const model = book([{ rows: rowsOf({ A1: { value: null, styleId: CURRENCY }, B1: "Anchor", C1: "Covered" }), merges: [range("B1", "C1")] }]);
    expect(text(model, "A1", "values")).toBeNull();
    expect(text(model, "A1", "formulas")).toBeNull();
    expect(text(model, "B1", "values")).toBe("Anchor");
    expect(text(model, "C1", "values")).toBeNull();
    expect(text(model, "C1", "formulas")).toBeNull();
  });

  it("shows a zero as nothing on a sheet hiding zeros, and finds the row itself without one", () => {
    const model = book([{ rows: rowsOf({ A1: 0 }), showZeros: false }]);
    expect(text(model, "A1", "values")).toBe("");
    const found = cell(model, "A1");
    expect(searchText(model, found.sheet, found.cell, "formulas", ctx)).toBe("0");
  });
});

describe("match rules", () => {
  it("is case-insensitive unless Match case", () => {
    const model = salesBook();
    expect(find(model, "north")).toEqual(["sales!A2", "sales!A4", "sales!A6"]);
    expect(find(model, "north", { matchCase: true })).toEqual(["sales!A6"]);
  });

  it("anchors the whole search text for Match entire cell", () => {
    expect(find(salesBook(), "north", { entireCell: true })).toEqual(["sales!A2", "sales!A6"]);
    expect(query("a.b", { entireCell: true, regex: true }).global.source).toBe("^(?:a.b)$");
    expect(query("north", { entireCell: true }).test.flags).toBe("i");
  });

  it("keeps a non-regex query literal: * and ? are no wildcards", () => {
    const model = book([{ rows: rowsOf({ A1: "a*b", A2: "axxb", A3: "what?", A4: "whatever" }) }]);
    expect(find(model, "a*b")).toEqual(["s1!A1"]);
    expect(find(model, "what?")).toEqual(["s1!A3"]);
  });

  it("uses regular expressions per SPEC-017 — and refuses an invalid one", () => {
    const model = book([{ rows: rowsOf({ A1: "a1", A2: "b22", A3: "c" }) }]);
    expect(find(model, "^[a-z]\\d+$", { regex: true })).toEqual(["s1!A1", "s1!A2"]);
    expect(compileCellQuery("(open", { ...OPTIONS, regex: true })).toEqual({ error: "Invalid regular expression" });
    expect(compileCellQuery("", OPTIONS)).toBeNull();
  });

  it("matches a cell only on a non-empty match", () => {
    const q = query("x*", { regex: true });
    expect(cellMatches("abc", q)).toBe(false);
    expect(cellMatches("abxc", q)).toBe(true);
    expect(cellMatches("", query("^$", { regex: true }))).toBe(false);
  });

  it("counts cells, not occurrences", () => {
    const model = book([{ rows: rowsOf({ A1: "north north north", A2: "north" }) }]);
    expect(find(model, "north")).toHaveLength(2);
  });

  it("looks in formulas when asked: the fixture's 6 of 8 in the workbook", () => {
    const model = salesBook();
    expect(find(model, "north", { lookIn: "formulas" })).toEqual(["sales!A2", "sales!C2", "sales!C3", "sales!A4", "sales!C4", "sales!A6", "sales!C6"]);
    expect(find(model, "north", { lookIn: "formulas", within: "workbook" })).toHaveLength(8);
  });
});

describe("scope and order", () => {
  it("searches the whole active sheet, row-major, for a single selected cell", () => {
    const model = book([{ rows: rowsOf({ C1: "x", A2: "x", B1: "x" }) }, { rows: rowsOf({ A1: "x" }) }]);
    expect(find(model, "x")).toEqual(["s1!B1", "s1!C1", "s1!A2"]);
    expect(find(model, "x", {}, singleCell({ row: 0, col: 0 }), "s2")).toEqual(["s2!A1"]);
  });

  it("searches only a multi-cell selection's ranges, in ranges order, a shared cell once", () => {
    const model = book([{ rows: rowsOf({ A1: "x", B2: "x", C3: "x", D4: "x", B3: "x" }) }]);
    const selection = select("C3", range("C3", "D4"), range("A1", "B3"), range("B2", "C3"));
    expect(find(model, "x", {}, selection)).toEqual(["s1!C3", "s1!D4", "s1!A1", "s1!B2", "s1!B3"]);
    // A merged single cell is one cell: the whole sheet is searched.
    const merged = book([{ rows: rowsOf({ A1: "x", D9: "x" }), merges: [range("A1", "B2")] }]);
    expect(find(merged, "x", {}, select("A1", range("A1", "B2")))).toEqual(["s1!A1", "s1!D9"]);
  });

  it("searches every visible grid sheet in tab order for Within = Workbook, ignoring the selection", () => {
    const model = book([
      { rows: rowsOf({ B2: "x" }) },
      { rows: rowsOf({ A1: "x" }), visibility: "hidden" },
      { rows: rowsOf({ A1: "x" }), visibility: "veryHidden" },
      { rows: rowsOf({ A1: "x" }), kind: "foreign" },
      { rows: rowsOf({ A1: "x" }) },
    ]);
    expect(find(model, "x", { within: "workbook" }, select("B2", range("B2", "C3")))).toEqual(["s1!B2", "s5!A1"]);
    expect(findAreas(model, "workbook", "s1", singleCell({ row: 0, col: 0 })).map((area) => area.sheet.id)).toEqual(["s1", "s5"]);
    expect(findAreas(model, "sheet", "s4", singleCell({ row: 0, col: 0 }))).toEqual([]);
  });

  it("searches hidden rows and columns, and knows a hit is in one", () => {
    const sheet: Partial<Sheet> = {
      rows: [{ index: 0, cells: [{ col: 0, value: "x" }] }, { index: 1, hidden: true, cells: [{ col: 0, value: "x" }] }, { index: 2, filterHidden: true, cells: [{ col: 0, value: "x" }, { col: 2, value: "x" }] }],
      columns: [{ min: 2, max: 2, hidden: true }],
    };
    const model = book([sheet]);
    expect(find(model, "x")).toEqual(["s1!A1", "s1!A2", "s1!A3", "s1!C3"]);
    expect([hit("s1", "A1"), hit("s1", "A2"), hit("s1", "A3"), hit("s1", "C1")].map((h) => hitIsHidden(model, h))).toEqual([false, true, true, true]);
  });

  it("locates the first hit at or after, after, and before a cell, wrapping at either end", () => {
    const model = salesBook();
    const search = createSheetSearch(model, query("north"), OPTIONS, "sales", singleCell({ row: 0, col: 0 }), ctx);
    search.step(() => false);
    expect(search.locate(hit("sales", "A1"), "atOrAfter")).toEqual({ kind: "hit", index: 0, wrapped: false });
    expect(search.locate(hit("sales", "A2"), "atOrAfter")).toEqual({ kind: "hit", index: 0, wrapped: false });
    expect(search.locate(hit("sales", "A2"), "after")).toEqual({ kind: "hit", index: 1, wrapped: false });
    expect(search.locate(hit("sales", "A6"), "after")).toEqual({ kind: "hit", index: 0, wrapped: true });
    expect(search.locate(hit("sales", "A4"), "before")).toEqual({ kind: "hit", index: 0, wrapped: false });
    expect(search.locate(hit("sales", "A2"), "before")).toEqual({ kind: "hit", index: 2, wrapped: true });
    expect(search.locate(hit("sales", "Z99"), "after")).toEqual({ kind: "hit", index: 0, wrapped: true });
    // A cell outside the scope stands before it.
    expect(search.locate(hit("locked", "A1"), "after")).toEqual({ kind: "hit", index: 0, wrapped: false });
  });

  it("locates in a multi-range selection by range order, and answers none for no hit", () => {
    const model = book([{ rows: rowsOf({ A1: "x", D4: "x", C3: "y" }) }]);
    const selection = select("D4", range("C3", "D4"), range("A1", "B2"));
    const search = createSheetSearch(model, query("x"), OPTIONS, "s1", selection, ctx);
    search.step(() => false);
    // D4 is in the first range, A1 in the second: after D4 comes A1.
    expect(search.hits.map((h) => h.row)).toEqual([3, 0]);
    expect(search.locate(hit("s1", "D4"), "after")).toEqual({ kind: "hit", index: 1, wrapped: false });
    expect(search.locate(hit("s1", "A1"), "after")).toEqual({ kind: "hit", index: 0, wrapped: true });
    const none = createSheetSearch(model, query("zzz"), OPTIONS, "s1", selection, ctx);
    none.step(() => false);
    expect(none.locate(hit("s1", "A1"), "after")).toEqual({ kind: "none" });
    expect(none.locate(hit("s1", "A1"), "before")).toEqual({ kind: "none" });
  });

  it("answers \"pending\" until the walk has gone far enough", () => {
    const rows: Row[] = Array.from({ length: 2000 }, (_, index) => ({ index, cells: [{ col: 0, value: index === 1500 ? "target" : "filler" }] }));
    const model = book([{ rows }]);
    const search = createSheetSearch(model, query("target"), OPTIONS, "s1", singleCell({ row: 0, col: 0 }), ctx);
    let asked = 0;
    search.step(() => (asked += 1) >= 2);
    expect(search.done).toBe(false);
    expect(search.hits).toEqual([]);
    expect(search.locate(hit("s1", "A1"), "after")).toEqual({ kind: "pending" });
    expect(search.locate(hit("s1", "A1"), "before")).toEqual({ kind: "pending" });
    expect(search.locate(hit("s1", "A5"), "before")).toEqual({ kind: "pending" });
    search.step(() => false);
    expect(search.locate(hit("s1", "A1"), "after")).toEqual({ kind: "hit", index: 0, wrapped: false });
  });
});

describe("the walk", () => {
  it("never overruns its deadline by more than one 256-cell check interval (injected clock)", () => {
    // The clock is a counter every read of a cell's value advances: time
    // passes only as the walk does work, deterministically.
    let reads = 0;
    const COST_MS = 0.004;
    const now = (): number => reads * COST_MS;
    const rows: Row[] = Array.from({ length: 4000 }, (_, index) => ({
      index,
      cells: Array.from({ length: 6 }, (__, col) => {
        const cell = { col } as Cell;
        const value = (index + col) % 97 === 0 ? "needle" : `hay ${index}`;
        Object.defineProperty(cell, "value", { enumerable: true, get: () => ((reads += 1), value) });
        return cell;
      }),
    }));
    const model = book([{ rows }]);
    // Calibrate: how many value reads one visited cell costs.
    const calibration = createSheetSearch(model, query("needle"), OPTIONS, "s1", singleCell({ row: 0, col: 0 }), ctx);
    reads = 0;
    calibration.step(() => false);
    const readsPerCell = reads / (4000 * 6);
    const search = createSheetSearch(model, query("needle"), OPTIONS, "s1", singleCell({ row: 0, col: 0 }), ctx);
    let slots = 0;
    while (!search.done) {
      const start = now();
      search.step(() => now() - start >= FIND_SLICE_BUDGET_MS);
      const overrun = now() - start - FIND_SLICE_BUDGET_MS;
      expect(overrun).toBeLessThanOrEqual(FIND_CHECK_INTERVAL * readsPerCell * COST_MS + 1e-9);
      slots += 1;
    }
    expect(slots).toBeGreaterThan(5);
    expect(search.hits.map((h) => `${h.row}:${h.col}`)).toEqual(calibration.hits.map((h) => `${h.row}:${h.col}`));
  });

  it("visits stored cells only: a sparse whole column costs its stored rows, never its area", () => {
    const rows: Row[] = [0, 400_000, MAX_ROWS - 1].map((index) => ({ index, cells: [{ col: 0, value: "a" }, { col: 1, value: "x" }, { col: MAX_COLS - 1, value: "x" }] }));
    const model = book([{ rows }]);
    const selection = select("B1", range("B1", `B${MAX_ROWS}`), range("C1", `C${MAX_ROWS}`));
    const search = createSheetSearch(model, query("x"), OPTIONS, "s1", selection, ctx);
    search.step(() => false);
    expect(search.hits.map((h) => [h.row, h.col])).toEqual([
      [0, 1],
      [400_000, 1],
      [MAX_ROWS - 1, 1],
    ]);
    // Three rows examined and three cells visited for B, three rows for C.
    expect(search.work).toBe(9);
    const whole = createSheetSearch(model, query("x"), OPTIONS, "s1", singleCell({ row: 0, col: 0 }), ctx);
    whole.step(() => false);
    expect(whole.hits).toHaveLength(6);
    expect(whole.work).toBe(3 + 9);
  });

  it("is resumable: slots of any size find the same hits as one pass", () => {
    const model = salesBook();
    const options: SheetFindOptions = { ...OPTIONS, lookIn: "formulas", within: "workbook" };
    const search = createSheetSearch(model, query("north"), options, "sales", singleCell({ row: 0, col: 0 }), ctx);
    let calls = 0;
    while (!search.done) {
      search.step(() => (calls += 1) % 1 === 0);
    }
    expect(search.hits).toHaveLength(8);
  });
});

describe("the replace plan", () => {
  it("re-enters a constant's text as a value in Values mode and counts formula cells as formula results", () => {
    const model = salesBook();
    const hits = [hit("sales", "A2"), hit("sales", "C2"), hit("sales", "A4")];
    const plan = planReplace(model, hits, query("north"), "N.", { lookIn: "values" }, ctx);
    expect(plan.edits).toEqual([
      { sheetId: "sales", row: 1, col: 0, text: "N.", asValue: true },
      { sheetId: "sales", row: 3, col: 0, text: "N.west", asValue: true },
    ]);
    expect(plan.occurrences).toBe(2);
    expect(plan.skipped).toEqual({ breaksFormula: 0, protected: 0, array: 0, formulaResult: 1 });
  });

  it("replaces every occurrence in a cell's formula text in Formulas mode, as typed", () => {
    const model = book([{ rows: rowsOf({ A1: { formula: 'IF(B1="North","North",0)', value: "North" }, A2: "north north" }) }]);
    const plan = planReplace(model, [hit("s1", "A1"), hit("s1", "A2")], query("north"), "N.", { lookIn: "formulas" }, ctx);
    expect(plan.edits).toEqual([
      { sheetId: "s1", row: 0, col: 0, text: '=IF(B1="N.","N.",0)', asValue: false },
      { sheetId: "s1", row: 1, col: 0, text: "N. N.", asValue: false },
    ]);
    expect(plan.occurrences).toBe(4);
  });

  it("refuses a formula the replacement would break, and skips protected cells and arrays", () => {
    const model = book([
      {
        rows: rowsOf({
          A1: { formula: "SUM(B1:B3)", value: 6 },
          A2: { formula: "ROW(A2:A3)", value: 2, arrayRange: range("A2", "A3") },
          A3: 3,
          A4: "SUM here",
        }),
      },
      { rows: rowsOf({ A1: "SUM locked", A2: { value: "SUM open", styleId: UNLOCKED } }), protection: { sheet: true } },
    ]);
    const hits = [hit("s1", "A1"), hit("s1", "A2"), hit("s1", "A3"), hit("s1", "A4"), hit("s2", "A1"), hit("s2", "A2")];
    const plan = planReplace(model, hits, query("SUM"), "SUM(", { lookIn: "formulas" }, ctx);
    expect(plan.skipped).toEqual({ breaksFormula: 1, protected: 1, array: 2, formulaResult: 0 });
    expect(plan.edits.map((edit) => [edit.sheetId, edit.row, edit.text])).toEqual([
      ["s1", 3, "SUM( here"],
      ["s2", 1, "SUM( open"],
    ]);
    expect(replaceReport(plan, plan.edits.length)).toBe("Replaced 2 occurrences in 2 cells · skipped 1 that would break a formula, 1 protected, 2 in array formulas");
  });

  it("refuses a new formula over the 8,192-character limit", () => {
    const model = book([{ rows: rowsOf({ A1: { formula: "1+X", value: 1 } }) }]);
    const plan = planReplace(model, [hit("s1", "A1")], query("X"), "1".repeat(8192), { lookIn: "formulas" }, ctx);
    expect(plan.skipped.breaksFormula).toBe(1);
  });

  it("replaces literally: $1 and $& mean nothing", () => {
    const model = book([{ rows: rowsOf({ A1: "cost: 5" }) }]);
    const plan = planReplace(model, [hit("s1", "A1")], query("(\\d)", { regex: true }), "$1$&$$", { lookIn: "values" }, ctx);
    expect(plan.edits[0]!.text).toBe("cost: $1$&$$");
  });

  it("replaces only non-empty matches, mutates nothing, and ignores stale hits", () => {
    const model = book([{ rows: rowsOf({ A1: "abc" }) }]);
    const before = structuredClone(model);
    const plan = planReplace(model, [hit("s1", "A1"), hit("s1", "Z9")], query("b*", { regex: true }), "-", { lookIn: "values" }, ctx);
    expect(plan.edits).toEqual([{ sheetId: "s1", row: 0, col: 0, text: "a-c", asValue: true }]);
    expect(plan.occurrences).toBe(1);
    expect(model).toEqual(before);
  });

  it("reports with singular forms and the skip parts in the Messages table's order", () => {
    const skipped = { breaksFormula: 0, protected: 1, array: 0, formulaResult: 0 };
    expect(replaceReport({ occurrences: 7, skipped }, 7)).toBe("Replaced 7 occurrences in 7 cells · skipped 1 protected");
    expect(replaceReport({ occurrences: 1, skipped: { ...skipped, protected: 0 } }, 1)).toBe("Replaced 1 occurrence in 1 cell");
    expect(replaceReport({ occurrences: 0, skipped: { breaksFormula: 0, protected: 0, array: 0, formulaResult: 3 } }, 0)).toBe(
      "Replaced 0 occurrences in 0 cells · skipped 3 formula results",
    );
    expect(breaksFormulaMessage("B4")).toBe("That replacement would break the formula in B4");
    expect(findAllCapMessage(12345)).toBe("Showing the first 5,000 of 12,345 matches");
    expect(FIND_ALL_LIMIT).toBe(5000);
  });
});

describe("resolveGoToTarget — the Go To grammar", () => {
  function goToBook(): SpreadsheetModel {
    const table: SheetTable = {
      id: "t1",
      name: "Orders",
      ref: range("A7", "B10"),
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Item" }, { name: "Qty" }],
      style: { showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    };
    const model = book([
      { id: "a", name: "Sheet1" },
      { id: "b", name: "My Sheet", tables: [table] },
      { id: "h", name: "Archive", visibility: "hidden" },
      { id: "f", name: "Dialog", kind: "foreign" },
    ]);
    model.names = [
      { name: "Totals", formula: "Sheet1!$B$10" },
      { name: "Pair", formula: "Sheet1!$A$1,Sheet1!$C$3:$D$4" },
      { name: "TaxRate", formula: "0.2" },
      { name: "Local", scopeSheetId: "b", formula: "'My Sheet'!$C$1" },
      { name: "Above", formula: "A1048576" },
      { name: "Secret", formula: "Sheet1!$A$1", hidden: true },
      { name: "_xlnm.Print_Area", scopeSheetId: "a", formula: "Sheet1!$A$1:$D$10" },
      { name: "Ghost", formula: "Archive!$A$1" },
    ];
    return model;
  }

  it("accepts cells, ranges, whole columns and rows — case-insensitive, $-tolerant, sheet-qualified", () => {
    const model = goToBook();
    expect(resolveGoToTarget(model, "B12", "a")).toEqual({ sheetId: "a", ranges: [range("B12")] });
    expect(resolveGoToTarget(model, "c3:a1", "a")).toEqual({ sheetId: "a", ranges: [range("A1", "C3")] });
    expect(resolveGoToTarget(model, "$B$3:$C4", "a")).toEqual({ sheetId: "a", ranges: [range("B3", "C4")] });
    expect(resolveGoToTarget(model, "A:C", "a")).toEqual({ sheetId: "a", ranges: [{ top: 0, left: 0, bottom: MAX_ROWS - 1, right: 2 }] });
    expect(resolveGoToTarget(model, "3:3", "a")).toEqual({ sheetId: "a", ranges: [{ top: 2, left: 0, bottom: 2, right: MAX_COLS - 1 }] });
    expect(resolveGoToTarget(model, "'My Sheet'!B4", "a")).toEqual({ sheetId: "b", ranges: [range("B4")] });
    expect(resolveGoToTarget(model, "sheet1!a1:c3", "b")).toEqual({ sheetId: "a", ranges: [range("A1", "C3")] });
    expect(resolveGoToTarget(model, "XFD1048576", "a")).toEqual({ sheetId: "a", ranges: [{ top: MAX_ROWS - 1, left: MAX_COLS - 1, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }] });
  });

  it("splits on top-level commas: every part on one sheet", () => {
    const model = goToBook();
    expect(resolveGoToTarget(model, "A1, C3:D4,Sheet1!F6", "a")).toEqual({ sheetId: "a", ranges: [range("A1"), range("C3", "D4"), range("F6")] });
    expect(resolveGoToTarget(model, "A1,'My Sheet'!B2", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
    expect(resolveGoToTarget(model, "A1,,B2", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
    expect(resolveGoToTarget(model, "(A1,B2)", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
  });

  it("accepts a visible name (its range, or ranges) and a table (its data body)", () => {
    const model = goToBook();
    expect(resolveGoToTarget(model, "Totals", "b")).toEqual({ sheetId: "a", ranges: [range("B10")] });
    expect(resolveGoToTarget(model, "totals", "a")).toEqual({ sheetId: "a", ranges: [range("B10")] });
    expect(resolveGoToTarget(model, "Pair", "a")).toEqual({ sheetId: "a", ranges: [range("A1"), range("C3", "D4")] });
    expect(resolveGoToTarget(model, "Local", "b")).toEqual({ sheetId: "b", ranges: [range("C1")] });
    expect(resolveGoToTarget(model, "'My Sheet'!Local", "a")).toEqual({ sheetId: "b", ranges: [range("C1")] });
    // A relative name evaluates at the active cell (A1 by default).
    expect(resolveGoToTarget(model, "Above", "a", at("C6"))).toEqual({ sheetId: "a", ranges: [range("C5")] });
    expect(resolveGoToTarget(model, "orders", "a")).toEqual({ sheetId: "b", ranges: [range("A8", "B10")] });
    expect(resolveGoToTarget(model, "Totals,A1", "a")).toEqual({ sheetId: "a", ranges: [range("B10"), range("A1")] });
  });

  it("rejects 3-D, external and R1C1 references, intersections, the grid's outside, unknown sheets and names", () => {
    const model = goToBook();
    for (const text of ["Sheet1:'My Sheet'!A1", "[1]Sheet1!A1", "R1C1", "R[1]C[1]", "A1:B2 B1:C3", "XFE1", "A1048577", "A1:XFD1048577", "Nowhere!A1", "Nobody", "Sheet1!Totals2", "Nowhere!", "=A1", "", "A1+B1", "Local"]) {
      expect(resolveGoToTarget(model, text, "a"), text).toEqual({ error: GO_TO_INVALID_MESSAGE });
    }
  });

  it("never resolves hidden or reserved names", () => {
    const model = goToBook();
    expect(resolveGoToTarget(model, "Secret", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
    expect(resolveGoToTarget(model, "Print_Area", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
    expect(resolveGoToTarget(model, "_xlnm.Print_Area", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
  });

  it("refuses a name that is no reference, a hidden sheet's target and a foreign sheet", () => {
    const model = goToBook();
    expect(resolveGoToTarget(model, "TaxRate", "a")).toEqual({ error: GO_TO_NOT_A_RANGE_MESSAGE });
    expect(resolveGoToTarget(model, "Archive!A1", "a")).toEqual({ error: GO_TO_HIDDEN_SHEET_MESSAGE });
    expect(resolveGoToTarget(model, "Ghost", "a")).toEqual({ error: GO_TO_HIDDEN_SHEET_MESSAGE });
    expect(resolveGoToTarget(model, "Dialog!A1", "a")).toEqual({ error: GO_TO_INVALID_MESSAGE });
  });
});

const special = (model: SpreadsheetModel, kind: Parameters<typeof selectSpecial>[3], types: SpecialTypes = ALL_TYPES, selection: SheetSelection = singleCell({ row: 0, col: 0 })) =>
  selectSpecial(model, model.sheets[0]!, selection, kind, types);

describe("selectSpecial — Go To Special", () => {
  function specialBook(): SpreadsheetModel {
    return book([
      {
        rows: rowsOf({
          A1: "Region",
          B1: 12,
          C1: true,
          A2: { error: "#N/A" },
          B2: { formula: "B1*2", value: 24 },
          C2: { formula: 'IF(B1>0,"","x")', value: "" },
          A3: { value: null, styleId: CURRENCY },
          B3: { formula: "C1", value: false },
          C3: { formula: "1/0", value: { error: "#DIV/0!" } },
          A4: { formula: "NOW()", value: null },
          B4: { formula: "ROW(B4:B5)", value: 4, arrayRange: range("B4", "B5") },
          B5: 5,
          D5: 45000,
        }),
      },
    ]);
  }

  it("selects blanks — absent or null without a formula; a formula returning \"\" is no blank", () => {
    // D1:D3 stops where row 4's run widens to C4:D4.
    expect(special(specialBook(), "blanks")).toEqual([range("D1", "D3"), range("A3"), range("C4", "D4"), range("A5"), range("C5")]);
  });

  it("selects constants by type: numbers (dates included), text, logicals, errors — never array members", () => {
    const model = specialBook();
    expect(special(model, "constants")).toEqual([range("A1", "C1"), range("A2"), range("D5")]);
    expect(special(model, "constants", { ...ALL_TYPES, text: false, logicals: false, errors: false })).toEqual([range("B1"), range("D5")]);
    expect(special(model, "constants", { numbers: false, text: false, logicals: false, errors: true })).toEqual([range("A2")]);
  });

  it("selects formulas and array members by their value's type: \"\" is text, an uncomputed null a number", () => {
    const model = specialBook();
    expect(special(model, "formulas")).toEqual([range("B2", "C3"), range("A4", "B4"), range("B5")]);
    expect(special(model, "formulas", { numbers: false, text: true, logicals: false, errors: false })).toEqual([range("C2")]);
    expect(special(model, "formulas", { numbers: true, text: false, logicals: false, errors: false })).toEqual([range("B2"), range("A4", "B4"), range("B5")]);
    expect(special(model, "formulas", { numbers: false, text: false, logicals: true, errors: false })).toEqual([range("B3")]);
  });

  it("selects every error, constant or formula, and the last cell", () => {
    const model = specialBook();
    expect(special(model, "errors")).toEqual([range("A2"), range("C3")]);
    expect(special(model, "lastCell", ALL_TYPES, select("A1", range("A1", "B2")))).toEqual([range("D5")]);
    expect(special(book([{}]), "lastCell")).toEqual([range("A1")]);
  });

  it("searches a multi-cell selection's ranges clipped to the used range, else the used range", () => {
    const model = specialBook();
    expect(special(model, "constants", ALL_TYPES, select("A1", range("A1", "A1048576")))).toEqual([range("A1", "A2")]);
    expect(special(model, "blanks", ALL_TYPES, select("C4", { top: 3, left: 2, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }))).toEqual([range("C4", "D4"), range("C5")]);
  });

  it("selects visible cells only as the cross product of visible row and column runs", () => {
    const model = book([
      {
        rows: [
          { index: 0, cells: [{ col: 0, value: 1 }, { col: 4, value: 1 }] },
          { index: 1, hidden: true, cells: [] },
          { index: 2, filterHidden: true, cells: [] },
          { index: 4, cells: [{ col: 0, value: 1 }] },
        ],
        columns: [{ min: 2, max: 2, hidden: true }],
      },
    ]);
    expect(special(model, "visible")).toEqual([range("A1", "B1"), range("D1", "E1"), range("A4", "B5"), range("D4", "E5")]);
  });

  it("finds nothing on an empty sheet or with no match, and refuses more than 10,000 areas", () => {
    expect(special(book([{}]), "blanks")).toEqual({ error: GO_TO_SPECIAL_NONE_MESSAGE });
    expect(special(specialBook(), "constants", { numbers: false, text: false, logicals: false, errors: false })).toEqual({ error: GO_TO_SPECIAL_NONE_MESSAGE });
    // A checkerboard of 20,001 constants is 20,001 rectangles.
    const rows: Row[] = Array.from({ length: 10_001 }, (_, index) => ({ index, cells: [{ col: index % 2, value: 1 }, { col: 2 + (index % 2), value: 1 }] }));
    const model = book([{ rows }]);
    expect(GO_TO_SPECIAL_MAX_RANGES).toBe(10_000);
    expect(special(model, "constants")).toEqual({ error: GO_TO_SPECIAL_TOO_MANY_MESSAGE });
    expect(special(model, "blanks")).toEqual({ error: GO_TO_SPECIAL_TOO_MANY_MESSAGE });
  });

  it("computes blanks per row, never per position: a huge empty block is one rectangle", () => {
    const model = book([{ rows: rowsOf({ A1: 1, C1048576: 1 }) }]);
    expect(special(model, "blanks")).toEqual([range("B1", "C1"), { top: 1, left: 0, bottom: MAX_ROWS - 2, right: 2 }, range("A1048576", "B1048576")]);
  });
});

describe("coalesceCells", () => {
  it("sweeps maximal horizontal runs, extending the rectangle directly above with the same left and right", () => {
    const cells = ["A1", "B1", "A2", "B2", "A3", "C3", "A4", "B4", "B5"].map(at);
    expect(coalesceCells(cells)).toEqual([range("A1", "B2"), range("A3"), range("C3"), range("A4", "B4"), range("B5")]);
  });

  it("sorts and deduplicates its input, and never joins rows that are not adjacent", () => {
    expect(coalesceCells(["B2", "A2", "B2", "A4", "B4"].map(at))).toEqual([range("A2", "B2"), range("A4", "B4")]);
    expect(coalesceCells([])).toEqual([]);
  });
});
