// SPEC-128 "Sort semantics", "The header guess", "The sort record": the type
// order both ways, collation in both sensitivities across every
// SpreadsheetLocale, stability, the permutation over hidden rows, formula
// shifting, every refusal, the permutation hook and the header guess table.
import { afterEach, describe, expect, it } from "vitest";
import { getCell, type GridRange } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import { cellKey } from "./formula/dependencies";
import type { Cell, CellValue, Row, Sheet, SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALE_TAGS } from "./numberFormat/locale";
import {
  compareSortValues,
  duplicateKeyMessage,
  guessHasHeader,
  MAX_SORT_KEYS,
  ROW_PERMUTATION_ADJUSTERS,
  SORT_REFUSAL_MESSAGES,
  sortCollator,
  sortRange,
  type SortOptions,
} from "./sort";

function sheet(fields: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], validations: [], id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(s: Sheet, fields: Partial<SpreadsheetModel> = {}): SpreadsheetModel {
  return { dxfs: [], kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 9, sheets: [s], activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}, { font: { bold: true } }], defaultStyleId: 0, ...fields };
}

const cell = (col: number, value: CellValue, extra: Partial<Cell> = {}): Cell => ({ col, value, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const err = (code: "#N/A" | "#DIV/0!" | "#REF!"): CellValue => ({ error: code });

/** One column A of `values` from row 0, no header. */
function column(values: CellValue[]): SpreadsheetModel {
  return model(sheet({ rows: values.map((v, i) => row(i, v === null ? [] : [cell(0, v)])).filter((r) => r.cells.length > 0) }));
}

const valuesOf = (m: SpreadsheetModel, col: number, top: number, bottom: number): CellValue[] =>
  Array.from({ length: bottom - top + 1 }, (_, i) => getCell(m.sheets[0]!, top + i, col)?.value ?? null);

const options = (keys: SortOptions["keys"], extra: Partial<SortOptions> = {}): SortOptions => ({ keys, hasHeader: false, caseSensitive: false, localeTag: "en-US", ...extra });

afterEach(() => {
  ROW_PERMUTATION_ADJUSTERS.length = 0;
});

describe("MAX_SORT_KEYS", () => {
  it("is 64", () => {
    expect(MAX_SORT_KEYS).toBe(64);
  });
});

describe("compareSortValues — the type order", () => {
  const collator = sortCollator("en-US", false);
  const mixed: CellValue[] = [true, err("#N/A"), "b", 3, false, "a", -1, err("#DIV/0!")];

  it("ascending: numbers < text < FALSE < TRUE < errors, all errors equal", () => {
    const sorted = mixed.toSorted((a, b) => compareSortValues(a, b, collator));
    expect(sorted.slice(0, 6)).toEqual([-1, 3, "a", "b", false, true]);
    expect(compareSortValues(err("#N/A"), err("#DIV/0!"), collator)).toBe(0);
    expect(compareSortValues(1e9, "0", collator)).toBeLessThan(0);
    expect(compareSortValues(false, true, collator)).toBeLessThan(0);
  });

  it("sortRange ascending and descending: every class reversed, blanks last both ways", () => {
    const values: CellValue[] = [true, null, "b", 3, false, err("#N/A"), "a", -1, null];
    const asc = column(values);
    expect(sortRange(asc, "s1", range(0, 0, 8, 0), options([{ col: 0, descending: false }])).ok).toBe(true);
    expect(valuesOf(asc, 0, 0, 8)).toEqual([-1, 3, "a", "b", false, true, err("#N/A"), null, null]);
    const desc = column(values);
    expect(sortRange(desc, "s1", range(0, 0, 8, 0), options([{ col: 0, descending: true }])).ok).toBe(true);
    expect(valuesOf(desc, 0, 0, 8)).toEqual([err("#N/A"), true, false, "b", "a", 3, -1, null, null]);
  });

  it("a formula returning \"\" is text, not blank; a formula with no value is blank", () => {
    const m = model(sheet({ rows: [row(0, [cell(0, null, { formula: "B1" })]), row(1, [cell(0, "", { formula: "\"\"" })]), row(2, [cell(0, 5)])] }));
    sortRange(m, "s1", range(0, 0, 2, 0), options([{ col: 0, descending: false }]));
    expect(valuesOf(m, 0, 0, 2)).toEqual([5, "", null]);
  });

  it("dates are numbers", () => {
    expect(compareSortValues(45000, 45001, collator)).toBeLessThan(0);
  });
});

describe("sortCollator — collation in both sensitivities, every locale", () => {
  for (const tag of SPREADSHEET_LOCALE_TAGS) {
    it(`${tag}: A10 before A2; apple ties Apple by default; apple before Apple when case-sensitive`, () => {
      const insensitive = sortCollator(tag, false);
      const sensitive = sortCollator(tag, true);
      expect(insensitive.compare("A10", "A2")).toBeLessThan(0);
      expect(insensitive.compare("apple", "Apple")).toBe(0);
      expect(sensitive.compare("apple", "Apple")).toBeLessThan(0);
      expect(insensitive.resolvedOptions()).toMatchObject({ usage: "sort", numeric: false, sensitivity: "accent", ignorePunctuation: false });
      expect(sensitive.resolvedOptions()).toMatchObject({ usage: "sort", numeric: false, sensitivity: "variant", caseFirst: "lower" });
    });
  }

  it("ties in a case-insensitive sort keep their order; a case-sensitive sort puts lower case first", () => {
    const m = column(["Apple", "apple", "APPLE"]);
    sortRange(m, "s1", range(0, 0, 2, 0), options([{ col: 0, descending: false }]));
    expect(valuesOf(m, 0, 0, 2)).toEqual(["Apple", "apple", "APPLE"]);
    const cs = column(["Apple", "apple"]);
    sortRange(cs, "s1", range(0, 0, 1, 0), options([{ col: 0, descending: false }], { caseSensitive: true }));
    expect(valuesOf(cs, 0, 0, 1)).toEqual(["apple", "Apple"]);
  });
});

describe("sortRange — stability and keys", () => {
  const people = (): SpreadsheetModel =>
    model(
      sheet({
        rows: [
          row(0, [cell(0, "Name"), cell(1, "Team"), cell(2, "Score")]),
          row(1, [cell(0, "Ann"), cell(1, "B"), cell(2, 3)]),
          row(2, [cell(0, "Bob"), cell(1, "A"), cell(2, 3)]),
          row(3, [cell(0, "Cy"), cell(1, "B"), cell(2, 1)]),
          row(4, [cell(0, "Di"), cell(1, "A"), cell(2, 3)]),
        ],
      }),
    );

  it("ties fall through to the next key, then to the original order", () => {
    const m = people();
    const outcome = sortRange(m, "s1", range(0, 0, 4, 2), options([{ col: 2, descending: true }], { hasHeader: true }));
    expect(outcome.ok).toBe(true);
    expect(valuesOf(m, 0, 1, 4)).toEqual(["Ann", "Bob", "Di", "Cy"]);
    const two = people();
    sortRange(two, "s1", range(0, 0, 4, 2), options([{ col: 2, descending: true }, { col: 1, descending: false }], { hasHeader: true }));
    expect(valuesOf(two, 0, 1, 4)).toEqual(["Bob", "Di", "Ann", "Cy"]);
    expect(valuesOf(two, 0, 0, 0)).toEqual(["Name"]);
  });

  it("applying the same keys twice changes nothing", () => {
    const m = people();
    const keys = options([{ col: 1, descending: false }, { col: 2, descending: true }], { hasHeader: true });
    sortRange(m, "s1", range(0, 0, 4, 2), keys);
    const once = structuredClone(m.sheets[0]!.rows);
    const again = sortRange(m, "s1", range(0, 0, 4, 2), keys);
    expect(m.sheets[0]!.rows).toEqual(once);
    expect(again.ok && again.permutation).toEqual([]);
  });

  it("returns the permutation and the changed keys, and runs the permutation adjusters with exactly that permutation", () => {
    const m = people();
    const seen: unknown[] = [];
    ROW_PERMUTATION_ADJUSTERS.push({ id: "test", permute: (_m, sheetId, sorted, permutation) => seen.push({ sheetId, sorted, permutation }) });
    const outcome = sortRange(m, "s1", range(0, 0, 4, 2), options([{ col: 0, descending: true }], { hasHeader: true }));
    expect(outcome.ok).toBe(true);
    if (!outcome.ok) {
      return;
    }
    expect(outcome.permutation).toEqual([
      { from: 4, to: 1 },
      { from: 3, to: 2 },
      { from: 2, to: 3 },
      { from: 1, to: 4 },
    ]);
    expect(seen).toEqual([{ sheetId: "s1", sorted: range(1, 0, 4, 2), permutation: outcome.permutation }]);
    expect(outcome.changed).toContain(cellKey("s1", 1, 0));
    expect(outcome.changed).toHaveLength(12);
  });
});

describe("sortRange — what moves", () => {
  it("permutes whole records among the visible rows only; row properties stay; cells outside the columns never move", () => {
    const m = model(
      sheet({
        rows: [
          row(0, [cell(0, 3, { styleId: 1, runs: undefined }), cell(1, "c"), cell(3, "x0")], { heightPt: 30 }),
          row(1, [cell(0, 9), cell(1, "h")], { hidden: true }),
          row(2, [cell(0, 1), cell(1, "a", { styleId: 1 })], { styleId: 1 }),
          row(3, [cell(0, 7), cell(1, "f")], { filterHidden: true }),
          row(4, [cell(0, 2), cell(1, "b"), cell(3, "x4")]),
        ],
      }),
    );
    const outcome = sortRange(m, "s1", range(0, 0, 4, 1), options([{ col: 0, descending: false }]));
    expect(outcome.ok).toBe(true);
    const rows = m.sheets[0]!.rows;
    expect(valuesOf(m, 0, 0, 4)).toEqual([1, 9, 2, 7, 3]);
    expect(valuesOf(m, 1, 0, 4)).toEqual(["a", "h", "b", "f", "c"]);
    expect(getCell(m.sheets[0]!, 0, 1)?.styleId).toBe(1); // "a" travelled with its style
    expect(getCell(m.sheets[0]!, 4, 0)?.styleId).toBe(1); // 3 travelled with its style
    expect(rows.find((r) => r.index === 0)?.heightPt).toBe(30);
    expect(rows.find((r) => r.index === 2)?.styleId).toBe(1);
    expect(rows.find((r) => r.index === 1)?.hidden).toBe(true);
    expect(rows.find((r) => r.index === 3)?.filterHidden).toBe(true);
    expect(valuesOf(m, 3, 0, 4)).toEqual(["x0", null, null, null, "x4"]);
  });

  it("rows whose cells did not change keep their identity", () => {
    const m = column([1, 2, 4, 3]);
    const before = m.sheets[0]!.rows;
    sortRange(m, "s1", range(0, 0, 3, 0), options([{ col: 0, descending: false }]));
    expect(m.sheets[0]!.rows[0]).toBe(before[0]);
    expect(m.sheets[0]!.rows[1]).toBe(before[1]);
    expect(m.sheets[0]!.rows[2]).not.toBe(before[2]);
  });

  it("a moved formula is shifted like a copy; an unparsed one moves verbatim; a single-row array moves with its anchor", () => {
    const m = model(
      sheet({
        rows: [
          row(0, [cell(0, 30), cell(1, 60, { formula: "A1*2" }), cell(2, 1, { formula: "$A$1+A2" })]),
          row(1, [cell(0, 10), cell(1, 20, { formula: "A2*2" }), cell(2, 5, { formula: "SUM(", formulaUnparsed: true })]),
          row(2, [cell(0, 20), cell(1, 40, { formula: "{1,2}", arrayRange: range(2, 1, 2, 2) }), cell(2, 2)]),
          row(5, [cell(0, 0, { formula: "B1" })]),
        ],
      }),
    );
    sortRange(m, "s1", range(0, 0, 2, 2), options([{ col: 0, descending: false }]));
    const s = m.sheets[0]!;
    expect(valuesOf(m, 0, 0, 2)).toEqual([10, 20, 30]);
    expect(getCell(s, 0, 1)?.formula).toBe("A1*2");
    expect(getCell(s, 0, 2)).toMatchObject({ formula: "SUM(", formulaUnparsed: true });
    expect(getCell(s, 1, 1)).toMatchObject({ formula: "{1,2}", arrayRange: range(1, 1, 1, 2) });
    expect(getCell(s, 2, 1)?.formula).toBe("A3*2");
    expect(getCell(s, 2, 2)?.formula).toBe("$A$1+A4"); // Excel's known surprise: the relative reference moved too
    expect(getCell(s, 5, 0)?.formula).toBe("B1"); // outside the range: not adjusted
  });

  it("a single-row merge wholly inside the range moves with its row", () => {
    const m = model(sheet({ rows: [row(0, [cell(0, 2), cell(1, "x")]), row(1, [cell(0, 1)])], merges: [range(0, 1, 0, 2)] }));
    expect(sortRange(m, "s1", range(0, 0, 1, 2), options([{ col: 0, descending: false }])).ok).toBe(true);
    expect(m.sheets[0]!.merges).toEqual([range(1, 1, 1, 2)]);
  });

  it("marks the sheet dirty", () => {
    const m = column([2, 1]);
    sortRange(m, "s1", range(0, 0, 1, 0), options([{ col: 0, descending: false }]));
    expect(m.sheets[0]!.dirty).toBe(true);
  });
});

/** A refused sort's message, asserting it left the sheet untouched. */
function refusal(m: SpreadsheetModel, target: GridRange | GridRange[], opts: SortOptions): string | undefined {
  const before = m.sheets[0];
  const outcome = sortRange(m, "s1", target, opts);
  expect(m.sheets[0]).toBe(before);
  return outcome.ok ? undefined : outcome.message;
}

describe("sortRange — refusals change nothing", () => {
  const key = [{ col: 0, descending: false }];

  it("a merge spanning more than one data row, or crossing the left/right edge", () => {
    expect(refusal(model(sheet({ rows: [row(0, [cell(0, 2)]), row(1, [cell(0, 1)])], merges: [range(0, 1, 1, 1)] })), range(0, 0, 1, 1), options(key))).toBe(
      SORT_REFUSAL_MESSAGES.merges,
    );
    expect(refusal(model(sheet({ rows: [row(0, [cell(0, 2)]), row(1, [cell(0, 1)])], merges: [range(0, 1, 0, 2)] })), range(0, 0, 1, 1), options(key))).toBe(
      SORT_REFUSAL_MESSAGES.merges,
    );
    expect(SORT_REFUSAL_MESSAGES.merges).toBe("Sorting isn't possible here — the range contains merged cells that span rows. Unmerge them first.");
  });

  it("part of a multi-row array formula", () => {
    const m = model(sheet({ rows: [row(0, [cell(0, 2), cell(1, 1, { formula: "{1;2}", arrayRange: range(0, 1, 1, 1) })]), row(1, [cell(0, 1), cell(1, 2)])] }));
    expect(refusal(m, range(0, 0, 1, 1), options(key))).toBe("Sorting would split an array formula — the range contains part of one.");
  });

  it("a multi-area selection", () => {
    expect(refusal(column([2, 1]), [range(0, 0, 1, 0), range(0, 2, 1, 2)], options(key))).toBe("Sorting works on one range at a time — select a single range.");
  });

  it("fewer than two data rows", () => {
    expect(refusal(column([2, 1]), range(0, 0, 1, 0), options(key, { hasHeader: true }))).toBe(
      "There's nothing to sort — select a cell inside a range of at least two rows.",
    );
  });

  it("a duplicated key column", () => {
    expect(refusal(column([2, 1]), range(0, 0, 1, 1), options([{ col: 1, descending: false }, { col: 1, descending: true }]))).toBe(duplicateKeyMessage(1));
    expect(duplicateKeyMessage(1)).toBe("Column B is used by more than one level — remove the duplicate.");
  });

  it("a protected sheet without sort=\"0\"; one with it sorts", () => {
    expect(refusal(model(sheet({ rows: [row(0, [cell(0, 2)]), row(1, [cell(0, 1)])], protection: { sheet: true } })), range(0, 0, 1, 0), options(key))).toBe(
      PROTECTED_CELL_MESSAGE,
    );
    const allowed = model(sheet({ rows: [row(0, [cell(0, 2)]), row(1, [cell(0, 1)])], protection: { sheet: true, allowSort: true } }));
    expect(sortRange(allowed, "s1", range(0, 0, 1, 0), options(key)).ok).toBe(true);
  });
});

describe("the sort record", () => {
  it("a sort outside an autofilter records Sheet.sortState", () => {
    const m = column(["b", "a", "c"]);
    sortRange(m, "s1", range(0, 0, 2, 0), options([{ col: 0, descending: true }], { caseSensitive: true }));
    expect(m.sheets[0]!.sortState).toEqual({ ref: range(0, 0, 2, 0), keys: [{ col: 0, descending: true }], caseSensitive: true });
  });

  it("a sort of an autofilter's data rows records autoFilter.sort and drops its source text", () => {
    const m = model(
      sheet({
        rows: [row(0, [cell(0, "H")]), row(1, [cell(0, 2)]), row(2, [cell(0, 1)])],
        autoFilter: { range: range(0, 0, 2, 0), columns: [], sourceXml: "<autoFilter ref=\"A1:A3\"/>" },
      }),
    );
    sortRange(m, "s1", range(0, 0, 2, 0), options([{ col: 0, descending: false }], { hasHeader: true }));
    expect(m.sheets[0]!.autoFilter).toEqual({ range: range(0, 0, 2, 0), columns: [], sort: { ref: range(1, 0, 2, 0), keys: [{ col: 0, descending: false }] } });
    expect(m.sheets[0]!.sortState).toBeUndefined();
  });
});

describe("guessHasHeader — the pinned table", () => {
  const styles = model(sheet()).styles;
  const guess = (rows: Row[], target: GridRange, extra: Partial<Sheet> = {}): boolean => guessHasHeader({ styles }, sheet({ rows, ...extra }), target);
  const bold = { styleId: 1 };

  it.each([
    ["fewer than two rows", [row(0, [cell(0, "Name")])], range(0, 0, 0, 0), false],
    ["a string above a number", [row(0, [cell(0, "Name"), cell(1, "Qty")]), row(1, [cell(0, "Ann"), cell(1, 3)])], range(0, 0, 1, 1), true],
    ["a string above a boolean", [row(0, [cell(0, "Done")]), row(1, [cell(0, true)])], range(0, 0, 1, 0), true],
    ["a string above an error", [row(0, [cell(0, "Ratio")]), row(1, [cell(0, err("#DIV/0!"))])], range(0, 0, 1, 0), true],
    ["strings above strings, nothing bold", [row(0, [cell(0, "a")]), row(1, [cell(0, "b")])], range(0, 0, 1, 0), false],
    ["a bold first row over a plain second", [row(0, [cell(0, "Name", bold)]), row(1, [cell(0, "Ann")])], range(0, 0, 1, 0), true],
    ["bold over bold", [row(0, [cell(0, "Name", bold)]), row(1, [cell(0, "Ann", bold)])], range(0, 0, 1, 0), false],
    ["a bold row through the row style", [row(0, [cell(0, "Name")], { styleId: 1 }), row(1, [cell(0, "Ann")])], range(0, 0, 1, 0), true],
    ["a number in the first row", [row(0, [cell(0, "Name"), cell(1, 2026)]), row(1, [cell(0, "Ann"), cell(1, 3)])], range(0, 0, 1, 1), false],
    ["an empty first row", [row(1, [cell(0, 1)])], range(0, 0, 1, 0), false],
  ])("%s", (_name, rows, target, expected) => {
    expect(guess(rows, target)).toBe(expected);
  });

  it("the range of an autofilter (its first row is its header)", () => {
    expect(guess([row(4, [cell(0, 1)]), row(5, [cell(0, 2)])], range(4, 0, 5, 0), { autoFilter: { range: range(4, 0, 9, 3), columns: [] } })).toBe(true);
  });
});
