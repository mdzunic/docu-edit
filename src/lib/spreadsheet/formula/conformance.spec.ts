// SPEC-120 "Conformance": every row of the spec's table, each labelled with
// its source per "Where expected values come from" — **Excel** (Microsoft's
// documented behaviour: "Calculation operators and precedence in Excel"'s
// conversion table, the function pages and their examples) or **pinned**
// (computed by hand from the rule the spec states, where Excel's behaviour
// is undocumented). No expected value here was produced by running this
// engine. Then the `.ods` settings variants (case sensitivity, plain and
// regular-expression criteria, partial-cell matching), and spec120-calc.xlsx,
// whose column G holds the self-contained rows with Excel's cached values:
// F9 over it must reproduce every one.
//
// SPEC-127 "Conformance tables": then every row of the category files under
// conformance/ (one per category; conformance/row.ts says where a value may
// come from), each run through the real parser, evaluator and a full
// recalculation of its own workbook — the formula at Z100 of Sheet1 (or over
// its legacy array range), the clock at 2026-10-08 09:30, RAND drawing 0.5 — and
// named by its file and formula when it fails. Numbers compare at the row's
// significant digits (15 by default). Then the completeness rules: a row's
// root function (the call at the formula's root) belongs to its file's
// category, and every registered function — SPEC-120's included — roots at
// least three rows, one of them an error case and, unless its Microsoft page
// carries no example table, one of them an Excel row.
import { describe, expect, it } from "vitest";
import { getCell, parseA1, parseRange } from "../address";
import { buildSpec120CalcXlsxBytes, SPEC120_CONFORMANCE } from "../fixtures/build";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import type { CalcSettings, CellValue } from "../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { readSpreadsheetBytes } from "../open";
import type { FormulaNode } from "./ast";
import { DATE_ROWS } from "./conformance/date";
import { ENGINEERING_ROWS } from "./conformance/engineering";
import { FINANCIAL_ROWS } from "./conformance/financial";
import { INFORMATION_ROWS } from "./conformance/information";
import { LOGICAL_ROWS } from "./conformance/logical";
import { LOOKUP_ROWS } from "./conformance/lookup";
import { MATH_ROWS } from "./conformance/math";
import type { ConformanceRow } from "./conformance/row";
import { STATISTICAL_ROWS } from "./conformance/statistical";
import { TEXT_ROWS } from "./conformance/text";
import { evaluateCell, isUncomputable, storedValueContext } from "./evaluate";
import { parseFormula } from "./parse";
import { recalculate } from "./recalc";
import { getFunction, registeredFunctionNames, type FunctionCategory } from "./registry";

type Source = "Excel" | "pinned";

interface Row {
  formula: string;
  expected: CellValue;
  source: Source;
  cells?: Record<string, CellSpec>;
  host?: string;
  locale?: SpreadsheetLocaleTag;
  label?: string;
}

const VALUE = { error: "#VALUE!" } as const;

const ROWS: Row[] = [
  { formula: '"1"+"2"', expected: 3, source: "Excel" },
  { formula: '"1"+"$4.00"', expected: 5, source: "Excel", label: "en-US" },
  { formula: '"6/1/2001"-"5/1/2001"', expected: 31, source: "Excel", label: "en-US" },
  { formula: 'SQRT("8+1")', expected: VALUE, source: "Excel" },
  { formula: '"A"&TRUE', expected: "ATRUE", source: "Excel" },
  { formula: '1+""', expected: VALUE, source: "Excel" },
  { formula: '"TRUE"+0', expected: VALUE, source: "Excel" },
  { formula: '"a"*1', expected: VALUE, source: "Excel" },
  { formula: "TRUE+TRUE", expected: 2, source: "Excel" },
  { formula: "A1+1", expected: 6, source: "Excel", cells: { A1: "5" }, label: 'A1 the text "5"' },
  { formula: "A1", expected: 0, source: "Excel", label: "A1 empty" },
  { formula: "A1=0", expected: true, source: "Excel", label: "A1 empty" },
  { formula: 'A1=""', expected: true, source: "Excel", label: "A1 empty" },
  { formula: "A1=FALSE", expected: true, source: "Excel", label: "A1 empty" },
  { formula: '1="1"', expected: false, source: "Excel" },
  { formula: '1<"1"', expected: true, source: "Excel" },
  { formula: '"z"<TRUE', expected: true, source: "Excel" },
  { formula: "FALSE<TRUE", expected: true, source: "Excel" },
  { formula: '"A"="a"', expected: true, source: "Excel" },
  { formula: '"a"<"B"', expected: true, source: "Excel" },
  { formula: "0.1+0.2=0.3", expected: true, source: "Excel" },
  { formula: "0.1+0.2-0.3", expected: 0, source: "pinned" },
  { formula: "-2^2", expected: 4, source: "Excel" },
  { formula: "0-2^2", expected: -4, source: "Excel" },
  { formula: "2^3^2", expected: 64, source: "Excel" },
  { formula: "0^0", expected: { error: "#NUM!" }, source: "Excel" },
  { formula: '1/3&""', expected: "0.333333333333333", source: "Excel" },
  { formula: "#N/A+1/0", expected: { error: "#N/A" }, source: "pinned" },
  { formula: "1/0+#N/A", expected: { error: "#DIV/0!" }, source: "pinned" },
  { formula: "IF(TRUE,1,1/0)", expected: 1, source: "Excel" },
  { formula: 'SUM("3",TRUE)', expected: 4, source: "Excel" },
  { formula: "SUM(A1:A2)", expected: 0, source: "Excel", cells: { A1: "3", A2: true }, label: 'A1 "3" text, A2 TRUE' },
  { formula: 'MAX("a")', expected: VALUE, source: "Excel" },
  { formula: "MIN(A1:A3)", expected: 0, source: "Excel", label: "all empty" },
  { formula: "AVERAGE(A1:A3)", expected: { error: "#DIV/0!" }, source: "Excel", label: "all empty" },
  { formula: "AND(A1:A2)", expected: true, source: "pinned", cells: { A1: "x", A2: true }, label: 'A1 "x" text, A2 TRUE' },
  { formula: 'AND("x")', expected: VALUE, source: "pinned" },
  { formula: "ISBLANK(A1)", expected: false, source: "Excel", cells: { A1: { formula: '""', value: "" } }, label: 'A1 =""' },
  { formula: "COUNTBLANK(A1:A2)", expected: 2, source: "Excel", cells: { A1: { formula: '""', value: "" } }, label: 'A1 ="", A2 empty' },
  { formula: "COUNTA(A1:A2)", expected: 1, source: "Excel", cells: { A1: { formula: '""', value: "" } }, label: 'A1 ="", A2 empty' },
  { formula: '"1,5"+0', expected: VALUE, source: "pinned", locale: "en-US" },
  { formula: '"1,5"+0', expected: 1.5, source: "pinned", locale: "de-DE" },
  { formula: '"2026-10-08"+0', expected: 46303, source: "pinned" },
  { formula: '"12:00"+0', expected: 0.5, source: "pinned" },
  { formula: '"(5)"+0', expected: -5, source: "pinned" },
  { formula: '"10%"+0', expected: 0.1, source: "pinned" },
  { formula: "ROUND(2.5,0)", expected: 3, source: "Excel" },
  { formula: "ROUND(-2.5,0)", expected: -3, source: "Excel" },
  { formula: "ROUND(1.005,2)", expected: 1.01, source: "pinned" },
  { formula: "MOD(-3,2)", expected: 1, source: "Excel" },
  { formula: "MOD(3,-2)", expected: -1, source: "Excel" },
  { formula: "INT(-8.9)", expected: -9, source: "Excel" },
  { formula: "TRUNC(-8.9)", expected: -8, source: "Excel" },
  { formula: "CEILING(-2.5,2)", expected: -2, source: "Excel" },
  { formula: "FLOOR(-2.5,-2)", expected: -2, source: "Excel" },
  { formula: "MROUND(10,3)", expected: 9, source: "Excel" },
  { formula: "FLOOR(0.3,0.1)", expected: 0.3, source: "pinned" },
  { formula: "DATE(108,1,2)", expected: 39449, source: "Excel", label: "2008-01-02" },
  { formula: "DATE(2008,14,2)", expected: 39846, source: "Excel", label: "2009-02-02" },
  { formula: "DATE(1900,2,29)", expected: 60, source: "Excel" },
  { formula: "WEEKDAY(1)", expected: 1, source: "Excel" },
  { formula: "WEEKDAY(46303)", expected: 5, source: "pinned" },
  { formula: 'VLOOKUP("a",{"c",1;"a",2},2,)', expected: 2, source: "Excel" },
  { formula: 'MATCH("B*",{"apple","banana"},0)', expected: 2, source: "Excel" },
  { formula: "COUNTIF(A1:A3,1)", expected: 2, source: "pinned", cells: { A1: "1", A2: 1, A3: 2 }, label: 'A1 the text "1", A2 1, A3 2' },
  { formula: 'LEN("😀")', expected: 2, source: "pinned" },
  { formula: "PROPER(\"it's\")", expected: "It'S", source: "pinned" },
  { formula: 'TRIM("  a   b  ")', expected: "a b", source: "Excel" },
  { formula: 'TEXT(0.5,"0%")', expected: "50%", source: "Excel" },
  { formula: "A2:A4", expected: 30, source: "Excel", cells: { A2: 20, A3: 30, A4: 40 }, host: "B3", label: "in row 3, non-array: A3's value" },
  { formula: "A2:A4", expected: VALUE, source: "Excel", cells: { A2: 20, A3: 30, A4: 40 }, host: "B9", label: "in row 9" },
];

/** The value a formula at `host` (default Z100) of Sheet1 stores. */
function computed(formula: string, cells: Record<string, CellSpec> = {}, host = "Z100", locale: SpreadsheetLocaleTag = "en-US", calc?: CalcSettings, format: "xlsx" | "ods" = "xlsx") {
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: { formula } } }, { calc, format });
  const at = parseA1(host)!;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, storedValueContext(model, { locale: SPREADSHEET_LOCALES[locale] }));
}

describe("the conformance table", () => {
  for (const row of ROWS) {
    const where = [row.label, row.locale].filter(Boolean).join(", ");
    it(`=${row.formula}${where ? ` (${where})` : ""} is ${JSON.stringify(row.expected)} — ${row.source}`, () => {
      expect(computed(row.formula, row.cells, row.host, row.locale)).toEqual(row.expected);
    });
  }

  it("{=A1:A2*{10,100}} over B1:C3: the 2×2 products, row 3 #N/A — Excel", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: 2, B1: { formula: "A1:A2*{10,100}", array: "B1:C3" } } });
    recalculate(model, [], { full: true, reason: "manual" });
    const value = (a1: string) => getCell(model.sheets[0]!, parseA1(a1)!.row, parseA1(a1)!.col)?.value;
    expect([value("B1"), value("C1"), value("B2"), value("C2")]).toEqual([10, 100, 20, 200]);
    expect([value("B3"), value("C3")]).toEqual([{ error: "#N/A" }, { error: "#N/A" }]);
  });
});

describe("the .ods settings variants (SPEC-120 \"Calculation settings\")", () => {
  const cells: Record<string, CellSpec> = { A1: "Apple", A2: "apple", A3: "a.c", A4: "abc" };

  it("text compares case-sensitively when the workbook says so — pinned", () => {
    expect(computed('"A"="a"', {}, "Z100", "en-US", { caseSensitive: true }, "ods")).toBe(false);
    expect(computed('"A"<"a"', {}, "Z100", "en-US", { caseSensitive: true }, "ods")).toBe(false);
    expect(computed('COUNTIF(A1:A2,"apple")', cells, "Z100", "en-US", { caseSensitive: true }, "ods")).toBe(1);
    expect(computed('COUNTIF(A1:A2,"apple")', cells)).toBe(2);
  });

  it("plain criteria match literally — pinned", () => {
    expect(computed('COUNTIF(A1:A4,"a*")', cells, "Z100", "en-US", { criteria: "plain" }, "ods")).toBe(0);
    // With wildcards (xlsx), case-insensitively: every one of the four.
    expect(computed('COUNTIF(A1:A4,"a*")', cells)).toBe(4);
  });

  it("a regular-expression criterion matches literally without a special character, and is uncomputable with one — pinned", () => {
    expect(computed('COUNTIF(A1:A4,"abc")', cells, "Z100", "en-US", { criteria: "regex" }, "ods")).toBe(1);
    const uncomputable = computed('COUNTIF(A1:A4,"a.c")', cells, "Z100", "en-US", { criteria: "regex" }, "ods");
    expect(isUncomputable(uncomputable) && uncomputable.uncomputable).toEqual({ kind: "regex" });
  });

  it("partial-cell matching makes = text criteria match substrings — pinned", () => {
    expect(computed('COUNTIF(A1:A4,"pp")', cells, "Z100", "en-US", { partialCellMatch: true, caseSensitive: true }, "ods")).toBe(2);
    expect(computed('COUNTIF(A1:A4,"pp")', cells)).toBe(0);
  });

  it("an .ods workbook's February 1900 has 28 days — pinned", () => {
    expect(computed("DATE(1900,2,29)", {}, "Z100", "en-US", undefined, "ods")).toBe(61);
  });
});

describe("spec120-calc.xlsx's conformance rows (cached with Excel's values)", () => {
  it("F9 reproduces every cached value in column G", async () => {
    const model = await readSpreadsheetBytes(await buildSpec120CalcXlsxBytes(), "xlsx", "spec120-calc");
    const calc = model.sheets[0]!;
    const cached = SPEC120_CONFORMANCE.map((row) => getCell(calc, parseA1(row.ref)!.row, parseA1(row.ref)!.col)?.value);
    expect(cached).toEqual(SPEC120_CONFORMANCE.map((row) => row.value));
    recalculate(model, [], { full: true, reason: "manual", now: () => new Date(2026, 9, 8, 9, 30) });
    const after = model.sheets[0]!;
    expect(SPEC120_CONFORMANCE.map((row) => getCell(after, parseA1(row.ref)!.row, parseA1(row.ref)!.col)?.value)).toEqual(
      SPEC120_CONFORMANCE.map((row) => row.value),
    );
  });
});

// --- SPEC-127: the category files ------------------------------------------------------

/** Each category file, and the categories whose functions its rows root in. */
const CATEGORY_FILES: readonly { file: string; rows: readonly ConformanceRow[]; categories: readonly FunctionCategory[] }[] = [
  { file: "lookup.ts", rows: LOOKUP_ROWS, categories: ["Lookup & Reference"] },
  { file: "text.ts", rows: TEXT_ROWS, categories: ["Text"] },
  { file: "date.ts", rows: DATE_ROWS, categories: ["Date & Time"] },
  { file: "math.ts", rows: MATH_ROWS, categories: ["Math & Trig"] },
  { file: "statistical.ts", rows: STATISTICAL_ROWS, categories: ["Statistical", "Compatibility"] },
  { file: "financial.ts", rows: FINANCIAL_ROWS, categories: ["Financial"] },
  { file: "information.ts", rows: INFORMATION_ROWS, categories: ["Information"] },
  { file: "engineering.ts", rows: ENGINEERING_ROWS, categories: ["Engineering"] },
  { file: "logical.ts", rows: LOGICAL_ROWS, categories: ["Logical"] },
];

/** Functions whose Microsoft reference page carries no example table, so no
 * Excel row can exist for them (each checked against its page). */
const NO_EXAMPLE_TABLE: ReadonlySet<string> = new Set<string>([
  // Their pages print no fixed result: the values "vary" (NOW and TODAY's
  // with the date the page is read on).
  "RAND",
  "RANDBETWEEN",
  "NOW",
  "TODAY",
  // Excel's `@` operator in file form: no function page.
  "SINGLE",
  // Examples printed only as screenshots or prose (no literal data).
  "DATE",
  "IF",
  "TRUE",
  "FALSE",
  "IFNA",
  "VLOOKUP",
  "XLOOKUP",
  "TRANSPOSE",
  "CELL",
  "INFO",
  "SUM",
  "SUMPRODUCT",
  "MMULT",
  "DOLLAR",
  // Every example row needs what a row cannot hold: an external link
  // (HYPERLINK), a named range or table or several sheets (SHEET, SHEETS), or
  // a root that is no call of the function (SINH's `2.868*SINH(…)`).
  "HYPERLINK",
  "SHEET",
  "SHEETS",
  "SINH",
]);

/** The conformance clock: 2026-10-08 09:30, local time. */
const CONFORMANCE_NOW = new Date(2026, 9, 8, 9, 30);

/** The conformance RAND: every draw is 0.5, so a row over RAND or
 * RANDBETWEEN is derived by hand from SPEC-120's rule. */
const CONFORMANCE_RANDOM = (): number => 0.5;

/** The value a row's formula computes, and its array members' values. */
function runRow(row: ConformanceRow): { value: CellValue; members?: CellValue[][] } {
  const range = row.array === undefined ? undefined : parseRange(row.array);
  if (row.array !== undefined && !range) {
    throw new Error(`Not a range: ${row.array}`);
  }
  const host = range ? { row: range.top, col: range.left } : parseA1("Z100")!;
  const hostRef = range ? row.array!.split(":")[0]! : "Z100";
  const cells: Record<string, CellSpec> = { ...row.cells, [hostRef]: range ? { formula: row.formula, array: row.array } : { formula: row.formula } };
  const model = calcWorkbook({ Sheet1: cells });
  recalculate(model, [], { full: true, reason: "manual", now: () => CONFORMANCE_NOW, random: CONFORMANCE_RANDOM });
  const sheet = model.sheets[0]!;
  const at = (r: number, c: number): CellValue => getCell(sheet, r, c)?.value ?? null;
  if (!range) {
    return { value: at(host.row, host.col) };
  }
  const members: CellValue[][] = [];
  for (let r = range.top; r <= range.bottom; r += 1) {
    const line: CellValue[] = [];
    for (let c = range.left; c <= range.right; c += 1) {
      line.push(at(r, c));
    }
    members.push(line);
  }
  return { value: at(host.row, host.col), members };
}

/** A value as the row compares it: a number at `digits` significant digits. */
function comparable(value: CellValue, digits: number): CellValue {
  if (typeof value !== "number") {
    return value;
  }
  return value === 0 ? 0 : Number(value.toPrecision(digits));
}

/** The function at a formula's root (through parentheses), if any. */
function rootFunction(formula: string): string | undefined {
  let node: FormulaNode = parseFormula(formula);
  while (node.kind === "paren") {
    node = node.inner;
  }
  return node.kind === "function" ? node.name.toUpperCase() : undefined;
}

const isErrorValue = (value: CellValue): boolean => typeof value === "object" && value !== null && "error" in value;

describe("SPEC-127's conformance tables (formula/conformance/*.ts)", () => {
  for (const { file, rows } of CATEGORY_FILES) {
    rows.forEach((row, index) => {
      const label = `${file} #${index + 1}: =${row.formula}${row.array ? ` over ${row.array}` : ""} is ${JSON.stringify(row.expected)} — ${row.source}`;
      it(label, () => {
        const digits = row.digits ?? 15;
        const { value, members } = runRow(row);
        expect(comparable(value, digits), `${label} (${row.note})`).toEqual(comparable(row.expected, digits));
        if (row.expectedArray) {
          expect(
            members?.map((line) => line.map((member) => comparable(member, digits))),
            `${label}: the array's members`,
          ).toEqual(row.expectedArray.map((line) => line.map((member) => comparable(member, digits))));
        }
      });
    });
  }
});

describe("SPEC-127's conformance completeness", () => {
  const all = CATEGORY_FILES.flatMap(({ file, rows, categories }) => rows.map((row) => ({ file, row, categories, root: rootFunction(row.formula) })));

  it("labels every row Excel, pinned or reference, with a note, and gives `expectedArray` exactly to array rows", () => {
    for (const { file, row } of all) {
      expect(["Excel", "pinned", "reference"], `${file}: =${row.formula}`).toContain(row.source);
      expect(row.note.trim().length, `${file}: =${row.formula} has a note`).toBeGreaterThan(0);
      expect(row.expectedArray === undefined, `${file}: =${row.formula}`).toBe(row.array === undefined);
    }
  });

  it("files every row under its root function's category", () => {
    for (const { file, row, categories, root } of all) {
      expect(root, `${file}: =${row.formula} roots in a function call`).toBeDefined();
      const spec = getFunction(root!);
      expect(spec, `${file}: =${row.formula} roots in a registered function`).toBeDefined();
      expect(categories, `${file}: =${row.formula} (${root} is ${spec?.category})`).toContain(spec!.category);
    }
  });

  it.each(registeredFunctionNames().filter((name) => name !== "ZZTESTFN"))("%s roots at least three rows, an error row among them, and an Excel row where its page has an example", (name) => {
    const rows = all.filter((entry) => entry.root === name).map((entry) => entry.row);
    expect(rows.length, `${name}'s rows`).toBeGreaterThanOrEqual(3);
    expect(rows.some((row) => isErrorValue(row.expected)), `${name} has an error row`).toBe(true);
    if (!NO_EXAMPLE_TABLE.has(name)) {
      expect(rows.some((row) => row.source === "Excel"), `${name} has an Excel row`).toBe(true);
    }
  });
});
