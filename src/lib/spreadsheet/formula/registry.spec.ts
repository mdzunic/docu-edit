// SPEC-120 "The registry" and "The core functions": registeredFunctionNames()
// is exactly the core table's names, each file (functions/index.ts's
// CORE_FUNCTION_FILES) registering the functions the table gives it; every
// FunctionSpec's metadata holds the invariants SPEC-120 states — a category
// from FunctionCategory, 0 ≤ minArgs ≤ maxArgs ≤ 255, argument kinds from
// ArgKind with one per entry of `args` (a variadic function's repeating group
// once in each, marked repeating), `repeat` within argKinds, a one-line
// summary ending in ".",
// snake_case argument names, required before optional; volatile and lazy are
// exactly the functions the spec names; argKindAt cycles through `repeat`;
// every name resolves through SPEC-119's functionNames.ts in both wire
// dialects; the evaluator — not each function — applies arity and argument
// kinds (#VALUE! for a wrong count or a "range" argument that is no
// reference, implicit intersection for a "value" argument outside array mode
// and lifting inside it, an "any" argument as evaluated, an "array" argument
// evaluated in array mode, value-kind errors before the function runs); and
// registerFunction refuses a duplicate while getFunction ignores case.
//
// The registry is module-global: the one test that registers a function
// (ZZTESTFN) runs last, and the name lists below are read before it.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../address";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import { evaluateCell, storedValueContext } from "./evaluate";
import { canonicalFromOds, canonicalFromXlsx, odsFromCanonical, xlsxFromCanonical } from "./functionNames";
import { CORE_FUNCTION_FILES } from "./functions";
import { argKindAt, getFunction, registerFunction, registeredFunctionNames, type ArgKind, type FunctionCategory, type FunctionSpec } from "./registry";

/** SPEC-120 "The core functions", transcribed. (The spec calls them "the 100
 * core functions"; its table lists 101 names — the table is what is
 * asserted.) */
const CORE_TABLE: Readonly<Record<string, string>> = {
  "math.ts":
    "SUM SUMIF SUMIFS SUMPRODUCT PRODUCT ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD ABS SIGN SQRT POWER EXP LN LOG LOG10 PI CEILING FLOOR MROUND RAND RANDBETWEEN",
  "aggregate.ts": "AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MIN MAX MINIFS MAXIFS",
  "logical.ts": "IF IFS IFERROR IFNA AND OR NOT XOR TRUE FALSE SWITCH CHOOSE",
  "information.ts": "ISBLANK ISNUMBER ISTEXT ISERROR ISERR ISNA ISLOGICAL NA",
  "text.ts": "CONCAT CONCATENATE TEXTJOIN LEN LEFT RIGHT MID UPPER LOWER PROPER TRIM SUBSTITUTE REPLACE FIND SEARCH TEXT VALUE REPT EXACT",
  "dateTime.ts": "DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND TODAY NOW WEEKDAY EDATE EOMONTH DAYS",
  "lookup.ts": "VLOOKUP HLOOKUP INDEX MATCH ROW COLUMN ROWS COLUMNS OFFSET INDIRECT SINGLE",
};

const TABLE_NAMES: readonly string[] = Object.values(CORE_TABLE).flatMap((names) => names.split(" "));

const CATEGORIES: readonly FunctionCategory[] = [
  "Financial",
  "Date & Time",
  "Math & Trig",
  "Statistical",
  "Lookup & Reference",
  "Text",
  "Logical",
  "Information",
  "Engineering",
  "Compatibility",
];

const ARG_KINDS: readonly ArgKind[] = ["value", "range", "any", "array"];

/** Read once, before the last describe registers ZZTESTFN. */
const REGISTERED: readonly string[] = registeredFunctionNames();

function specOf(name: string): FunctionSpec {
  const spec = getFunction(name);
  if (!spec) {
    throw new Error(`${name} is not registered`);
  }
  return spec;
}

/** Evaluates the formula cell at `address` of a one-sheet workbook. */
function evalAt(cells: Record<string, CellSpec>, address: string): ReturnType<typeof evaluateCell> {
  const model = calcWorkbook({ Sheet1: cells });
  const at = parseA1(address);
  if (!at) {
    throw new Error(`Not an A1 address: ${address}`);
  }
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, storedValueContext(model));
}

describe("the core library", () => {
  it("registers exactly the names of the core functions table, sorted", () => {
    expect(TABLE_NAMES).toHaveLength(101);
    expect(new Set(TABLE_NAMES).size).toBe(TABLE_NAMES.length);
    expect(REGISTERED).toEqual(TABLE_NAMES.toSorted());
  });

  it("registers each function from the file the table names", () => {
    expect(Object.keys(CORE_FUNCTION_FILES).toSorted()).toEqual(Object.keys(CORE_TABLE).toSorted());
    for (const [file, names] of Object.entries(CORE_TABLE)) {
      const specs = CORE_FUNCTION_FILES[file] ?? [];
      expect(specs.map((spec) => spec.name).toSorted(), file).toEqual(names.split(" ").toSorted());
      for (const spec of specs) {
        expect(getFunction(spec.name), `${file}: ${spec.name}`).toBe(spec);
      }
    }
  });

  it("marks exactly NOW, TODAY, RAND, RANDBETWEEN, OFFSET and INDIRECT volatile", () => {
    expect(REGISTERED.filter((name) => specOf(name).volatile === true)).toEqual(["INDIRECT", "NOW", "OFFSET", "RAND", "RANDBETWEEN", "TODAY"]);
    expect(REGISTERED.filter((name) => specOf(name).volatile !== undefined)).toHaveLength(6);
  });

  it("makes exactly IF, IFS, IFERROR, IFNA, SWITCH and CHOOSE lazy", () => {
    expect(REGISTERED.filter((name) => specOf(name).lazy === true)).toEqual(["CHOOSE", "IF", "IFERROR", "IFNA", "IFS", "SWITCH"]);
    expect(REGISTERED.filter((name) => specOf(name).lazy !== undefined)).toHaveLength(6);
  });
});

describe("getFunction", () => {
  it("looks a name up case-insensitively", () => {
    expect(getFunction("SUM")?.name).toBe("SUM");
    expect(getFunction("sum")).toBe(getFunction("SUM"));
    expect(getFunction("VLookup")).toBe(getFunction("VLOOKUP"));
    expect(getFunction("log10")).toBe(getFunction("LOG10"));
  });

  it("is undefined for a name nobody registered", () => {
    expect(getFunction("NOSUCHFUNCTION")).toBeUndefined();
    // SPEC-127's library and the dynamic-array backlog are not core.
    expect(getFunction("XLOOKUP")).toBeUndefined();
    expect(getFunction("FILTER")).toBeUndefined();
    expect(getFunction("CELL")).toBeUndefined();
  });
});

describe("every registered function's metadata", () => {
  it.each(REGISTERED)("%s: a canonical upper-case name and a FunctionCategory", (name) => {
    const spec = specOf(name);
    expect(spec.name).toBe(name);
    expect(name).toBe(name.toUpperCase());
    expect(CATEGORIES).toContain(spec.category);
    expect(typeof spec.impl).toBe("function");
  });

  it.each(REGISTERED)("%s: 0 ≤ minArgs ≤ maxArgs ≤ 255", (name) => {
    const { minArgs, maxArgs } = specOf(name);
    expect(Number.isInteger(minArgs)).toBe(true);
    expect(Number.isInteger(maxArgs)).toBe(true);
    expect(minArgs).toBeGreaterThanOrEqual(0);
    expect(minArgs).toBeLessThanOrEqual(maxArgs);
    expect(maxArgs).toBeLessThanOrEqual(255);
  });

  it.each(REGISTERED)("%s: argument kinds from ArgKind, one per entry of args", (name) => {
    const spec = specOf(name);
    for (const kind of spec.argKinds) {
      expect(ARG_KINDS).toContain(kind);
    }
    // One kind per entry of `args` — for a variadic function too: its
    // repeating group appears once in each (SPEC-120: SUMIFS ["range",
    // "range", "value"] with repeat { from: 1, count: 2 }), marked repeating.
    expect(spec.argKinds).toHaveLength(spec.args.length);
    if (spec.repeat === undefined) {
      // Every argument a call may have has a declared kind.
      expect(spec.maxArgs).toBeLessThanOrEqual(spec.argKinds.length);
      expect(spec.args.filter((arg) => arg.repeating === true)).toEqual([]);
    } else {
      const { from, count } = spec.repeat;
      expect(spec.args.map((arg) => arg.repeating === true)).toEqual(spec.args.map((_, i) => i >= from && i < from + count));
    }
  });

  it.each(REGISTERED.filter((name) => specOf(name).repeat !== undefined))("%s: `repeat` lies within argKinds", (name) => {
    const spec = specOf(name);
    const { from, count } = spec.repeat!;
    expect(Number.isInteger(from)).toBe(true);
    expect(Number.isInteger(count)).toBe(true);
    expect(from).toBeGreaterThanOrEqual(0);
    expect(count).toBeGreaterThanOrEqual(1);
    expect(from + count).toBeLessThanOrEqual(spec.argKinds.length);
  });

  it.each(REGISTERED)("%s: a one-line summary ending in \".\"", (name) => {
    const { summary } = specOf(name);
    expect(summary).not.toMatch(/[\r\n]/);
    expect(summary.trim()).toBe(summary);
    expect(summary.length).toBeGreaterThan(1);
    expect(summary.endsWith(".")).toBe(true);
  });

  it.each(REGISTERED)("%s: snake_case argument names, required before optional", (name) => {
    const { args } = specOf(name);
    for (const arg of args) {
      expect(arg.name).toMatch(/^[a-z][a-z0-9_]*$/);
    }
    const firstOptional = args.findIndex((arg) => arg.optional === true);
    if (firstOptional >= 0) {
      expect(args.slice(firstOptional).filter((arg) => arg.optional !== true)).toEqual([]);
    }
  });
});

describe("argKindAt", () => {
  it("cycles an argument past argKinds through `repeat` (SPEC-120's SUMIFS shape)", () => {
    const sumifs = { argKinds: ["range", "range", "value"] as ArgKind[], repeat: { from: 1, count: 2 } };
    expect([0, 1, 2, 3, 4, 5, 6, 7].map((index) => argKindAt(sumifs, index))).toEqual([
      "range",
      "range",
      "value",
      "range",
      "value",
      "range",
      "value",
      "range",
    ]);
  });

  it("cycles a one-kind group", () => {
    const choose = { argKinds: ["value", "any"] as ArgKind[], repeat: { from: 1, count: 1 } };
    expect([0, 1, 2, 3, 9].map((index) => argKindAt(choose, index))).toEqual(["value", "any", "any", "any", "any"]);
  });

  it("gives the registered SUMIFS range, range, value, then range/value pairs", () => {
    const sumifs = specOf("SUMIFS");
    expect([0, 1, 2, 3, 4, 5, 6].map((index) => argKindAt(sumifs, index))).toEqual(["range", "range", "value", "range", "value", "range", "value"]);
  });
});

describe("the wire dialects", () => {
  it.each(REGISTERED)("%s resolves through functionNames.ts in xlsx and ods", (name) => {
    expect(canonicalFromXlsx(xlsxFromCanonical(name))).toBe(name);
    expect(canonicalFromOds(odsFromCanonical(name))).toBe(name);
  });
});

describe("the evaluator applies arity and argument kinds", () => {
  it.each(["ROUND(1)", "ROUND(1,2,3)", "ABS()", "ABS(1,2)", "PI(1)", "TODAY(1)", 'MID("abc",1)', "NOT(TRUE,FALSE)", "IFERROR(1)", "SUMIFS(A1:A2,A1:A2)"])(
    "=%s, a wrong argument count, is #VALUE!",
    (formula) => {
      expect(evalAt({ A1: 1, A2: 2, C1: `=${formula}` }, "C1")).toEqual({ error: "#VALUE!" });
    },
  );

  it.each(["COUNTIF(1,1)", "OFFSET(1,0,0)", "SUMIF(5,5)", "AVERAGEIF(1,1)", 'COUNTBLANK("a")', "COUNTIF({1,2},1)"])(
    "=%s gives a \"range\" argument no reference: #VALUE!",
    (formula) => {
      expect(evalAt({ A1: 1, A2: 2, C1: `=${formula}` }, "C1")).toEqual({ error: "#VALUE!" });
    },
  );

  it("passes a reference to a \"range\" argument", () => {
    expect(evalAt({ A1: 1, A2: 2, A3: 1, C1: "=COUNTIF(A1:A3,1)" }, "C1")).toBe(2);
  });

  const column: Record<string, CellSpec> = { A1: -1, A2: -5, A3: -9 };

  it("intersects a \"value\" argument's one-column range with the host's row outside array mode", () => {
    expect(evalAt({ ...column, B2: "=ABS(A1:A3)" }, "B2")).toBe(5);
    expect(evalAt({ ...column, B3: "=ABS(A1:A3)" }, "B3")).toBe(9);
    expect(evalAt({ ...column, B9: "=ABS(A1:A3)" }, "B9")).toEqual({ error: "#VALUE!" });
  });

  it("intersects a one-row range with the host's column", () => {
    const row: Record<string, CellSpec> = { A5: -1, B5: -2, C5: -3 };
    expect(evalAt({ ...row, B7: "=ABS(A5:C5)" }, "B7")).toBe(2);
    expect(evalAt({ ...row, E7: "=ABS(A5:C5)" }, "E7")).toEqual({ error: "#VALUE!" });
  });

  it("reduces a 2-D range in a value position to #VALUE! and an array to its top-left element", () => {
    expect(evalAt({ ...column, B1: -2, B2: -6, B3: -10, D2: "=ABS(A1:B3)" }, "D2")).toEqual({ error: "#VALUE!" });
    expect(evalAt({ D2: "=ABS({-4,5})" }, "D2")).toBe(4);
  });

  it("lifts a \"value\" argument element-wise in array mode", () => {
    expect(evalAt({ ...column, B1: { formula: "=ABS(A1:A3)", array: "B1:B3" } }, "B1")).toEqual({ kind: "array", rows: 3, cols: 1, data: [1, 5, 9] });
  });

  it("hands an \"any\" argument its range un-intersected", () => {
    expect(evalAt({ ...column, B2: "=ROWS(A1:A3)" }, "B2")).toBe(3);
    expect(evalAt({ ...column, B2: "=SUM(A1:A3)" }, "B2")).toBe(-15);
  });

  it("evaluates an \"array\" argument's subtree in array mode, without CSE", () => {
    expect(evalAt({ A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30, D1: "=SUMPRODUCT((A1:A3>1)*B1:B3)" }, "D1")).toBe(50);
  });

  it("returns the first error among value arguments, left to right, before the function runs", () => {
    expect(evalAt({ C1: "=ROUND(1/0,NA())" }, "C1")).toEqual({ error: "#DIV/0!" });
    expect(evalAt({ C1: "=ROUND(NA(),1/0)" }, "C1")).toEqual({ error: "#N/A" });
  });

  it.each<[string, CellSpec]>([
    // Microsoft's IF reference: an omitted value_if_false is FALSE, an empty
    // one (a comma, then the parenthesis) 0; so is an empty value_if_true.
    ["IF(FALSE,1)", false],
    ["IF(FALSE,1,)", 0],
    ["IF(TRUE,)", 0],
    // The conformance table's exact-mode VLOOKUP through an empty slot.
    ['VLOOKUP("a",{"c",1;"a",2},2,)', 2],
    ['LEFT("abc")', "a"],
    ['LEFT("abc",)', ""],
    ["ROUND(2.5,)", 3],
  ])("reads an empty slot as EMPTY and gives an omitted trailing argument its default: =%s → %j", (formula, expected) => {
    expect(evalAt({ C1: `=${formula}` }, "C1")).toEqual(expected);
  });

  it.each<[string, number]>([
    ["IF(TRUE,1,RAND())", 0],
    ["IF(FALSE,1,RAND())", 1],
    ["IFS(TRUE,1,TRUE,RAND())", 0],
    ["IFERROR(1,RAND())", 0],
    ["IFNA(1,RAND())", 0],
    ["SWITCH(1,1,2,3,RAND())", 0],
    ["CHOOSE(1,2,RAND())", 0],
    ["CHOOSE(2,2,RAND())", 1],
    ["ROUND(1,RAND())", 1],
  ])("evaluates a lazy function's arguments only when it reads them: =%s draws %i random numbers", (formula, draws) => {
    const model = calcWorkbook({ Sheet1: { C1: `=${formula}` } });
    let drawn = 0;
    const base = {
      ...storedValueContext(model),
      random: (): number => {
        drawn += 1;
        return 0.25;
      },
    };
    expect(evaluateCell(model, "sheet-Sheet1", 0, 2, base)).not.toEqual({ error: "#VALUE!" });
    expect(drawn).toBe(draws);
  });
});

// Last: this registers ZZTESTFN in the module-global registry.
describe("registerFunction", () => {
  const spec: FunctionSpec = {
    name: "ZZTESTFN",
    category: "Information",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    impl: () => 42,
    summary: "Returns 42, for the registry spec.",
    args: [],
  };

  it("registers a new function, which the evaluator then computes, and throws on a duplicate name", () => {
    expect(getFunction("ZZTESTFN")).toBeUndefined();
    registerFunction(spec);
    expect(getFunction("zztestfn")).toBe(spec);
    expect(registeredFunctionNames()).toContain("ZZTESTFN");
    expect(evalAt({ A1: "=ZZTESTFN()" }, "A1")).toBe(42);
    expect(() => registerFunction(spec)).toThrow();
    expect(() => registerFunction({ ...spec, impl: () => 7 })).toThrow();
    expect(() => registerFunction({ ...specOf("SUM") })).toThrow();
    expect(getFunction("ZZTESTFN")).toBe(spec);
    expect(registeredFunctionNames().filter((name) => name === "ZZTESTFN")).toHaveLength(1);
  });
});
