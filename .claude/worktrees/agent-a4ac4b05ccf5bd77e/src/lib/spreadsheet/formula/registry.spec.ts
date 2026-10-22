// SPEC-120 "The registry" and "The core functions", SPEC-127 "Function
// metadata": registeredFunctionNames() is exactly the core table's names —
// each file (functions/index.ts's CORE_FUNCTION_FILES) registering the
// functions the table gives it — and SPEC-127's library, each in the
// category its acceptance criteria file it under; every FunctionSpec's
// metadata holds SPEC-127's invariants (a test per function, so a failure
// names it): a canonical name, a 10–100 character summary, snake_case
// argument names, unique, each described in 5–120 characters, required
// before optional with minArgs their count, a variadic function's repeating
// group the trailing run of optional numbered arguments — `repeat` the same
// group, maxArgs ≤ 255 — and otherwise maxArgs = args.length; argument kinds
// from ArgKind, one per entry of `args`; FUNCTION_CATEGORY_ORDER lists every
// FunctionCategory once; the legacy names are "Compatibility", sharing their
// modern function's implementation; volatile, lazy and visibilitySensitive
// are exactly the functions the specs name; no dynamic-array function is
// registered; argKindAt cycles through `repeat`; every name resolves through
// SPEC-119's functionNames.ts in both wire dialects; the evaluator — not
// each function — applies arity and argument kinds (#VALUE! for a wrong
// count or a "range" argument that is no reference, implicit intersection
// for a "value" argument outside array mode and lifting inside it, an "any"
// argument as evaluated, an "array" argument evaluated in array mode,
// value-kind errors before the function runs); and registerFunction refuses
// a duplicate while getFunction ignores case.
//
// The registry is module-global: the one test that registers a function
// (ZZTESTFN) runs last, and the name lists below are read before it.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../address";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import { evaluateCell, storedValueContext } from "./evaluate";
import { canonicalFromOds, canonicalFromXlsx, odsFromCanonical, xlsxFromCanonical } from "./functionNames";
import { CORE_FUNCTION_FILES } from "./functions";
import {
  argKindAt,
  FUNCTION_CATEGORY_ORDER,
  getFunction,
  registerFunction,
  registeredFunctionNames,
  type ArgKind,
  type FunctionCategory,
  type FunctionSpec,
} from "./registry";

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

/** SPEC-127's library by category, transcribed from its acceptance criteria
 * (the trigonometric set from "Function pins", the twelve base conversions
 * spelled out); the legacy names are "Compatibility", as Excel files them. */
const LIBRARY_TABLE: Readonly<Record<FunctionCategory, string>> = {
  Financial: "PMT IPMT PPMT FV PV NPER RATE NPV XNPV IRR XIRR MIRR SLN SYD DB DDB CUMIPMT CUMPRINC EFFECT NOMINAL",
  "Date & Time": "DATEVALUE TIMEVALUE WEEKNUM ISOWEEKNUM WORKDAY WORKDAY.INTL NETWORKDAYS NETWORKDAYS.INTL YEARFRAC DATEDIF DAYS360",
  "Math & Trig":
    "SUMSQ GCD LCM FACT COMBIN PERMUT QUOTIENT EVEN ODD SIN COS TAN ASIN ACOS ATAN ATAN2 SINH COSH TANH ASINH ACOSH ATANH COT CSC SEC ACOT " +
    "DEGREES RADIANS CEILING.MATH FLOOR.MATH CEILING.PRECISE FLOOR.PRECISE ISO.CEILING ROMAN ARABIC BASE DECIMAL MMULT SUBTOTAL AGGREGATE",
  Statistical:
    "MEDIAN MODE.SNGL STDEV.S STDEV.P VAR.S VAR.P LARGE SMALL RANK.EQ RANK.AVG PERCENTILE.INC PERCENTILE.EXC QUARTILE.INC QUARTILE.EXC " +
    "PERCENTRANK.INC PERCENTRANK.EXC CORREL PEARSON RSQ COVARIANCE.P COVARIANCE.S SLOPE INTERCEPT STEYX FORECAST.LINEAR NORM.DIST NORM.INV " +
    "NORM.S.DIST NORM.S.INV STANDARDIZE T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T BINOM.DIST POISSON.DIST EXPON.DIST CONFIDENCE.NORM AVERAGEA " +
    "MAXA MINA STDEVA STDEVPA VARA VARPA GEOMEAN HARMEAN AVEDEV DEVSQ KURT SKEW SKEW.P TRIMMEAN FREQUENCY GAMMA GAMMALN GAMMALN.PRECISE",
  "Lookup & Reference": "XLOOKUP XMATCH LOOKUP ADDRESS AREAS TRANSPOSE HYPERLINK FORMULATEXT",
  Text: "CHAR CODE UNICHAR UNICODE CLEAN FIXED DOLLAR NUMBERVALUE T TEXTBEFORE TEXTAFTER VALUETOTEXT ARRAYTOTEXT",
  Logical: "",
  Information: "ISEVEN ISODD ISREF ISFORMULA TYPE ERROR.TYPE N SHEET SHEETS CELL INFO",
  Engineering:
    "BIN2DEC BIN2HEX BIN2OCT DEC2BIN DEC2HEX DEC2OCT HEX2BIN HEX2DEC HEX2OCT OCT2BIN OCT2DEC OCT2HEX DELTA GESTEP ERF ERF.PRECISE ERFC ERFC.PRECISE",
  Compatibility:
    "STDEV VAR PERCENTILE QUARTILE RANK MODE NORMDIST NORMINV NORMSDIST NORMSINV TDIST TINV BINOMDIST POISSON EXPONDIST CONFIDENCE COVAR " +
    "FORECAST PERCENTRANK STDEVP VARP",
};

const LIBRARY_NAMES: readonly string[] = Object.values(LIBRARY_TABLE).flatMap((names) => (names === "" ? [] : names.split(" ")));

/** Each legacy name and the modern function whose implementation it shares
 * (SPEC-127: TDIST's `tails` picks between T.DIST.RT and T.DIST.2T, so it
 * shares their kernels rather than one `impl`). */
const LEGACY_MODERN: Readonly<Record<string, string>> = {
  STDEV: "STDEV.S",
  VAR: "VAR.S",
  PERCENTILE: "PERCENTILE.INC",
  QUARTILE: "QUARTILE.INC",
  RANK: "RANK.EQ",
  MODE: "MODE.SNGL",
  NORMDIST: "NORM.DIST",
  NORMINV: "NORM.INV",
  NORMSDIST: "NORM.S.DIST",
  NORMSINV: "NORM.S.INV",
  TINV: "T.INV.2T",
  BINOMDIST: "BINOM.DIST",
  POISSON: "POISSON.DIST",
  EXPONDIST: "EXPON.DIST",
  CONFIDENCE: "CONFIDENCE.NORM",
  COVAR: "COVARIANCE.P",
  FORECAST: "FORECAST.LINEAR",
  PERCENTRANK: "PERCENTRANK.INC",
  STDEVP: "STDEV.P",
  VARP: "VAR.P",
};

/** Every FunctionCategory — a Record, so the compiler holds it to the
 * union. */
const EVERY_CATEGORY: Readonly<Record<FunctionCategory, true>> = {
  Financial: true,
  "Date & Time": true,
  "Math & Trig": true,
  Statistical: true,
  "Lookup & Reference": true,
  Text: true,
  Logical: true,
  Information: true,
  Engineering: true,
  Compatibility: true,
};

const CATEGORIES: readonly FunctionCategory[] = Object.keys(EVERY_CATEGORY) as FunctionCategory[];

/** SPEC-127: the dynamic-array functions stay unregistered (they spill), as
 * does Google Sheets' COUNTUNIQUE. */
const NEVER_REGISTERED: readonly string[] = [
  "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "LET", "LAMBDA", "BYCOL", "BYROW", "MAKEARRAY", "MAP", "REDUCE", "SCAN",
  "ISOMITTED", "TEXTSPLIT", "VSTACK", "HSTACK", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "EXPAND",
  "COUNTUNIQUE",
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

describe("the core library and SPEC-127's", () => {
  it("registers exactly the core functions table and SPEC-127's library, sorted", () => {
    expect(TABLE_NAMES).toHaveLength(101);
    expect(LIBRARY_NAMES).toHaveLength(200);
    const all = [...TABLE_NAMES, ...LIBRARY_NAMES];
    expect(new Set(all).size).toBe(all.length);
    expect(REGISTERED).toEqual(all.toSorted());
  });

  it.each(Object.entries(LIBRARY_TABLE).flatMap(([category, names]) => (names === "" ? [] : names.split(" ").map((name) => [name, category]))))(
    "files %s under %s",
    (name, category) => {
      expect(specOf(name).category).toBe(category);
    },
  );

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

  it("marks exactly NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, CELL and INFO volatile", () => {
    expect(REGISTERED.filter((name) => specOf(name).volatile === true)).toEqual(["CELL", "INDIRECT", "INFO", "NOW", "OFFSET", "RAND", "RANDBETWEEN", "TODAY"]);
    expect(REGISTERED.filter((name) => specOf(name).volatile !== undefined)).toHaveLength(8);
  });

  it("marks exactly SUBTOTAL and AGGREGATE visibilitySensitive", () => {
    expect(REGISTERED.filter((name) => specOf(name).visibilitySensitive === true)).toEqual(["AGGREGATE", "SUBTOTAL"]);
    expect(REGISTERED.filter((name) => specOf(name).visibilitySensitive !== undefined)).toHaveLength(2);
  });

  it("registers no dynamic-array function, and not Google Sheets' COUNTUNIQUE", () => {
    for (const name of NEVER_REGISTERED) {
      expect(getFunction(name), name).toBeUndefined();
    }
  });

  it.each(Object.entries(LEGACY_MODERN))("files the legacy %s under Compatibility, sharing %s's implementation", (legacy, modern) => {
    expect(specOf(legacy).category).toBe("Compatibility");
    expect(specOf(modern).category).not.toBe("Compatibility");
    expect(specOf(legacy).impl, `${legacy} shares ${modern}'s impl`).toBe(specOf(modern).impl);
  });

  it("files TDIST under Compatibility too", () => {
    expect(specOf("TDIST").category).toBe("Compatibility");
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
    // The dynamic-array backlog stays unregistered; SPEC-127's library is in.
    expect(getFunction("FILTER")).toBeUndefined();
    expect(getFunction("SEQUENCE")).toBeUndefined();
    expect(getFunction("xlookup")).toBe(getFunction("XLOOKUP"));
    expect(getFunction("Cell")?.name).toBe("CELL");
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

  it.each(REGISTERED)("%s: a one-line summary of 10–100 characters, capitalised, ending in \".\"", (name) => {
    const { summary } = specOf(name);
    expect(summary, name).not.toMatch(/[\r\n]/);
    expect(summary.trim(), name).toBe(summary);
    expect(summary.length, `${name}'s summary length`).toBeGreaterThanOrEqual(10);
    expect(summary.length, `${name}'s summary length`).toBeLessThanOrEqual(100);
    expect(summary, `${name}'s summary starts with a capital`).toMatch(/^[A-Z]/);
    expect(summary.endsWith("."), `${name}'s summary ends with "."`).toBe(true);
  });

  it.each(REGISTERED)("%s: a name matching SPEC-127's pattern", (name) => {
    expect(name).toMatch(/^[A-Z][A-Z0-9]*(\.[A-Z0-9]+)*$/);
  });

  it.each(REGISTERED)("%s: unique snake_case argument names, each described in 5–120 characters ending in \".\"", (name) => {
    const { args } = specOf(name);
    for (const arg of args) {
      expect(arg.name, `${name}: ${arg.name}`).toMatch(/^[a-z][a-z0-9_]*$/);
      expect(arg.description.length, `${name}: ${arg.name}'s description length`).toBeGreaterThanOrEqual(5);
      expect(arg.description.length, `${name}: ${arg.name}'s description length`).toBeLessThanOrEqual(120);
      expect(arg.description.endsWith("."), `${name}: ${arg.name}'s description ends with "."`).toBe(true);
    }
    expect(new Set(args.map((arg) => arg.name)).size, `${name}'s argument names are unique`).toBe(args.length);
  });

  it.each(REGISTERED)("%s: required arguments before optional ones, minArgs their count", (name) => {
    const { args, minArgs } = specOf(name);
    const firstOptional = args.findIndex((arg) => arg.optional === true);
    const required = firstOptional < 0 ? args.length : firstOptional;
    expect(args.slice(required).filter((arg) => arg.optional !== true), `${name}: a required argument after an optional one`).toEqual([]);
    expect(minArgs, `${name}: minArgs is the number of required arguments`).toBe(required);
  });

  it.each(REGISTERED)("%s: a repeating group is the trailing run of optional numbered arguments (maxArgs ≤ 255), else maxArgs = args.length", (name) => {
    const { args, maxArgs, repeat } = specOf(name);
    const first = args.findIndex((arg) => arg.repeating === true);
    if (first < 0) {
      expect(maxArgs, `${name}: maxArgs equals args.length`).toBe(args.length);
      expect(repeat, `${name}: no repeat without repeating args`).toBeUndefined();
      return;
    }
    const group = args.slice(first);
    expect(group.every((arg) => arg.repeating === true), `${name}: the repeating args are the trailing run`).toBe(true);
    expect(group.every((arg) => arg.optional === true), `${name}: the repeating args are optional`).toBe(true);
    for (const arg of group) {
      expect(arg.name, `${name}: a repeating name ends in a digit`).toMatch(/[0-9]$/);
    }
    expect(maxArgs, `${name}: maxArgs ≤ 255`).toBeLessThanOrEqual(255);
    expect(maxArgs, `${name}: maxArgs reaches the group`).toBeGreaterThanOrEqual(args.length);
    expect(repeat, `${name}: repeat names the same group`).toEqual({ from: first, count: group.length });
  });
});

describe("FUNCTION_CATEGORY_ORDER", () => {
  it("lists every FunctionCategory exactly once, in the dialog's order", () => {
    expect(FUNCTION_CATEGORY_ORDER).toEqual([
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
    ]);
    expect(new Set(FUNCTION_CATEGORY_ORDER).size).toBe(FUNCTION_CATEGORY_ORDER.length);
    expect(FUNCTION_CATEGORY_ORDER.toSorted()).toEqual(CATEGORIES.toSorted());
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
  it.each(["ROUND(1)", "ROUND(1,2,3)", "ABS()", "ABS(1,2)", "PI(1)", "TODAY(1)", 'MID("abc",1)', "NOT(TRUE,FALSE)", "IFERROR(1)", "SUMIFS(A1:A2,A1:A2)", "PMT(1,2)"])(
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
