import { describe, expect, it } from "vitest";
import { parseA1, parseRange, type GridRange } from "./address";
import { storedValueContext } from "./formula/evaluate";
import type { Cell, CellValue, Row, Sheet, SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import { dateToSerial } from "./serial";
import {
  clearValidation,
  DATA_VALIDATION_LABEL,
  defaultValidationSettings,
  formatListLiteral,
  formulaCandidateValue,
  isEmptyValidation,
  itemIndexOf,
  LITERAL_LIST_MAX_CHARS,
  listItems,
  listSourceOf,
  mapValidationFormulas,
  matchesItem,
  parseListLiteral,
  sameValidationSettings,
  setValidation,
  sortItemsAscending,
  subtractFromValidations,
  validateValue,
  validationAt,
  validationsInRanges,
  VALIDATION_LIST_MAX,
  withX14Placement,
  type DataValidation,
  type ValidationDeps,
  type ValidationOperator,
  type ValidationSettings,
  type ValidationType,
} from "./validation";

// --- builders ------------------------------------------------------------------------------------

function range(text: string): GridRange {
  return parseRange(text)!;
}

function sheetOf(id: string, name: string, values: Record<string, CellValue>, validations: DataValidation[] = []): Sheet {
  const rows = new Map<number, Cell[]>();
  for (const [address, value] of Object.entries(values)) {
    const ref = parseA1(address)!;
    rows.set(ref.row, [...(rows.get(ref.row) ?? []), { col: ref.col, value }]);
  }
  const sorted: Row[] = [...rows.entries()].toSorted(([a], [b]) => a - b).map(([index, cells]) => ({ index, cells: cells.toSorted((a, b) => a.col - b.col) }));
  return { id, name, visibility: "visible", kind: "grid", rows: sorted, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], conditionalFormats: [], validations };
}

function book(values: Record<string, CellValue>, validations: DataValidation[] = [], lists: Record<string, CellValue> = {}): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "book",
    title: "Book",
    format: "xlsx",
    readerVersion: 11,
    sheets: [sheetOf("s1", "Form", values, validations), sheetOf("s2", "Lists", lists)],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}, { numberFormat: "0.00" }],
    defaultStyleId: 0,
    names: [],
    dxfs: [],
  };
}

let ids = 0;
/** A rule over `ranges` (space-separated A1 ranges). */
function rule(ranges: string, fields: Partial<DataValidation> & { type: ValidationType }): DataValidation {
  ids += 1;
  return { id: `dv${ids}`, ranges: ranges.split(" ").map(range), ...defaultValidationSettings(), ...fields };
}

function deps(model: SpreadsheetModel): ValidationDeps {
  const locale = SPREADSHEET_LOCALES["en-US"];
  return { base: storedValueContext(model, { locale }), ctx: workbookFormatContext(model, locale) };
}

/** Whether `candidate` may be entered at `address` on the first sheet. */
function valid(model: SpreadsheetModel, address: string, candidate: CellValue): boolean {
  return validateValue(model, model.sheets[0]!, parseA1(address)!, candidate, deps(model)).ok;
}

function itemsAt(model: SpreadsheetModel, address: string) {
  const sheet = model.sheets[0]!;
  const at = parseA1(address)!;
  return listItems(model, sheet, validationAt(sheet, at)!, at, deps(model));
}

function bounded(type: ValidationType, operator: ValidationOperator, formula1: string, formula2?: string): SpreadsheetModel {
  return book({}, [rule("A1:A10", { type, operator, formula1, ...(formula2 === undefined ? {} : { formula2 }) })]);
}

// --- the model -----------------------------------------------------------------------------------

describe("the rule model", () => {
  it("defaults a new rule to Excel's dialog defaults", () => {
    expect(defaultValidationSettings()).toEqual({
      type: "any",
      operator: "between",
      allowBlank: true,
      showDropdown: true,
      input: { show: true, title: "", text: "" },
      error: { show: true, style: "stop", title: "", text: "" },
    });
    expect(DATA_VALIDATION_LABEL).toBe("Data validation");
    expect(VALIDATION_LIST_MAX).toBe(32_767);
    expect(LITERAL_LIST_MAX_CHARS).toBe(255);
  });

  it("finds the first rule in model order covering a cell", () => {
    const first = rule("B2:B5", { type: "whole", formula1: "1", formula2: "2" });
    const second = rule("B4:C9", { type: "list", formula1: '"a,b"' });
    const sheet = sheetOf("s", "S", {}, [first, second]);
    expect(validationAt(sheet, parseA1("B4")!)).toBe(first);
    expect(validationAt(sheet, parseA1("C4")!)).toBe(second);
    expect(validationAt(sheet, parseA1("A1")!)).toBeUndefined();
  });

  it("says an Any-value rule without an input message is no rule", () => {
    expect(isEmptyValidation(defaultValidationSettings())).toBe(true);
    expect(isEmptyValidation({ ...defaultValidationSettings(), input: { show: true, title: "Hint", text: "" } })).toBe(false);
    expect(isEmptyValidation({ ...defaultValidationSettings(), input: { show: false, title: "Hint", text: "x" } })).toBe(true);
    expect(isEmptyValidation({ ...defaultValidationSettings(), type: "whole" })).toBe(false);
  });
});

// --- validation semantics -------------------------------------------------------------------------

describe("validateValue — numbers", () => {
  it("accepts any value under Any value, and every cell no rule covers", () => {
    const model = book({}, [rule("A1", { type: "any", allowBlank: false })]);
    expect(valid(model, "A1", "text")).toBe(true);
    expect(valid(model, "A1", null)).toBe(true);
    expect(valid(model, "B1", "anything")).toBe(true);
  });

  it("whole: a number without a fraction satisfying the operator", () => {
    const model = bounded("whole", "between", "1", "10");
    expect(valid(model, "A1", 5)).toBe(true);
    expect(valid(model, "A1", 1)).toBe(true);
    expect(valid(model, "A1", 10)).toBe(true);
    expect(valid(model, "A1", 11)).toBe(false);
    expect(valid(model, "A1", 0)).toBe(false);
    expect(valid(model, "A1", 5.5)).toBe(false);
    expect(valid(model, "A1", "5")).toBe(false);
    expect(valid(model, "A1", true)).toBe(false);
    expect(valid(model, "A1", { error: "#N/A" })).toBe(false);
  });

  it("between and not between take their bounds in either order", () => {
    expect(valid(bounded("whole", "between", "10", "1"), "A1", 5)).toBe(true);
    expect(valid(bounded("whole", "between", "10", "1"), "A1", 12)).toBe(false);
    expect(valid(bounded("whole", "notBetween", "10", "1"), "A1", 5)).toBe(false);
    expect(valid(bounded("whole", "notBetween", "10", "1"), "A1", 12)).toBe(true);
    expect(valid(bounded("whole", "notBetween", "1", "10"), "A1", 0)).toBe(true);
    expect(valid(bounded("whole", "notBetween", "1", "10"), "A1", 10)).toBe(false);
  });

  const comparisons: [ValidationOperator, number, boolean][] = [
    ["equal", 5, true],
    ["equal", 6, false],
    ["notEqual", 5, false],
    ["notEqual", 6, true],
    ["greaterThan", 5, false],
    ["greaterThan", 6, true],
    ["lessThan", 5, false],
    ["lessThan", 4, true],
    ["greaterThanOrEqual", 5, true],
    ["greaterThanOrEqual", 4, false],
    ["lessThanOrEqual", 5, true],
    ["lessThanOrEqual", 6, false],
  ];
  for (const type of ["whole", "decimal", "date", "time"] as const) {
    for (const [operator, value, expected] of comparisons) {
      it(`${type} ${operator} 5: ${value} is ${expected ? "valid" : "invalid"}`, () => {
        expect(valid(bounded(type, operator, "5"), "A1", value)).toBe(expected);
      });
    }
  }

  it("decimal: any number satisfying the operator; text never passes", () => {
    const model = bounded("decimal", "greaterThan", "0");
    expect(valid(model, "A1", 0.5)).toBe(true);
    expect(valid(model, "A1", -0.5)).toBe(false);
    expect(valid(model, "A1", "0.5")).toBe(false);
  });

  it("date and time: serials compared as numbers", () => {
    const start = dateToSerial(2026, 1, 1, "1900");
    const end = dateToSerial(2026, 12, 31, "1900");
    const dates = bounded("date", "between", String(start), String(end));
    expect(valid(dates, "A1", dateToSerial(2026, 6, 15, "1900"))).toBe(true);
    expect(valid(dates, "A1", dateToSerial(2027, 1, 1, "1900"))).toBe(false);
    expect(valid(dates, "A1", "2026-06-15")).toBe(false);
    const times = bounded("time", "between", "0.25", "0.75");
    expect(valid(times, "A1", 0.5)).toBe(true);
    expect(valid(times, "A1", 0.8)).toBe(false);
  });

  it("evaluates formula bounds at the target", () => {
    const model = book({ C1: 3, C2: 7 }, [rule("A1:A2", { type: "whole", operator: "lessThan", formula1: "C1" })]);
    // Relative to the origin A1: A1 compares with C1, A2 with C2.
    expect(valid(model, "A1", 5)).toBe(false);
    expect(valid(model, "A2", 5)).toBe(true);
    const absolute = book({ C1: 3 }, [rule("A1:A2", { type: "whole", operator: "lessThan", formula1: "$C$1" })]);
    expect(valid(absolute, "A2", 5)).toBe(false);
  });

  it("reads a blank bound as 0 and accepts the entry when a bound is an error, text or unevaluable", () => {
    const blank = book({}, [rule("A1", { type: "decimal", operator: "greaterThan", formula1: "$Z$1" })]);
    expect(valid(blank, "A1", 1)).toBe(true);
    expect(valid(blank, "A1", -1)).toBe(false);
    const text = book({ Z1: "ten" }, [rule("A1", { type: "whole", operator: "lessThan", formula1: "$Z$1" })]);
    expect(valid(text, "A1", 99)).toBe(true);
    const error = book({}, [rule("A1", { type: "whole", operator: "lessThan", formula1: "1/0" })]);
    expect(valid(error, "A1", 99)).toBe(true);
    const unknownName = book({}, [rule("A1", { type: "whole", operator: "lessThan", formula1: "Limit" })]);
    expect(valid(unknownName, "A1", 99)).toBe(true);
    const missingSecond = book({}, [rule("A1", { type: "whole", operator: "between", formula1: "1" })]);
    expect(valid(missingSecond, "A1", 99)).toBe(true);
    const syntax = book({}, [rule("A1", { type: "whole", operator: "lessThan", formula1: "1+" })]);
    expect(valid(syntax, "A1", 99)).toBe(true);
  });
});

describe("validateValue — text length", () => {
  it("measures the value's text in UTF-16 code units, as LEN does", () => {
    const model = bounded("textLength", "lessThanOrEqual", "5");
    expect(valid(model, "A1", "abcde")).toBe(true);
    expect(valid(model, "A1", "abcdef")).toBe(false);
    expect(valid(model, "A1", 12345)).toBe(true);
    expect(valid(model, "A1", 123456)).toBe(false);
    expect(valid(model, "A1", 1.5)).toBe(true);
    expect(valid(model, "A1", "😀😀")).toBe(true); // 4 code units
    expect(valid(model, "A1", "😀😀😀")).toBe(false); // 6 code units
    expect(valid(model, "A1", true)).toBe(true); // "TRUE"
    expect(valid(model, "A1", false)).toBe(true); // "FALSE": 5
  });

  it("supports every operator", () => {
    expect(valid(bounded("textLength", "between", "2", "4"), "A1", "abc")).toBe(true);
    expect(valid(bounded("textLength", "notBetween", "2", "4"), "A1", "abc")).toBe(false);
    expect(valid(bounded("textLength", "equal", "3"), "A1", "abc")).toBe(true);
    expect(valid(bounded("textLength", "notEqual", "3"), "A1", "abc")).toBe(false);
    expect(valid(bounded("textLength", "greaterThan", "3"), "A1", "abcd")).toBe(true);
    expect(valid(bounded("textLength", "lessThan", "3"), "A1", "ab")).toBe(true);
    expect(valid(bounded("textLength", "greaterThanOrEqual", "3"), "A1", "ab")).toBe(false);
  });
});

describe("validateValue — blanks", () => {
  for (const type of ["whole", "decimal", "date", "time", "textLength", "list", "custom"] as const) {
    it(`${type}: a blank is valid iff allowBlank`, () => {
      const fields = { type, operator: "greaterThan" as const, formula1: type === "list" ? '"a,b"' : type === "custom" ? "FALSE" : "5" };
      const allowing = book({}, [rule("A1", { ...fields, allowBlank: true })]);
      const refusing = book({}, [rule("A1", { ...fields, allowBlank: false })]);
      expect(valid(allowing, "A1", null)).toBe(true);
      expect(valid(allowing, "A1", "")).toBe(true);
      expect(valid(refusing, "A1", null)).toBe(false);
      expect(valid(refusing, "A1", "")).toBe(false);
    });
  }
});

function sourceOf(formula1: string) {
  return listSourceOf({ type: "list", formula1 });
}

describe("list sources", () => {
  it("parses and formats a literal list", () => {
    expect(parseListLiteral('"Low,Medium,High"')).toEqual(["Low", "Medium", "High"]);
    expect(parseListLiteral('" a , b ,,c "')).toEqual(["a", "b", "", "c"]);
    expect(parseListLiteral('"say ""hi"",bye"')).toEqual(['say "hi"', "bye"]);
    expect(parseListLiteral("$E$2:$E$6")).toBeNull();
    expect(parseListLiteral('"a"&"b"')).toBeNull();
    expect(formatListLiteral(["Yes", "No"])).toBe('"Yes,No"');
    expect(formatListLiteral(['say "hi"', "bye"])).toBe('"say ""hi"",bye"');
    expect(parseListLiteral(formatListLiteral(['say "hi"', "bye"]))).toEqual(['say "hi"', "bye"]);
  });

  it("names a rule's source kind", () => {
    expect(sourceOf('"a,b"')).toEqual({ kind: "literal", items: ["a", "b"] });
    expect(sourceOf("$E$2:$E$6")).toEqual({ kind: "reference", formula: "$E$2:$E$6" });
    expect(sourceOf("Lists!$A$1:$A$4")).toEqual({ kind: "reference", formula: "Lists!$A$1:$A$4" });
    expect(sourceOf("Colours")).toEqual({ kind: "name", formula: "Colours" });
    expect(sourceOf("OFFSET($E$2,0,0,3,1)")).toEqual({ kind: "formula", formula: "OFFSET($E$2,0,0,3,1)" });
    expect(listSourceOf({ type: "whole", formula1: '"a"' })).toBeNull();
    expect(listSourceOf({ type: "list" })).toBeNull();
  });

  it("a literal list: items in order, each item's value what typing it enters", () => {
    const model = book({}, [rule("B2", { type: "list", formula1: '"Low, Medium ,High,10,,Low"' })]);
    expect(itemsAt(model, "B2")).toEqual({
      items: [
        { value: "Low", text: "Low" },
        { value: "Medium", text: "Medium" },
        { value: "High", text: "High" },
        { value: 10, text: "10" },
        { value: null, text: "" },
        { value: "Low", text: "Low" },
      ],
    });
    expect(valid(model, "B2", "medium")).toBe(true);
    expect(valid(model, "B2", "HIGH")).toBe(true);
    expect(valid(model, "B2", 10)).toBe(true);
    expect(valid(model, "B2", "10")).toBe(true);
    expect(valid(model, "B2", "Urgent")).toBe(false);
    expect(valid(model, "B2", 11)).toBe(false);
  });

  it("a range on the same sheet: stored values, text through each cell's format, blanks as empty items, duplicates kept", () => {
    const model = book({ E2: "Red", E3: 5, E5: "Red", E6: 2.5 }, [rule("B8", { type: "list", formula1: "$E$2:$E$6" })]);
    const sheet = model.sheets[0]!;
    // E6 formatted 0.00
    sheet.rows = sheet.rows.map((row) => (row.index === 5 ? { ...row, cells: row.cells.map((cell) => ({ ...cell, styleId: 1 })) } : row));
    expect(itemsAt(model, "B8").items).toEqual([
      { value: "Red", text: "Red" },
      { value: 5, text: "5" },
      { value: null, text: "" },
      { value: "Red", text: "Red" },
      { value: 2.5, text: "2.50" },
    ]);
    expect(valid(model, "B8", 5)).toBe(true);
    expect(valid(model, "B8", 5.000000000000001)).toBe(true); // 15 significant digits
    expect(valid(model, "B8", "red")).toBe(true);
    expect(valid(model, "B8", "2.50")).toBe(true);
    expect(valid(model, "B8", "Blue")).toBe(false);
  });

  it("a range on another sheet", () => {
    const model = book({}, [rule("B9", { type: "list", formula1: "Lists!$A$1:$A$4" })], { A1: "North", A2: "South", A3: "East", A4: "West" });
    expect(itemsAt(model, "B9").items.map((item) => item.text)).toEqual(["North", "South", "East", "West"]);
    expect(valid(model, "B9", "east")).toBe(true);
    expect(valid(model, "B9", "Up")).toBe(false);
  });

  it("a row range lists across", () => {
    const model = book({ E1: "a", F1: "b", G1: "c" }, [rule("A5", { type: "list", formula1: "$E$1:$G$1" })]);
    expect(itemsAt(model, "A5").items.map((item) => item.text)).toEqual(["a", "b", "c"]);
  });

  it("a name the source file does not define (a workbook with no source here) is unavailable — and accepts the entry", () => {
    const model = book({}, [rule("B10", { type: "list", formula1: "Colours" })]);
    expect(itemsAt(model, "B10")).toEqual({ items: [], unavailable: true });
    expect(valid(model, "B10", "anything")).toBe(true);
  });

  it("a formula whose result is one row or one column", () => {
    const model = book({ E2: "x", E3: "y", E4: "z", E5: "w" }, [
      rule("A1", { type: "list", formula1: "OFFSET($E$2,0,0,3,1)" }),
      rule("A2", { type: "list", formula1: '{"p","q"}' }),
      rule("A3", { type: "list", formula1: '"solo"&""' }),
    ]);
    expect(itemsAt(model, "A1").items.map((item) => item.text)).toEqual(["x", "y", "z"]);
    expect(valid(model, "A1", "w")).toBe(false);
    expect(itemsAt(model, "A2").items).toEqual([
      { value: "p", text: "p" },
      { value: "q", text: "q" },
    ]);
    expect(itemsAt(model, "A3").items).toEqual([{ value: "solo", text: "solo" }]);
  });

  it("a 2-D result, an error or an unparsable source is unavailable and accepts the entry", () => {
    const model = book({ E2: 1, F3: 2 }, [
      rule("A1", { type: "list", formula1: "$E$2:$F$3" }),
      rule("A2", { type: "list", formula1: "{1,2;3,4}" }),
      rule("A3", { type: "list", formula1: "1/0" }),
      rule("A4", { type: "list", formula1: "$E$2:" }),
    ]);
    for (const address of ["A1", "A2", "A3", "A4"]) {
      expect(itemsAt(model, address).unavailable).toBe(true);
      expect(valid(model, address, "anything")).toBe(true);
    }
  });

  it("caps a long source at VALIDATION_LIST_MAX items", () => {
    const model = book({ A40000: "last" }, [rule("C1", { type: "list", formula1: "$A$1:$A$40000" })]);
    const items = itemsAt(model, "C1").items;
    expect(items).toHaveLength(VALIDATION_LIST_MAX);
    expect(items.every((item) => item.value === null)).toBe(true);
  });

  it("an ODS sort-ascending list sorts numbers numerically before text by the locale's collation", () => {
    const model = book({}, [rule("A1", { type: "list", formula1: '"pear,10,Apple,2,banana"', odsListOrder: "sort-ascending" })]);
    expect(itemsAt(model, "A1").items.map((item) => item.text)).toEqual(["2", "10", "Apple", "banana", "pear"]);
    expect(sortItemsAscending([{ value: null, text: "" }, { value: "b", text: "b" }], "en-US").map((item) => item.text)).toEqual(["b", ""]);
  });

  it("matches items numerically when both are numbers, else by text case-insensitively", () => {
    expect(matchesItem(10, { value: 10, text: "10.00" })).toBe(true);
    expect(matchesItem("ABC", { value: "abc", text: "abc" })).toBe(true);
    expect(matchesItem(true, { value: true, text: "TRUE" })).toBe(true);
    expect(matchesItem(3, { value: "3", text: "3" })).toBe(true);
    expect(matchesItem(3, { value: 4, text: "3" })).toBe(false);
    const items = [
      { value: "Low", text: "Low" },
      { value: "Medium", text: "Medium" },
    ];
    expect(itemIndexOf(items, "medium")).toBe(1);
    expect(itemIndexOf(items, "Urgent")).toBe(-1);
    expect(itemIndexOf(items, null)).toBe(-1);
  });
});

describe("validateValue — custom", () => {
  it("evaluates formula1 at the target with the candidate overlaid on that cell", () => {
    const model = book({}, [rule("B7", { type: "custom", formula1: 'ISNUMBER(SEARCH("@",B7))' })]);
    expect(valid(model, "B7", "me@example.com")).toBe(true);
    expect(valid(model, "B7", "nobody")).toBe(false);
  });

  it("is relative to the rule's origin, and the overlay reaches ranges over the target", () => {
    const model = book({ Z1: 3, B2: 1, B3: 1 }, [rule("B2:B10", { type: "custom", formula1: "B2>$Z$1" }), rule("C1:C3", { type: "custom", formula1: "SUM($C$1:$C$3)<=10" })]);
    expect(valid(model, "B4", 5)).toBe(true);
    expect(valid(model, "B4", 2)).toBe(false);
    expect(valid(model, "C2", 10)).toBe(true);
    expect(valid(model, "C2", 11)).toBe(false);
  });

  it("is true for TRUE or a non-zero number, and accepts what it cannot evaluate", () => {
    expect(valid(book({}, [rule("A1", { type: "custom", formula1: "2" })]), "A1", "x")).toBe(true);
    expect(valid(book({}, [rule("A1", { type: "custom", formula1: "0" })]), "A1", "x")).toBe(false);
    expect(valid(book({}, [rule("A1", { type: "custom", formula1: '"TRUE"' })]), "A1", "x")).toBe(false);
    expect(valid(book({}, [rule("A1", { type: "custom", formula1: "Unknown>0" })]), "A1", "x")).toBe(true);
  });

  it("names the rule an invalid entry breaks", () => {
    const broken = rule("A1", { type: "whole", operator: "lessThan", formula1: "3" });
    const model = book({}, [broken]);
    expect(validateValue(model, model.sheets[0]!, parseA1("A1")!, 5, deps(model))).toEqual({ ok: false, rule: broken });
    expect(validateValue(model, model.sheets[0]!, parseA1("A1")!, 2, deps(model))).toEqual({ ok: true });
  });
});

describe("formula candidates", () => {
  it("evaluates a formula entry at the target", () => {
    const model = book({ C1: 4 });
    const sheet = model.sheets[0]!;
    expect(formulaCandidateValue(sheet, parseA1("A1")!, "C1*2", deps(model))).toBe(8);
    expect(formulaCandidateValue(sheet, parseA1("A1")!, '"a"&"b"', deps(model))).toBe("ab");
    expect(formulaCandidateValue(sheet, parseA1("A1")!, "Unknown", deps(model))).toBeUndefined();
  });
});

// --- commands ------------------------------------------------------------------------------------

const LIST: ValidationSettings = { ...defaultValidationSettings(), type: "list", formula1: '"a,b"' };
const WHOLE: ValidationSettings = { ...defaultValidationSettings(), type: "whole", formula1: "1", formula2: "10" };

describe("setValidation and clearValidation", () => {
  it("creates a rule over the target, copy-on-write, marking the sheet dirty", () => {
    const model = book({ A1: 1 });
    const before = model.sheets[0]!;
    const other = model.sheets[1]!;
    const id = setValidation(model, "s1", [range("B2:B5")], LIST);
    const after = model.sheets[0]!;
    expect(after).not.toBe(before);
    expect(after.rows).toBe(before.rows);
    expect(model.sheets[1]).toBe(other);
    expect(after.dirty).toBe(true);
    expect(after.validations).toEqual([{ id, ranges: [range("B2:B5")], ...LIST }]);
    expect(before.validations).toEqual([]);
  });

  it("subtracts the target from every other rule, re-anchoring one that lost its origin and removing one left empty", () => {
    const custom = { ...rule("B2:B5", { type: "custom", formula1: "B2>0" }), sourceXml: "<dataValidation/>" };
    const gone = rule("D1", { type: "whole", formula1: "1", formula2: "2" });
    const untouched = rule("F1:F3", { type: "list", formula1: '"x"' });
    const model = book({}, [custom, gone, untouched]);
    setValidation(model, "s1", [range("B2"), range("D1")], LIST);
    const rules = model.sheets[0]!.validations;
    expect(rules).toHaveLength(3);
    expect(rules[0]).toEqual({ id: custom.id, ranges: [range("B3:B5")], ...defaultValidationSettings(), type: "custom", formula1: "B3>0" });
    expect(rules[1]).toBe(untouched);
    expect(rules[2]!.ranges).toEqual([range("B2"), range("D1")]);
  });

  it("adds the target to an existing rule with identical settings — formulas compared relative to each origin", () => {
    const existing = { ...rule("B2:B5", { type: "custom", formula1: "B2>$Z$1" }), sourceXml: "<x/>", odsName: "val1" };
    const model = book({}, [existing]);
    const id = setValidation(model, "s1", [range("D2:D3")], { ...defaultValidationSettings(), type: "custom", formula1: "D2>$Z$1" });
    expect(id).toBe(existing.id);
    const merged = model.sheets[0]!.validations;
    expect(merged).toHaveLength(1);
    expect(merged[0]).toEqual({ id: existing.id, odsName: "val1", ranges: [range("B2:B5"), range("D2:D3")], ...defaultValidationSettings(), type: "custom", formula1: "B2>$Z$1" });
    // Same text, different meaning (it reads B2 from D2): a new rule.
    setValidation(model, "s1", [range("F2")], { ...defaultValidationSettings(), type: "custom", formula1: "B2>$Z$1" });
    expect(model.sheets[0]!.validations).toHaveLength(2);
  });

  it("ignores id, ranges, sourceXml and odsName when comparing — but not messages or flags", () => {
    const a = { ...rule("A1", { type: "list", formula1: '"a,b"' }), sourceXml: "<a/>", odsName: "n1" };
    const b = { ...rule("C3:C4", { type: "list", formula1: '"a,b"' }), sourceXml: "<b/>", odsName: "n2" };
    expect(sameValidationSettings(a, b)).toBe(true);
    expect(sameValidationSettings(a, { ...b, allowBlank: false })).toBe(false);
    expect(sameValidationSettings(a, { ...b, input: { ...b.input, title: "T" } })).toBe(false);
    expect(sameValidationSettings(a, { ...b, odsListOrder: "sort-ascending" })).toBe(false);
  });

  it("re-setting a rule's own cells keeps one rule", () => {
    const model = book({});
    setValidation(model, "s1", [range("B2:B5")], WHOLE);
    setValidation(model, "s1", [range("B3")], WHOLE);
    const rules = model.sheets[0]!.validations;
    expect(rules).toHaveLength(1);
    expect(rules[0]!.ranges).toEqual([range("B2"), range("B4:B5"), range("B3")]);
  });

  it("places a new cross-sheet rule in the x14 ext of an xlsx workbook", () => {
    const model = book({});
    setValidation(model, "s1", [range("B9")], { ...LIST, formula1: "Lists!$A$1:$A$4" });
    expect(model.sheets[0]!.validations[0]!.x14).toBe(true);
    setValidation(model, "s1", [range("B2")], LIST);
    expect(model.sheets[0]!.validations[0]!.x14).toBeUndefined();
    // A new base rule goes before the x14 one — the order an xlsx file lists them in.
    expect(model.sheets[0]!.validations.map((dv) => dv.formula1)).toEqual(['"a,b"', "Lists!$A$1:$A$4"]);
    expect(withX14Placement({ formula1: "Lists!A1", x14: true as const }, "ods", "Form")).toEqual({ formula1: "Lists!A1" });
  });

  it("clears the rules from ranges, reporting whether anything changed", () => {
    const model = book({}, [rule("B2:B5", { type: "whole", formula1: "1", formula2: "9" })]);
    const before = model.sheets[0]!;
    expect(clearValidation(model, "s1", [range("Z9")])).toBe(false);
    expect(model.sheets[0]).toBe(before);
    expect(clearValidation(model, "s1", [range("B2:B4")])).toBe(true);
    expect(model.sheets[0]!.validations.map((dv) => dv.ranges)).toEqual([[range("B5")]]);
    expect(model.sheets[0]!.dirty).toBe(true);
    expect(clearValidation(model, "s1", [range("A1:Z99")])).toBe(true);
    expect(model.sheets[0]!.validations).toEqual([]);
  });

  it("subtractFromValidations returns the same array when the cut misses every rule", () => {
    const rules = [rule("B2", { type: "any" })];
    expect(subtractFromValidations(rules, [range("C3")])).toBe(rules);
  });

  it("refuses an empty target", () => {
    expect(() => setValidation(book({}), "s1", [], LIST)).toThrow();
  });
});

describe("helpers for the dialog and the adjusters", () => {
  it("reports the rules a selection touches and whether part of it has none", () => {
    const a = rule("B2:B5", { type: "list", formula1: '"a"' });
    const b = rule("C2", { type: "whole", formula1: "1", formula2: "2" });
    const sheet = sheetOf("s", "S", {}, [a, b]);
    expect(validationsInRanges(sheet, [range("B3:B4")])).toEqual({ rules: [a], uncovered: false });
    expect(validationsInRanges(sheet, [range("B5:C5")])).toEqual({ rules: [a], uncovered: true });
    expect(validationsInRanges(sheet, [range("B2:C2")])).toEqual({ rules: [a, b], uncovered: false });
    expect(validationsInRanges(sheet, [range("Z1")])).toEqual({ rules: [], uncovered: true });
  });

  it("maps a rule's formulas copy-on-write, dropping its source when one changes", () => {
    const dv = { ...rule("A1", { type: "whole", formula1: "Lists!A1", formula2: "2" }), sourceXml: "<x/>" };
    expect(mapValidationFormulas(dv, (formula) => formula)).toBe(dv);
    const renamed = mapValidationFormulas(dv, (formula) => formula.replace("Lists", "Data"));
    expect(renamed).toEqual({ ...dv, formula1: "Data!A1", sourceXml: undefined });
    expect("sourceXml" in renamed).toBe(false);
  });
});
