// SPEC-130 "The dialog" — validationFields.ts: the labels each value field
// carries per type and operator, a rule's formulas shown relative to the
// active cell in the user's locale, and the fields read back — formulas
// canonical and moved onto the selection's origin, constants through
// parseInput as invariant number text, a list Source without "=" as a
// literal — with the dialog's five errors, verbatim; the mixed-selection
// test and the disjoint target.
import { describe, expect, it } from "vitest";
import type { GridRange } from "../../lib/spreadsheet/address";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { defaultValidationSettings, isEmptyValidation, sameValidationSettings, type DataValidation } from "../../lib/spreadsheet/validation";
import {
  BOUNDS_ORDER_ERROR,
  defaultValidationFields,
  disjointRanges,
  fieldsFromRule,
  FORMULA_ERROR,
  LIST_SOURCE_ERROR,
  LIST_TOO_LONG_ERROR,
  MIXED_SELECTION_NOTICE,
  sameValidationFields,
  selectionDisagrees,
  settingsFromFields,
  TEXT_LENGTH_ERROR,
  valueLabels,
  VALIDATION_ALLOW_OPTIONS,
  VALIDATION_OPERATOR_OPTIONS,
  type ValidationFields,
} from "./validationFields";

const US: FormatContext = { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" };
const DE: FormatContext = { locale: SPREADSHEET_LOCALES["de-DE"], dateSystem: "1900" };

const B2 = { row: 1, col: 1 };

function range(top: number, left: number, bottom = top, right = left): GridRange {
  return { top, left, bottom, right };
}

function rule(partial: Partial<DataValidation> & Pick<DataValidation, "type">): DataValidation {
  return { id: "r", ranges: [range(1, 1, 9, 1)], ...defaultValidationSettings(), ...partial };
}

function fields(partial: Partial<ValidationFields>): ValidationFields {
  return { ...defaultValidationFields(), ...partial };
}

/** The settings the fields give at B2 (typed and stored there), or the error. */
function read(partial: Partial<ValidationFields>, format: FormatContext = US) {
  return settingsFromFields(fields(partial), { active: B2, origin: B2, format });
}

describe("validationFields — choices and labels (SPEC-130)", () => {
  it("offers Excel's Allow types and operators", () => {
    expect(VALIDATION_ALLOW_OPTIONS.map((option) => option.label)).toEqual([
      "Any value",
      "Whole number",
      "Decimal",
      "List",
      "Date",
      "Time",
      "Text length",
      "Custom",
    ]);
    expect(VALIDATION_OPERATOR_OPTIONS.map((option) => option.label)).toEqual([
      "between",
      "not between",
      "equal to",
      "not equal to",
      "greater than",
      "less than",
      "greater than or equal to",
      "less than or equal to",
    ]);
  });

  it("labels the value fields per type and operator", () => {
    expect(valueLabels("any", "between")).toEqual({ value1: null, value2: null });
    expect(valueLabels("list", "between")).toEqual({ value1: "Source", value2: null });
    expect(valueLabels("custom", "between")).toEqual({ value1: "Formula", value2: null });
    expect(valueLabels("whole", "between")).toEqual({ value1: "Minimum", value2: "Maximum" });
    expect(valueLabels("decimal", "notBetween")).toEqual({ value1: "Minimum", value2: "Maximum" });
    expect(valueLabels("textLength", "between")).toEqual({ value1: "Minimum", value2: "Maximum" });
    expect(valueLabels("date", "between")).toEqual({ value1: "Start date", value2: "End date" });
    expect(valueLabels("time", "notBetween")).toEqual({ value1: "Start time", value2: "End time" });
    expect(valueLabels("whole", "equal").value1).toBe("Value");
    expect(valueLabels("decimal", "notEqual").value1).toBe("Value");
    expect(valueLabels("date", "equal").value1).toBe("Date");
    expect(valueLabels("time", "notEqual").value1).toBe("Time");
    expect(valueLabels("textLength", "equal").value1).toBe("Length");
    expect(valueLabels("whole", "greaterThan")).toEqual({ value1: "Minimum", value2: null });
    expect(valueLabels("decimal", "greaterThanOrEqual")).toEqual({ value1: "Minimum", value2: null });
    expect(valueLabels("whole", "lessThan")).toEqual({ value1: "Maximum", value2: null });
    expect(valueLabels("textLength", "lessThanOrEqual")).toEqual({ value1: "Maximum", value2: null });
  });

  it("pins the mixed-selection notice", () => {
    expect(MIXED_SELECTION_NOTICE).toBe("The selection contains cells with different validation settings — OK applies these settings to all of them");
  });
});

describe("validationFields — a rule as the dialog shows it (SPEC-130)", () => {
  it("starts a cell without a rule from Excel's defaults", () => {
    expect(fieldsFromRule(undefined, B2, US)).toEqual(defaultValidationFields());
    expect(defaultValidationFields()).toMatchObject({
      type: "any",
      operator: "between",
      value1: "",
      value2: "",
      allowBlank: true,
      showDropdown: true,
      inputShow: true,
      errorShow: true,
      errorStyle: "stop",
    });
  });

  it("shows a literal list as its comma-joined items, other sources with '=' relative to the active cell", () => {
    const literal = rule({ type: "list", formula1: '"Low, Medium,High"', showDropdown: false });
    expect(fieldsFromRule(literal, B2, US)).toMatchObject({ type: "list", value1: "Low,Medium,High", showDropdown: false });
    expect(fieldsFromRule(rule({ type: "list", formula1: "$E$2:$E$6" }), { row: 5, col: 1 }, US).value1).toBe("=$E$2:$E$6");
    // Relative to the rule's origin B2, typed at B4: two rows down.
    expect(fieldsFromRule(rule({ type: "list", formula1: "E2:E6" }), { row: 3, col: 1 }, US).value1).toBe("=E4:E8");
    expect(fieldsFromRule(rule({ type: "list", formula1: "Colours" }), B2, US).value1).toBe("=Colours");
  });

  it("shows a custom formula relative to the active cell", () => {
    const custom = rule({ type: "custom", formula1: 'ISNUMBER(SEARCH("@",B2))' });
    expect(fieldsFromRule(custom, { row: 6, col: 1 }, US).value1).toBe('=ISNUMBER(SEARCH("@",B7))');
  });

  it("shows numbers in the user's locale and date or time bounds as dates and times", () => {
    expect(fieldsFromRule(rule({ type: "whole", formula1: "1", formula2: "10" }), B2, US)).toMatchObject({ value1: "1", value2: "10" });
    expect(fieldsFromRule(rule({ type: "decimal", operator: "greaterThan", formula1: "1.5" }), B2, DE)).toMatchObject({ operator: "greaterThan", value1: "1,5" });
    expect(fieldsFromRule(rule({ type: "date", formula1: "45306", formula2: "45657" }), B2, US)).toMatchObject({ value1: "1/15/2024", value2: "12/31/2024" });
    expect(fieldsFromRule(rule({ type: "date", formula1: "45306", formula2: "45657" }), B2, DE)).toMatchObject({ value1: "15.01.2024", value2: "31.12.2024" });
    expect(fieldsFromRule(rule({ type: "time", formula1: "0.375", formula2: "0.75" }), B2, US)).toMatchObject({ value1: "9:00:00 AM", value2: "6:00:00 PM" });
    expect(fieldsFromRule(rule({ type: "date", operator: "lessThanOrEqual", formula1: "TODAY()" }), B2, US).value1).toBe("=TODAY()");
    expect(fieldsFromRule(rule({ type: "textLength", operator: "lessThanOrEqual", formula1: "$Z$1" }), B2, US).value1).toBe("=$Z$1");
  });

  it("carries the messages, and a list's LibreOffice sort order", () => {
    const loaded = fieldsFromRule(
      rule({
        type: "list",
        formula1: '"a,b"',
        odsListOrder: "sort-ascending",
        allowBlank: false,
        input: { show: true, title: "Priority", text: "Pick one" },
        error: { show: false, style: "warning", title: "Oops", text: "Pick from the list" },
      }),
      B2,
      US,
    );
    expect(loaded).toEqual({
      type: "list",
      operator: "between",
      value1: "a,b",
      value2: "",
      allowBlank: false,
      showDropdown: true,
      inputShow: true,
      inputTitle: "Priority",
      inputText: "Pick one",
      errorShow: false,
      errorStyle: "warning",
      errorTitle: "Oops",
      errorText: "Pick from the list",
      odsListOrder: "sort-ascending",
    });
    const result = read(loaded);
    expect("settings" in result && result.settings.odsListOrder).toBe("sort-ascending");
  });

  it("reads back what it shows, in any locale", () => {
    const rules = [
      rule({ type: "list", formula1: '"Low,Medium,High"' }),
      rule({ type: "list", formula1: "Lists!$A$1:$A$4" }),
      rule({ type: "whole", formula1: "1", formula2: "10", error: { show: true, style: "warning", title: "", text: "" } }),
      rule({ type: "decimal", operator: "greaterThan", formula1: "0.25" }),
      rule({ type: "date", formula1: "45306", formula2: "45306.5" }),
      rule({ type: "time", operator: "notBetween", formula1: "0.375", formula2: "0.75" }),
      rule({ type: "textLength", operator: "lessThanOrEqual", formula1: "5" }),
      rule({ type: "custom", formula1: "B2>$Z$1" }),
      rule({ type: "any", input: { show: true, title: "Hint", text: "Anything" } }),
    ];
    for (const format of [US, DE]) {
      for (const each of rules) {
        const active = { row: 4, col: 1 };
        const result = settingsFromFields(fieldsFromRule(each, active, format), { active, origin: B2, format });
        expect("settings" in result && sameValidationSettings(result.settings, each), `${format.locale.tag} ${each.type} ${each.formula1}`).toBe(true);
      }
    }
  });
});

describe("validationFields — the fields read back (SPEC-130)", () => {
  it("makes a Source without '=' a literal list: split on commas, trimmed", () => {
    expect(read({ type: "list", value1: "Yes, No" })).toEqual({
      settings: { ...defaultValidationSettings(), type: "list", formula1: '"Yes,No"' },
    });
    expect(read({ type: "list", value1: 'Say "hi",Bye' })).toMatchObject({ settings: { formula1: '"Say ""hi"",Bye"' } });
    expect(read({ type: "list", value1: "=\"a,b\"" })).toMatchObject({ settings: { formula1: '"a,b"' } });
  });

  it("refuses a literal list over 255 characters", () => {
    const at255 = Array.from({ length: 51 }, (_, i) => `i${String(i).padStart(3, "0")}`).join(",");
    expect(at255.length).toBe(254);
    expect(read({ type: "list", value1: `${at255}x` })).toMatchObject({ settings: { type: "list" } });
    expect(read({ type: "list", value1: `${at255}xy` })).toEqual({ error: LIST_TOO_LONG_ERROR, field: "value1" });
    expect(LIST_TOO_LONG_ERROR).toBe("The list source can't be longer than 255 characters");
  });

  it("takes a reference Source of one row or one column, a name or a formula — and refuses an empty or 2-D one", () => {
    expect(read({ type: "list", value1: "=$E$2:$E$6" })).toMatchObject({ settings: { formula1: "$E$2:$E$6" } });
    expect(read({ type: "list", value1: "=Lists!A1:D1" })).toMatchObject({ settings: { formula1: "Lists!A1:D1" } });
    expect(read({ type: "list", value1: "=E:E" })).toMatchObject({ settings: { formula1: "E:E" } });
    expect(read({ type: "list", value1: "=$3:$3" })).toMatchObject({ settings: { formula1: "$3:$3" } });
    expect(read({ type: "list", value1: "=E2" })).toMatchObject({ settings: { formula1: "E2" } });
    expect(read({ type: "list", value1: "=Colours" })).toMatchObject({ settings: { formula1: "Colours" } });
    expect(read({ type: "list", value1: "=OFFSET($E$2,0,0,5,1)" })).toMatchObject({ settings: { formula1: "OFFSET($E$2,0,0,5,1)" } });
    for (const source of ["", "  ", "=A1:B2", "=E:F", "=1:2", "=Sheet1:Sheet2!A1:A3"]) {
      expect(read({ type: "list", value1: source }), source).toEqual({ error: LIST_SOURCE_ERROR, field: "value1" });
    }
    expect(LIST_SOURCE_ERROR).toBe("The list source must be a delimited list, or a reference to a single row or column");
    expect(read({ type: "list", value1: "=SUM(" })).toEqual({ error: FORMULA_ERROR, field: "value1" });
  });

  it("stores a formula typed at the active cell relative to the selection's origin", () => {
    // Selection D2:D10 typed at D5: E5:E9 means E2:E6 from the origin D2.
    const ctx = { active: { row: 4, col: 3 }, origin: { row: 1, col: 3 }, format: US };
    expect(settingsFromFields(fields({ type: "list", value1: "=E5:E9" }), ctx)).toMatchObject({ settings: { formula1: "E2:E6" } });
    expect(settingsFromFields(fields({ type: "custom", value1: "=D5>$Z$1" }), ctx)).toMatchObject({ settings: { formula1: "D2>$Z$1" } });
    expect(settingsFromFields(fields({ type: "whole", value1: "=C5", value2: "=$C$1" }), ctx)).toMatchObject({ settings: { formula1: "C2", formula2: "$C$1" } });
  });

  it("reads a custom Formula with or without '=', canonical; refuses one that does not parse", () => {
    expect(read({ type: "custom", value1: "=isnumber( B2 )" })).toMatchObject({ settings: { formula1: "ISNUMBER(B2)" } });
    expect(read({ type: "custom", value1: "LEN(B2)<5" })).toMatchObject({ settings: { formula1: "LEN(B2)<5" } });
    for (const text of ["", "=", "=1+", "=SUM(B2"]) {
      expect(read({ type: "custom", value1: text }), text).toEqual({ error: FORMULA_ERROR, field: "value1" });
    }
    expect(FORMULA_ERROR).toBe("The formula contains an error");
  });

  it("stores numbers, dates and times as invariant number text, in the user's locale", () => {
    expect(read({ type: "whole", value1: "1", value2: "10" })).toEqual({
      settings: { ...defaultValidationSettings(), type: "whole", formula1: "1", formula2: "10" },
    });
    expect(read({ type: "decimal", operator: "greaterThan", value1: "1,5" }, DE)).toMatchObject({ settings: { operator: "greaterThan", formula1: "1.5" } });
    expect(read({ type: "decimal", operator: "lessThan", value1: "50%" })).toMatchObject({ settings: { formula1: "0.5" } });
    expect(read({ type: "date", value1: "1/15/2024", value2: "12/31/2024" })).toMatchObject({ settings: { formula1: "45306", formula2: "45657" } });
    expect(read({ type: "date", operator: "equal", value1: "15.01.2024" }, DE)).toMatchObject({ settings: { operator: "equal", formula1: "45306" } });
    expect(read({ type: "time", value1: "9:00", value2: "18:00" })).toMatchObject({ settings: { formula1: "0.375", formula2: "0.75" } });
    // One bound for the comparisons: no formula2.
    expect(read({ type: "whole", operator: "lessThan", value1: "3", value2: "7" })).toEqual({
      settings: { ...defaultValidationSettings(), type: "whole", operator: "lessThan", formula1: "3" },
    });
  });

  it("refuses an empty or non-numeric bound as a formula error", () => {
    for (const [type, value] of [
      ["whole", ""],
      ["whole", "abc"],
      ["decimal", "TRUE"],
      ["date", "someday"],
      ["time", "#N/A"],
    ] as const) {
      expect(read({ type, operator: "greaterThan", value1: value }), `${type} ${value}`).toEqual({ error: FORMULA_ERROR, field: "value1" });
    }
    expect(read({ type: "whole", value1: "1", value2: "" })).toEqual({ error: FORMULA_ERROR, field: "value2" });
  });

  it("refuses a Maximum below the Minimum when both are constants", () => {
    expect(read({ type: "whole", value1: "10", value2: "1" })).toEqual({ error: BOUNDS_ORDER_ERROR, field: "value2" });
    expect(read({ type: "date", operator: "notBetween", value1: "12/31/2024", value2: "1/1/2024" })).toEqual({ error: BOUNDS_ORDER_ERROR, field: "value2" });
    expect(read({ type: "whole", value1: "5", value2: "5" })).toMatchObject({ settings: { formula1: "5", formula2: "5" } });
    // A formula bound is not compared.
    expect(read({ type: "whole", value1: "10", value2: "=A1" })).toMatchObject({ settings: { formula1: "10", formula2: "A1" } });
    expect(BOUNDS_ORDER_ERROR).toBe("The Maximum must be greater than or equal to the Minimum");
  });

  it("takes a text length of a whole number of 0 or more, or a formula", () => {
    expect(read({ type: "textLength", operator: "lessThanOrEqual", value1: "5" })).toMatchObject({ settings: { formula1: "5" } });
    expect(read({ type: "textLength", value1: "0", value2: "=$Z$1" })).toMatchObject({ settings: { formula1: "0", formula2: "$Z$1" } });
    for (const value of ["", "-1", "2.5", "abc"]) {
      expect(read({ type: "textLength", operator: "equal", value1: value }), value).toEqual({ error: TEXT_LENGTH_ERROR, field: "value1" });
    }
    expect(read({ type: "textLength", value1: "3", value2: "1" })).toEqual({ error: BOUNDS_ORDER_ERROR, field: "value2" });
    expect(TEXT_LENGTH_ERROR).toBe("Enter a whole number of 0 or more");
  });

  it("keeps the canonical defaults for what a type does not use", () => {
    expect(read({ type: "list", operator: "lessThan", value1: "a,b" })).toMatchObject({ settings: { operator: "between" } });
    expect(read({ type: "whole", showDropdown: false, value1: "1", value2: "2" })).toMatchObject({ settings: { showDropdown: true } });
    expect(read({ type: "list", showDropdown: false, value1: "a" })).toMatchObject({ settings: { showDropdown: false } });
    const any = read({ type: "any", value1: "ignored" });
    expect(any).toEqual({ settings: defaultValidationSettings() });
    expect("settings" in any && isEmptyValidation(any.settings)).toBe(true);
    const withMessage = read({ type: "any", inputTitle: "Hint" });
    expect("settings" in withMessage && isEmptyValidation(withMessage.settings)).toBe(false);
  });

  it("compares field sets", () => {
    expect(sameValidationFields(defaultValidationFields(), defaultValidationFields())).toBe(true);
    expect(sameValidationFields(fields({ value1: "a" }), defaultValidationFields())).toBe(false);
    expect(sameValidationFields(fields({ odsListOrder: "sort-ascending" }), defaultValidationFields())).toBe(false);
  });
});

describe("validationFields — the selection (SPEC-130)", () => {
  const list = (id: string, ranges: GridRange[], formula1: string): DataValidation => ({ ...rule({ type: "list", formula1 }), id, ranges });

  it("disagrees when some cells are covered and some not, or rules differ", () => {
    const sheet = { validations: [list("a", [range(1, 1, 9, 1)], '"Low,High"'), list("b", [range(1, 3, 9, 3)], '"Yes,No"')] };
    expect(selectionDisagrees({ validations: [] }, [range(1, 1, 3, 3)])).toBe(false);
    expect(selectionDisagrees(sheet, [range(2, 1, 4, 1)])).toBe(false);
    expect(selectionDisagrees(sheet, [range(5, 5)])).toBe(false);
    expect(selectionDisagrees(sheet, [range(8, 1, 10, 1)])).toBe(true);
    expect(selectionDisagrees(sheet, [range(2, 1), range(2, 3)])).toBe(true);
  });

  it("agrees over two rules with the same settings (formulas compared from one origin)", () => {
    const sheet = {
      validations: [
        { ...rule({ type: "custom", formula1: "B2>0" }), id: "a", ranges: [range(1, 1)] },
        { ...rule({ type: "custom", formula1: "B3>0" }), id: "b", ranges: [range(2, 1)] },
      ],
    };
    expect(selectionDisagrees(sheet, [range(1, 1, 2, 1)])).toBe(false);
    sheet.validations[1] = { ...sheet.validations[1]!, formula1: "B2>0" };
    expect(selectionDisagrees(sheet, [range(1, 1, 2, 1)])).toBe(true);
  });

  it("makes a selection's ranges disjoint, keeping the first range's top-left first", () => {
    expect(disjointRanges([range(1, 1, 4, 1)])).toEqual([range(1, 1, 4, 1)]);
    const out = disjointRanges([range(3, 1, 5, 1), range(1, 1, 9, 1), range(3, 1, 4, 1)]);
    expect(out[0]).toEqual(range(3, 1, 5, 1));
    expect(out).toEqual([range(3, 1, 5, 1), range(1, 1, 2, 1), range(6, 1, 9, 1)]);
  });
});
