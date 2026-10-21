// SPEC-130 "Entry gate" and the grid chrome's state: useDataValidation's
// checkEntry driven with each alert style and each button (the verdicts),
// the defaults for an untitled rule, the first failure deciding across
// several targets (an `error.show` false rule committing silently), a
// formula candidate checked by its evaluated result, valid entries resolving
// "commit" without an alert; the sparse fill targets and the per-rule memo;
// the input message (prompts disabled, hidden while the list is open); the
// dropdown (showDropdown, the listed items and the current value's item,
// Alt+Down, the unavailable source, a pick entering the item's value through
// the editor and returning focus to the grid, closing on another cell). Small
// hand-built models with `validations` set directly.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef } from "vue";
import { parseA1, parseSqref, type CellRef } from "../../lib/spreadsheet/address";
import type { Cell, CellValue, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { defaultValidationSettings, type DataValidation } from "../../lib/spreadsheet/validation";
import {
  ALERT_VERDICTS,
  DEFAULT_ALERT_TEXT,
  DEFAULT_ALERT_TITLE,
  firstValidationFailure,
  positionIndependent,
  useDataValidation,
  validatedFillTargets,
  validationDepsFor,
  type EntryVerdict,
  type ValidationEditor,
} from "./useDataValidation";

const scopes: Array<() => void> = [];
afterEach(() => {
  for (const stop of scopes.splice(0)) {
    stop();
  }
});

let ids = 0;
/** A rule on `sqref` (Excel's defaults, then `settings`). */
function rule(sqref: string, settings: Partial<DataValidation> = {}): DataValidation {
  ids += 1;
  return { id: `dv${ids}`, ranges: parseSqref(sqref), ...defaultValidationSettings(), ...settings };
}

/** A stop/warning/information error part. */
function error(style: DataValidation["error"]["style"], title = "", text = "", show = true): DataValidation["error"] {
  return { show, style, title, text };
}

function rowsOf(cells: Record<string, CellValue>): Row[] {
  const byRow = new Map<number, Cell[]>();
  for (const [address, value] of Object.entries(cells)) {
    const at = parseA1(address)!;
    byRow.set(at.row, [...(byRow.get(at.row) ?? []), { col: at.col, value }]);
  }
  return [...byRow].toSorted((a, b) => a[0] - b[0]).map(([index, cells]) => ({ index, cells: cells.toSorted((a, b) => a.col - b.col) }));
}

/** A blank workbook whose Sheet1 holds `cells` and `validations`. */
function book(cells: Record<string, CellValue>, validations: DataValidation[], sheet: Partial<Sheet> = {}): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows: rowsOf(cells), validations, ...sheet }];
  return model;
}

const at = (address: string): CellRef => {
  const ref = parseA1(address)!;
  return { row: ref.row, col: ref.col };
};

function setup(model: SpreadsheetModel, active = "A1") {
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const current = ref<SheetSelection>({ active: at(active), ranges: [{ top: at(active).row, left: at(active).col, bottom: at(active).row, right: at(active).col }], activeRangeIndex: 0 });
  const entered: Array<{ value: CellValue; text: string }> = [];
  const editor: ValidationEditor & { focused: number } = {
    focused: 0,
    editing: ref(false),
    enterValue: (value, text) => (entered.push({ value, text }), true),
    focusCell: () => {
      editor.focused += 1;
    },
    commit: () => true,
  };
  const scope = effectScope();
  const validation = scope.run(() =>
    useDataValidation({ book: { workbook, activeSheet }, selection: { selection: computed(() => current.value) }, locale: () => SPREADSHEET_LOCALES["en-US"], editor: () => editor }),
  )!;
  scopes.push(() => scope.stop());
  const select = (address: string): void => {
    const ref = at(address);
    current.value = { active: ref, ranges: [{ top: ref.row, left: ref.col, bottom: ref.row, right: ref.col }], activeRangeIndex: 0 };
  };
  return { validation, editor, entered, select, workbook };
}

/** checkEntry's promise, answered with `choice` once its alert opened. */
async function answered(validation: ReturnType<typeof setup>["validation"], promise: Promise<EntryVerdict>, choice: keyof typeof ALERT_VERDICTS): Promise<EntryVerdict> {
  expect(validation.alert.value).not.toBeNull();
  validation.answer(choice);
  expect(validation.alert.value).toBeNull();
  return promise;
}

describe("checkEntry — the verdicts", () => {
  it("a valid entry resolves \"commit\" without an alert", async () => {
    const { validation } = setup(book({}, [rule("B2", { type: "whole", operator: "between", formula1: "1", formula2: "10", error: error("stop") })]));
    await expect(validation.checkEntry([at("B2")], { value: 5 })).resolves.toBe("commit");
    expect(validation.alert.value).toBeNull();
    // A cell no rule covers is always valid.
    await expect(validation.checkEntry([at("C9")], { value: "anything" })).resolves.toBe("commit");
  });

  it("stop: the default title and text; Retry → retry, Cancel → cancel", async () => {
    const { validation } = setup(book({}, [rule("B2", { type: "list", formula1: '"Low,Medium,High"', error: error("stop") })]));
    const first = validation.checkEntry([at("B2")], { value: "Urgent" });
    expect(validation.alert.value).toEqual({ style: "stop", title: DEFAULT_ALERT_TITLE, text: DEFAULT_ALERT_TEXT });
    expect(DEFAULT_ALERT_TITLE).toBe("Invalid entry");
    expect(DEFAULT_ALERT_TEXT).toBe("This value doesn't match the data validation rules for this cell.");
    await expect(answered(validation, first, "retry")).resolves.toBe("retry");
    const second = validation.checkEntry([at("B2")], { value: "Urgent" });
    await expect(answered(validation, second, "cancel")).resolves.toBe("cancel");
  });

  it("warning: its own title and text; Yes → commit, No → retry, Cancel → cancel", async () => {
    const { validation } = setup(book({}, [rule("B3", { type: "whole", formula1: "1", formula2: "10", error: error("warning", "Out of range", "Between 1 and 10") })]));
    const outcomes: EntryVerdict[] = [];
    for (const choice of ["yes", "no", "cancel"] as const) {
      const promise = validation.checkEntry([at("B3")], { value: 42 });
      expect(validation.alert.value).toEqual({ style: "warning", title: "Out of range", text: "Between 1 and 10" });
      outcomes.push(await answered(validation, promise, choice));
    }
    expect(outcomes).toEqual(["commit", "retry", "cancel"]);
  });

  it("information: OK → commit, Cancel → cancel", async () => {
    const { validation } = setup(book({}, [rule("B4", { type: "decimal", operator: "greaterThan", formula1: "0", error: error("information") })]));
    const ok = validation.checkEntry([at("B4")], { value: -1 });
    expect(validation.alert.value?.style).toBe("information");
    await expect(answered(validation, ok, "ok")).resolves.toBe("commit");
    const cancel = validation.checkEntry([at("B4")], { value: -1 });
    await expect(answered(validation, cancel, "cancel")).resolves.toBe("cancel");
  });

  it("a rule whose error.show is false commits an invalid entry silently", async () => {
    const { validation } = setup(book({}, [rule("B2", { type: "whole", formula1: "1", formula2: "10", error: error("stop", "", "", false) })]));
    await expect(validation.checkEntry([at("B2")], { value: 99 })).resolves.toBe("commit");
    expect(validation.alert.value).toBeNull();
  });

  it("the first failing target decides, in the order given; silent rules are passed over", async () => {
    const model = book({}, [
      rule("A1", { type: "whole", formula1: "1", formula2: "100", error: error("stop", "A1 rule") }), // 42 passes
      rule("B1", { type: "whole", formula1: "1", formula2: "10", error: error("stop", "silent", "", false) }), // fails, silently
      rule("C1", { type: "whole", formula1: "1", formula2: "10", error: error("warning", "C1 rule") }), // fails: decides
      rule("D1", { type: "whole", formula1: "1", formula2: "10", error: error("stop", "D1 rule") }), // fails, later
    ]);
    const { validation } = setup(model);
    const promise = validation.checkEntry([at("A1"), at("B1"), at("C1"), at("D1")], { value: 42 });
    expect(validation.alert.value?.title).toBe("C1 rule");
    await expect(answered(validation, promise, "yes")).resolves.toBe("commit");
    // Reordered, D1 decides.
    const reordered = validation.checkEntry([at("D1"), at("C1")], { value: 42 });
    expect(validation.alert.value?.title).toBe("D1 rule");
    await expect(answered(validation, reordered, "cancel")).resolves.toBe("cancel");
  });

  it("a formula entry is checked by its evaluated result at the target; an uncomputable one is accepted", async () => {
    const model = book({ A1: 3, A2: 30 }, [rule("B1:B2", { type: "whole", formula1: "1", formula2: "10", error: error("stop") })]);
    const { validation } = setup(model);
    await expect(validation.checkEntry([at("B1")], { value: null, formula: "A1*2" })).resolves.toBe("commit"); // 6
    const promise = validation.checkEntry([at("B1")], { value: null, formula: "A2*2" }); // 60
    await expect(answered(validation, promise, "cancel")).resolves.toBe("cancel");
    // As a fill writes it: shifted from the edited cell (B1) to B2 → A2*2 = 60.
    const shifted = validation.checkEntry([at("B2")], { value: null, formula: "A1*2" }, { origin: at("B1") });
    await expect(answered(validation, shifted, "cancel")).resolves.toBe("cancel");
    await expect(validation.checkEntry([at("B1")], { value: null, formula: "NOSUCHFUNCTION(1)" })).resolves.toBe("commit");
  });

  it("blank entries follow allowBlank", async () => {
    const model = book({}, [rule("B1", { type: "whole", formula1: "1", formula2: "10", allowBlank: false, error: error("stop") }), rule("B2", { type: "whole", formula1: "1", formula2: "10" })]);
    const { validation } = setup(model);
    await expect(validation.checkEntry([at("B2")], { value: null })).resolves.toBe("commit");
    const promise = validation.checkEntry([at("B1")], { value: null });
    await expect(answered(validation, promise, "retry")).resolves.toBe("retry");
  });

  it("a second alert cancels one still open", async () => {
    const { validation } = setup(book({}, [rule("B2", { type: "whole", formula1: "1", formula2: "10" })]));
    const first = validation.checkEntry([at("B2")], { value: 50 });
    const second = validation.checkEntry([at("B2")], { value: 60 });
    await expect(first).resolves.toBe("cancel");
    await expect(answered(validation, second, "retry")).resolves.toBe("retry");
  });
});

describe("the fill's targets", () => {
  it("lists only the cells a rule covers, in fill order, skipping a merge's covered cells", () => {
    const sheet = book({ A1: 1 }, [rule("C2:C4", { type: "whole", formula1: "1", formula2: "2" }), rule("B3 E3", { type: "whole", formula1: "1", formula2: "2" })], {
      merges: [{ top: 3, left: 2, bottom: 4, right: 2 }], // C4:C5
    }).sheets[0]!;
    const targets = [...validatedFillTargets(sheet, [{ top: 0, left: 0, bottom: 9, right: 9 }], at("A1"))];
    expect(targets).toEqual([at("C2"), at("B3"), at("C3"), at("E3"), at("C4")]);
  });

  it("is sparse: a rule far down a huge selection is reached without walking the rows between", () => {
    const sheet = book({}, [rule("A1"), rule("A900000:A900001", { type: "whole", formula1: "1", formula2: "2" })]).sheets[0]!;
    const targets = [...validatedFillTargets(sheet, [{ top: 0, left: 0, bottom: 999_999, right: 0 }], at("A1"))];
    expect(targets).toEqual([at("A1"), at("A900000"), at("A900001")]);
  });

  it("checks a position-independent rule once per fill, a relative one at every cell", () => {
    expect(positionIndependent(rule("A1", { type: "list", formula1: '"a,b"' }))).toBe(true);
    expect(positionIndependent(rule("A1", { type: "list", formula1: "$E$2:$E$6" }))).toBe(true);
    expect(positionIndependent(rule("A1", { type: "whole", formula1: "B1", formula2: "10" }))).toBe(false);
    expect(positionIndependent(rule("A1", { type: "whole", formula1: "ROW()", formula2: "10" }))).toBe(false);
    expect(positionIndependent(rule("A1", { type: "custom", formula1: "TRUE" }))).toBe(false);
    // A relative bound: B1 holds 5, B2 holds 50 — 20 fails at A1 and passes at A2.
    const model = book({ B1: 5, B2: 50 }, [rule("A1:A2", { type: "whole", operator: "lessThan", formula1: "B1", error: error("stop") })]);
    const sheet = model.sheets[0]!;
    const deps = validationDepsFor(model, SPREADSHEET_LOCALES["en-US"]);
    expect(firstValidationFailure(model, sheet, [at("A2"), at("A1")], { value: 20 }, deps)?.at).toEqual(at("A1"));
  });
});

describe("the input message", () => {
  it("shows the active cell's rule's title and text while input.show and not disabled", () => {
    const model = book({}, [rule("B2", { input: { show: true, title: "Priority", text: "Pick one" } }), rule("B3", { input: { show: false, title: "Hidden", text: "x" } }), rule("B4")]);
    const { validation, select, workbook } = setup(model, "B2");
    expect(validation.inputMessage.value).toEqual({ title: "Priority", text: "Pick one" });
    select("B3"); // input.show false
    expect(validation.inputMessage.value).toBeNull();
    select("B4"); // no title or text
    expect(validation.inputMessage.value).toBeNull();
    select("C9");
    expect(validation.inputMessage.value).toBeNull();
    select("B2");
    workbook.value = { ...model, sheets: [{ ...model.sheets[0]!, validationPromptsDisabled: true }] };
    expect(validation.inputMessage.value).toBeNull();
  });

  it("hides while the dropdown list is open", () => {
    const model = book({}, [rule("B2", { type: "list", formula1: '"Low,Medium,High"', input: { show: true, title: "Priority", text: "" } })]);
    const { validation } = setup(model, "B2");
    expect(validation.inputMessage.value).not.toBeNull();
    expect(validation.openList()).toBe(true);
    expect(validation.inputMessage.value).toBeNull();
    validation.closeList(false);
    expect(validation.inputMessage.value).not.toBeNull();
  });
});

describe("the dropdown", () => {
  it("shows for a list rule with showDropdown only", () => {
    const model = book({}, [rule("B2", { type: "list", formula1: '"a,b"' }), rule("B8", { type: "list", formula1: '"a,b"', showDropdown: false }), rule("B3", { type: "whole", formula1: "1", formula2: "2" })]);
    const { validation, select } = setup(model, "B2");
    expect(validation.dropdown.value).toEqual({ expanded: false });
    select("B8");
    expect(validation.dropdown.value).toBeNull();
    expect(validation.openList()).toBe(false);
    select("B3");
    expect(validation.dropdown.value).toBeNull();
  });

  it("Alt+Down opens the list in source order with the current value's item selected", () => {
    const model = book({ B2: "medium" }, [rule("B2:B3", { type: "list", formula1: '"Low,Medium,High"' })]);
    const { validation, select } = setup(model, "B2");
    const plain = new KeyboardEvent("keydown", { key: "ArrowDown" });
    expect(validation.onGridKey(plain)).toBe(false);
    expect(validation.onGridKey(new KeyboardEvent("keydown", { key: "ArrowDown", altKey: true }))).toBe(true);
    expect(validation.dropdown.value).toEqual({ expanded: true });
    expect(validation.listView.value).toEqual({ texts: ["Low", "Medium", "High"], selected: 1, unavailable: false });
    // Another active cell closes it; there, nothing is selected (B3 is empty).
    select("B3");
    expect(validation.listView.value).toBeNull();
    validation.openList();
    expect(validation.listView.value?.selected).toBe(-1);
    validation.selectItem(2);
    expect(validation.listView.value?.selected).toBe(2);
  });

  it("an unavailable source lists nothing and says so", () => {
    const model = book({}, [rule("B10", { type: "list", formula1: "Colours" })]);
    const { validation } = setup(model, "B10");
    validation.openList();
    expect(validation.listView.value).toEqual({ texts: [], selected: -1, unavailable: true });
  });

  it("a pick enters the item's value through the editor and returns focus to the grid; Escape only returns focus", () => {
    const model = book({ E2: 10, E3: 20, E4: "thirty" }, [rule("B8", { type: "list", formula1: "$E$2:$E$4" })]);
    const { validation, editor, entered } = setup(model, "B8");
    validation.openList();
    expect(validation.listView.value?.texts).toEqual(["10", "20", "thirty"]);
    expect(validation.pick(1)).toBe(true);
    expect(entered).toEqual([{ value: 20, text: "20" }]); // a number stays a number
    expect(validation.listView.value).toBeNull();
    expect(editor.focused).toBe(1);
    validation.openList();
    validation.closeList(true);
    expect(validation.listView.value).toBeNull();
    expect(editor.focused).toBe(2);
    expect(entered).toHaveLength(1);
  });

  it("an edit in progress commits before the list opens; a held commit keeps it closed", () => {
    const model = book({}, [rule("B2", { type: "list", formula1: '"a,b"' })]);
    const { validation, editor } = setup(model, "B2");
    editor.editing = ref(true);
    const commit = vi.fn(() => false);
    editor.commit = commit;
    expect(validation.openList()).toBe(false);
    expect(commit).toHaveBeenCalledTimes(1);
    commit.mockReturnValue(true);
    expect(validation.openList()).toBe(true);
  });

  it("the list closes when another sheet shows", async () => {
    const model = book({}, [rule("A1", { type: "list", formula1: '"a,b"' })]);
    const { validation, workbook } = setup(model, "A1");
    validation.openList();
    workbook.value = { ...model, sheets: [{ ...model.sheets[0]!, id: "other" }] };
    await nextTick();
    expect(validation.listView.value).toBeNull();
  });
});
