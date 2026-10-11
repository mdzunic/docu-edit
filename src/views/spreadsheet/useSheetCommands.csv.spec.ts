// SPEC-126: adding or duplicating a sheet in a csv/tsv workbook is allowed,
// and raises CSV_SECOND_SHEET_NOTICE once per view session; a package
// workbook never sees it.
import { afterEach, describe, expect, it, vi } from "vitest";
import { effectScope, ref, shallowRef, type EffectScope } from "vue";
import { useToasts } from "../../composables/useToasts";
import { CSV_SECOND_SHEET_NOTICE } from "../../lib/spreadsheet/csv/messages";
import { readCsvBytes } from "../../lib/spreadsheet/csv/parse";
import { buildSpec126SepCsvBytes } from "../../lib/spreadsheet/fixtures/build";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useSheetCommands } from "./useSheetCommands";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

const scopes: EffectScope[] = [];

afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  const { toasts, dismissToast } = useToasts();
  for (const toast of toasts.value) {
    dismissToast(toast.id);
  }
});

function commandsFor(model: SpreadsheetModel) {
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const undo = { record: vi.fn(), commit: vi.fn() } as unknown as SpreadsheetUndo;
  const scope = effectScope();
  scopes.push(scope);
  return scope.run(() =>
    useSheetCommands({
      book: { workbook, activeIndex, selectSheet: (index: number) => (activeIndex.value = index) },
      undo,
      syncViewState: () => {},
    }),
  )!;
}

function notices(): string[] {
  return useToasts().toasts.value.map((toast) => toast.message).filter((message) => message === CSV_SECOND_SHEET_NOTICE);
}

describe("CSV_SECOND_SHEET_NOTICE", () => {
  it("reads as the spec says", () => {
    expect(CSV_SECOND_SHEET_NOTICE).toBe("CSV files keep only the first sheet — download a copy as .xlsx to keep the others");
  });

  it("is raised once per view session when a csv workbook gains a sheet by insert or duplicate", () => {
    const model = readCsvBytes(buildSpec126SepCsvBytes(), "csv", "Fruit", { userLocaleTag: "en-US" });
    const commands = commandsFor(model);
    commands.insert();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Fruit", "Sheet1"]);
    expect(notices()).toEqual([CSV_SECOND_SHEET_NOTICE]);
    commands.openMenu({ index: 0, x: 0, y: 0 });
    commands.runMenu("duplicate");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Fruit", "Fruit (2)", "Sheet1"]);
    expect(notices()).toEqual([CSV_SECOND_SHEET_NOTICE]);

    // A new view session says it again.
    const again = commandsFor(model);
    again.openMenu({ index: 0, x: 0, y: 0 });
    again.runMenu("duplicate");
    expect(notices()).toEqual([CSV_SECOND_SHEET_NOTICE, CSV_SECOND_SHEET_NOTICE]);
  });

  it("is never raised for a package workbook", () => {
    const commands = commandsFor(createBlankWorkbook());
    commands.insert();
    expect(notices()).toEqual([]);
  });
});
