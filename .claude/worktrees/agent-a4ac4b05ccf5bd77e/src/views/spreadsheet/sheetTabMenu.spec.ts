// The sheet tab menu's items (SPEC-121 "UI"; SPEC-122 "Protection"): every
// item is present, one that cannot apply disabled with its reason — and on a
// workbook whose structure is protected, every one of them disabled with
// "The workbook's structure is protected".
import { describe, expect, it } from "vitest";
import { addSheet, createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { asSheetTabMenuItems, sheetTabMenuItems, STRUCTURE_PROTECTED_MESSAGE } from "./sheetTabMenu";

function twoSheets() {
  const model = createBlankWorkbook();
  addSheet(model, 0);
  return model;
}

describe("sheetTabMenuItems", () => {
  it("offers every command, disabling only those that cannot apply", () => {
    const items = sheetTabMenuItems(twoSheets(), 0);
    expect(items.map((item) => item.action)).toEqual([
      "insert",
      "rename",
      "duplicate",
      "delete",
      "move-left",
      "move-right",
      "hide",
      "unhide",
      "tab-color",
    ]);
    expect(items.filter((item) => item.disabled).map((item) => [item.action, item.reason])).toEqual([
      ["move-left", "This sheet is already first"],
      ["unhide", "No sheets are hidden"],
    ]);
  });

  it("disables every command with the protection reason when the structure is protected", () => {
    const model = { ...twoSheets(), structureProtected: true as const };
    const items = asSheetTabMenuItems(sheetTabMenuItems(model, 1));
    expect(items).toHaveLength(9);
    for (const item of items) {
      expect(item.disabled, item.action).toBe(true);
      expect(item.reason, item.action).toBe(STRUCTURE_PROTECTED_MESSAGE);
    }
    expect(STRUCTURE_PROTECTED_MESSAGE).toBe("The workbook's structure is protected");
  });
});
