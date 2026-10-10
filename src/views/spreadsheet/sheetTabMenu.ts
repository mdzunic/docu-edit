// The sheet tab context menu's items (SPEC-121 "UI"): what each command
// offers for a sheet — every item present, one that cannot apply disabled
// with its reason; SPEC-122: all of them disabled with
// STRUCTURE_PROTECTED_MESSAGE on a workbook whose structure is protected.
// SheetTabMenu takes its items as a runtime `Array` prop (the espree
// constraint on .vue scripts bars a typed PropType there — see AppShell.vue),
// so it coerces them through here. Framework-free. Item test ids are
// `sheet-tab-menu-<action>` — distinct from the menu bar's `sheet-menu-*`
// and SPEC-124's `sheet-delete-dialog`.
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { cannotDuplicateReason, isLastVisibleSheet, LAST_VISIBLE_SHEET_MESSAGE } from "../../lib/spreadsheet/sheets";

/** One command in the tab menu. An item that cannot apply is `disabled` with
 * the stated `reason` (shown as its title and inline) — never hidden. */
export interface SheetTabMenuItem {
  action: string;
  label: string;
  testid: string;
  disabled?: boolean;
  reason?: string;
  /** Destructive (Delete): drawn in the danger ink, as MenuItem does. */
  danger?: boolean;
  /** Draw a separator above this item, to group the commands. */
  separatorBefore?: boolean;
}

/** Coerces the untyped `items` prop into the typed list the menu renders. */
export function asSheetTabMenuItems(value: unknown): SheetTabMenuItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((entry) => {
    const item = (entry ?? {}) as Partial<SheetTabMenuItem>;
    return {
      action: String(item.action ?? ""),
      label: String(item.label ?? ""),
      testid: String(item.testid ?? ""),
      disabled: item.disabled === true,
      reason: typeof item.reason === "string" ? item.reason : "",
      danger: item.danger === true,
      separatorBefore: item.separatorBefore === true,
    };
  });
}

export const FIRST_SHEET_MESSAGE = "This sheet is already first";
export const LAST_SHEET_MESSAGE = "This sheet is already last";
export const NO_HIDDEN_SHEETS_MESSAGE = "No sheets are hidden";
/** SPEC-122: why every sheet command is disabled on a workbook whose
 * structure is protected. */
export const STRUCTURE_PROTECTED_MESSAGE = "The workbook's structure is protected";

/** `Sheet "<name>" moved to position <k>`. */
export function movedAnnouncement(name: string, position: number): string {
  return `Sheet "${name}" moved to position ${position}`;
}

/** The visible sheet before / after `index` in model order, or -1. */
export function visibleNeighbour(model: SpreadsheetModel, index: number, step: -1 | 1): number {
  for (let i = index + step; i >= 0 && i < model.sheets.length; i += step) {
    if (model.sheets[i]!.visibility === "visible") {
      return i;
    }
  }
  return -1;
}

/** The tab menu's actions and labels, in menu order. */
const SHEET_TAB_MENU_ACTIONS: readonly (readonly [string, string])[] = [
  ["insert", "Insert sheet"],
  ["rename", "Rename"],
  ["duplicate", "Duplicate"],
  ["delete", "Delete"],
  ["move-left", "Move left"],
  ["move-right", "Move right"],
  ["hide", "Hide"],
  ["unhide", "Unhide…"],
  ["tab-color", "Tab colour"],
];

/** One menu item: `sheet-tab-menu-<action>`, disabled with its reason. */
function item(action: string, label: string, reason: string | null = null): SheetTabMenuItem {
  return reason === null ? { action, label, testid: `sheet-tab-menu-${action}` } : { action, label, testid: `sheet-tab-menu-${action}`, disabled: true, reason };
}

/** The tab context menu's items for the sheet at `index` (SPEC-121 "UI");
 * all of them disabled with STRUCTURE_PROTECTED_MESSAGE when the workbook's
 * structure is protected (SPEC-122). */
export function sheetTabMenuItems(model: SpreadsheetModel, index: number): SheetTabMenuItem[] {
  if (model.structureProtected) {
    return SHEET_TAB_MENU_ACTIONS.map(([action, label]) => item(action, label, STRUCTURE_PROTECTED_MESSAGE));
  }
  const lastVisible = isLastVisibleSheet(model, index);
  const duplicateReason = cannotDuplicateReason(model, index);
  const hasHidden = model.sheets.some((sheet) => sheet.visibility === "hidden");
  return [
    item("insert", "Insert sheet"),
    item("rename", "Rename"),
    item("duplicate", "Duplicate", duplicateReason),
    item("delete", "Delete", lastVisible ? LAST_VISIBLE_SHEET_MESSAGE : null),
    item("move-left", "Move left", visibleNeighbour(model, index, -1) < 0 ? FIRST_SHEET_MESSAGE : null),
    item("move-right", "Move right", visibleNeighbour(model, index, 1) < 0 ? LAST_SHEET_MESSAGE : null),
    item("hide", "Hide", lastVisible ? LAST_VISIBLE_SHEET_MESSAGE : null),
    item("unhide", "Unhide…", hasHidden ? null : NO_HIDDEN_SHEETS_MESSAGE),
    item("tab-color", "Tab colour"),
  ];
}
