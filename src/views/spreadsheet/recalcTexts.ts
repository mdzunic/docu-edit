// The texts of recalculation in the workbook view (SPEC-120 "Stale cells",
// "Cycles", "Opening a workbook"): the manual-mode toast, the status's
// "Calculating…" and circular-reference text, a stale cell's title, and the
// F9 shortcut — framework-free, beside useRecalc, which drives the passes
// and re-exports them.
import { formatA1 } from "../../lib/spreadsheet/address";
import { parseCellKey, type CellKey } from "../../lib/spreadsheet/formula/dependencies";
import type { UncomputableReason } from "../../lib/spreadsheet/formula/evaluate";
import { quoteSheetName } from "../../lib/spreadsheet/formula/print";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";

export const MANUAL_MODE_MESSAGE = "This workbook was set to calculate manually. It recalculates automatically here.";
export const CALCULATING_TEXT = "Calculating…";
export const STALE_TITLE_PREFIX = "This value couldn't be recalculated here — ";

/** A stale cell's title (SPEC-120 "Stale cells"). */
export function staleTitle(reason: UncomputableReason): string {
  const suffix = ((): string => {
    switch (reason.kind) {
      case "function":
        return `${reason.name} isn't supported yet`;
      case "name":
        return "named ranges aren't supported yet";
      case "structured":
        return "table references aren't supported yet";
      case "external":
        return "it links to another workbook";
      case "dynamicArray":
        return "dynamic array formulas aren't supported yet";
      case "unparsed":
        return "its formula couldn't be read";
      case "iterative":
        return "iterative calculation isn't supported";
      case "regex":
        return "regular-expression criteria aren't supported";
      case "unsupportedCase":
        return `this use of ${reason.name} isn't supported yet`;
      default:
        return "data tables aren't supported";
    }
  })();
  return STALE_TITLE_PREFIX + suffix;
}

/** "Circular reference: Loop!A1" for a cycle's first cell. */
export function circularText(model: SpreadsheetModel, key: CellKey): string {
  const { sheetId, row, col } = parseCellKey(key);
  const sheet = model.sheets.find((candidate) => candidate.id === sheetId);
  return `Circular reference: ${quoteSheetName(sheet?.name ?? "")}!${formatA1({ row, col })}`;
}

/** F9 alone: a full recalculation. */
export function isRecalculateShortcut(event: Pick<KeyboardEvent, "key" | "ctrlKey" | "shiftKey" | "altKey" | "metaKey">): boolean {
  return event.key === "F9" && !event.ctrlKey && !event.shiftKey && !event.altKey && !event.metaKey;
}
