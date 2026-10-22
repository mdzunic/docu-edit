// The placeholder a foreign sheet shows in place of the grid (SPEC-116): a
// chartsheet, dialogsheet or macrosheet is carried verbatim but not shown
// yet. A chart sheet is told apart by its part's declared content type in
// the retained `[Content_Types].xml` (the package's own declaration), with
// Excel's `chartsheets/` folder as the fallback for a package that declares
// none for it.
import { contentTypeOf } from "../core/packageDetect";
import { base64ToBytes } from "../core/packageParts";
import type { Sheet, SpreadsheetModel } from "./model";

const CHARTSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

export const CHART_SHEET_MESSAGE = "This chart sheet can't be shown yet";
export const FOREIGN_SHEET_MESSAGE = "This sheet type can't be shown yet";

/** Whether a foreign sheet is a chart sheet. */
export function isChartSheet(model: SpreadsheetModel, sheet: Sheet): boolean {
  if (sheet.kind !== "foreign" || !sheet.partName) {
    return false;
  }
  const contentTypes = model.opaque.find((part) => part.name === "[Content_Types].xml");
  const declared = contentTypes
    ? contentTypeOf(new Map([["[Content_Types].xml", base64ToBytes(contentTypes.data)]]), sheet.partName)
    : undefined;
  return declared !== undefined && declared !== "application/xml"
    ? declared === CHARTSHEET_CONTENT_TYPE
    : sheet.partName.includes("/chartsheets/");
}

/** The sheet-foreign panel's text for a foreign sheet. */
export function foreignSheetMessage(model: SpreadsheetModel, sheet: Sheet): string {
  return isChartSheet(model, sheet) ? CHART_SHEET_MESSAGE : FOREIGN_SHEET_MESSAGE;
}
