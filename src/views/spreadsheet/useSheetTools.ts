// The grid's editing tools, composed in one call so SpreadsheetView stays a
// wiring hub (SpreadsheetView.size.spec.ts; extracted by SPEC-134 to make
// room for find and Go To): SPEC-123's formatting and the format painter,
// SPEC-124's rows and columns, header resizing and merges, SPEC-125's
// clipboard, fill handle and drag-move, SPEC-128's sort and filter, SPEC-133's
// names and tables — each over the view's book, view state, selection and
// editing exactly as the view called them before, in the same order — and
// the menu bar that drives them (useSpreadsheetMenu). SPEC-136: page setup,
// the print preview and PDF export (usePageSetup, useSheetPrint,
// useSheetPdfExport) — the menu's File ▸ Page setup… / Print… / Print area,
// View ▸ Page breaks and Insert ▸ page breaks — returned as `printing`.
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { useFillHandle } from "./useFillHandle";
import { useFormatPainter } from "./useFormatPainter";
import { useFormatting } from "./useFormatting";
import type { FormulaDisplay } from "./useFormulaDisplay";
import { useHeaderResize } from "./useHeaderResize";
import { useMerges } from "./useMerges";
import { useNames } from "./useNames";
import { usePageSetup } from "./usePageSetup";
import { useSheetPdfExport } from "./useSheetPdfExport";
import { useSheetPrint } from "./useSheetPrint";
import { useRangeMove } from "./useRangeMove";
import type { Recalc } from "./useRecalc";
import { useSheetClipboard } from "./useSheetClipboard";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import { useSheetTables } from "./useSheetTables";
import type { SheetViewState } from "./useSheetViewState";
import type { SheetZoom } from "./useSheetZoom";
import { useSortFilter } from "./useSortFilter";
import { useSpreadsheetMenu } from "./useSpreadsheetMenu";
import { useStructure } from "./useStructure";
import type { WorkbookEditing } from "./useWorkbookEditing";
import type { asWorkbookBook } from "./useWorkbookLoad";

export interface SheetToolsDeps {
  book: ReturnType<typeof asWorkbookBook>;
  viewState: SheetViewState;
  selection: SheetSelectionController;
  sheetMetrics: SheetMetrics;
  zoom: SheetZoom;
  formulas: FormulaDisplay;
  recalc: Recalc;
  editing: WorkbookEditing;
  locale: () => SpreadsheetLocale;
}

export function useSheetTools(deps: SheetToolsDeps) {
  const { book, viewState, selection, sheetMetrics, zoom, formulas, recalc, editing, locale } = deps;
  const formatting = useFormatting({ book, viewState, selection, editing, locale });
  const structure = useStructure({ book, viewState, selection, sheetMetrics, editing });
  const headerResize = useHeaderResize({ book, structure, selection, sheetMetrics, zoom });
  const merges = useMerges({ book, viewState, selection, sheetMetrics, editing });
  const painter = useFormatPainter({ book, selection, formatting, merges });
  const clipboard = useSheetClipboard({ book, viewState, selection, editing, locale });
  const fillHandle = useFillHandle({ book, viewState, selection, editing, locale });
  const rangeMove = useRangeMove({ book, viewState, selection, editing });
  const sortFilter = useSortFilter({ book, viewState, selection, sheetMetrics, editing, locale });
  const names = useNames({ book, viewState, selection, editing });
  const tables = useSheetTables({ book, viewState, selection, sheetMetrics, editing, locale });
  const pageSetup = usePageSetup({ book, viewState, selection, editing });
  const print = useSheetPrint({ book, selection, pageSetup, locale });
  const printing = { pageSetup, print, pdf: useSheetPdfExport({ book, print }) };
  const menu = useSpreadsheetMenu({ book, viewState, zoom, formulas, recalc, editing, formatting, structure, clipboard, fillHandle, sortFilter, names, tables, printing });
  return { formatting, structure, headerResize, merges, painter, clipboard, fillHandle, rangeMove, sortFilter, names, tables, menu, printing };
}
