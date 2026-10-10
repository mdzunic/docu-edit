// Cell formatting in the workbook view (SPEC-123): the commands the sheet
// formatting strip, the Format menu, the cell context menu and the
// SHEET_SHORTCUTS run — each ONE undo step through SPEC-121's record/commit
// (labels "Bold", "Fill colour", "Borders", "Number format", "Format cells",
// "Cell style", "Clear formats", "Clear all"; "Clear contents" is SPEC-122's
// Delete, useCellEditing's) over lib/spreadsheet/formatting.ts, written
// through by the edit pipeline. A refused command (protection, the
// materialisation cap) records nothing and shows the refusal as a toast.
//
// - State is the ACTIVE cell's effective style (sheetToolbar.ts): a toggle
//   sets its property on the whole selection when the active cell lacks it
//   and removes it when the active cell has it; indent, rotation and decimal
//   stepping are computed from the active cell and applied to the selection.
// - The view session keeps the fill button's last fill (yellow until one is
//   chosen), the borders button's last preset (Bottom) and the border line
//   style and colour (thin, automatic); recent colours are SPEC-013's.
// - Shortcuts (sheetShortcuts.ts) act only while focus is in the grid and no
//   cell is being edited — never in the formula bar or a dialog — and run on
//   the window's capture phase so a layout where a sheet binding is also a
//   text-registry key (Swiss Ctrl+Shift+1 is "+") reaches this table first.
// - The cell context menu opens on a right-click in the grid, Shift+F10 or
//   the Menu key; Format Cells (useFormatCellsDialog) on Ctrl/Cmd+1, the
//   Format menu, the context menu and "More formats…".
// - SPEC-124: a header's right-click, and Shift+F10 / the Menu key on whole
//   rows or columns, are the header menu's (useStructure), as are the
//   sheet.insertCells / sheet.deleteCells keys. A command changing a font
//   family or size, wrap or rotation (a patch, Format Cells, a cell style,
//   Clear formats/all) runs SPEC-122's refitRows inside its step.
import { computed, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { loadRecentColors, pushRecentColor } from "../../lib/core/colors";
import { isMacPlatform } from "../../lib/core/platform";
import { DEFAULT_THEME, type ModelColor } from "../../lib/core/theme";
import { findShortcut } from "../../lib/shortcuts";
import { getCell } from "../../lib/spreadsheet/address";
import { refitRows, type CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import {
  applyNamedStyle,
  applyStylePatch,
  applyStylePatches,
  borderPresetPatches,
  clearRange,
  commaButtonFormat,
  currencyButtonFormat,
  numberFormatChoiceOf,
  percentButtonFormat,
  presetNumberFormat,
  shortcutNumberFormat,
  stepDecimals,
  type BorderPreset,
  type FormatResult,
  type NumberFormatChoice,
  type StylePatch,
  type StylePatchEntry,
  patchChangesRowHeight,
} from "../../lib/spreadsheet/formatting";
import type { CellBorderEdge, Sheet, SheetBorderStyle, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { formatCodeOf } from "../../lib/spreadsheet/numberFormat/builtin";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import type { GridRange } from "../../lib/spreadsheet/address";
import { cellColorHex, cellFontName, cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";
import type { CellEditing } from "./useCellEditing";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { headerAt, inGrid, selectionAxis } from "./headerMenu";
import { isFormattingShortcut, NUMBER_SHORTCUTS, SHEET_SHORTCUTS, TOGGLE_LABELS, TOGGLE_SHORTCUTS } from "./sheetShortcuts";
import { cellStyleGroups, documentFonts, numberFormatItems, numberFormatLabel, type SheetToolbarCommand, type SheetToolbarState } from "./sheetToolbar";
import { useFormatCellsDialog } from "./useFormatCellsDialog";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface FormattingDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection" | "formatContext">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing" | "clearSelection"> };
  locale: () => SpreadsheetLocale;
  /** SPEC-124: row refit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
}

const DEFAULT_FILL = "#ffff00";
const FONT_SIZE_RANGE = [1, 409] as const;
const MAX_INDENT = 250;


export function useFormatting(deps: FormattingDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  const lastFill = shallowRef<ModelColor | null>(DEFAULT_FILL);
  const lastBorder = ref<BorderPreset>("bottom");
  const lineStyle = ref<SheetBorderStyle>("thin");
  const lineColor = shallowRef<ModelColor | null>(null);
  let measurer = deps.measurer;
  const recentColors = ref<string[]>(loadRecentColors());
  /** The cell context menu's position (null: closed). */
  const contextMenu = ref<{ x: number; y: number } | null>(null);

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const theme = computed(() => model()?.theme ?? DEFAULT_THEME);

  /** The active cell: its effective style and value. */
  const active = computed(() => {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    if (!m || !sheet || sheet.kind !== "grid") {
      return null;
    }
    const at = deps.selection.selection.value.active;
    return { style: cellStyle(m, effectiveStyleId(sheet, at.row, at.col)), value: getCell(sheet, at.row, at.col)?.value ?? null };
  });

  /** Runs one formatting command as ONE undo step (see the module header);
   * `refit` (SPEC-124): the command can change a row's height, so SPEC-122's
   * refitRows runs inside the step over the rows it touched. */
  function command(label: string, mutate: (m: SpreadsheetModel, sheetId: string, ranges: GridRange[]) => FormatResult, refit = false): void {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    if (!m || !sheet || sheet.kind !== "grid" || deps.editing.cell.editing.value) {
      return;
    }
    deps.viewState.syncInto(m);
    deps.editing.undo.record(label, { sheetId: sheet.id });
    const ranges = deps.selection.selection.value.ranges;
    const before = m.sheets.find((each) => each.id === sheet.id)?.rows;
    const result = mutate(m, sheet.id, ranges);
    if (result.ok && refit) {
      measurer ??= createCanvasCellMeasurer();
      refitRows(m, sheet.id, ranges, measurer, deps.selection.formatContext.value, before);
    }
    deps.editing.undo.commit();
    if (!result.ok) {
      toast(result.message);
    }
  }

  const patch = (label: string, value: StylePatch): void => command(label, (m, id, ranges) => applyStylePatch(m, id, ranges, value), patchChangesRowHeight(value));
  const numberFormat = (choice: NumberFormatChoice): void => patch("Number format", { numberFormat: choice.numberFormat, builtinFormatId: choice.builtinFormatId ?? null });

  function remember(color: ModelColor | null): void {
    if (typeof color === "string") {
      recentColors.value = pushRecentColor(color, recentColors.value);
    }
  }

  function borders(preset: BorderPreset): void {
    lastBorder.value = preset;
    const edge: CellBorderEdge = { style: lineStyle.value };
    if (lineColor.value !== null) {
      edge.color = lineColor.value;
    }
    command("Borders", (m, id, ranges) => applyStylePatches(m, id, ranges.flatMap((range): StylePatchEntry[] => borderPresetPatches(range, preset, edge))));
  }

  const dialog = useFormatCellsDialog({
    target: () => {
      const current = active.value;
      const m = model();
      if (!current || !m) {
        return null;
      }
      return {
        style: current.style,
        value: current.value,
        locale: deps.locale(),
        format: deps.selection.formatContext.value,
        theme: theme.value,
        documentFonts: documentFonts(m),
        styles: m.styles,
        ranges: deps.selection.selection.value.ranges,
        recentColors: recentColors.value,
      };
    },
    apply: (entries) => command("Format cells", (m, id) => applyStylePatches(m, id, entries), entries.some((entry) => patchChangesRowHeight(entry.patch))),
  });

  function run(cmd: SheetToolbarCommand): void {
    const style = active.value?.style ?? {};
    const font = style.font ?? {};
    const alignment = style.alignment ?? {};
    switch (cmd.kind) {
      case "toggle": {
        const on = cmd.format === "underline" ? font.underline !== undefined : cmd.format === "strikethrough" ? font.strike === true : font[cmd.format] === true;
        const key = cmd.format === "strikethrough" ? "strike" : cmd.format;
        patch(TOGGLE_LABELS[cmd.format], { font: { [key]: on ? null : cmd.format === "underline" ? "single" : true } });
        break;
      }
      case "font-family":
        patch("Font", { font: { family: cmd.family } });
        break;
      case "font-size":
        patch("Font size", { font: { sizePt: Math.min(FONT_SIZE_RANGE[1], Math.max(FONT_SIZE_RANGE[0], cmd.sizePt)) } });
        break;
      case "text-color":
        remember(cmd.color);
        patch("Text colour", { font: { color: cmd.color } });
        break;
      case "fill":
        if (cmd.color !== null) {
          lastFill.value = cmd.color;
        }
        remember(cmd.color);
        patch("Fill colour", { fill: cmd.color === null ? null : { kind: "solid", color: cmd.color } });
        break;
      case "fill-last":
        run({ kind: "fill", color: lastFill.value });
        break;
      case "border":
        borders(cmd.preset);
        break;
      case "border-last":
        borders(lastBorder.value);
        break;
      case "border-line-style":
        lineStyle.value = cmd.style;
        break;
      case "border-line-color":
        lineColor.value = cmd.color;
        remember(cmd.color);
        break;
      case "align":
        patch("Alignment", { alignment: { horizontal: alignment.horizontal === cmd.horizontal ? null : cmd.horizontal } });
        break;
      case "valign":
        patch("Alignment", { alignment: { vertical: cmd.vertical === "bottom" ? null : cmd.vertical } });
        break;
      case "wrap":
        patch("Wrap text", { alignment: { wrap: alignment.wrap ? null : true } });
        break;
      case "indent": {
        const indent = Math.min(MAX_INDENT, Math.max(0, (alignment.indent ?? 0) + cmd.delta));
        const leftAlign = cmd.delta > 0 && (alignment.horizontal === undefined || alignment.horizontal === "general" || alignment.horizontal === "center");
        patch("Indent", { alignment: { indent: indent === 0 ? null : indent, ...(leftAlign ? { horizontal: "left" as const } : {}) } });
        break;
      }
      case "rotation":
        patch("Text rotation", { alignment: { textRotation: alignment.textRotation === cmd.degrees ? null : cmd.degrees } });
        break;
      case "format-cells":
        dialog.show(cmd.tab);
        break;
      case "number-format":
        numberFormat(presetNumberFormat(cmd.preset, deps.locale()));
        break;
      case "currency":
        numberFormat(currencyButtonFormat(deps.locale()));
        break;
      case "percent":
        numberFormat(percentButtonFormat());
        break;
      case "comma":
        numberFormat(commaButtonFormat(deps.locale()));
        break;
      case "decimals": {
        const code = formatCodeOf(style, deps.locale());
        const stepped = stepDecimals(code, active.value?.value ?? null, cmd.delta);
        if (stepped !== code) {
          numberFormat(numberFormatChoiceOf(stepped));
        }
        break;
      }
      case "cell-style":
        command("Cell style", (m, id, ranges) => applyNamedStyle(m, id, ranges, cmd.name), true);
        break;
      case "clear":
        if (cmd.what === "contents") {
          deps.editing.cell.clearSelection();
        } else {
          command(cmd.what === "all" ? "Clear all" : "Clear formats", (m, id, ranges) => clearRange(m, id, ranges, cmd.what), true);
        }
        break;
    }
  }

  const toolbar = computed((): SheetToolbarState => {
    const m = model();
    const current = active.value;
    const style = current?.style ?? {};
    const font = { ...m?.styles[0]?.font, ...style.font };
    const alignment = style.alignment ?? {};
    const locale = deps.locale();
    const horizontal = alignment.horizontal;
    const vertical = alignment.vertical;
    const hex = (color: ModelColor | null | undefined): string => cellColorHex(color, theme.value);
    return {
      disabled: current === null || deps.editing.cell.editing.value,
      fontFamily: cellFontName(font.family, theme.value),
      fontSizePt: font.sizePt ?? 11,
      documentFonts: m ? documentFonts(m) : [],
      bold: style.font?.bold === true,
      italic: style.font?.italic === true,
      underline: style.font?.underline !== undefined,
      strikethrough: style.font?.strike === true,
      textColor: hex(style.font?.color),
      theme: theme.value,
      recentColors: recentColors.value,
      fillColor: hex(lastFill.value),
      borderLineStyle: lineStyle.value,
      borderLineColor: hex(lineColor.value),
      horizontal: horizontal === "left" || horizontal === "center" || horizontal === "right" ? horizontal : null,
      vertical: vertical === "top" || vertical === "center" ? vertical : "bottom",
      wrap: alignment.wrap === true,
      textRotation: alignment.textRotation ?? 0,
      indent: alignment.indent ?? 0,
      numberFormatLabel: numberFormatLabel(style, locale),
      numberFormats: numberFormatItems(current?.value ?? null, locale, deps.selection.formatContext.value),
      cellStyles: m ? cellStyleGroups(m) : [],
    };
  });

  /** Opens the cell context menu at a point (a right-click), or under the
   * active cell (Shift+F10, the Menu key). */
  function openContextMenu(x?: number, y?: number): void {
    if (x === undefined || y === undefined) {
      const box = document.querySelector('[data-testid="cell-editor"]')?.getBoundingClientRect();
      contextMenu.value = { x: box?.left ?? 0, y: box?.bottom ?? 0 };
    } else {
      contextMenu.value = { x, y };
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || deps.editing.cell.editing.value || active.value === null) {
      return;
    }
    if (event.key === "ContextMenu" || (event.key === "F10" && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey)) {
      // SPEC-124: on whole rows or columns these keys open the header menu.
      if (selectionAxis(deps.selection.selection.value) !== null) {
        return;
      }
      event.preventDefault();
      openContextMenu();
      return;
    }
    const id = findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id;
    // Insert/delete (SPEC-124), paste special and fills (SPEC-125) are others'.
    if (id === undefined || !isFormattingShortcut(id)) {
      return;
    }
    event.preventDefault();
    if (id === "sheet.formatCells") {
      dialog.show();
    } else if (TOGGLE_SHORTCUTS[id]) {
      run({ kind: "toggle", format: TOGGLE_SHORTCUTS[id] });
    } else if (NUMBER_SHORTCUTS[id]) {
      numberFormat(shortcutNumberFormat(NUMBER_SHORTCUTS[id], deps.locale()));
    }
  }

  function onContextMenu(event: MouseEvent): void {
    // A header's right-click opens SPEC-124's header menu (useStructure).
    if (!inGrid(event.target) || active.value === null || headerAt(event.target) !== null) {
      return;
    }
    event.preventDefault();
    openContextMenu(event.clientX, event.clientY);
  }

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
    window.addEventListener("contextmenu", onContextMenu);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
      window.removeEventListener("contextmenu", onContextMenu);
    }
  });

  return {
    toolbar,
    run,
    dialog,
    contextMenu,
    openContextMenu,
    closeContextMenu: (): void => {
      contextMenu.value = null;
    },
    /** One undo step for another formatting surface (the format painter). */
    command,
    onKeydown,
  };
}

export type Formatting = ReturnType<typeof useFormatting>;

/** Narrows a component's runtime-typed `formatting` prop (the .vue parser
 * bars casts). */
export function asFormatting(value: unknown): Formatting {
  return value as Formatting;
}
