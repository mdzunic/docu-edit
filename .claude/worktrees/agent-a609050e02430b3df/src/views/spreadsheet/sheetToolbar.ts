// The sheet formatting strip's contract (SPEC-123 "The toolbar"): the state
// SheetFormatToolbar.vue (and the Format menu, the cell context menu and the
// shortcuts) render from, the commands they hand useFormatting, and the pure
// helpers that build the state — everything here reads the ACTIVE cell's
// effective style (cell → row → column → 0, SPEC-117), never a scan of the
// selection, which can be a million cells (the stated deviation from
// SPEC-061's mixed state; Excel and Sheets behave the same way).
import { DEFAULT_THEME, type DocumentTheme, type ModelColor } from "../../lib/core/theme";
import { borderShorthand, cellCss, type CssDeclarations } from "../../lib/spreadsheet/cellRender";
import type { BorderPreset } from "../../lib/spreadsheet/formatting";
import type { CellStyle, CellValue, SheetBorderStyle, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { describeNumberCode, NUMBER_CATEGORIES, NUMBER_FORMAT_PRESETS, presetIdOf, presetNumberFormat, type NumberPresetId } from "../../lib/spreadsheet/numberCodes";
import { formatCodeOf } from "../../lib/spreadsheet/numberFormat/builtin";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { formatValue, plainText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { BUILTIN_CELL_STYLES, cellColorHex, cellFontName, CELL_STYLE_GROUPS, CUSTOM_CELL_STYLE_GROUP, resolveNamedStyle } from "../../lib/spreadsheet/styles";

export type FormatCellsTab = "number" | "alignment" | "font" | "border" | "fill" | "protection";

/** What the strip, the Format menu, the context menu and the shortcuts ask
 * useFormatting to do. */
export type SheetToolbarCommand =
  | { kind: "font-family"; family: string }
  | { kind: "font-size"; sizePt: number }
  | { kind: "toggle"; format: "bold" | "italic" | "underline" | "strikethrough" }
  | { kind: "text-color"; color: ModelColor }
  | { kind: "fill"; color: ModelColor | null }
  | { kind: "fill-last" }
  | { kind: "border"; preset: BorderPreset }
  | { kind: "border-last" }
  | { kind: "border-line-style"; style: SheetBorderStyle }
  | { kind: "border-line-color"; color: ModelColor | null }
  | { kind: "align"; horizontal: "left" | "center" | "right" }
  | { kind: "valign"; vertical: "top" | "center" | "bottom" }
  | { kind: "wrap" }
  | { kind: "indent"; delta: 1 | -1 }
  | { kind: "rotation"; degrees: number }
  | { kind: "format-cells"; tab?: FormatCellsTab }
  | { kind: "number-format"; preset: NumberPresetId }
  | { kind: "currency" }
  | { kind: "percent" }
  | { kind: "comma" }
  | { kind: "decimals"; delta: 1 | -1 }
  | { kind: "cell-style"; name: string }
  | { kind: "clear"; what: "all" | "formats" | "contents" };

/** One number-format dropdown item: its id, label and the active cell's value
 * in that format. */
export interface NumberFormatItem {
  id: NumberPresetId;
  label: string;
  preview: string;
}

/** One gallery entry, drawn in its own style. */
export interface CellStyleEntry {
  name: string;
  box: CssDeclarations;
  text: CssDeclarations;
}

export interface CellStyleGroup {
  label: string;
  styles: CellStyleEntry[];
}

export interface SheetToolbarState {
  /** No grid sheet on show, or a cell is being edited. */
  disabled: boolean;
  fontFamily: string;
  fontSizePt: number;
  documentFonts: string[];
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  /** The active cell's text colour as hex; "" for automatic. */
  textColor: string;
  theme: DocumentTheme;
  recentColors: string[];
  /** The fill the fill button applies (the last used) as hex; "" = none. */
  fillColor: string;
  borderLineStyle: SheetBorderStyle;
  /** "" = automatic. */
  borderLineColor: string;
  horizontal: "left" | "center" | "right" | null;
  vertical: "top" | "center" | "bottom";
  wrap: boolean;
  textRotation: number;
  indent: number;
  /** The active cell's category name ("General", "Currency", "Custom" …). */
  numberFormatLabel: string;
  numberFormats: NumberFormatItem[];
  cellStyles: CellStyleGroup[];
}

/** The families the workbook's styles name, resolved, first-seen order. */
export function documentFonts(model: Pick<SpreadsheetModel, "styles" | "theme">): string[] {
  const theme = model.theme ?? DEFAULT_THEME;
  const families = new Set<string>();
  for (const style of model.styles) {
    if (style.font?.family !== undefined) {
      families.add(cellFontName(style.font.family, theme));
    }
  }
  return [...families];
}

/** The dropdown's label for a style's format: the preset it matches
 * exactly, else its category. */
export function numberFormatLabel(style: CellStyle, locale: SpreadsheetLocale): string {
  // An ods data style reads a currency symbol back bare ("$#,##0.00"): the
  // same code as the quoted spelling the presets write.
  const quoted = style.numberFormat?.replace(/(^|;)([$€£¥])(?=[#0])/g, '$1"$2"');
  const preset = presetIdOf(style, locale) ?? (quoted === style.numberFormat ? null : presetIdOf({ ...style, numberFormat: quoted }, locale));
  if (preset !== null) {
    return NUMBER_FORMAT_PRESETS.find((entry) => entry.id === preset)!.label;
  }
  const category = describeNumberCode(quoted, locale).category;
  return NUMBER_CATEGORIES.find((entry) => entry.id === category)?.label ?? "Custom";
}

/** The dropdown's items, each previewing `value` in its format. */
export function numberFormatItems(value: CellValue, locale: SpreadsheetLocale, ctx: FormatContext): NumberFormatItem[] {
  return NUMBER_FORMAT_PRESETS.map(({ id, label }) => {
    const choice = presetNumberFormat(id, locale);
    const code = choice.numberFormat === null ? undefined : formatCodeOf({ numberFormat: choice.numberFormat, builtinFormatId: choice.builtinFormatId }, locale);
    return { id, label, preview: value === null ? "" : plainText(formatValue(value, code, ctx)) };
  });
}

/** A style's look as gallery CSS (the cell's box and text, its bottom and
 * top borders when it has them). */
function previewCss(style: CellStyle, theme: DocumentTheme): { box: CssDeclarations; text: CssDeclarations } {
  const { box, text } = cellCss(style, theme, "string");
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const edge = style.borders?.[side];
    if (edge) {
      box[`border-${side}`] = borderShorthand(edge.style, cellColorHex(edge.color, theme) || "#000000");
    }
  }
  return { box, text };
}

/** The gallery: the built-ins in Excel's groups, each in the workbook's
 * definition when it has one, then the workbook's own (unhidden) names. */
export function cellStyleGroups(model: Pick<SpreadsheetModel, "styles" | "theme" | "namedStyles">): CellStyleGroup[] {
  const theme = model.theme ?? DEFAULT_THEME;
  const normal = model.styles[0] ?? {};
  const entry = (name: string): CellStyleEntry | null => {
    const named = resolveNamedStyle(model, name);
    if (!named) {
      return null;
    }
    // Drawn as applied: the definition over Normal (a font keeps Normal's
    // missing keys).
    const style: CellStyle = { ...normal, ...named.style, font: { ...normal.font, ...named.style.font } };
    return { name, ...previewCss(style, theme) };
  };
  const groups = CELL_STYLE_GROUPS.map((group) => ({
    label: group.label,
    styles: group.names.map(entry).filter((each): each is CellStyleEntry => each !== null),
  }));
  const builtinIds = new Set(BUILTIN_CELL_STYLES.map((style) => style.builtinId));
  const builtinNames = new Set(BUILTIN_CELL_STYLES.map((style) => style.name));
  const custom = (model.namedStyles ?? [])
    .filter((style) => !style.hidden && style.builtinId === undefined && !builtinNames.has(style.name) && !builtinIds.has(style.builtinId ?? -1))
    .map((style) => entry(style.name))
    .filter((each): each is CellStyleEntry => each !== null);
  if (custom.length > 0) {
    groups.push({ label: CUSTOM_CELL_STYLE_GROUP, styles: custom });
  }
  return groups;
}
