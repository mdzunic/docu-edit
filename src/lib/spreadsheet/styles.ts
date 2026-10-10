// Cell-style resolution (SPEC-117 "Model additions"): which style a cell
// shows and the legacy indexed palette both readers and renderers share.
// Style resolution is ONE function, effectiveStyleId — the cell's own
// `styleId`, else its row's (present only when the source row carried a
// format), else the column span covering it, else the default style 0 —
// each step a binary search over sparse, sorted data, never a scan.
import { resolveFont, type DocumentTheme } from "../core/theme";
import { findRow, getCell } from "./address";
import { maxDigitWidthPx } from "./geometry";
import type { CellBorderEdge, CellBorders, CellFont, CellStyle, ColumnSpan, Sheet, SpreadsheetModel } from "./model";

/** The default style table: one empty ("Normal") style. */
export function defaultStyles(): CellStyle[] {
  return [{}];
}

/** ECMA-376 §18.8.27's 64-entry legacy palette (lowercase), the colour an
 * `indexed="n"` attribute names unless the file overrides it with
 * `<colors><indexedColors>`. 64 is "system foreground" (automatic) and 65
 * "system background" — the readers map those, not this table. */
export const INDEXED_COLORS: readonly string[] = [
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080",
  "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff",
  "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff",
  "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99",
  "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/** The column span covering `col` — a binary search over the sorted,
 * non-overlapping spans; undefined when the column has the sheet default. */
export function columnSpanAt(columns: readonly ColumnSpan[] | undefined, col: number): ColumnSpan | undefined {
  if (!columns) {
    return undefined;
  }
  let lo = 0;
  let hi = columns.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const span = columns[mid]!;
    if (col < span.min) {
      hi = mid - 1;
    } else if (col > span.max) {
      lo = mid + 1;
    } else {
      return span;
    }
  }
  return undefined;
}

/** The style a position shows: the cell's `styleId`, else its row's
 * `styleId`, else the covering column span's, else 0. */
export function effectiveStyleId(sheet: Sheet, row: number, col: number): number {
  const cellStyleId = getCell(sheet, row, col)?.styleId;
  if (cellStyleId !== undefined) {
    return cellStyleId;
  }
  const rowStyleId = findRow(sheet, row)?.styleId;
  if (rowStyleId !== undefined) {
    return rowStyleId;
  }
  return columnSpanAt(sheet.columns, col)?.styleId ?? 0;
}

/** `model.styles[id]`, or the default style for an id out of range. */
export function cellStyle(model: Pick<SpreadsheetModel, "styles">, id: number): CellStyle {
  const styles = model.styles;
  if (!styles || styles.length === 0) {
    return {};
  }
  return (Number.isInteger(id) && id >= 0 ? styles[id] : undefined) ?? styles[0]!;
}

// --- SPEC-122: interning (THE normative source for internStyle) -----------------------------
// internStyle(model, style): the key is the style as JSON with keys sorted
// (at every level) and `sourceXf`/`sourceStyleName` omitted, an absent and an
// undefined field alike; the key → index table is cached in a WeakMap keyed
// by the `styles` array (arrays are replaced, never mutated, so a table never
// goes stale). An equal entry's index is returned — a source-backed one
// (`sourceXf`/`sourceStyleName`) first, else the first equal one — else
// `model.styles = [...model.styles, style]` and the new index. An appended
// entry never carries `sourceXf`/`sourceStyleName`: a style equal to no
// entry is not the source's, so a style passed with them (a source style
// copied and changed) is appended without them, and the writers never
// mistake it for the source entry. Entries are immutable once interned.

/** A style's interning key: the style as JSON with keys sorted at every
 * level, `sourceXf`/`sourceStyleName` omitted (where a style was read from
 * says nothing about how it looks). */
export function styleKey(style: CellStyle): string {
  const { sourceXf: _xf, sourceStyleName: _name, ...rest } = style;
  return stableJson(rest);
}

function stableJson(value: unknown): string {
  if (Array.isArray(value)) {
    return `[${value.map((item) => (item === undefined ? "null" : stableJson(item))).join(",")}]`;
  }
  if (value !== null && typeof value === "object") {
    const entries = Object.keys(value)
      .filter((key) => (value as Record<string, unknown>)[key] !== undefined)
      .toSorted()
      .map((key) => `${JSON.stringify(key)}:${stableJson((value as Record<string, unknown>)[key])}`);
    return `{${entries.join(",")}}`;
  }
  return JSON.stringify(value);
}

function isSourceBacked(style: CellStyle): boolean {
  return style.sourceXf !== undefined || style.sourceStyleName !== undefined;
}

/** Each `styles` array's key → index, built once per array (arrays are
 * replaced, never mutated, so an array's table never goes stale). */
const internTables = new WeakMap<readonly CellStyle[], Map<string, number>>();

function internTable(styles: readonly CellStyle[]): Map<string, number> {
  let table = internTables.get(styles);
  if (!table) {
    table = new Map();
    styles.forEach((style, index) => {
      const key = styleKey(style);
      const found = table!.get(key);
      // The first equal entry wins, a source-backed one over an appended one.
      if (found === undefined || (!isSourceBacked(styles[found]!) && isSourceBacked(style))) {
        table!.set(key, index);
      }
    });
    internTables.set(styles, table);
  }
  return table;
}

/** The index of a style structurally equal to `style` (SPEC-122 "Styles"):
 * an existing entry's — a source-backed one first — else `style` is appended
 * copy-on-write (`model.styles` replaced, never mutated; a new entry carries
 * no `sourceXf`/`sourceStyleName`, since no source wrote it) and its index
 * returned. Entries are immutable once interned. */
export function internStyle(model: Pick<SpreadsheetModel, "styles">, style: CellStyle): number {
  const styles = model.styles;
  const table = internTable(styles);
  const key = styleKey(style);
  const found = table.get(key);
  if (found !== undefined) {
    return found;
  }
  let fresh = style;
  if (isSourceBacked(style)) {
    const { sourceXf: _xf, sourceStyleName: _name, ...rest } = style;
    fresh = rest;
  }
  const next = [...styles, fresh];
  const nextTable = new Map(table);
  nextTable.set(key, next.length - 1);
  internTables.set(next, nextTable);
  model.styles = next;
  return next.length - 1;
}

/** The maximum digit width of a cell font (the pinned MDW table — never a
 * measurement), for the text-width estimates of overflow, shrink-to-fit and
 * fill: its family resolved against the theme (the theme's minor font when
 * the font names none), its size 11 pt when it states none. */
export function fontMaxDigitWidth(font: CellFont | undefined, theme: DocumentTheme): number {
  const family = font?.family !== undefined ? resolveFont(font.family, theme) : theme.fonts.minor;
  return maxDigitWidthPx(family, font?.sizePt ?? 11);
}

// --- SPEC-123: named cell styles (THE normative source for NamedCellStyle and
// BUILTIN_CELL_STYLES) ----------------------------------------------------------
// A named style ("Good", "Heading 1", a workbook's own "Invoice total") is a
// definition — only the aspects it includes — that a cell style names through
// `CellStyle.namedStyle`. `model.namedStyles` holds the workbook's own, in
// file order (xlsx `cellStyleXfs` + `cellStyles`, ods `office:styles` common
// table-cell styles — xlsx/styles.ts and ods/styles.ts map the wire both
// ways). A built-in the workbook defines (xlsx by `builtinId` or name, ods by
// display name) renders and applies with the workbook's definition; one it
// lacks renders with the pinned definition below — Office-theme colours as
// ThemeColorRefs, so it follows the workbook's theme — and is appended to
// `namedStyles` the first time it is applied (formatting.ts), after which the
// writers add it to the file. Applying sets, for each included aspect, exactly
// the definition's value (a font wholesale, its missing keys Normal's).

/** The six aspects a named style can include (xlsx's `applyX` flags). */
export type StyleAspect = "numberFormat" | "alignment" | "font" | "borders" | "fill" | "protection";

/** Every aspect, in xlsx's `applyX` order. */
export const STYLE_ASPECTS: readonly StyleAspect[] = ["numberFormat", "font", "fill", "borders", "alignment", "protection"];

export interface NamedCellStyle {
  /** The display name ("Good", "Heading 1", "Invoice total"). */
  name: string;
  /** xlsx `cellStyle@builtinId`. */
  builtinId?: number;
  /** The definition: only the included aspects; never `namedStyle`,
   * `sourceXf` or `sourceStyleName`. */
  style: CellStyle;
  includes: StyleAspect[];
  /** xlsx `cellStyle@hidden` — kept, never shown in the gallery. */
  hidden?: true;
  /** xlsx: the `cellStyleXfs` index the definition was read from. */
  sourceXfId?: number;
  /** ods: the common style's `style:name` ("Heading_20_1"). */
  sourceStyleName?: string;
}

/** One pinned built-in: this product's definition for a workbook that lacks
 * the style. Normal's definition is the workbook's own `styles[0]`. */
export interface BuiltinCellStyle {
  name: string;
  builtinId: number;
  includes: StyleAspect[];
  style: CellStyle;
}

const DK1 = { theme: "dk1" } as const;
const DK2 = { theme: "dk2" } as const;
const LT1 = { theme: "lt1" } as const;
const ACCENT1 = { theme: "accent1" } as const;

/** All four outer edges in one line style and colour. */
function allEdges(edge: CellBorderEdge): CellBorders {
  return { top: { ...edge }, right: { ...edge }, bottom: { ...edge }, left: { ...edge } };
}

/** The eighteen built-ins the gallery offers (SPEC-123 "Named cell styles"). */
export const BUILTIN_CELL_STYLES: readonly BuiltinCellStyle[] = [
  { name: "Normal", builtinId: 0, includes: [...STYLE_ASPECTS], style: {} },
  { name: "Good", builtinId: 26, includes: ["font", "fill"], style: { font: { color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } } },
  { name: "Bad", builtinId: 27, includes: ["font", "fill"], style: { font: { color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" } } },
  { name: "Neutral", builtinId: 28, includes: ["font", "fill"], style: { font: { color: "#9c5700" }, fill: { kind: "solid", color: "#ffeb9c" } } },
  { name: "Title", builtinId: 15, includes: ["font"], style: { font: { family: { theme: "major" }, sizePt: 18, color: DK2 } } },
  {
    name: "Heading 1",
    builtinId: 16,
    includes: ["font", "borders"],
    style: { font: { sizePt: 15, bold: true, color: DK2 }, borders: { bottom: { style: "thick", color: ACCENT1 } } },
  },
  {
    name: "Heading 2",
    builtinId: 17,
    includes: ["font", "borders"],
    style: { font: { sizePt: 13, bold: true, color: DK2 }, borders: { bottom: { style: "thick", color: { theme: "accent1", lumMod: 0.5, lumOff: 0.5 } } } },
  },
  {
    name: "Heading 3",
    builtinId: 18,
    includes: ["font", "borders"],
    style: { font: { bold: true, color: DK2 }, borders: { bottom: { style: "medium", color: { theme: "accent1", lumMod: 0.6, lumOff: 0.4 } } } },
  },
  { name: "Heading 4", builtinId: 19, includes: ["font"], style: { font: { bold: true, color: DK2 } } },
  {
    name: "Total",
    builtinId: 25,
    includes: ["font", "borders"],
    style: { font: { bold: true, color: DK1 }, borders: { top: { style: "thin", color: ACCENT1 }, bottom: { style: "double", color: ACCENT1 } } },
  },
  {
    name: "Input",
    builtinId: 20,
    includes: ["font", "fill", "borders"],
    style: { font: { color: "#3f3f76" }, fill: { kind: "solid", color: "#ffcc99" }, borders: allEdges({ style: "thin", color: "#7f7f7f" }) },
  },
  {
    name: "Output",
    builtinId: 21,
    includes: ["font", "fill", "borders"],
    style: { font: { bold: true, color: "#3f3f3f" }, fill: { kind: "solid", color: "#f2f2f2" }, borders: allEdges({ style: "thin", color: "#3f3f3f" }) },
  },
  {
    name: "Calculation",
    builtinId: 22,
    includes: ["font", "fill", "borders"],
    style: { font: { bold: true, color: "#fa7d00" }, fill: { kind: "solid", color: "#f2f2f2" }, borders: allEdges({ style: "thin", color: "#7f7f7f" }) },
  },
  {
    name: "Check Cell",
    builtinId: 23,
    includes: ["font", "fill", "borders"],
    style: { font: { bold: true, color: LT1 }, fill: { kind: "solid", color: "#a5a5a5" }, borders: allEdges({ style: "double", color: "#3f3f3f" }) },
  },
  { name: "Note", builtinId: 10, includes: ["fill", "borders"], style: { fill: { kind: "solid", color: "#ffffcc" }, borders: allEdges({ style: "thin", color: "#b2b2b2" }) } },
  { name: "Warning Text", builtinId: 11, includes: ["font"], style: { font: { color: "#ff0000" } } },
  { name: "Explanatory Text", builtinId: 53, includes: ["font"], style: { font: { italic: true, color: "#7f7f7f" } } },
  {
    name: "Linked Cell",
    builtinId: 24,
    includes: ["font", "borders"],
    style: { font: { color: "#fa7d00" }, borders: { bottom: { style: "double", color: "#ff8001" } } },
  },
];

/** The gallery's groups, as Excel groups the built-ins; workbook-defined
 * names follow in CUSTOM_CELL_STYLE_GROUP. */
export const CELL_STYLE_GROUPS: readonly { label: string; names: readonly string[] }[] = [
  { label: "Good, Bad and Neutral", names: ["Normal", "Bad", "Good", "Neutral"] },
  {
    label: "Data and model",
    names: ["Calculation", "Check Cell", "Explanatory Text", "Input", "Linked Cell", "Note", "Output", "Warning Text"],
  },
  { label: "Titles and headings", names: ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Title", "Total"] },
];

export const CUSTOM_CELL_STYLE_GROUP = "Custom";

/** The pinned built-in called `name`, or undefined. */
export function builtinCellStyle(name: string): BuiltinCellStyle | undefined {
  return BUILTIN_CELL_STYLES.find((entry) => entry.name === name);
}

/** The pinned built-in with an xlsx `builtinId`, or undefined. */
export function builtinCellStyleById(builtinId: number): BuiltinCellStyle | undefined {
  return BUILTIN_CELL_STYLES.find((entry) => entry.builtinId === builtinId);
}

/** The workbook's own definition of `name`: for a built-in name, the entry
 * with its `builtinId` (an xlsx file names "Good" whatever its language) or
 * its name; else the entry with that (display) name. */
export function workbookNamedStyle(model: Pick<SpreadsheetModel, "namedStyles">, name: string): NamedCellStyle | undefined {
  const styles = model.namedStyles ?? [];
  const builtin = builtinCellStyle(name);
  return (builtin !== undefined ? styles.find((entry) => entry.builtinId === builtin.builtinId) : undefined) ?? styles.find((entry) => entry.name === name);
}

/** A named style as it renders and applies: the workbook's definition when
 * it has one, else the pinned built-in's (Normal's is `styles[0]`). */
export interface ResolvedNamedStyle {
  name: string;
  builtinId?: number;
  style: CellStyle;
  includes: StyleAspect[];
  /** The workbook defines it (in `namedStyles`). */
  defined: boolean;
}

/** The definition `name` renders and applies with (see ResolvedNamedStyle);
 * undefined for a name neither the workbook nor the built-ins know. */
export function resolveNamedStyle(model: Pick<SpreadsheetModel, "namedStyles" | "styles">, name: string): ResolvedNamedStyle | undefined {
  const own = workbookNamedStyle(model, name);
  const builtin = builtinCellStyle(name);
  if (name === "Normal" || own?.builtinId === 0 || builtin?.builtinId === 0) {
    const { namedStyle: _named, sourceXf: _xf, sourceStyleName: _name, quotePrefix: _quote, ...normal } = cellStyle(model, 0);
    return { name: "Normal", builtinId: 0, style: normal, includes: [...STYLE_ASPECTS], defined: own !== undefined };
  }
  if (own) {
    return { name, ...(own.builtinId !== undefined ? { builtinId: own.builtinId } : {}), style: own.style, includes: [...own.includes], defined: true };
  }
  if (builtin) {
    return { name, builtinId: builtin.builtinId, style: structuredClone(builtin.style), includes: [...builtin.includes], defined: false };
  }
  return undefined;
}
