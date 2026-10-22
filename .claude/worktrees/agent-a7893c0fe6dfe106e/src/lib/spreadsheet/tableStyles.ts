// Table styles (SPEC-133 "Table styles"): how a table's style renders — THE
// normative source for this product's renderings of Excel's 60 built-in
// table styles, which SPEC-133 pins as family templates so screenshots and
// tests are deterministic (Excel and LibreOffice draw their own from the
// name, and every name round-trips unchanged whatever is drawn here). The
// grid (useSheetCells), the gallery (TableStyleGallery.vue), Convert to range
// and the ODS table writer all take a cell's look from tableCellLook, so none
// can drift from another.
//
// - A look is ECMA-376's table-style model (§18.8.40 tableStyleElement): one
//   part per region — the whole table, the column and row stripes, the last
//   and first column, the header and totals rows — each a font, a fill and
//   lines, layered in that order (a later part's property wins). A part's
//   lines are relative to its region: `top`/`bottom`/`left`/`right` on the
//   region's outer edges, `horizontal`/`vertical` between its rows/columns.
//   So wholeTable's are the outline and the inner lines, headerRow's `bottom`
//   the line under the header, totalsRow's `top` the line over the totals.
// - The templates: within a family of seven, member 1 colours with dk1 and
//   members 2–7 with accent1–accent6 (C); T80/T60/T40 are C lightened 80, 60
//   and 40% (lumMod 0.2/0.4/0.6 with lumOff 0.8/0.6/0.4), W is lt1 and K
//   dk1. Every colour stays a theme reference, resolved by the renderer
//   against the workbook theme, so a look follows the theme and one look
//   serves every theme. Readings of the spec's table: Light 8–14's stripes
//   frame their band — top and bottom for a row stripe, left and right for a
//   column stripe; Medium 22–28's inner lines are thin T40 where 8–14's are
//   W; Dark 1–7's body, "C shaded 25%", is `lumMod 0.75` — Excel's "Darker
//   25%" (SpreadsheetML tint −0.25, which xlsx/themeColor.ts writes back as
//   a tint; a DrawingML `shade` could only be written as a literal rgb).
//   First/last column, when on, are bold in every look but the empty one.
// - Dark 8–11, unknown names and workbook-defined styles (custom table
//   styles are out of scope) take the neutral fallback; no name ("None") is
//   the empty look, which paints nothing.
// - Row stripes alternate over the VISIBLE data rows (rowIsHidden: hidden by
//   the user or by the filter), the first visible one the first stripe;
//   column stripes alternate by column from the table's first. Each `rows`
//   array's hidden rows per data body are cached against the array, which
//   is replaced, never mutated (model.ts invariant 1).
// - tableCellLook memoises one CellStyle per look and position class, so a
//   cell's look is an identity-stable object a renderer can key a cache by.
//   Looks and cell looks are shared: never mutate one.
// - styleOverLook lays a cell's own style over its look (SPEC-123: the look
//   sits beneath the cell's own style); SPEC-129's conditional formats go
//   over the result.
import type { DocumentTheme, ThemeColorRef, ThemeColorSlot } from "../core/theme";
import { MAX_ROWS, rowIsHidden } from "./model";
import type { CellBorderEdge, CellBorders, CellFill, CellFont, CellStyle, Row, Sheet, SheetBorderStyle, SheetTable } from "./model";

function numbered(prefix: string, count: number): string[] {
  return Array.from({ length: count }, (_, i) => `${prefix}${i + 1}`);
}

/** Excel's 60 built-in table styles, in its gallery order: Light 1–21,
 * Medium 1–28, Dark 1–11. */
export const BUILT_IN_TABLE_STYLE_NAMES: readonly string[] = [
  ...numbered("TableStyleLight", 21),
  ...numbered("TableStyleMedium", 28),
  ...numbered("TableStyleDark", 11),
];

/** The lines a part draws, relative to the region it covers: `top`/`bottom`/
 * `left`/`right` on the region's outer edges, `horizontal`/`vertical`
 * between its rows and its columns. */
export interface TableStyleLines {
  top?: CellBorderEdge;
  bottom?: CellBorderEdge;
  left?: CellBorderEdge;
  right?: CellBorderEdge;
  horizontal?: CellBorderEdge;
  vertical?: CellBorderEdge;
}

/** One part of a table style: what it paints over the parts beneath it. */
export interface TableStylePart {
  font?: CellFont;
  fill?: CellFill;
  lines?: TableStyleLines;
}

/** A table style's parts, in the order they layer — a later part's font
 * properties, fill and edges win over an earlier one's. */
export interface TableStyleLook {
  /** The whole `ref`, header and totals rows included: its outline and inner lines. */
  wholeTable?: TableStylePart;
  /** Banded columns: the data body's columns alternating from the table's first. */
  firstColumnStripe?: TableStylePart;
  secondColumnStripe?: TableStylePart;
  /** Banded rows: the visible data rows alternating from the first visible one. */
  firstRowStripe?: TableStylePart;
  secondRowStripe?: TableStylePart;
  /** The data body's last / first column, when the table's option is on. */
  lastColumn?: TableStylePart;
  firstColumn?: TableStylePart;
  headerRow?: TableStylePart;
  totalsRow?: TableStylePart;
}

// --- the templates ---------------------------------------------------------------------

type Family = "light1" | "light8" | "light15" | "medium1" | "medium8" | "medium15" | "medium22" | "dark1" | "neutral";

/** Each kind's families of seven, in name order, and how many it has. */
const KINDS: Readonly<Record<string, { families: readonly Family[]; count: number }>> = {
  Light: { families: ["light1", "light8", "light15"], count: 21 },
  Medium: { families: ["medium1", "medium8", "medium15", "medium22"], count: 28 },
  Dark: { families: ["dark1", "neutral"], count: 11 },
};

/** C for members 1–7 of a family. */
const MEMBER_COLORS: readonly ThemeColorSlot[] = ["dk1", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

const BUILT_IN_NAME = /^TableStyle(Light|Medium|Dark)([1-9]\d*)$/;

const W: ThemeColorRef = { theme: "lt1" };
const K: ThemeColorRef = { theme: "dk1" };
const BOLD: CellFont = { bold: true };

function line(style: SheetBorderStyle, color: ThemeColorRef): CellBorderEdge {
  return { style, color };
}

function solid(color: ThemeColorRef): CellFill {
  return { kind: "solid", color };
}

/** A look with first/last column bold — every look but the empty one. */
function withColumns(look: TableStyleLook): TableStyleLook {
  return { ...look, lastColumn: { font: BOLD }, firstColumn: { font: BOLD } };
}

/** Dark 8–11, unknown names and workbook-defined styles. */
const NEUTRAL_LOOK: TableStyleLook = withColumns({
  headerRow: { font: BOLD, lines: { bottom: line("thin", K) } },
  totalsRow: { font: BOLD },
});

/** No style ("None"). */
const NO_LOOK: TableStyleLook = {};

/** A family's look for its member coloured `slot` (the spec's table, row by row). */
function familyLook(family: Family, slot: ThemeColorSlot): TableStyleLook {
  const c: ThemeColorRef = { theme: slot };
  const t80: ThemeColorRef = { theme: slot, lumMod: 0.2, lumOff: 0.8 };
  const t60: ThemeColorRef = { theme: slot, lumMod: 0.4, lumOff: 0.6 };
  const t40: ThemeColorRef = { theme: slot, lumMod: 0.6, lumOff: 0.4 };
  const boldWhiteOn = (fill: ThemeColorRef): TableStylePart => ({ fill: solid(fill), font: { bold: true, color: W } });
  const boldOverDouble: TableStylePart = { font: BOLD, lines: { top: line("double", c) } };
  const boldUnderThin: TableStylePart = { font: BOLD, lines: { bottom: line("thin", c) } };
  switch (family) {
    case "light1":
      return withColumns({
        wholeTable: { lines: { top: line("thin", c), bottom: line("thin", c) } },
        firstColumnStripe: { fill: solid(t80) },
        firstRowStripe: { fill: solid(t80) },
        headerRow: boldUnderThin,
        totalsRow: boldOverDouble,
      });
    case "light8": {
      const thin = line("thin", c);
      return withColumns({
        wholeTable: { lines: { top: thin, bottom: thin, left: thin, right: thin } },
        firstColumnStripe: { lines: { left: thin, right: thin } },
        firstRowStripe: { lines: { top: thin, bottom: thin } },
        headerRow: boldWhiteOn(c),
        totalsRow: boldOverDouble,
      });
    }
    case "light15": {
      const thin = line("thin", c);
      return withColumns({
        wholeTable: { lines: { top: thin, bottom: thin, left: thin, right: thin, horizontal: thin, vertical: thin } },
        firstColumnStripe: { fill: solid(t80) },
        firstRowStripe: { fill: solid(t80) },
        headerRow: boldUnderThin,
        totalsRow: boldOverDouble,
      });
    }
    case "medium1":
    case "medium15": {
      const thin = line("thin", t40);
      return withColumns({
        wholeTable: { lines: { top: thin, bottom: thin, left: thin, right: thin, horizontal: thin } },
        firstColumnStripe: { fill: solid(t80) },
        firstRowStripe: { fill: solid(t80) },
        headerRow: boldWhiteOn(family === "medium15" ? K : c),
        totalsRow: boldOverDouble,
      });
    }
    case "medium8":
    case "medium22": {
      const inner = line("thin", family === "medium22" ? t40 : W);
      return withColumns({
        wholeTable: { fill: solid(t80), lines: { horizontal: inner, vertical: inner } },
        firstColumnStripe: { fill: solid(t60) },
        firstRowStripe: { fill: solid(t60) },
        headerRow: boldWhiteOn(c),
        totalsRow: boldWhiteOn(c),
      });
    }
    case "dark1":
      return withColumns({
        wholeTable: { fill: solid({ theme: slot, lumMod: 0.75 }), font: { color: W } },
        firstColumnStripe: { fill: solid(c) },
        firstRowStripe: { fill: solid(c) },
        headerRow: boldWhiteOn(K),
        totalsRow: boldWhiteOn(K),
      });
    default:
      return NEUTRAL_LOOK;
  }
}

/** The built-in looks generated so far, by name. */
const builtInLooks = new Map<string, TableStyleLook>();

/** How a table style looks (SPEC-133 "Table styles"): a built-in name's
 * family template, the neutral fallback for Dark 8–11 and any other name
 * (unknown or workbook-defined), the empty look for none (`undefined` or "").
 * Colours are theme references, so a look serves every theme: `_theme` is
 * part of SPEC-133's signature, not an input. Memoised per name; shared —
 * never mutate the result. */
export function tableStyleLook(name: string | undefined, _theme: DocumentTheme): TableStyleLook {
  if (name === undefined || name === "") {
    return NO_LOOK;
  }
  let look = builtInLooks.get(name);
  if (look === undefined) {
    const match = BUILT_IN_NAME.exec(name);
    const kind = match ? KINDS[match[1]!] : undefined;
    const n = match ? Number(match[2]) : 0;
    if (kind === undefined || n > kind.count) {
      return NEUTRAL_LOOK;
    }
    look = familyLook(kind.families[Math.floor((n - 1) / 7)]!, MEMBER_COLORS[(n - 1) % 7]!);
    builtInLooks.set(name, look);
  }
  return look;
}

// --- one cell's look ------------------------------------------------------------------

// The position facts a cell's look depends on, as the bits of its memo key.
const HEADER = 1 << 0;
const TOTALS = 1 << 1;
const COLUMN_STRIPE = 1 << 2; // a data row with banded columns on…
const SECOND_COLUMN = 1 << 3; // …in a second-stripe column
const ROW_STRIPE = 1 << 4; // a data row with banded rows on…
const SECOND_ROW = 1 << 5; // …that is a second stripe
const LAST_COLUMN = 1 << 6; // a data row's last column, the option on
const FIRST_COLUMN = 1 << 7;
const DATA_TOP = 1 << 8; // the data body's first / last row
const DATA_BOTTOM = 1 << 9;
const TABLE_TOP = 1 << 10; // on the table's outline
const TABLE_BOTTOM = 1 << 11;
const TABLE_LEFT = 1 << 12;
const TABLE_RIGHT = 1 << 13;

type Side = "top" | "right" | "bottom" | "left";
const SIDES: readonly Side[] = ["top", "right", "bottom", "left"];

/** Lays `part` over `style`; `outer` says which of the cell's edges lie on
 * the outer edge of the part's region (the rest take its inner lines). */
function layer(style: CellStyle, part: TableStylePart | undefined, outer: Readonly<Record<Side, boolean>>): void {
  if (part === undefined) {
    return;
  }
  if (part.font !== undefined) {
    style.font = { ...style.font, ...part.font };
  }
  if (part.fill !== undefined) {
    style.fill = part.fill;
  }
  const lines = part.lines;
  if (lines === undefined) {
    return;
  }
  for (const side of SIDES) {
    const edge = outer[side] ? lines[side] : lines[side === "top" || side === "bottom" ? "horizontal" : "vertical"];
    if (edge !== undefined) {
      const borders: CellBorders = { ...style.borders };
      borders[side] = edge;
      style.borders = borders;
    }
  }
}

/** The look of a cell whose position facts are `key` — undefined when the
 * look paints nothing there. */
function composeCellLook(look: TableStyleLook, key: number): CellStyle | undefined {
  const on = (bit: number): boolean => (key & bit) !== 0;
  const table = { top: on(TABLE_TOP), bottom: on(TABLE_BOTTOM), left: on(TABLE_LEFT), right: on(TABLE_RIGHT) };
  // A single data-body column (a column stripe, the first/last column) and a
  // single row across the table (a row stripe, the header, the totals).
  const column = { top: on(DATA_TOP), bottom: on(DATA_BOTTOM), left: true, right: true };
  const row = { top: true, bottom: true, left: table.left, right: table.right };
  const style: CellStyle = {};
  layer(style, look.wholeTable, table);
  if (on(COLUMN_STRIPE)) {
    layer(style, on(SECOND_COLUMN) ? look.secondColumnStripe : look.firstColumnStripe, column);
  }
  if (on(ROW_STRIPE)) {
    layer(style, on(SECOND_ROW) ? look.secondRowStripe : look.firstRowStripe, row);
  }
  if (on(LAST_COLUMN)) {
    layer(style, look.lastColumn, column);
  }
  if (on(FIRST_COLUMN)) {
    layer(style, look.firstColumn, column);
  }
  if (on(HEADER)) {
    layer(style, look.headerRow, row);
  }
  if (on(TOTALS)) {
    layer(style, look.totalsRow, row);
  }
  return style.font !== undefined || style.fill !== undefined || style.borders !== undefined ? style : undefined;
}

/** Each look's cell looks by position key (null: paints nothing). */
const cellLooks = new WeakMap<TableStyleLook, Map<number, CellStyle | null>>();

function memoisedCellLook(look: TableStyleLook, key: number): CellStyle | undefined {
  let byKey = cellLooks.get(look);
  if (byKey === undefined) {
    byKey = new Map();
    cellLooks.set(look, byKey);
  }
  let style = byKey.get(key);
  if (style === undefined) {
    style = composeCellLook(look, key) ?? null;
    byKey.set(key, style);
  }
  return style ?? undefined;
}

/** The first index in [0, length) that is not `before` the target — a binary
 * search over an ascending list. */
function lowerBound(length: number, before: (index: number) => boolean): number {
  let lo = 0;
  let hi = length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (before(mid)) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** Each `rows` array's hidden rows (ascending) per data body, keyed
 * `top * MAX_ROWS + bottom`. */
const hiddenRowCache = new WeakMap<readonly Row[], Map<number, readonly number[]>>();

function hiddenRowsIn(rows: readonly Row[], top: number, bottom: number): readonly number[] {
  let byBody = hiddenRowCache.get(rows);
  if (byBody === undefined) {
    byBody = new Map();
    hiddenRowCache.set(rows, byBody);
  }
  const key = top * MAX_ROWS + bottom;
  let hidden = byBody.get(key);
  if (hidden === undefined) {
    const found: number[] = [];
    for (let i = lowerBound(rows.length, (k) => rows[k]!.index < top); i < rows.length && rows[i]!.index <= bottom; i += 1) {
      if (rowIsHidden(rows[i])) {
        found.push(rows[i]!.index);
      }
    }
    hidden = found;
    byBody.set(key, hidden);
  }
  return hidden;
}

/** A data row's place among the data body's visible rows (0 = the first). */
function visibleOrdinal(rows: readonly Row[], top: number, bottom: number, row: number): number {
  const hidden = hiddenRowsIn(rows, top, bottom);
  return row - top - lowerBound(hidden.length, (i) => hidden[i]! < row);
}

/** The look of one cell of `table` (SPEC-133 "Table styles"): undefined
 * outside `table.ref`, or where its style paints nothing. The header row
 * takes the header part, the totals row the totals part, a data row the
 * stripes (rows alternating over the visible data rows, columns by column)
 * and the first/last column parts the table's options turn on — each over
 * wholeTable, with the edges and inner lines of every part. Memoised per
 * style and position class: the same position class is the same object. */
export function tableCellLook(
  sheet: Pick<Sheet, "rows">,
  table: Pick<SheetTable, "ref" | "headerRow" | "totalsRow" | "style">,
  row: number,
  col: number,
  theme: DocumentTheme,
): CellStyle | undefined {
  const { ref, style } = table;
  if (row < ref.top || row > ref.bottom || col < ref.left || col > ref.right) {
    return undefined;
  }
  const look = tableStyleLook(style.name, theme);
  if (look === NO_LOOK) {
    return undefined;
  }
  let key =
    (row === ref.top ? TABLE_TOP : 0) |
    (row === ref.bottom ? TABLE_BOTTOM : 0) |
    (col === ref.left ? TABLE_LEFT : 0) |
    (col === ref.right ? TABLE_RIGHT : 0);
  if (table.headerRow && row === ref.top) {
    key |= HEADER;
  } else if (table.totalsRow && row === ref.bottom) {
    key |= TOTALS;
  } else {
    const top = ref.top + (table.headerRow ? 1 : 0);
    const bottom = ref.bottom - (table.totalsRow ? 1 : 0);
    key |= (row === top ? DATA_TOP : 0) | (row === bottom ? DATA_BOTTOM : 0);
    if (style.showColumnStripes) {
      key |= COLUMN_STRIPE | ((col - ref.left) % 2 === 1 ? SECOND_COLUMN : 0);
    }
    if (style.showRowStripes) {
      key |= ROW_STRIPE | (visibleOrdinal(sheet.rows, top, bottom, row) % 2 === 1 ? SECOND_ROW : 0);
    }
    if (style.showLastColumn && col === ref.right) {
      key |= LAST_COLUMN;
    }
    if (style.showFirstColumn && col === ref.left) {
      key |= FIRST_COLUMN;
    }
  }
  return memoisedCellLook(look, key);
}

/** A sample to preview a style on (the gallery): a blank sheet holding a
 * `rows` × `cols` table at A1 styled `name` — a header row over banded data
 * rows, the other options off. */
export function sampleTable(name: string | undefined, rows = 5, cols = 4): { sheet: Sheet; table: SheetTable } {
  return {
    sheet: { id: "sample", name: "Sample", visibility: "visible", kind: "grid", rows: [], columns: [], merges: [], conditionalFormats: [], validations: [], defaultColWidthPx: 64, defaultRowHeightPt: 15 },
    table: {
      id: "sample",
      name: "Table1",
      ref: { top: 0, left: 0, bottom: rows - 1, right: cols - 1 },
      headerRow: true,
      totalsRow: false,
      columns: Array.from({ length: cols }, (_, i) => ({ name: `Column${i + 1}` })),
      style: { name, showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    },
  };
}

// --- the look beneath a cell's own style ------------------------------------------------

/** Structural equality of two style values (plain JSON), a colour's source
 * `xlsxTint` spelling ignored. */
function sameValue(a: unknown, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) {
    return false;
  }
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
    if (key !== "xlsxTint" && !sameValue(left[key], right[key])) {
      return false;
    }
  }
  return true;
}

/** Whether a cell's own value is its own setting: present, and not merely
 * Normal's (`normalValue`, undefined when no Normal style is given). */
function setByCell(value: unknown, normalValue: unknown): boolean {
  return value !== undefined && !sameValue(value, normalValue);
}

function copyKey<T extends object, K extends keyof T>(target: T, source: T, key: K): void {
  target[key] = source[key];
}

/** A cell's own style over its table look (SPEC-133; SPEC-123's table look
 * beneath a cell's own style): every property the cell sets wins — the font
 * property by property (its bold beside the look's colour), the fill, each
 * border edge — and the look shows wherever the cell sets nothing. `normal`
 * is the workbook's default style (`styles[0]`): a value equal to Normal's
 * is inherited, not set — every style of a read or blank workbook carries
 * Normal's font colour, and a header's white text must still show. Without
 * `normal` every present property counts as set. SPEC-129's conditional
 * formats layer over the result. Pure: `own` itself when the look changes
 * nothing, else a new style; neither input is mutated. */
export function styleOverLook(look: CellStyle | undefined, own: CellStyle, normal?: CellStyle): CellStyle {
  if (look === undefined) {
    return own;
  }
  const style: CellStyle = { ...own };
  let changed = false;
  const lookFont = look.font;
  if (lookFont !== undefined) {
    let font: CellFont | undefined;
    for (const key of Object.keys(lookFont) as (keyof CellFont)[]) {
      if (lookFont[key] !== undefined && !setByCell(own.font?.[key], normal?.font?.[key])) {
        font ??= { ...own.font };
        copyKey(font, lookFont, key);
      }
    }
    if (font !== undefined) {
      style.font = font;
      changed = true;
    }
  }
  if (look.fill !== undefined && !setByCell(own.fill, normal?.fill)) {
    style.fill = look.fill;
    changed = true;
  }
  const lookBorders = look.borders;
  if (lookBorders !== undefined) {
    let borders: CellBorders | undefined;
    for (const side of SIDES) {
      if (lookBorders[side] !== undefined && !setByCell(own.borders?.[side], normal?.borders?.[side])) {
        borders ??= { ...own.borders };
        copyKey(borders, lookBorders, side);
      }
    }
    if (borders !== undefined) {
      style.borders = borders;
      changed = true;
    }
  }
  return changed ? style : own;
}

const lookKeys = new WeakMap<CellStyle, number>();
let nextLookKey = 1;

/** A number per look object (tableCellLook's are identity-stable) — for a
 * css cache key, so a cell under a look never shares a plain cell's entry. */
export function lookKey(look: CellStyle): number {
  let key = lookKeys.get(look);
  if (key === undefined) {
    key = nextLookKey++;
    lookKeys.set(look, key);
  }
  return key;
}
