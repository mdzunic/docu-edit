// Formatting commands over ranges (SPEC-123 "Patches", "Range classes",
// "Borders", "Format painter and Clear", "Named cell styles") — THE
// normative source for StylePatch, patchStyle and applyStylePatch, the border
// presets, Clear, applying a named cell style and the format painter's
// patterns. Every command works on the sheet it names, keeps the model
// invariants through restyle.ts's range-class engine (copy-on-write rows,
// sparseness, explicit style ids, the MAX_FORMAT_MATERIALISED cap) and
// returns a FormatResult; the view wraps each in ONE undo step (SPEC-121's
// record/commit) and writes it through.
//
// - patchStyle(style, patch) is pure. A patch's object aspects merge key by
//   key; a `null` removes a field or a whole aspect; an aspect left empty is
//   dropped; borders replace edge by edge. A numberFormat without a
//   builtinFormatId clears the style's built-in id (SPEC-118's rule;
//   "General" is no format). A result that differs from its input never
//   carries sourceXf/sourceStyleName — it is not the source's style; a patch
//   that changes nothing returns the input object itself. Patch values are
//   copied in as plain data, a caller's reactive Proxy included.
// - applyStylePatch interns each changed style once per base style id within
//   one command (a 200,000-cell range sharing one base style calls internStyle
//   once), writes the sheet copy-on-write with `dirty: true` (that sheet
//   only), and — when the patch carries a font — applies the font to every
//   rich-text run of the cells it rewrites too, so the change shows
//   (formatting part of a cell's text is backlog).
// - Refusals, checked before the model is touched: a range holding a locked
//   cell on a protected sheet (SPEC-122's message) unless the protection
//   permits formatting for the range's class — xlsx `formatCells="0"` for an
//   ordinary range, `formatRows="0"` for whole rows, `formatColumns="0"` for
//   whole columns, both of the latter for the whole sheet; and a command that
//   would materialise more than MAX_FORMAT_MATERIALISED cells and rows
//   (FORMAT_TOO_LARGE_MESSAGE). Overlapping ranges are harmless: patches are
//   idempotent, and each range is applied to the result of the one before.
// - Borders: borderPresetPatches turns a preset into at most four range
//   patches (whole bands of the range, so a whole column's sides land on its
//   span and its top and bottom on rows 1 and 1,048,576) plus, for each outer
//   edge the preset sets or removes, a `storedOnly` patch removing the facing
//   edge of the stored neighbours just outside — never materialising one.
// - Clear: contents is SPEC-122's Delete (clearContents, with its protection
//   and array refusals); formats removes the stored cells' style ids (and the
//   rows' and spans' for whole rows and columns — everything for the whole
//   sheet) and flattens rich-text runs, the value unchanged; all does both.
//   SPEC-124: formats and all also unmerge the merges inside the ranges.
// - Named styles: applyNamedStyle sets, for each aspect the style includes,
//   exactly the definition's value (a font wholesale, keys it lacks falling
//   back to Normal's), leaves the others, and sets namedStyle; Normal is Clear
//   formats. A built-in the workbook lacks is appended to `namedStyles` the
//   first time it is applied (and stays — undo never removes one).
// - The format painter: capturePattern records each position's style id (for
//   a whole-row or whole-column source, the rows' or spans' ids);
//   applyPattern tiles it from the target's top-left — a single-cell target
//   grows to the source's size (to whole rows or columns for such a source),
//   a larger target repeats the pattern, a smaller one takes what fits.
//   SPEC-124: an ordinary source's pattern also records the merges inside
//   it, which mergeCells.ts's applyPatternMerges reproduces in the target.
import { ARRAY_PART_MESSAGE, arrayRangeTouched, clearContents, firstLockedCell, isEmptyRow, PROTECTED_CELL_MESSAGE } from "./cells";
import type { GridRange } from "./address";
import { findRow } from "./address";
import {
  dropValueBound,
  MAX_COLS,
  MAX_ROWS,
  type Cell,
  type CellAlignment,
  type CellBorderEdge,
  type CellFill,
  type CellFont,
  type CellStyle,
  type ColumnSpan,
  type Row,
  type Sheet,
  type SpreadsheetModel,
  type TextRun,
} from "./model";
import {
  columnRuns,
  firstCellAtOrAfter,
  firstRowAtOrAfter,
  holdsOnlyStyle,
  MaterialiseBudget,
  MAX_FORMAT_MATERIALISED,
  rangeClass,
  restylePositions,
  restyleRange,
  TooManyPositions,
  type RestyleOp,
} from "./restyle";
import { cellStyle, effectiveStyleId, internStyle, resolveNamedStyle, styleKey, type ResolvedNamedStyle } from "./styles";

export { MAX_FORMAT_MATERIALISED, rangeClass, type RangeClass } from "./restyle";

type Nullable<T> = { [K in keyof T]?: T[K] | null };

export type DiagonalEdge = CellBorderEdge & { up?: true; down?: true };

export interface StylePatch {
  /** Merged key by key; null removes the aspect. */
  font?: Nullable<CellFont> | null;
  fill?: CellFill | null;
  /** Edge by edge; null removes every edge. */
  borders?: {
    top?: CellBorderEdge | null;
    right?: CellBorderEdge | null;
    bottom?: CellBorderEdge | null;
    left?: CellBorderEdge | null;
    diagonal?: DiagonalEdge | null;
  } | null;
  alignment?: Nullable<CellAlignment> | null;
  /** null = General. */
  numberFormat?: string | null;
  /** SPEC-118's locale-dependent built-in that goes with `numberFormat` (the
   * short date, the en-US accounting presets …); a numberFormat patch
   * without one clears the style's. */
  builtinFormatId?: number | null;
  protection?: Nullable<{ locked: false; hidden: true }> | null;
  namedStyle?: string | null;
}

export type FormatResult = { ok: true } | { ok: false; message: string };

/** One range and its patch; a `storedOnly` entry touches only what exists
 * (the border presets' neighbour clearing). */
export interface StylePatchEntry {
  range: GridRange;
  patch: StylePatch;
  storedOnly?: true;
}

export const FORMAT_TOO_LARGE_MESSAGE = "That's too many cells to format one by one — select whole columns or rows instead";
/** A command naming a sheet that is not a worksheet (a chart sheet …). */
export const NOT_A_WORKSHEET_MESSAGE = "Only a worksheet's cells can be formatted";

// --- patchStyle ------------------------------------------------------------------

const FONT_KEYS: readonly (keyof CellFont)[] = ["family", "sizePt", "bold", "italic", "underline", "strike", "color", "vertAlign"];
const ALIGNMENT_KEYS: readonly (keyof CellAlignment)[] = ["horizontal", "vertical", "wrap", "indent", "textRotation", "shrinkToFit"];
const EDGE_KEYS = ["top", "right", "bottom", "left", "diagonal"] as const;

/** A plain-data copy of a patch value. The model holds plain data only
 * (model.ts), and a value a view kept in a Vue `ref` arrives as a reactive
 * Proxy, which structuredClone refuses — read key by key, it copies plain
 * whatever the caller hands in. */
function plainCopy<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map((item: unknown) => plainCopy(item)) as T;
  }
  if (typeof value === "object" && value !== null) {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, plainCopy(item)])) as T;
  }
  return value;
}

/** `current` with `patch` merged key by key (null removes a key); undefined
 * when nothing is left. */
function mergeKeys<T extends object>(current: T | undefined, patch: Nullable<T>): T | undefined {
  const next: Record<string, unknown> = { ...current };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) {
      continue;
    }
    if (value === null) {
      delete next[key];
    } else {
      next[key] = plainCopy(value);
    }
  }
  return Object.keys(next).length > 0 ? (next as T) : undefined;
}

/** An aspect patched: unchanged for an absent patch, removed for null. */
function patchAspect<T extends object>(current: T | undefined, patch: Nullable<T> | null | undefined): T | undefined {
  if (patch === undefined) {
    return current;
  }
  return patch === null ? undefined : mergeKeys(current, patch);
}

function setOrDelete<K extends keyof CellStyle>(style: CellStyle, key: K, value: CellStyle[K] | undefined): void {
  if (value === undefined) {
    delete style[key];
  } else {
    style[key] = value;
  }
}

/** A style with `patch` applied (see the module header) — pure. */
export function patchStyle(style: CellStyle, patch: StylePatch): CellStyle {
  const next: CellStyle = { ...style };
  setOrDelete(next, "font", patchAspect(style.font, patch.font));
  if (patch.fill !== undefined) {
    setOrDelete(next, "fill", patch.fill === null ? undefined : plainCopy(patch.fill));
  }
  setOrDelete(next, "borders", patchAspect(style.borders, patch.borders));
  let alignment = patchAspect(style.alignment, patch.alignment);
  if (alignment && (alignment.indent === 0 || alignment.textRotation === 0)) {
    // 0 is no indent and no rotation: the readers never store either.
    const { indent, textRotation, ...rest } = alignment;
    alignment = { ...rest, ...(indent ? { indent } : {}), ...(textRotation ? { textRotation } : {}) };
    alignment = Object.keys(alignment).length > 0 ? alignment : undefined;
  }
  setOrDelete(next, "alignment", alignment);
  if (patch.numberFormat !== undefined) {
    const code = patch.numberFormat;
    if (code === null || /^general$/i.test(code.trim())) {
      delete next.numberFormat;
      delete next.builtinFormatId;
    } else {
      next.numberFormat = code;
      setOrDelete(next, "builtinFormatId", patch.builtinFormatId ?? undefined);
    }
  } else if (patch.builtinFormatId !== undefined) {
    setOrDelete(next, "builtinFormatId", patch.builtinFormatId ?? undefined);
  }
  setOrDelete(next, "protection", patchAspect(style.protection, patch.protection) as CellStyle["protection"]);
  if (patch.namedStyle !== undefined) {
    setOrDelete(next, "namedStyle", patch.namedStyle ?? undefined);
  }
  if (styleKey(next) === styleKey(style)) {
    return style;
  }
  delete next.sourceXf;
  delete next.sourceStyleName;
  return next;
}

/** A cell whose rich-text runs take a font patch too (an unchanged cell is
 * returned as is); changed runs drop what is bound to the old text — the
 * shared-string index, the ods preserved text and an xlsx ` vm` (SPEC-122's
 * drop rules). */
function withPatchedRuns(cell: Cell, font: Nullable<CellFont> | null): Cell {
  if (!cell.runs) {
    return cell;
  }
  let changed = false;
  const runs = cell.runs.map((run): TextRun => {
    const next = font === null ? undefined : mergeKeys(run.font, font);
    if (styleKey({ font: next }) === styleKey({ font: run.font })) {
      return run;
    }
    changed = true;
    return next === undefined ? { text: run.text } : { text: run.text, font: next };
  });
  if (!changed) {
    return cell;
  }
  const next: Cell = { ...cell };
  dropValueBound(next);
  next.runs = runs;
  return next;
}

/** The engine's rewrite for a patch: patched styles interned once per base
 * id into `holder` (a scratch view of the model's styles). */
function patchOp(holder: Pick<SpreadsheetModel, "styles">, patch: StylePatch, storedOnly?: boolean): RestyleOp {
  const memo = new Map<number, number>();
  const op: RestyleOp = {
    next(base) {
      let id = memo.get(base);
      if (id === undefined) {
        const style = cellStyle(holder, base);
        const patched = patchStyle(style, patch);
        id = patched === style ? base : internStyle(holder, patched);
        memo.set(base, id);
      }
      return id;
    },
  };
  const font = patch.font;
  if (font !== undefined) {
    op.cell = (cell) => withPatchedRuns(cell, font);
  }
  if (storedOnly) {
    op.storedOnly = true;
  }
  return op;
}

// --- refusals and the commit -------------------------------------------------------

/** The sheet a command names, or the refusal. */
function sheetFor(model: SpreadsheetModel, sheetId: string): { index: number; sheet: Sheet } | { message: string } {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  return sheet && sheet.kind === "grid" ? { index, sheet } : { message: NOT_A_WORKSHEET_MESSAGE };
}

/** Whether a protected sheet's protection lets formatting reach a range's
 * locked cells (see the module header). */
function formattingPermitted(sheet: Sheet, range: GridRange): boolean {
  const protection = sheet.protection;
  if (!protection) {
    return true;
  }
  switch (rangeClass(range)) {
    case "ordinary":
      return protection.allowFormatCells === true;
    case "rows":
      return protection.allowFormatRows === true;
    case "columns":
      return protection.allowFormatColumns === true;
    case "sheet":
      return protection.allowFormatRows === true && protection.allowFormatColumns === true;
  }
}

/** Whether formatting `range` is refused: a locked cell on a protected sheet
 * whose protection does not permit formatting this class. */
function formattingRefused(model: SpreadsheetModel, index: number, range: GridRange): boolean {
  const sheet = model.sheets[index]!;
  return !formattingPermitted(sheet, range) && firstLockedCell(model, index, [range]) !== null;
}

/** Writes a command's sheet into the model (copy-on-write, `dirty: true`)
 * with the styles it interned — nothing at all when the sheet is unchanged. */
function commitSheet(model: SpreadsheetModel, index: number, sheet: Sheet, styles: SpreadsheetModel["styles"]): void {
  if (sheet === model.sheets[index]) {
    return;
  }
  if (styles !== model.styles) {
    model.styles = styles;
  }
  const next: Sheet = { ...sheet, dirty: true };
  model.sheets = model.sheets.map((each, i) => (i === index ? next : each));
}

// --- applyStylePatch ---------------------------------------------------------------

/** Applies every entry in order as ONE command (see the module header). */
export function applyStylePatches(model: SpreadsheetModel, sheetId: string, entries: readonly StylePatchEntry[]): FormatResult {
  const found = sheetFor(model, sheetId);
  if ("message" in found) {
    return { ok: false, message: found.message };
  }
  if (entries.some((entry) => !entry.storedOnly && formattingRefused(model, found.index, entry.range))) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  const holder = { styles: model.styles };
  const budget = new MaterialiseBudget();
  let sheet = found.sheet;
  try {
    for (const entry of entries) {
      // A protected neighbour keeps its edge rather than refusing the command.
      if (entry.storedOnly && formattingRefused(model, found.index, entry.range)) {
        continue;
      }
      sheet = restyleRange(sheet, entry.range, patchOp(holder, entry.patch, entry.storedOnly), budget);
    }
  } catch (error) {
    if (error instanceof TooManyPositions) {
      return { ok: false, message: FORMAT_TOO_LARGE_MESSAGE };
    }
    throw error;
  }
  commitSheet(model, found.index, sheet, holder.styles);
  return { ok: true };
}

/** Applies `patch` to every position of `ranges` on the sheet `sheetId`, by
 * range class (see the module header). */
export function applyStylePatch(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], patch: StylePatch): FormatResult {
  return applyStylePatches(
    model,
    sheetId,
    ranges.map((range) => ({ range, patch })),
  );
}

// --- borders ---------------------------------------------------------------------

export type BorderPreset = "bottom" | "top" | "left" | "right" | "all" | "outside" | "thick-outside" | "none";

/** The presets in the Borders menu's order. */
export const BORDER_PRESETS: readonly BorderPreset[] = ["bottom", "top", "left", "right", "all", "outside", "thick-outside", "none"];

/** The outer edges each preset sets or removes (whose stored neighbours lose
 * their facing edge). */
const OUTER_EDGES: Readonly<Record<BorderPreset, readonly ("top" | "bottom" | "left" | "right")[]>> = {
  bottom: ["bottom"],
  top: ["top"],
  left: ["left"],
  right: ["right"],
  all: ["top", "bottom", "left", "right"],
  outside: ["top", "bottom", "left", "right"],
  "thick-outside": ["top", "bottom", "left", "right"],
  none: ["top", "bottom", "left", "right"],
};

const sides = (borders: NonNullable<StylePatch["borders"]>): StylePatch => ({ borders });

const band = (range: GridRange, side: "top" | "bottom" | "left" | "right"): GridRange =>
  side === "top"
    ? { ...range, bottom: range.top }
    : side === "bottom"
      ? { ...range, top: range.bottom }
      : side === "left"
        ? { ...range, right: range.left }
        : { ...range, left: range.right };

/** A preset's patches for `range` (SPEC-123 "Borders"): `edge` is the
 * session's line style and colour; Thick outside borders draws `medium`
 * (what Excel's "Thick Box Border" writes); No border removes every edge but
 * the diagonals. */
export function borderPresetPatches(range: GridRange, preset: BorderPreset, edge: CellBorderEdge): StylePatchEntry[] {
  const line: CellBorderEdge = preset === "thick-outside" ? { ...edge, style: "medium" } : edge;
  const entries: StylePatchEntry[] = [];
  switch (preset) {
    case "bottom":
    case "top":
    case "left":
    case "right":
      entries.push({ range: band(range, preset), patch: sides({ [preset]: line }) });
      break;
    case "all":
      entries.push({ range, patch: sides({ top: line, right: line, bottom: line, left: line }) });
      break;
    case "none":
      entries.push({ range, patch: sides({ top: null, right: null, bottom: null, left: null }) });
      break;
    case "outside":
    case "thick-outside":
      if (range.left === range.right) {
        entries.push({ range, patch: sides({ left: line, right: line }) });
      } else {
        entries.push({ range: band(range, "left"), patch: sides({ left: line }) }, { range: band(range, "right"), patch: sides({ right: line }) });
      }
      if (range.top === range.bottom) {
        entries.push({ range, patch: sides({ top: line, bottom: line }) });
      } else {
        entries.push({ range: band(range, "top"), patch: sides({ top: line }) }, { range: band(range, "bottom"), patch: sides({ bottom: line }) });
      }
      break;
  }
  for (const side of OUTER_EDGES[preset]) {
    const neighbour = neighbourOf(range, side);
    if (neighbour) {
      const facing = side === "top" ? "bottom" : side === "bottom" ? "top" : side === "left" ? "right" : "left";
      entries.push({ range: neighbour, patch: sides({ [facing]: null }), storedOnly: true });
    }
  }
  return entries;
}

/** The row or column just outside one side of a range, or null at the
 * grid's edge. */
function neighbourOf(range: GridRange, side: "top" | "bottom" | "left" | "right"): GridRange | null {
  switch (side) {
    case "top":
      return range.top > 0 ? { ...range, top: range.top - 1, bottom: range.top - 1 } : null;
    case "bottom":
      return range.bottom < MAX_ROWS - 1 ? { ...range, top: range.bottom + 1, bottom: range.bottom + 1 } : null;
    case "left":
      return range.left > 0 ? { ...range, left: range.left - 1, right: range.left - 1 } : null;
    case "right":
      return range.right < MAX_COLS - 1 ? { ...range, left: range.right + 1, right: range.right + 1 } : null;
  }
}

// --- Clear -----------------------------------------------------------------------------

/** A cell without its format: no styleId, rich-text runs flattened (the value
 * is their text already) — the cell itself when it has neither; undefined
 * when nothing is left. */
function withoutFormat(cell: Cell): Cell | undefined {
  if (cell.styleId === undefined && cell.runs === undefined) {
    return cell;
  }
  const next: Cell = { ...cell };
  delete next.styleId;
  if (next.runs !== undefined) {
    dropValueBound(next);
  }
  return holdsOnlyStyle(next) ? undefined : next;
}

/** The spans with [left, right]'s styles removed (split at the edges; a span
 * left holding nothing dropped) — the same array when none had one. */
function spansWithoutStyle(columns: ColumnSpan[], left: number, right: number): ColumnSpan[] {
  if (!columns.some((span) => span.styleId !== undefined && span.max >= left && span.min <= right)) {
    return columns;
  }
  const out: ColumnSpan[] = [];
  for (const span of columns) {
    if (span.styleId === undefined || span.max < left || span.min > right) {
      out.push(span);
      continue;
    }
    if (span.min < left) {
      out.push({ ...span, max: left - 1 });
    }
    const { styleId: _style, ...inside } = { ...span, min: Math.max(span.min, left), max: Math.min(span.max, right) };
    if (Object.keys(inside).some((key) => key !== "min" && key !== "max" && (inside as Record<string, unknown>)[key] !== undefined)) {
      out.push(inside);
    }
    if (span.max > right) {
      out.push({ ...span, min: right + 1 });
    }
  }
  return out;
}

/** Clear formats over one range (see the module header). */
function clearFormatsIn(sheet: Sheet, range: GridRange): Sheet {
  const kind = rangeClass(range);
  const rowsToo = kind === "rows" || kind === "sheet";
  const replaced = new Map<number, Row | undefined>();
  const rows = sheet.rows;
  for (let r = firstRowAtOrAfter(rows, range.top); r < rows.length && rows[r]!.index <= range.bottom; r += 1) {
    const row = rows[r]!;
    let cells: Cell[] | null = null;
    for (let c = firstCellAtOrAfter(row.cells, range.left); c < row.cells.length && row.cells[c]!.col <= range.right; c += 1) {
      const cell = row.cells[c]!;
      const next = withoutFormat(cell);
      if (next !== cell) {
        cells ??= row.cells.slice(0, c);
        if (next) {
          cells.push(next);
        }
      } else if (cells) {
        cells.push(cell);
      }
    }
    const dropStyle = rowsToo && row.styleId !== undefined;
    if (!cells && !dropStyle) {
      continue;
    }
    const next: Row = { ...row };
    if (cells) {
      const after = firstCellAtOrAfter(row.cells, range.right + 1);
      next.cells = [...cells, ...row.cells.slice(after)];
    }
    if (dropStyle) {
      delete next.styleId;
    }
    replaced.set(row.index, isEmptyRow(next) ? undefined : next);
  }
  const columns = kind === "columns" || kind === "sheet" ? spansWithoutStyle(sheet.columns, range.left, range.right) : sheet.columns;
  if (replaced.size === 0 && columns === sheet.columns) {
    return sheet;
  }
  const next: Sheet = { ...sheet, columns };
  if (replaced.size > 0) {
    next.rows = rows.flatMap((row) => (replaced.has(row.index) ? (replaced.get(row.index) ?? []) : [row]));
  }
  return next;
}

/** Clear ▸ All / Formats / Contents over `ranges` (see the module header). */
export function clearRange(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], what: "all" | "formats" | "contents"): FormatResult {
  const found = sheetFor(model, sheetId);
  if ("message" in found) {
    return { ok: false, message: found.message };
  }
  const list = [...ranges];
  if (what !== "formats") {
    if (firstLockedCell(model, found.index, list)) {
      return { ok: false, message: PROTECTED_CELL_MESSAGE };
    }
    if (arrayRangeTouched(model, found.index, list)) {
      return { ok: false, message: ARRAY_PART_MESSAGE };
    }
  } else if (list.some((range) => formattingRefused(model, found.index, range))) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  if (what !== "formats") {
    clearContents(model, found.index, list);
  }
  let sheet = model.sheets[found.index]!;
  if (what !== "contents") {
    for (const range of list) {
      sheet = clearFormatsIn(sheet, range);
    }
    // SPEC-124: clearing formats also unmerges the merges inside the ranges.
    const merges = sheet.merges.filter((merge) => !list.some((range) => rangeContains(range, merge)));
    if (merges.length !== sheet.merges.length) {
      sheet = { ...sheet, merges };
    }
  }
  if (sheet !== found.sheet) {
    // clearContents may already have replaced the sheet: commit against the
    // command's starting point.
    model.sheets = model.sheets.map((each, i) => (i === found.index ? { ...sheet, dirty: true as const } : each));
  }
  return { ok: true };
}

// --- named styles ------------------------------------------------------------------

/** The patch applying a named style sets (see the module header). */
export function namedStylePatch(model: Pick<SpreadsheetModel, "styles">, named: ResolvedNamedStyle): StylePatch {
  const definition = named.style;
  const patch: StylePatch = { namedStyle: named.name };
  for (const aspect of named.includes) {
    switch (aspect) {
      case "numberFormat":
        patch.numberFormat = definition.numberFormat ?? null;
        patch.builtinFormatId = definition.builtinFormatId ?? null;
        break;
      case "font": {
        const font: CellFont = { ...cellStyle(model, 0).font, ...definition.font };
        patch.font = Object.fromEntries(FONT_KEYS.map((key) => [key, font[key] ?? null])) as Nullable<CellFont>;
        break;
      }
      case "fill":
        patch.fill = definition.fill ?? null;
        break;
      case "borders": {
        const borders = definition.borders;
        patch.borders = borders ? Object.fromEntries(EDGE_KEYS.map((key) => [key, borders[key] ?? null])) : null;
        break;
      }
      case "alignment": {
        const alignment = definition.alignment;
        patch.alignment = alignment ? (Object.fromEntries(ALIGNMENT_KEYS.map((key) => [key, alignment[key] ?? null])) as Nullable<CellAlignment>) : null;
        break;
      }
      case "protection": {
        const protection = definition.protection;
        patch.protection = protection ? { locked: protection.locked ?? null, hidden: protection.hidden ?? null } : null;
        break;
      }
    }
  }
  return patch;
}

/** Applies the cell style `name` to `ranges` (see the module header). */
export function applyNamedStyle(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], name: string): FormatResult {
  const named = resolveNamedStyle(model, name);
  if (!named) {
    return { ok: false, message: `There's no cell style called "${name}"` };
  }
  if (named.builtinId === 0) {
    return clearRange(model, sheetId, ranges, "formats");
  }
  const result = applyStylePatch(model, sheetId, ranges, namedStylePatch(model, named));
  if (result.ok && !named.defined) {
    const entry = { name: named.name, style: structuredClone(named.style), includes: [...named.includes] };
    model.namedStyles = [...(model.namedStyles ?? []), named.builtinId === undefined ? entry : { ...entry, builtinId: named.builtinId }];
  }
  return result;
}

// --- the format painter ------------------------------------------------------------------

/** A source's formats: `rows × cols` style ids, row-major. A whole-row source
 * holds its rows' ids (`axis: "rows"`, one column), a whole-column source
 * its spans' ids (`axis: "columns"`, one row). */
export interface StylePattern {
  rows: number;
  cols: number;
  ids: number[];
  axis?: "rows" | "columns";
  /** SPEC-124: the source's merges, relative to its top-left. */
  merges?: GridRange[];
}

/** The source's pattern (see the module header); null for a source of more
 * than MAX_FORMAT_MATERIALISED positions (refused with
 * FORMAT_TOO_LARGE_MESSAGE). */
export function capturePattern(sheet: Sheet, range: GridRange): StylePattern | null {
  const kind = rangeClass(range);
  if (kind === "columns" || kind === "sheet") {
    const ids: number[] = [];
    for (const run of columnRuns(sheet.columns, range.left, range.right)) {
      for (let col = run.min; col <= run.max; col += 1) {
        ids.push(run.styleId ?? 0);
      }
    }
    return { rows: 1, cols: ids.length, ids, axis: "columns" };
  }
  const rows = range.bottom - range.top + 1;
  if (kind === "rows") {
    if (rows > MAX_FORMAT_MATERIALISED) {
      return null;
    }
    const ids = Array.from({ length: rows }, (_, i) => findRow(sheet, range.top + i)?.styleId ?? 0);
    return { rows, cols: 1, ids, axis: "rows" };
  }
  const cols = range.right - range.left + 1;
  if (rows * cols > MAX_FORMAT_MATERIALISED) {
    return null;
  }
  const ids: number[] = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      ids.push(effectiveStyleId(sheet, row, col));
    }
  }
  // SPEC-124: the merges inside the source, relative to its top-left — the
  // painter reproduces them in the target (mergeCells.ts applyPatternMerges).
  const merges = sheet.merges
    .filter((merge) => rangeContains(range, merge))
    .map((merge) => ({ top: merge.top - range.top, left: merge.left - range.left, bottom: merge.bottom - range.top, right: merge.right - range.left }));
  return merges.length > 0 ? { rows, cols, ids, merges } : { rows, cols, ids };
}

const singleCell = (range: GridRange): boolean => range.top === range.bottom && range.left === range.right;

/** Whether `outer` contains all of `inner`. */
function rangeContains(outer: GridRange, inner: GridRange): boolean {
  return outer.top <= inner.top && outer.left <= inner.left && outer.bottom >= inner.bottom && outer.right >= inner.right;
}

/** The range a pattern lands on: a single-cell target grows to the source's
 * size (to whole rows or columns for such a source). */
export function patternTarget(target: GridRange, pattern: StylePattern): GridRange {
  if (!singleCell(target)) {
    return target;
  }
  if (pattern.axis === "columns") {
    return { top: 0, bottom: MAX_ROWS - 1, left: target.left, right: Math.min(MAX_COLS - 1, target.left + pattern.cols - 1) };
  }
  if (pattern.axis === "rows") {
    return { top: target.top, bottom: Math.min(MAX_ROWS - 1, target.top + pattern.rows - 1), left: 0, right: MAX_COLS - 1 };
  }
  return { ...target, bottom: Math.min(MAX_ROWS - 1, target.top + pattern.rows - 1), right: Math.min(MAX_COLS - 1, target.left + pattern.cols - 1) };
}

/** Runs of consecutive equal ids along [from, to] of a tiled axis pattern. */
function idRuns(from: number, to: number, idAt: (index: number) => number): { min: number; max: number; id: number }[] {
  const runs: { min: number; max: number; id: number }[] = [];
  for (let index = from; index <= to; index += 1) {
    const id = idAt(index);
    const last = runs.at(-1);
    if (last && last.id === id && last.max === index - 1) {
      last.max = index;
    } else {
      runs.push({ min: index, max: index, id });
    }
  }
  return runs;
}

/** Applies a captured pattern to `target` (see the module header). */
export function applyPattern(model: SpreadsheetModel, sheetId: string, target: GridRange, pattern: StylePattern): FormatResult {
  const found = sheetFor(model, sheetId);
  if ("message" in found) {
    return { ok: false, message: found.message };
  }
  const range = patternTarget(target, pattern);
  if (formattingRefused(model, found.index, range)) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  const kind = rangeClass(range);
  const holder = { styles: model.styles };
  const budget = new MaterialiseBudget();
  let sheet = found.sheet;
  const tileAt = (row: number, col: number): number =>
    pattern.ids[((row - range.top) % pattern.rows) * pattern.cols + ((col - range.left) % pattern.cols)]!;
  try {
    if (pattern.axis === "columns" && (kind === "columns" || kind === "sheet")) {
      for (const run of idRuns(range.left, range.right, (col) => pattern.ids[(col - range.left) % pattern.cols]!)) {
        sheet = restyleRange(sheet, { ...range, left: run.min, right: run.max }, { next: () => run.id }, budget);
      }
    } else if (pattern.axis === "rows" && (kind === "rows" || kind === "sheet")) {
      for (const run of idRuns(range.top, range.bottom, (row) => pattern.ids[(row - range.top) % pattern.rows]!)) {
        sheet = restyleRange(sheet, { ...range, top: run.min, bottom: run.max }, { next: () => run.id }, budget);
      }
    } else {
      if ((range.bottom - range.top + 1) * (range.right - range.left + 1) > MAX_FORMAT_MATERIALISED) {
        throw new TooManyPositions();
      }
      sheet = restylePositions(sheet, range, tileAt, budget);
    }
  } catch (error) {
    if (error instanceof TooManyPositions) {
      return { ok: false, message: FORMAT_TOO_LARGE_MESSAGE };
    }
    throw error;
  }
  commitSheet(model, found.index, sheet, holder.styles);
  return { ok: true };
}

// --- number formats ------------------------------------------------------------------
// The presets, decimal stepping and the Format Cells code builders live in
// numberCodes.ts; the formatting layer is their public face.
export {
  buildNumberCode,
  commaButtonFormat,
  currencyButtonFormat,
  describeNumberCode,
  LOCALE_CURRENCY,
  numberFormatChoiceOf,
  NUMBER_FORMAT_PRESETS,
  percentButtonFormat,
  presetIdOf,
  presetNumberFormat,
  shortcutNumberFormat,
  stepDecimals,
  type LocaleCurrency,
  type NumberCodeSpec,
  type NumberFormatChoice,
  type NumberPresetId,
} from "./numberCodes";
