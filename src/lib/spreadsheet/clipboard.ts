// Copying cells (SPEC-125 "Flavours", "The payload", "HTML out") — THIS FILE
// is the normative source for what Copy and Cut put on the clipboard: the
// copied rectangle, the internal payload (clipboardPayload.ts holds its type
// and validating parser) and the two flavours, plus TSV in. Pasting is
// paste.ts; external HTML, clipboardImport.ts.
//
// - The copied rectangle: one range (a multi-range selection is refused),
//   its bottom and right edges clipped to the sheet's used range — never
//   above its top-left, so copying column A copies A1 to its last used row —
//   and grown back over any merge the clip would cut. Over MAX_COPY_CELLS
//   positions it is refused.
// - The payload holds, per position, the cell's value (a formula's cached
//   value), its canonical formula (not an unparsed one), its array range
//   when the whole array was copied (a member copied without its anchor
//   gives its value only), its rich-text runs and its style. An ordinary
//   copy records each position's EFFECTIVE style (cell, row, column), so a
//   position styled only by its row or column still carries its formatting;
//   a whole-row or whole-column copy records the cells' own styles and the
//   rows' or columns' properties (`rowProps`/`colProps`) instead. Styles are
//   listed once each, without `sourceXf`/`sourceStyleName`; `styleIds` (kept
//   in memory by the view, never on the clipboard) maps them back to the
//   source workbook's ids so a same-workbook paste keeps them by identity.
// - text/plain is TSV of each position's SPEC-118 display text at full width
//   (no "########"): cells joined by \t, rows by \r\n, no trailing newline; a
//   field holding a tab, line break or `"` is quoted with `"` doubled.
// - text/html is `<meta charset="utf-8"><table data-docu-edit-cells="…">`:
//   one <td> per position (a merge's anchor with colspan/rowspan, its other
//   positions omitted), the escaped display text, and inline CSS from the
//   "HTML out" whitelist only — font-family, font-size (pt),
//   font-weight:700, font-style:italic, text-decoration (underline,
//   line-through), color, background-color, text-align, vertical-align,
//   white-space:pre-wrap for wrapped cells and border-top|right|bottom|left
//   (1px/2px/3px solid|dashed|dotted|double + colour), theme colours resolved
//   to hex. No class, no script, nothing a reader must execute.
import { DEFAULT_THEME, resolveColor, resolveFont, type DocumentTheme } from "../core/theme";
import { findRow, usedRange, type GridRange } from "./address";
import {
  encodePayload,
  MAX_COPY_CELLS,
  PAYLOAD_ATTRIBUTE,
  PAYLOAD_KIND,
  type ClipboardCell,
  type ClipboardColProps,
  type ClipboardRowProps,
  type SheetClipboardPayload,
} from "./clipboardPayload";
import { buildMergeIndex, mergeAt, mergesIn } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type CellBorderEdge, type CellStyle, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { formatValue, plainText, workbookFormatContext, type FormatContext } from "./numberFormat/render";
import { firstCellAtOrAfter, firstRowAtOrAfter } from "./restyle";
import { cellStyle, columnSpanAt } from "./styles";

export {
  encodePayload,
  MAX_COPY_CELLS,
  parseClipboardPayload,
  PAYLOAD_ATTRIBUTE,
  type ClipboardCell,
  type SheetClipboardPayload,
} from "./clipboardPayload";

// --- messages (SPEC-125's refusals and toasts) -------------------------------------------

export const MULTIPLE_SELECTIONS_MESSAGE = "That doesn't work on multiple selections";
export const TOO_MANY_TO_COPY_MESSAGE = "That's too many cells to copy — select fewer than 1,000,000 cells";
export const TOO_MANY_TO_CHANGE_MESSAGE = "That's too many cells to change at once";
export const PASTE_PAST_EDGE_MESSAGE = "The paste area goes past the edge of the sheet";
export const PASTE_MERGE_MESSAGE = "Can't paste over part of a merged cell";
/** SPEC-016's toast when the async clipboard cannot be read. */
export const CLIPBOARD_UNAVAILABLE_MESSAGE = "Clipboard unavailable — paste with the keyboard instead";
export const CLIPBOARD_BLOCKED_MESSAGE = "Clipboard access was blocked — allow it for this site to use Paste special";
export const MOVE_OVERWRITE_MESSAGE = "There's already data here. Do you want to replace it?";

/** "N formulas referring to other sheets were pasted as values" (one: "1
 * formula … was pasted as a value"). */
export function formulasAsValuesMessage(count: number): string {
  return count === 1
    ? "1 formula referring to other sheets was pasted as a value"
    : `${count} formulas referring to other sheets were pasted as values`;
}

// --- the copied rectangle ----------------------------------------------------------------

export interface CopyArea {
  range: GridRange;
  wholeRows: boolean;
  wholeCols: boolean;
}

const isWholeRows = (range: GridRange): boolean => range.left === 0 && range.right === MAX_COLS - 1;
const isWholeCols = (range: GridRange): boolean => range.top === 0 && range.bottom === MAX_ROWS - 1;

/** A selected range as Copy takes it (see the module header): its bottom
 * and right clipped to the used range, never above its top-left, grown back
 * over any merge the clip would cut. */
export function copyArea(sheet: Sheet, range: GridRange): CopyArea {
  const used = usedRange(sheet);
  let clipped: GridRange = {
    top: range.top,
    left: range.left,
    bottom: Math.max(range.top, Math.min(range.bottom, used?.bottom ?? range.top)),
    right: Math.max(range.left, Math.min(range.right, used?.right ?? range.left)),
  };
  const merges = buildMergeIndex(sheet.merges);
  for (let grown = true; grown; ) {
    grown = false;
    for (const merge of mergesIn(merges, clipped)) {
      const bottom = Math.min(range.bottom, Math.max(clipped.bottom, merge.bottom));
      const right = Math.min(range.right, Math.max(clipped.right, merge.right));
      if (bottom !== clipped.bottom || right !== clipped.right) {
        clipped = { ...clipped, bottom, right };
        grown = true;
      }
    }
  }
  return { range: clipped, wholeRows: isWholeRows(range), wholeCols: isWholeCols(range) };
}

/** Why Copy (or Cut) refuses a selection, or null: several ranges, or more
 * than MAX_COPY_CELLS positions once clipped. */
export function copyRefusal(sheet: Sheet, ranges: readonly GridRange[]): string | null {
  if (ranges.length !== 1) {
    return MULTIPLE_SELECTIONS_MESSAGE;
  }
  const { range } = copyArea(sheet, ranges[0]!);
  return (range.bottom - range.top + 1) * (range.right - range.left + 1) > MAX_COPY_CELLS ? TOO_MANY_TO_COPY_MESSAGE : null;
}

// --- the payload -------------------------------------------------------------------------

/** One copy: the payload, the source style id behind each of its styles
 * (in memory only — see the module header) and the area it covers. */
export interface ClipboardCopy {
  payload: SheetClipboardPayload;
  styleIds: number[];
  area: CopyArea;
}

/** A style as the clipboard carries it: no `sourceXf`/`sourceStyleName`. */
function clipboardStyle(style: CellStyle): CellStyle {
  const { sourceXf: _xf, sourceStyleName: _name, ...rest } = style;
  return structuredClone(rest);
}

/** Calls `visit` for every position of `range` (row-major) with its stored
 * row, its stored cell and its effective style id — walking each row's
 * cells with a cursor, never a search per position. */
export function forEachPosition(
  sheet: Sheet,
  range: GridRange,
  visit: (row: number, col: number, cell: Cell | undefined, styleId: number, stored: Row | undefined) => void,
): void {
  for (let row = range.top; row <= range.bottom; row += 1) {
    const stored = findRow(sheet, row);
    const cells = stored?.cells ?? [];
    let at = firstCellAtOrAfter(cells, range.left);
    for (let col = range.left; col <= range.right; col += 1) {
      while (at < cells.length && cells[at]!.col < col) {
        at += 1;
      }
      const cell = cells[at]?.col === col ? cells[at] : undefined;
      const styleId = cell?.styleId ?? stored?.styleId ?? columnSpanAt(sheet.columns, col)?.styleId ?? 0;
      visit(row, col, cell, styleId, stored);
    }
  }
}

/** The copy of `range` on `sheetId` (see the module header); a cut carries
 * `cutToken`. The caller checked copyRefusal. */
export function buildClipboardCopy(model: SpreadsheetModel, sheetId: string, range: GridRange, cutToken?: string): ClipboardCopy {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet) {
    throw new RangeError(`No sheet ${sheetId}`);
  }
  const area = copyArea(sheet, range);
  const { top, left, bottom, right } = area.range;
  const whole = area.wholeRows || area.wholeCols;
  const styleIds: number[] = [];
  const styleIndex = new Map<number, number>();
  const styles: CellStyle[] = [];
  const indexOf = (id: number): number => {
    let index = styleIndex.get(id);
    if (index === undefined) {
      index = styles.length;
      styleIndex.set(id, index);
      styleIds.push(id);
      styles.push(clipboardStyle(cellStyle(model, id)));
    }
    return index;
  };
  const cells: ClipboardCell[] = [];
  forEachPosition(sheet, area.range, (row, col, cell, effective) => {
    // A whole-axis copy carries the cells' own styles; the rows' and
    // columns' travel as their properties.
    const styleId = whole ? cell?.styleId : effective;
    const formula = cell?.formula !== undefined && !cell.formulaUnparsed ? cell.formula : undefined;
    // A blank position travels as nothing: no value, no formula, and no style
    // of its own (a whole-axis copy) or only the default (an ordinary one).
    if ((cell === undefined || (cell.value === null && formula === undefined)) && (whole ? styleId === undefined : styleId === 0)) {
      return;
    }
    const entry: ClipboardCell = { r: row - top, c: col - left, v: cell?.value ?? null };
    if (formula !== undefined) {
      entry.f = formula;
      const array = cell?.arrayRange;
      if (array && array.top === row && array.left === col && array.bottom <= bottom && array.right <= right) {
        entry.a = { top: array.top - top, left: array.left - left, bottom: array.bottom - top, right: array.right - left };
      }
    }
    if (styleId !== undefined && (styleId !== 0 || whole)) {
      entry.s = indexOf(styleId);
    }
    if (cell?.runs && typeof cell.value === "string") {
      entry.t = structuredClone(cell.runs);
    }
    cells.push(entry);
  });
  const payload: SheetClipboardPayload = {
    kind: PAYLOAD_KIND,
    version: 1,
    workbookId: model.id,
    sheetName: sheet.name,
    dateSystem: model.dateSystem,
    origin: { row: top, col: left },
    rows: bottom - top + 1,
    cols: right - left + 1,
    cells,
    styles,
    merges: sheet.merges
      .filter((merge) => merge.top >= top && merge.left >= left && merge.bottom <= bottom && merge.right <= right)
      .map((merge) => ({ top: merge.top - top, left: merge.left - left, bottom: merge.bottom - top, right: merge.right - left })),
  };
  if (area.wholeRows) {
    payload.wholeRows = true;
    payload.rowProps = rowPropsOf(sheet, area.range, indexOf);
  }
  if (area.wholeCols) {
    payload.wholeCols = true;
    payload.colProps = colPropsOf(sheet, area.range, indexOf);
  }
  if (cutToken !== undefined) {
    payload.cut = cutToken;
  }
  return { payload, styleIds, area };
}

/** The payload of one copy (see buildClipboardCopy). */
export function buildClipboardPayload(model: SpreadsheetModel, sheetId: string, range: GridRange, cutToken?: string): SheetClipboardPayload {
  return buildClipboardCopy(model, sheetId, range, cutToken).payload;
}

/** The copied rows' own properties (a whole-row copy). */
function rowPropsOf(sheet: Sheet, range: GridRange, indexOf: (id: number) => number): ClipboardRowProps[] {
  const out: ClipboardRowProps[] = [];
  for (let r = firstRowAtOrAfter(sheet.rows, range.top); r < sheet.rows.length && sheet.rows[r]!.index <= range.bottom; r += 1) {
    const row = sheet.rows[r]!;
    const props: ClipboardRowProps = { r: row.index - range.top };
    if (row.heightPt !== undefined) {
      props.heightPt = row.heightPt;
    }
    if (row.customHeight) {
      props.customHeight = true;
    }
    if (row.hidden) {
      props.hidden = true;
    }
    if (row.styleId !== undefined) {
      props.s = indexOf(row.styleId);
    }
    if (Object.keys(props).length > 1) {
      out.push(props);
    }
  }
  return out;
}

/** The copied columns' own properties (a whole-column copy), one entry per
 * column a span gives a width, `hidden` or a style. */
function colPropsOf(sheet: Sheet, range: GridRange, indexOf: (id: number) => number): ClipboardColProps[] {
  const out: ClipboardColProps[] = [];
  for (const span of sheet.columns) {
    if (span.max < range.left || span.min > range.right || (span.widthPx === undefined && !span.hidden && span.styleId === undefined)) {
      continue;
    }
    for (let col = Math.max(span.min, range.left); col <= Math.min(span.max, range.right); col += 1) {
      const props: ClipboardColProps = { c: col - range.left };
      if (span.widthPx !== undefined) {
        props.widthPx = span.widthPx;
      }
      if (span.hidden) {
        props.hidden = true;
      }
      if (span.styleId !== undefined) {
        props.s = indexOf(span.styleId);
      }
      out.push(props);
    }
  }
  return out;
}

// --- the flavours ------------------------------------------------------------------------

/** A TSV field, quoted when it holds a tab, a line break or a quote. */
export function tsvField(text: string): string {
  return /[\t\n\r"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

/** Rows of fields as TSV: \t between fields, \r\n between rows, no trailing
 * newline. */
export function tsvText(rows: readonly (readonly string[])[]): string {
  return rows.map((fields) => fields.map(tsvField).join("\t")).join("\r\n");
}

function escapeHtml(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

/** A font family as a CSS value, without anything that could leave it. */
function cssFamily(family: string): string {
  return `'${family.replace(/['"\\;<>{}()&]/g, "")}'`;
}

const BORDER_CSS: Readonly<Record<CellBorderEdge["style"], string>> = {
  thin: "1px solid",
  medium: "2px solid",
  thick: "3px solid",
  dashed: "1px dashed",
  mediumDashed: "2px dashed",
  dashDot: "1px dashed",
  mediumDashDot: "2px dashed",
  dashDotDot: "1px dashed",
  mediumDashDotDot: "2px dashed",
  slantDashDot: "2px dashed",
  dotted: "1px dotted",
  hair: "1px dotted",
  double: "3px double",
};

const TEXT_ALIGN: Readonly<Record<string, string>> = { left: "left", center: "center", centerContinuous: "center", right: "right", justify: "justify", distributed: "justify" };
const VERTICAL_ALIGN: Readonly<Record<string, string>> = { top: "top", center: "middle", bottom: "bottom" };

/** A style as "HTML out"'s inline CSS (see the module header). */
export function cellCss(style: CellStyle, theme: DocumentTheme): string {
  const css: string[] = [];
  const font = style.font;
  if (font?.family !== undefined) {
    css.push(`font-family:${cssFamily(resolveFont(font.family, theme))}`);
  }
  if (font?.sizePt !== undefined) {
    css.push(`font-size:${font.sizePt}pt`);
  }
  if (font?.bold) {
    css.push("font-weight:700");
  }
  if (font?.italic) {
    css.push("font-style:italic");
  }
  const decorations = [...(font?.underline ? ["underline"] : []), ...(font?.strike ? ["line-through"] : [])];
  if (decorations.length > 0) {
    css.push(`text-decoration:${decorations.join(" ")}`);
  }
  if (font?.color !== undefined) {
    css.push(`color:${resolveColor(font.color, theme)}`);
  }
  if (style.fill?.kind === "solid") {
    css.push(`background-color:${resolveColor(style.fill.color, theme)}`);
  }
  const horizontal = TEXT_ALIGN[style.alignment?.horizontal ?? ""];
  if (horizontal) {
    css.push(`text-align:${horizontal}`);
  }
  const vertical = VERTICAL_ALIGN[style.alignment?.vertical ?? ""];
  if (vertical) {
    css.push(`vertical-align:${vertical}`);
  }
  if (style.alignment?.wrap) {
    css.push("white-space:pre-wrap");
  }
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const edge = style.borders?.[side];
    if (edge) {
      css.push(`border-${side}:${BORDER_CSS[edge.style]} ${edge.color === undefined ? "#000000" : resolveColor(edge.color, theme)}`);
    }
  }
  return css.join(";");
}

/** The two flavours of a copy (see the module header); `ctx` is the view's
 * format context (default: the workbook's in en-US). */
export function copyFlavours(model: SpreadsheetModel, sheetId: string, copy: ClipboardCopy, ctx?: FormatContext): { text: string; html: string } {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet) {
    throw new RangeError(`No sheet ${sheetId}`);
  }
  const format = { ...(ctx ?? workbookFormatContext(model, SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE])), widthPx: undefined };
  const theme = model.theme ?? DEFAULT_THEME;
  const { range } = copy.area;
  const merges = buildMergeIndex(sheet.merges);
  const codes = new Map<number, string | undefined>();
  const css = new Map<number, string>();
  const lines: string[][] = [];
  const html: string[] = [];
  let line: string[] = [];
  let currentRow = -1;
  forEachPosition(sheet, range, (row, col, cell, styleId) => {
    if (row !== currentRow) {
      if (currentRow >= 0) {
        lines.push(line);
        html.push("</tr>");
      }
      line = [];
      currentRow = row;
      html.push("<tr>");
    }
    const merge = mergeAt(merges, row, col);
    const anchor = !merge || (merge.top === row && merge.left === col);
    let shown = "";
    if (anchor && cell && cell.value !== null) {
      if (!codes.has(styleId)) {
        codes.set(styleId, formatCodeOf(cellStyle(model, styleId), format.locale));
      }
      shown = plainText(formatValue(cell.value, codes.get(styleId), format));
    }
    line.push(shown);
    if (!anchor) {
      return;
    }
    let style = css.get(styleId);
    if (style === undefined) {
      style = cellCss(cellStyle(model, styleId), theme);
      css.set(styleId, style);
    }
    const span =
      merge && merge.bottom <= range.bottom && merge.right <= range.right
        ? `${merge.right > merge.left ? ` colspan="${merge.right - merge.left + 1}"` : ""}${merge.bottom > merge.top ? ` rowspan="${merge.bottom - merge.top + 1}"` : ""}`
        : "";
    html.push(`<td${span}${style ? ` style="${escapeHtml(style)}"` : ""}>${escapeHtml(shown).replaceAll("\n", "<br>")}</td>`);
  });
  lines.push(line);
  html.push("</tr>");
  return {
    text: tsvText(lines),
    html: `<meta charset="utf-8"><table ${PAYLOAD_ATTRIBUTE}="${encodePayload(copy.payload)}">${html.join("")}</table>`,
  };
}

/** Both flavours of copying `range` (see the module header). */
export function clipboardFlavours(
  model: SpreadsheetModel,
  sheetId: string,
  range: GridRange,
  cutToken?: string,
  ctx?: FormatContext,
): { text: string; html: string } {
  return copyFlavours(model, sheetId, buildClipboardCopy(model, sheetId, range, cutToken), ctx);
}

// --- TSV in ------------------------------------------------------------------------------

/** Plain text as rows of fields (SPEC-125 "TSV in"): fields split at tabs,
 * rows at \r\n, \n or \r, a field starting with `"` quoted (embedded tabs,
 * line breaks and doubled quotes; anything after its closing quote kept as
 * written), and a trailing empty line ignored. "" is no rows. */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  const n = text.length;
  if (n === 0) {
    return rows;
  }
  let row: string[] = [];
  let i = 0;
  for (;;) {
    let field = "";
    if (text[i] === '"') {
      let at = i + 1;
      for (;;) {
        const quote = text.indexOf('"', at);
        if (quote < 0) {
          field += text.slice(at);
          at = n;
          break;
        }
        field += text.slice(at, quote);
        if (text[quote + 1] === '"') {
          field += '"';
          at = quote + 2;
        } else {
          at = quote + 1;
          break;
        }
      }
      i = at;
    }
    const start = i;
    while (i < n && text[i] !== "\t" && text[i] !== "\n" && text[i] !== "\r") {
      i += 1;
    }
    field += text.slice(start, i);
    row.push(field);
    if (i >= n) {
      rows.push(row);
      break;
    }
    if (text[i] === "\t") {
      i += 1;
      continue;
    }
    i += text[i] === "\r" && text[i + 1] === "\n" ? 2 : 1;
    rows.push(row);
    row = [];
    if (i >= n) {
      break;
    }
  }
  return rows;
}
