// SPEC-135's cell links in the workbook view, framework-free (useCellLinks
// holds the state): the link dialog's fields and validation, the link a cell
// carries, the hyperlink look an inserted link gets, and the DOM test for a
// press over a linked cell's text.
//
// - The dialog (SheetLinkDialog): `text` (the active cell's display text),
//   `kind` web | place, `address` (web), `sheet` + `ref` (place: a visible
//   sheet and a cell or range parseRange accepts) or `name` (place: one of
//   SPEC-133's visible names — hidden and `_xlnm.*` never listed — which wins
//   over sheet and ref while chosen), `tooltip`. linkDialogResult validates:
//   a dangerous scheme is refused inline with LINK_SCHEME_NOT_ALLOWED_MESSAGE
//   (LinkDialog.vue's string, the same refusal), a reference parseRange
//   rejects with LINK_REF_INVALID_MESSAGE; an empty address saves nothing.
//   A web address without a scheme that reads as a domain gets `https://`,
//   one that reads as an e-mail address `mailto:` (a sheet only opens http,
//   https and mailto). An empty text means "show the target" (LinkDialog's
//   rule).
// - linkAtCell: a cell's stored hyperlink (Sheet.hyperlinks, model order)
//   first, else a top-level HYPERLINK() formula (SPEC-127; its target is
//   computed when followed, never stored).
// - The hyperlink look (AC-35): the theme `hlink` colour and a single
//   underline, applied by the dialog's insert to the cells of the linked
//   range whose font colour is automatic (isAutomaticColor; by cell up to
//   HYPERLINK_STYLE_CELL_LIMIT cells, beyond it the whole range when the
//   active cell's colour is automatic).
// - linkTextEntry: the dialog's text, when it differs from what the cell
//   shows (displayTextAt), as SPEC-122's cell entry (protection's refusal or
//   a formula error first) — the same path a typed commit takes.
// - linkTextRef: a press's target inside `.sheet-cell-link` — the span
//   SheetCellLayer wraps a linked cell's text in — gives that cell's
//   reference (its element's data-ref: a merge's anchor, the cell whose text
//   overflows there); anywhere else in the cell gives null.
import { getCell, parseA1, parseRange, type CellRef, type GridRange } from "../../lib/spreadsheet/address";
import { isThemeColor, type ModelColor } from "../../lib/core/theme";
import { isDangerousUrl, sanitizeHref, urlScheme } from "../../lib/core/urls";
import { canonicalPlaceLocation, externalHref, hyperlinkAt, hyperlinkFormulaCall, hyperlinkFormulaTarget } from "../../lib/spreadsheet/annotations";
import { cellDisplay } from "../../lib/spreadsheet/cellRender";
import type { CellEntry } from "../../lib/spreadsheet/cells";
import type { StylePatch, StylePatchEntry } from "../../lib/spreadsheet/formatting";
import { quoteSheetName } from "../../lib/spreadsheet/formula/print";
import type { CellHyperlink, HyperlinkTarget, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { visibleNames } from "../../lib/spreadsheet/names";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { plainText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";
import { entryFromText, refusalFor } from "./cellEntry";
import type { CellItem } from "./sheetCellItems";

/** LinkDialog.vue's refusal of a dangerous scheme, word for word. */
export const LINK_SCHEME_NOT_ALLOWED_MESSAGE = "This address type isn't allowed in links";
export const LINK_REF_INVALID_MESSAGE = "Enter a cell or range, such as A1 or B2:C5";
/** The look an inserted link gives cells whose font colour is automatic. */
export const HYPERLINK_STYLE_PATCH: StylePatch = { font: { color: { theme: "hlink" }, underline: "single" } };
export const HYPERLINK_STYLE_CELL_LIMIT = 10_000;

export type LinkKind = "web" | "place";

export interface LinkDialogFields {
  text: string;
  kind: LinkKind;
  address: string;
  sheet: string;
  ref: string;
  name: string;
  tooltip: string;
}

/** What a valid dialog saves. */
export interface LinkDialogResult {
  text: string;
  target: HyperlinkTarget;
  tooltip: string;
}

export interface LinkOption {
  value: string;
  label: string;
}

/** A cell's link: stored (`link`) or a HYPERLINK() formula's (`formula`). */
export type CellLink = { kind: "stored"; link: CellHyperlink } | { kind: "formula" };

/** The link `at` carries on `sheet` (see the module header), or null. */
export function linkAtCell(sheet: Sheet, at: CellRef): CellLink | null {
  if (sheet.kind !== "grid") {
    return null;
  }
  const link = hyperlinkAt(sheet, at);
  if (link) {
    return { kind: "stored", link };
  }
  const cell = getCell(sheet, at.row, at.col);
  return cell && hyperlinkFormulaCall(cell, at) ? { kind: "formula" } : null;
}

/** A cell's link target now: stored, or its HYPERLINK() call evaluated. */
export function linkTargetAt(model: SpreadsheetModel, sheet: Sheet, at: CellRef): HyperlinkTarget | null {
  const link = linkAtCell(sheet, at);
  if (link?.kind === "stored") {
    return link.link.target;
  }
  const cell = link ? getCell(sheet, at.row, at.col) : undefined;
  return cell ? hyperlinkFormulaTarget(model, sheet.id, at, cell) : null;
}

/** What a cell shows (SPEC-118's display text, a fill run as one character). */
export function displayTextAt(model: SpreadsheetModel, sheet: Sheet, at: CellRef, ctx: FormatContext): string {
  const value = getCell(sheet, at.row, at.col)?.value ?? null;
  return value === null ? "" : plainText(cellDisplay(value, cellStyle(model, effectiveStyleId(sheet, at.row, at.col)), sheet, ctx));
}

/** The cell entry the dialog's `text` writes into `at` (null: the cell
 * already shows it), or why it cannot be written. */
export function linkTextEntry(model: SpreadsheetModel, sheetIndex: number, at: CellRef, text: string, locale: SpreadsheetLocale, ctx: FormatContext): { entry: CellEntry | null } | { error: string } {
  const sheet = model.sheets[sheetIndex];
  if (!sheet || text === displayTextAt(model, sheet, at, ctx)) {
    return { entry: null };
  }
  const refusal = refusalFor(model, sheetIndex, [{ top: at.row, left: at.col, bottom: at.row, right: at.col }]);
  if (refusal !== null) {
    return { error: refusal };
  }
  const outcome = entryFromText(text, { locale, dateSystem: ctx.dateSystem, style: cellStyle(model, effectiveStyleId(sheet, at.row, at.col)) });
  return "error" in outcome ? { error: outcome.error } : { entry: outcome.entry };
}

/** The dialog's sheet choices: the visible grid sheets' names (plus
 * `current` when it is none of them, so editing never retargets silently). */
export function linkSheetOptions(model: SpreadsheetModel, current = ""): string[] {
  const names = model.sheets.filter((sheet) => sheet.kind === "grid" && sheet.visibility === "visible").map((sheet) => sheet.name);
  return current !== "" && !names.includes(current) ? [...names, current] : names;
}

/** The dialog's name choices: SPEC-133's visible names as a place writes
 * them (a sheet-scoped one qualified with its sheet). */
export function linkNameOptions(model: SpreadsheetModel): LinkOption[] {
  return visibleNames(model)
    .map((name) => {
      const scope = name.scopeSheetId === undefined ? undefined : model.sheets.find((sheet) => sheet.id === name.scopeSheetId);
      const value = scope ? `${quoteSheetName(scope.name)}!${name.name}` : name.name;
      return { value, label: value };
    })
    .toSorted((a, b) => a.label.localeCompare(b.label, "en", { sensitivity: "base" }));
}

/** A sheet name as a location quotes it, unquoted again. */
function unquote(text: string): string {
  return text.startsWith("'") && text.endsWith("'") && text.length >= 2 ? text.slice(1, -1).replaceAll("''", "'") : text;
}

/** A `place` location as the dialog's sheet, ref and name: `Sheet!A1:B2`
 * → sheet + ref, anything else (a name, `#REF!`) → name. */
export function placeFields(location: string, hostSheetName: string): Pick<LinkDialogFields, "sheet" | "ref" | "name"> {
  const bang = location.lastIndexOf("!");
  const sheet = bang > 0 ? unquote(location.slice(0, bang)) : hostSheetName;
  const ref = bang > 0 ? location.slice(bang + 1) : location;
  return parseRange(ref) ? { sheet, ref, name: "" } : { sheet: hostSheetName, ref: "A1", name: location };
}

/** The dialog's opening fields: the cell's display text, and the link's
 * target when it has one (else a web link, the reference A1 on `sheetName`). */
export function linkDialogFields(displayText: string, sheetName: string, target: HyperlinkTarget | null, tooltip = ""): LinkDialogFields {
  const fields: LinkDialogFields = { text: displayText, kind: "web", address: "", sheet: sheetName, ref: "A1", name: "", tooltip };
  if (target?.kind === "url") {
    fields.address = externalHref(target);
  } else if (target?.kind === "place") {
    Object.assign(fields, { kind: "place" }, placeFields(target.location, sheetName));
  }
  return fields;
}

/** A typed web address as a link stores it (see the module header). */
export function normaliseAddress(address: string): string {
  const trimmed = address.trim();
  if (urlScheme(trimmed) !== null) {
    return trimmed;
  }
  if (/^[^\s@/:]+@[^\s@/]+\.[^\s@/]+$/.test(trimmed)) {
    return `mailto:${trimmed}`;
  }
  return /^(www\.)?[\w-]+(\.[\w-]+)+([/?#]|$)/i.test(trimmed) ? `https://${trimmed}` : trimmed;
}

/** The place a valid sheet and reference name, as canonical text. */
export function placeLocation(sheet: string, ref: string): string {
  return canonicalPlaceLocation(`${quoteSheetName(sheet)}!${ref.trim()}`);
}

/** The dialog's verdict: a result to save, an error to show, or null while
 * it is incomplete (Save disabled, nothing said). */
export function linkDialogResult(fields: LinkDialogFields): LinkDialogResult | { error: string } | null {
  let target: HyperlinkTarget;
  if (fields.kind === "web") {
    if (fields.address.trim() === "") {
      return null;
    }
    if (isDangerousUrl(fields.address)) {
      return { error: LINK_SCHEME_NOT_ALLOWED_MESSAGE };
    }
    const href = sanitizeHref(normaliseAddress(fields.address));
    if (href === null) {
      return { error: LINK_SCHEME_NOT_ALLOWED_MESSAGE };
    }
    target = { kind: "url", href };
  } else if (fields.name !== "") {
    target = { kind: "place", location: fields.name };
  } else {
    if (fields.ref.trim() === "" || fields.sheet === "") {
      return null;
    }
    if (!parseRange(fields.ref.trim())) {
      return { error: LINK_REF_INVALID_MESSAGE };
    }
    target = { kind: "place", location: placeLocation(fields.sheet, fields.ref) };
  }
  const shown = target.kind === "url" ? target.href : target.location;
  return { text: fields.text.trim() === "" ? shown : fields.text, target, tooltip: fields.tooltip.trim() };
}

/** The range a save links: an existing link's own range when the selection
 * is that link's single cell (Edit keeps its extent), else the selection's
 * active range. */
export function linkRange(active: GridRange, existing: CellHyperlink | null): GridRange {
  const single = active.top === active.bottom && active.left === active.right;
  return existing && single ? existing.range : active;
}

/** Whether a font colour reads as automatic: none, or the theme's Text 1
 * (`dk1`) untransformed — what a Normal style carries for it (Excel's
 * `theme="1"`, createBlankWorkbook's). */
export function isAutomaticColor(color: ModelColor | undefined): boolean {
  return color === undefined || (isThemeColor(color) && color.theme === "dk1" && [color.lumMod, color.lumOff, color.shade, color.tint].every((each) => each === undefined));
}

/** Whether a position's effective font colour is automatic. */
function automaticColor(model: SpreadsheetModel, sheet: Sheet, row: number, col: number): boolean {
  return isAutomaticColor(cellStyle(model, effectiveStyleId(sheet, row, col)).font?.color);
}

/** The hyperlink look's entries for `range` (see the module header). */
export function hyperlinkStyleEntries(model: SpreadsheetModel, sheet: Sheet, range: GridRange, active: CellRef): StylePatchEntry[] {
  const cells = (range.bottom - range.top + 1) * (range.right - range.left + 1);
  if (cells > HYPERLINK_STYLE_CELL_LIMIT) {
    return automaticColor(model, sheet, active.row, active.col) ? [{ range, patch: HYPERLINK_STYLE_PATCH }] : [];
  }
  const entries: StylePatchEntry[] = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (automaticColor(model, sheet, row, col)) {
        entries.push({ range: { top: row, left: col, bottom: row, right: col }, patch: HYPERLINK_STYLE_PATCH });
      }
    }
  }
  return entries;
}

/** The cell whose linked text a press landed on (see the module header). */
export function linkTextRef(target: EventTarget | null): CellRef | null {
  const text = target instanceof Element ? target.closest(".sheet-cell-link") : null;
  const ref = text?.closest('[data-testid="sheet-cell"]')?.getAttribute("data-ref");
  const at = ref ? parseA1(ref) : null;
  return at ? { row: at.row, col: at.col } : null;
}

/** Whether a cell element is linked (SheetCellLayer's `linkAt`, useCellLinks'
 * isLinked): data-link and the text's `.sheet-cell-link` span. A merge's
 * paint-only copy is not. */
export function isLinkedCell(cell: Pick<CellItem, "row" | "col" | "paintOnly">, at: unknown): boolean {
  return !cell.paintOnly && typeof at === "function" && (at as (row: number, col: number) => boolean)(cell.row, cell.col) === true;
}
