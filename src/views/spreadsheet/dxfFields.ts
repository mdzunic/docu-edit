// Dxf <-> FormatCellsFields conversions (SPEC-129 "The rule editor") — the
// seam useFormatCellsDialog's differential mode uses to show a Dxf (a
// conditional format's partial style) in the same Number/Font/Border/Fill
// tabs SPEC-123 built for a cell's own style, and to read one back out.
// Framework-free and pure.
//
// - `dxfAsStyle` seeds the dialog's fields (fieldsOf reads a CellStyle);
//   `dxfTouchedFields` says which of them the loaded Dxf already set, so
//   re-saving an edited rule without touching the Font tab keeps its colour,
//   weight and underline rather than silently dropping them.
// - `fieldsAsDxf` reads them back: the tri-state font flags (SPEC-129
//   "Differential styles" — `bold: false` un-bolds, `underline: "none"`
//   removes one) only for a touched field, exactly as fieldsPatch does for a
//   cell's own style; a fill or a border edge is read unconditionally,
//   because the dialog's own fields already carry the loaded Dxf's value
//   when the user leaves them alone (there is no "untouched" fill/border
//   state to lose).
import type { DialogEdge, FormatCellsFields } from "../../lib/spreadsheet/formatCells";
import { fillOf } from "../../lib/spreadsheet/formatCells";
import type { CellBorderEdge, CellStyle } from "../../lib/spreadsheet/model";
import type { Dxf, DxfFont } from "../../lib/spreadsheet/conditionalFormats";

/** A Dxf's font/fill/borders/numberFormat as a CellStyle, for fieldsOf to
 * read — alignment and protection are absent (differential mode hides those
 * tabs), and a tri-state "false"/"none" removal reads the same as absent
 * (fieldsOf shows both as "Regular"/"None", which is correct either way). */
export function dxfAsStyle(dxf: Dxf): CellStyle {
  const style: CellStyle = {};
  if (dxf.font) {
    style.font = {
      ...(dxf.font.color !== undefined ? { color: dxf.font.color } : {}),
      ...(dxf.font.bold === true ? { bold: true } : {}),
      ...(dxf.font.italic === true ? { italic: true } : {}),
      ...(dxf.font.strike === true ? { strike: true } : {}),
      ...(dxf.font.underline !== undefined && dxf.font.underline !== "none" ? { underline: dxf.font.underline } : {}),
    };
  }
  if (dxf.fill !== undefined) {
    style.fill = dxf.fill;
  }
  if (dxf.borders !== undefined) {
    style.borders = dxf.borders;
  }
  if (dxf.numberFormat !== undefined) {
    style.numberFormat = dxf.numberFormat;
  }
  return style;
}

/** The FormatCellsFields keys an already-loaded `dxf` set — pre-seeded as
 * "touched" so saving without revisiting the Font tab preserves them. */
export function dxfTouchedFields(dxf: Dxf): Set<keyof FormatCellsFields> {
  const touched = new Set<keyof FormatCellsFields>();
  if (dxf.font?.color !== undefined) {
    touched.add("fontColor");
  }
  if (dxf.font?.bold !== undefined || dxf.font?.italic !== undefined) {
    touched.add("fontStyle");
  }
  if (dxf.font?.underline !== undefined) {
    touched.add("underline");
  }
  if (dxf.font?.strike !== undefined) {
    touched.add("strike");
  }
  return touched;
}

/** The four edges a Dxf may carry (no diagonal, no inside — those have no
 * meaning for a single differential cell). */
export const DIFFERENTIAL_EDGES = ["top", "right", "bottom", "left"] as const;

/** The dialog's current fields/edges as a Dxf (useFormatCellsDialog's
 * differential `ok()`); `edgeShown` is the dialog's own (the session's
 * change, else the loaded Dxf's edge), so borders need no touched-tracking
 * either. */
export function fieldsAsDxf(
  fields: FormatCellsFields,
  touched: ReadonlySet<keyof FormatCellsFields>,
  edgeShown: (edge: DialogEdge) => CellBorderEdge | null,
  numberFormat: string | undefined,
  numberTouched: boolean,
): Dxf {
  const dxf: Dxf = {};
  const font: DxfFont = {};
  let fontTouched = false;
  if (touched.has("fontColor")) {
    if (fields.fontColor !== null) {
      font.color = fields.fontColor;
    }
    fontTouched = true;
  }
  if (touched.has("fontStyle")) {
    font.bold = fields.fontStyle === "bold" || fields.fontStyle === "boldItalic";
    font.italic = fields.fontStyle === "italic" || fields.fontStyle === "boldItalic";
    fontTouched = true;
  }
  if (touched.has("underline")) {
    font.underline = fields.underline;
    fontTouched = true;
  }
  if (touched.has("strike")) {
    font.strike = fields.strike;
    fontTouched = true;
  }
  if (fontTouched) {
    dxf.font = font;
  }
  const fill = fillOf(fields);
  if (fill !== null) {
    dxf.fill = fill;
  }
  const borders: Dxf["borders"] = {};
  let hasBorder = false;
  for (const edge of DIFFERENTIAL_EDGES) {
    const value = edgeShown(edge);
    if (value !== null) {
      borders[edge] = value;
      hasBorder = true;
    }
  }
  if (hasBorder) {
    dxf.borders = borders;
  }
  if (numberTouched && numberFormat !== undefined) {
    dxf.numberFormat = numberFormat;
  }
  return dxf;
}
