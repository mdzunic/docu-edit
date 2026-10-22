// The Format Cells dialog's model (SPEC-123 "Format Cells"): its fields read
// from a style, and the ONE StylePatch (plus border entries) its OK applies —
// built from the fields the user changed only, so a dialog opened on a mixed
// selection never flattens it: an untouched field keeps each cell's own
// value. Pure, so the dialog's composable (useFormatCellsDialog) stays thin.
//
// - Rotation: the dialog's −90…90 degrees map to xlsx's 0…180 (a negative
//   angle d is 90 − d) and Vertical text to 255.
// - Borders: each edge the user set or removed becomes per-position patches
//   as in formatting.ts's presets — the outline on the range's boundary
//   cells, an inside edge on both cells sharing it (multi-cell selections
//   only), a diagonal on every cell — with the facing edge of each stored
//   neighbour outside a changed outer edge removed.
import type { GridRange } from "./address";
import type { StylePatch, StylePatchEntry } from "./formatting";
import { MAX_COLS, MAX_ROWS, type CellAlignment, type CellBorderEdge, type CellFill, type CellFont, type CellStyle, type SheetBorderStyle } from "./model";
import type { ModelColor } from "../core/theme";

export type FontStyleChoice = "regular" | "italic" | "bold" | "boldItalic";
export type UnderlineChoice = "none" | NonNullable<CellFont["underline"]>;
export type DialogEdge = "top" | "insideHorizontal" | "bottom" | "left" | "insideVertical" | "right" | "diagonalUp" | "diagonalDown";

/** The 13 line styles in the Border tab's order. */
export const LINE_STYLES: readonly SheetBorderStyle[] = [
  "hair",
  "dotted",
  "dashDotDot",
  "dashDot",
  "dashed",
  "thin",
  "mediumDashDotDot",
  "slantDashDot",
  "mediumDashDot",
  "mediumDashed",
  "medium",
  "thick",
  "double",
];

/** `solid` and the 17 xlsx hatch patterns, in Excel's order. */
export const PATTERN_STYLES: readonly string[] = [
  "solid",
  "gray0625",
  "gray125",
  "lightGray",
  "mediumGray",
  "darkGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
];

export const HORIZONTAL_CHOICES: readonly { value: NonNullable<CellAlignment["horizontal"]>; label: string }[] = [
  { value: "general", label: "General" },
  { value: "left", label: "Left (indent)" },
  { value: "center", label: "Centre" },
  { value: "right", label: "Right (indent)" },
  { value: "fill", label: "Fill" },
  { value: "justify", label: "Justify" },
  { value: "centerContinuous", label: "Centre across selection" },
  { value: "distributed", label: "Distributed (indent)" },
];

export const VERTICAL_CHOICES: readonly { value: NonNullable<CellAlignment["vertical"]>; label: string }[] = [
  { value: "top", label: "Top" },
  { value: "center", label: "Centre" },
  { value: "bottom", label: "Bottom" },
  { value: "justify", label: "Justify" },
  { value: "distributed", label: "Distributed" },
];

/** The dialog's alignment, font, fill and protection fields. */
export interface FormatCellsFields {
  horizontal: NonNullable<CellAlignment["horizontal"]>;
  vertical: NonNullable<CellAlignment["vertical"]>;
  indent: number;
  wrap: boolean;
  shrink: boolean;
  /** −90…90. */
  rotation: number;
  verticalText: boolean;
  fontFamily: string;
  fontStyle: FontStyleChoice;
  fontSize: number;
  underline: UnderlineChoice;
  fontColor: ModelColor | null;
  strike: boolean;
  superscript: boolean;
  subscript: boolean;
  fillColor: ModelColor | null;
  patternStyle: string;
  patternColor: ModelColor | null;
  locked: boolean;
  hidden: boolean;
}

/** xlsx textRotation → the dialog's degrees. */
export function rotationDegrees(textRotation: number | undefined): number {
  if (textRotation === undefined || textRotation === 255) {
    return 0;
  }
  return textRotation <= 90 ? textRotation : 90 - textRotation;
}

/** The dialog's degrees → xlsx textRotation (null for none). */
export function textRotationOf(degrees: number, verticalText: boolean): number | null {
  if (verticalText) {
    return 255;
  }
  const d = Math.max(-90, Math.min(90, Math.round(degrees)));
  return d === 0 ? null : d > 0 ? d : 90 - d;
}

/** A style's fields as the dialog opens on them; `family` is the resolved
 * font family name. */
export function fieldsOf(style: CellStyle, family: string, normalSizePt: number): FormatCellsFields {
  const alignment = style.alignment ?? {};
  const font = style.font ?? {};
  const fill = style.fill;
  return {
    horizontal: alignment.horizontal ?? "general",
    vertical: alignment.vertical ?? "bottom",
    indent: alignment.indent ?? 0,
    wrap: alignment.wrap === true,
    shrink: alignment.shrinkToFit === true,
    rotation: rotationDegrees(alignment.textRotation),
    verticalText: alignment.textRotation === 255,
    fontFamily: family,
    fontStyle: font.bold && font.italic ? "boldItalic" : font.bold ? "bold" : font.italic ? "italic" : "regular",
    fontSize: font.sizePt ?? normalSizePt,
    underline: font.underline ?? "none",
    fontColor: font.color ?? null,
    strike: font.strike === true,
    superscript: font.vertAlign === "superscript",
    subscript: font.vertAlign === "subscript",
    fillColor: fill?.kind === "solid" ? fill.color : fill?.kind === "pattern" ? (fill.bg ?? null) : null,
    patternStyle: fill?.kind === "pattern" ? fill.pattern : "solid",
    patternColor: fill?.kind === "pattern" ? (fill.fg ?? null) : null,
    locked: style.protection?.locked !== false,
    hidden: style.protection?.hidden === true,
  };
}

/** The fill the Fill tab's fields describe (null: no fill). */
export function fillOf(fields: Pick<FormatCellsFields, "fillColor" | "patternStyle" | "patternColor">): CellFill | null {
  if (fields.patternStyle !== "solid") {
    const fill: CellFill = { kind: "pattern", pattern: fields.patternStyle };
    if (fields.patternColor !== null) {
      fill.fg = fields.patternColor;
    }
    if (fields.fillColor !== null) {
      fill.bg = fields.fillColor;
    }
    return fill;
  }
  return fields.fillColor === null ? null : { kind: "solid", color: fields.fillColor };
}

/** The patch for the changed alignment, font, fill and protection fields. */
export function fieldsPatch(fields: FormatCellsFields, touched: ReadonlySet<keyof FormatCellsFields>): StylePatch {
  const patch: StylePatch = {};
  const alignment: NonNullable<StylePatch["alignment"]> = {};
  const font: NonNullable<StylePatch["font"]> = {};
  const has = (field: keyof FormatCellsFields): boolean => touched.has(field);
  if (has("horizontal")) {
    alignment.horizontal = fields.horizontal === "general" ? null : fields.horizontal;
  }
  if (has("vertical")) {
    alignment.vertical = fields.vertical === "bottom" ? null : fields.vertical;
  }
  if (has("indent")) {
    alignment.indent = fields.indent > 0 ? Math.min(250, Math.round(fields.indent)) : null;
  }
  if (has("wrap")) {
    alignment.wrap = fields.wrap ? true : null;
  }
  if (has("shrink")) {
    alignment.shrinkToFit = fields.shrink ? true : null;
  }
  if (has("rotation") || has("verticalText")) {
    alignment.textRotation = textRotationOf(fields.rotation, fields.verticalText);
  }
  if (has("fontFamily")) {
    font.family = fields.fontFamily;
  }
  if (has("fontStyle")) {
    font.bold = fields.fontStyle === "bold" || fields.fontStyle === "boldItalic" ? true : null;
    font.italic = fields.fontStyle === "italic" || fields.fontStyle === "boldItalic" ? true : null;
  }
  if (has("fontSize")) {
    font.sizePt = Math.min(409, Math.max(1, fields.fontSize));
  }
  if (has("underline")) {
    font.underline = fields.underline === "none" ? null : fields.underline;
  }
  if (has("fontColor")) {
    font.color = fields.fontColor;
  }
  if (has("strike")) {
    font.strike = fields.strike ? true : null;
  }
  if (has("superscript") || has("subscript")) {
    font.vertAlign = fields.superscript ? "superscript" : fields.subscript ? "subscript" : null;
  }
  if (Object.keys(alignment).length > 0) {
    patch.alignment = alignment;
  }
  if (Object.keys(font).length > 0) {
    patch.font = font;
  }
  if (has("fillColor") || has("patternStyle") || has("patternColor")) {
    patch.fill = fillOf(fields);
  }
  if (has("locked") || has("hidden")) {
    patch.protection = {};
    if (has("locked")) {
      patch.protection.locked = fields.locked ? null : false;
    }
    if (has("hidden")) {
      patch.protection.hidden = fields.hidden ? true : null;
    }
  }
  return patch;
}

/** The entries the Border tab's changed edges apply to one range (see the
 * module header); `edges` holds only the edges the user set (an edge) or
 * removed (null), and both diagonals when either changed. */
export function borderEdgeEntries(range: GridRange, edges: Partial<Record<DialogEdge, CellBorderEdge | null>>): StylePatchEntry[] {
  const entries: StylePatchEntry[] = [];
  const add = (target: GridRange, borders: NonNullable<StylePatch["borders"]>, storedOnly = false): void => {
    entries.push(storedOnly ? { range: target, patch: { borders }, storedOnly: true } : { range: target, patch: { borders } });
  };
  const outer: Record<"top" | "bottom" | "left" | "right", { band: GridRange; neighbour: GridRange | null; facing: "top" | "bottom" | "left" | "right" }> = {
    top: { band: { ...range, bottom: range.top }, neighbour: range.top > 0 ? { ...range, top: range.top - 1, bottom: range.top - 1 } : null, facing: "bottom" },
    bottom: {
      band: { ...range, top: range.bottom },
      neighbour: range.bottom < MAX_ROWS - 1 ? { ...range, top: range.bottom + 1, bottom: range.bottom + 1 } : null,
      facing: "top",
    },
    left: { band: { ...range, right: range.left }, neighbour: range.left > 0 ? { ...range, left: range.left - 1, right: range.left - 1 } : null, facing: "right" },
    right: {
      band: { ...range, left: range.right },
      neighbour: range.right < MAX_COLS - 1 ? { ...range, left: range.right + 1, right: range.right + 1 } : null,
      facing: "left",
    },
  };
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const edge = edges[side];
    if (edge === undefined) {
      continue;
    }
    add(outer[side].band, { [side]: edge });
    const neighbour = outer[side].neighbour;
    if (neighbour) {
      add(neighbour, { [outer[side].facing]: null }, true);
    }
  }
  const inH = edges.insideHorizontal;
  if (inH !== undefined && range.bottom > range.top) {
    add({ ...range, bottom: range.bottom - 1 }, { bottom: inH });
    add({ ...range, top: range.top + 1 }, { top: inH });
  }
  const inV = edges.insideVertical;
  if (inV !== undefined && range.right > range.left) {
    add({ ...range, right: range.right - 1 }, { right: inV });
    add({ ...range, left: range.left + 1 }, { left: inV });
  }
  // The two diagonals are one model edge: the caller passes both directions
  // when either changed (an untouched one at its current state).
  const up = edges.diagonalUp;
  const down = edges.diagonalDown;
  if (up !== undefined || down !== undefined) {
    const line = up ?? down;
    add(range, { diagonal: line ? { ...line, ...(up ? { up: true as const } : {}), ...(down ? { down: true as const } : {}) } : null });
  }
  return entries;
}
