// The workbook generation API's style vocabulary (SPEC-140): `CellStyleSpec`,
// a validated subset of SPEC-117's `CellStyle` with colours as `#rrggbb`
// only, mapped field for field onto SPEC-123's `StylePatch` — so a builder
// style and an editor style mean the same thing: both reach the model
// through formatting.ts's applyStylePatch.
//
// - `fill` is a solid fill (`{ kind: "solid", color }`), `rotation` is
//   `textRotation`, a border edge is a SheetBorderStyle or `{ style, color? }`
//   (no colour = automatic).
// - Booleans map to `true`, and `false` to the patch's `null` (remove), so a
//   later `bold: false` undoes an earlier `bold: true`; an `indent` or
//   `rotation` of 0 removes it too (patchStyle stores neither).
// - Borders: a named edge (`top`, `right`, `bottom`, `left`) is CellStyle's
//   — every cell's own edge — while the shorthands `all` set every edge of
//   every cell in the range and `outline` the range's outer edges only. The
//   shorthands apply first (all, then outline, then the named edges), so a
//   named edge wins where they meet.
// - Number format codes are validated by SPEC-118's parseFormatCode
//   (en-US code syntax, as the model stores them); "General" removes one.
//
// Imports lib/spreadsheet and generate's product-neutral validators only
// (workbookBoundary.spec.ts).
import type { GridRange } from "../spreadsheet/address";
import type { StylePatch } from "../spreadsheet/formatting";
import type { CellBorderEdge, CellFont, SheetBorderStyle } from "../spreadsheet/model";
import { FormatCodeError, parseFormatCode } from "../spreadsheet/numberFormat/parse";
import { DocumentGenerationError } from "./errors";
import {
  expectBoolean,
  expectEnum,
  expectHexColor,
  expectIntInRange,
  expectNumber,
  expectPlainObject,
  expectString,
  rejectUnknownKeys,
} from "./validate";

/** A border edge: a style alone (automatic colour) or with a `#rrggbb` colour. */
export type BorderSpec = SheetBorderStyle | { style: SheetBorderStyle; color?: string };

export interface CellStyleSpec {
  font?: {
    family?: string;
    sizePt?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: "single" | "double";
    strike?: boolean;
    color?: string;
  };
  /** A solid fill, `#rrggbb`. */
  fill?: string;
  borders?: Partial<Record<"top" | "right" | "bottom" | "left" | "all" | "outline", BorderSpec>>;
  alignment?: {
    horizontal?: "left" | "center" | "right" | "justify";
    vertical?: "top" | "center" | "bottom";
    wrap?: boolean;
    /** Indent levels, 0–250. */
    indent?: number;
    /** SPEC-117's `textRotation`: 1–90 degrees counter-clockwise, 91–180 the
     * same clockwise (90 + degrees), 255 stacked vertically; 0 none. */
    rotation?: number;
  };
  numberFormat?: string;
}

export const SHEET_BORDER_STYLES: readonly SheetBorderStyle[] = [
  "thin",
  "medium",
  "thick",
  "dashed",
  "dotted",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
];

export const STYLE_SPEC_KEYS = ["font", "fill", "borders", "alignment", "numberFormat"] as const;
export const FONT_SPEC_KEYS = ["family", "sizePt", "bold", "italic", "underline", "strike", "color"] as const;
export const BORDER_SPEC_KEYS = ["top", "right", "bottom", "left", "all", "outline"] as const;
export const BORDER_EDGE_KEYS = ["style", "color"] as const;
export const ALIGNMENT_SPEC_KEYS = ["horizontal", "vertical", "wrap", "indent", "rotation"] as const;

/** The largest font size Excel accepts. */
export const MAX_FONT_SIZE_PT = 409;
/** The deepest indent Excel's Format Cells dialog accepts. */
export const MAX_INDENT = 250;
/** SPEC-117's stacked-vertical text rotation. */
export const STACKED_ROTATION = 255;

/** A number format code SPEC-118 parses, as the model stores it. */
export function expectNumberFormat(value: unknown, path: string): string {
  const code = expectString(value, path);
  if (code.trim() === "") {
    throw new DocumentGenerationError(`${path}: expected a number format code`);
  }
  try {
    parseFormatCode(code);
  } catch (error) {
    if (error instanceof FormatCodeError) {
      throw new DocumentGenerationError(`${path}: invalid number format "${code}": ${error.message}`);
    }
    throw error;
  }
  return code;
}

function borderSpec(value: unknown, path: string): BorderSpec {
  if (typeof value === "string") {
    return expectEnum(value, SHEET_BORDER_STYLES, path);
  }
  const obj = expectPlainObject(value, path);
  rejectUnknownKeys(obj, BORDER_EDGE_KEYS, path);
  const edge: { style: SheetBorderStyle; color?: string } = { style: expectEnum(obj.style, SHEET_BORDER_STYLES, `${path}.style`) };
  if (obj.color !== undefined) {
    edge.color = expectHexColor(obj.color, `${path}.color`);
  }
  return edge;
}

/** Validates a style spec at `path` and returns a normalised copy (colours
 * lower-case) — the caller's object is never kept. */
export function validateStyleSpec(value: unknown, path: string): CellStyleSpec {
  const obj = expectPlainObject(value, path);
  rejectUnknownKeys(obj, STYLE_SPEC_KEYS, path);
  const spec: CellStyleSpec = {};
  if (obj.font !== undefined) {
    const fontPath = `${path}.font`;
    const font = expectPlainObject(obj.font, fontPath);
    rejectUnknownKeys(font, FONT_SPEC_KEYS, fontPath);
    const out: NonNullable<CellStyleSpec["font"]> = {};
    if (font.family !== undefined) {
      out.family = expectString(font.family, `${fontPath}.family`);
      if (out.family.trim() === "") {
        throw new DocumentGenerationError(`${fontPath}.family: expected a font family name`);
      }
    }
    if (font.sizePt !== undefined) {
      const size = expectNumber(font.sizePt, `${fontPath}.sizePt`);
      if (size < 1 || size > MAX_FONT_SIZE_PT) {
        throw new DocumentGenerationError(`${fontPath}.sizePt: expected a size between 1 and ${MAX_FONT_SIZE_PT} pt`);
      }
      out.sizePt = size;
    }
    for (const key of ["bold", "italic", "strike"] as const) {
      if (font[key] !== undefined) {
        out[key] = expectBoolean(font[key], `${fontPath}.${key}`);
      }
    }
    if (font.underline !== undefined) {
      out.underline = expectEnum(font.underline, ["single", "double"] as const, `${fontPath}.underline`);
    }
    if (font.color !== undefined) {
      out.color = expectHexColor(font.color, `${fontPath}.color`);
    }
    spec.font = out;
  }
  if (obj.fill !== undefined) {
    spec.fill = expectHexColor(obj.fill, `${path}.fill`);
  }
  if (obj.borders !== undefined) {
    const bordersPath = `${path}.borders`;
    const borders = expectPlainObject(obj.borders, bordersPath);
    rejectUnknownKeys(borders, BORDER_SPEC_KEYS, bordersPath);
    const out: NonNullable<CellStyleSpec["borders"]> = {};
    for (const key of BORDER_SPEC_KEYS) {
      if (borders[key] !== undefined) {
        out[key] = borderSpec(borders[key], `${bordersPath}.${key}`);
      }
    }
    spec.borders = out;
  }
  if (obj.alignment !== undefined) {
    const alignmentPath = `${path}.alignment`;
    const alignment = expectPlainObject(obj.alignment, alignmentPath);
    rejectUnknownKeys(alignment, ALIGNMENT_SPEC_KEYS, alignmentPath);
    const out: NonNullable<CellStyleSpec["alignment"]> = {};
    if (alignment.horizontal !== undefined) {
      out.horizontal = expectEnum(alignment.horizontal, ["left", "center", "right", "justify"] as const, `${alignmentPath}.horizontal`);
    }
    if (alignment.vertical !== undefined) {
      out.vertical = expectEnum(alignment.vertical, ["top", "center", "bottom"] as const, `${alignmentPath}.vertical`);
    }
    if (alignment.wrap !== undefined) {
      out.wrap = expectBoolean(alignment.wrap, `${alignmentPath}.wrap`);
    }
    if (alignment.indent !== undefined) {
      out.indent = expectIntInRange(alignment.indent, `${alignmentPath}.indent`, 0, MAX_INDENT);
    }
    if (alignment.rotation !== undefined) {
      const rotationPath = `${alignmentPath}.rotation`;
      const rotation = expectNumber(alignment.rotation, rotationPath);
      if (rotation !== STACKED_ROTATION) {
        expectIntInRange(rotation, rotationPath, 0, 180);
      }
      out.rotation = rotation;
    }
    spec.alignment = out;
  }
  if (obj.numberFormat !== undefined) {
    spec.numberFormat = expectNumberFormat(obj.numberFormat, `${path}.numberFormat`);
  }
  return spec;
}

/** A boolean spec field as a patch value: true, or null (remove). */
const flag = (value: boolean): true | null => (value ? true : null);

function edgeOf(spec: BorderSpec): CellBorderEdge {
  return typeof spec === "string" ? { style: spec } : spec.color === undefined ? { style: spec.style } : { style: spec.style, color: spec.color };
}

/** One applyStylePatch call: a patch over ranges. */
export interface StylePatchCall {
  ranges: GridRange[];
  patch: StylePatch;
}

/** The applyStylePatch calls a spec makes over `range`, in order: `all`,
 * then `outline` (one call per outer edge), then everything else with the
 * named edges (see the module header). Empty for an empty spec. */
export function styleSpecPatches(spec: CellStyleSpec, range: GridRange): StylePatchCall[] {
  const calls: StylePatchCall[] = [];
  const borders = spec.borders ?? {};
  if (borders.all !== undefined) {
    const edge = edgeOf(borders.all);
    calls.push({ ranges: [range], patch: { borders: { top: edge, right: edge, bottom: edge, left: edge } } });
  }
  if (borders.outline !== undefined) {
    const edge = edgeOf(borders.outline);
    calls.push(
      { ranges: [{ ...range, bottom: range.top }], patch: { borders: { top: edge } } },
      { ranges: [{ ...range, top: range.bottom }], patch: { borders: { bottom: edge } } },
      { ranges: [{ ...range, right: range.left }], patch: { borders: { left: edge } } },
      { ranges: [{ ...range, left: range.right }], patch: { borders: { right: edge } } },
    );
  }
  const patch: StylePatch = {};
  if (spec.font !== undefined) {
    const font: NonNullable<StylePatch["font"]> = {};
    const { family, sizePt, bold, italic, underline, strike, color } = spec.font;
    if (family !== undefined) {
      font.family = family;
    }
    if (sizePt !== undefined) {
      font.sizePt = sizePt;
    }
    if (bold !== undefined) {
      font.bold = flag(bold);
    }
    if (italic !== undefined) {
      font.italic = flag(italic);
    }
    if (underline !== undefined) {
      font.underline = underline satisfies NonNullable<CellFont["underline"]>;
    }
    if (strike !== undefined) {
      font.strike = flag(strike);
    }
    if (color !== undefined) {
      font.color = color;
    }
    if (Object.keys(font).length > 0) {
      patch.font = font;
    }
  }
  if (spec.fill !== undefined) {
    patch.fill = { kind: "solid", color: spec.fill };
  }
  const edges: NonNullable<StylePatch["borders"]> = {};
  for (const key of ["top", "right", "bottom", "left"] as const) {
    const edge = borders[key];
    if (edge !== undefined) {
      edges[key] = edgeOf(edge);
    }
  }
  if (Object.keys(edges).length > 0) {
    patch.borders = edges;
  }
  if (spec.alignment !== undefined) {
    const alignment: NonNullable<StylePatch["alignment"]> = {};
    const { horizontal, vertical, wrap, indent, rotation } = spec.alignment;
    if (horizontal !== undefined) {
      alignment.horizontal = horizontal;
    }
    if (vertical !== undefined) {
      alignment.vertical = vertical;
    }
    if (wrap !== undefined) {
      alignment.wrap = flag(wrap);
    }
    if (indent !== undefined) {
      alignment.indent = indent;
    }
    if (rotation !== undefined) {
      alignment.textRotation = rotation;
    }
    if (Object.keys(alignment).length > 0) {
      patch.alignment = alignment;
    }
  }
  if (spec.numberFormat !== undefined) {
    patch.numberFormat = spec.numberFormat;
  }
  if (Object.keys(patch).length > 0) {
    calls.push({ ranges: [range], patch });
  }
  return calls;
}
