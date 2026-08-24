// DrawingML (`a:` namespace) shared, product-neutral core (SPEC-079's
// two-phase move): the OOXML/ODF geometry-name tables and the pure,
// Element-reading helpers that both the `.docx` shape reader/writer
// (src/lib/docx/shape.ts, src/lib/odf/shape.ts — Phase one re-imports these)
// and the `.pptx`/`.odp` presentation readers (Phase two consumers) need.
// Everything here takes DOM `Element`s and returns plain data — no document
// model, no Vue, no format-layer types — so it can live in `core/` and be
// imported by both products (boundary.spec.ts).
import type { ShapeGeometry } from "./shapes";
import { SHAPE_DEFINITIONS, guideValueToAdjustment } from "./shapeCatalogue";

/** The DrawingML main namespace (`a:*` elements), shared by every OOXML
 * family — `.docx` drawings and `.pptx` slides alike. */
export const DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";

// The geometry-name maps, all four derived from the single SHAPE_DEFINITIONS
// table (SPEC-069 AC-15a) rather than maintained beside it. Lines are dropped
// from the ODF reverse map (a `draw:line` element carries no enhanced
// geometry, so only area shapes look up by `draw:type`).
/** `ShapeGeometry` → OOXML preset name (`a:prstGeom/@prst`). */
export const GEOMETRY_TO_PRST: Record<string, string> = Object.fromEntries(
  SHAPE_DEFINITIONS.map((def) => [def.geometry, def.prst]),
);
/** OOXML preset name → `ShapeGeometry`. */
export const PRST_TO_GEOMETRY: Record<string, ShapeGeometry> = Object.fromEntries(
  SHAPE_DEFINITIONS.map((def) => [def.prst, def.geometry]),
);
/** `ShapeGeometry` → ODF enhanced-geometry type (`draw:type`). */
export const GEOMETRY_TO_DRAW_TYPE: Record<string, string> = Object.fromEntries(
  SHAPE_DEFINITIONS.map((def) => [def.geometry, def.drawType]),
);
/** ODF enhanced-geometry type → `ShapeGeometry` (area shapes only). */
export const DRAW_TYPE_TO_GEOMETRY: Record<string, ShapeGeometry> = Object.fromEntries(
  SHAPE_DEFINITIONS.filter((def) => !def.isLine).map((def) => [def.drawType, def.geometry]),
);

/** The adjustment fractions carried by an `a:prstGeom/a:avLst`: every
 * `a:gd@fmla="val N"` in document order, normalised from the guide value to a
 * fraction (SPEC-069 AC-6a). Guides with any other formula are ignored (only
 * literal adjust values are modelled). */
export function parseAvLstAdjustments(prstGeom: Element): number[] {
  const avLst = firstA(prstGeom, "avLst");
  if (!avLst) {
    return [];
  }
  const adjustments: number[] = [];
  for (const gd of Array.from(avLst.children)) {
    if (gd.namespaceURI !== DRAWINGML_NAMESPACE || gd.localName !== "gd") {
      continue;
    }
    const fmla = gd.getAttribute("fmla") ?? "";
    const match = /^val\s+(-?\d+)$/.exec(fmla.trim());
    if (match) {
      adjustments.push(guideValueToAdjustment(Number(match[1])));
    }
  }
  return adjustments;
}

/** A DrawingML `a:xfrm` resolved to EMU offsets/extents and whole-degree
 * rotation. Absent `a:off`/`a:ext` yield `undefined` coordinates so a caller
 * can fall back to an inherited placeholder transform. */
export interface DrawingXfrm {
  xEmu?: number;
  yEmu?: number;
  cxEmu?: number;
  cyEmu?: number;
  rotationDeg: number;
}

/** OOXML rotation (`a:xfrm/@rot`) is 60000ths of a degree, clockwise. */
function ooxmlRotToDeg(rot: number): number {
  return Math.round(rot / 60000);
}

/** Reads an `a:xfrm` (from a `p:spPr`/`p:grpSpPr`/`a:graphicFrame`) into EMU
 * numbers. Returns `undefined` when the element carries no `a:xfrm` at all. */
export function parseXfrm(xfrm: Element | undefined): DrawingXfrm | undefined {
  if (!xfrm) {
    return undefined;
  }
  const off = firstA(xfrm, "off");
  const ext = firstA(xfrm, "ext");
  const rotAttr = xfrm.getAttribute("rot");
  const num = (value: string | null): number | undefined => {
    if (value === null) {
      return undefined;
    }
    const n = Number(value);
    return Number.isFinite(n) ? n : undefined;
  };
  return {
    xEmu: num(off?.getAttribute("x") ?? null),
    yEmu: num(off?.getAttribute("y") ?? null),
    cxEmu: num(ext?.getAttribute("cx") ?? null),
    cyEmu: num(ext?.getAttribute("cy") ?? null),
    rotationDeg: rotAttr ? ooxmlRotToDeg(Number(rotAttr)) : 0,
  };
}

/** The child `a:*` chain of a `p:grpSpPr`'s `a:xfrm` a group uses to map its
 * children into slide space: the group's own bounds (`a:off`/`a:ext`) and the
 * child coordinate space (`a:chOff`/`a:chExt`) children are authored in. */
export interface GroupChildMapping {
  xEmu: number;
  yEmu: number;
  cxEmu: number;
  cyEmu: number;
  chXEmu: number;
  chYEmu: number;
  chCxEmu: number;
  chCyEmu: number;
}

/** Parses a group's `a:xfrm` (with `a:chOff`/`a:chExt`) into the offset/extent
 * mapping that turns a child's authored EMU coordinates into slide-space EMU
 * (pptx `chOff`/`chExt`, AC-37). Returns `undefined` when the extents are
 * degenerate, so a caller renders children at their own coordinates. */
export function parseGroupChildMapping(grpSpPr: Element): GroupChildMapping | undefined {
  const xfrm = firstA(grpSpPr, "xfrm");
  if (!xfrm) {
    return undefined;
  }
  const off = firstA(xfrm, "off");
  const ext = firstA(xfrm, "ext");
  const chOff = firstA(xfrm, "chOff");
  const chExt = firstA(xfrm, "chExt");
  const n = (el: Element | undefined, attr: string): number => Number(el?.getAttribute(attr) ?? "NaN");
  const mapping: GroupChildMapping = {
    xEmu: n(off, "x"),
    yEmu: n(off, "y"),
    cxEmu: n(ext, "cx"),
    cyEmu: n(ext, "cy"),
    chXEmu: n(chOff, "x"),
    chYEmu: n(chOff, "y"),
    chCxEmu: n(chExt, "cx"),
    chCyEmu: n(chExt, "cy"),
  };
  if (!Number.isFinite(mapping.chCxEmu) || !Number.isFinite(mapping.chCyEmu) || mapping.chCxEmu === 0 || mapping.chCyEmu === 0) {
    return undefined;
  }
  return mapping;
}

/** Maps one child EMU coordinate/extent through a group's mapping into
 * slide-space EMU (the `chOff`/`chExt` → `off`/`ext` linear map). */
export function mapGroupChild(
  mapping: GroupChildMapping,
  childXEmu: number,
  childYEmu: number,
  childCxEmu: number,
  childCyEmu: number,
): { xEmu: number; yEmu: number; cxEmu: number; cyEmu: number } {
  const sx = mapping.cxEmu / mapping.chCxEmu;
  const sy = mapping.cyEmu / mapping.chCyEmu;
  return {
    xEmu: mapping.xEmu + (childXEmu - mapping.chXEmu) * sx,
    yEmu: mapping.yEmu + (childYEmu - mapping.chYEmu) * sy,
    cxEmu: childCxEmu * sx,
    cyEmu: childCyEmu * sy,
  };
}

/** First `a:{localName}` child of `parent`, or `undefined`. A local
 * convenience so callers do not each import the generic xml helper with the
 * DrawingML namespace spelled out. */
export function firstA(parent: Element, localName: string): Element | undefined {
  for (const child of Array.from(parent.children)) {
    if (child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === localName) {
      return child;
    }
  }
  return undefined;
}

/** All `a:{localName}` children of `parent`, in document order. */
export function childrenA(parent: Element, localName: string): Element[] {
  return Array.from(parent.children).filter(
    (child) => child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === localName,
  );
}
