// The .ods sheet drawing layer on the wire (SPEC-131 "Reading ODS drawings",
// "Writing" ods, "Canonical new drawings", "Media parts"): THIS FILE is the
// normative source for how a table's `table:shapes` children and its cells'
// `draw:*` children become Sheet.drawings, how a dirty table writes them
// back, and the reader upgrade's lift of drawings out of Cell.preservedXml.
// drawings.ts owns the model's geometry and commands; ods/drawingGolden.ts
// fixes the canonical forms.
//
// Reading (reader.ts captures each element's exact source slice — the
// element alone, without the whitespace around it — and hands it here):
// - Where. A `table:shapes` child → `absolute`: `posEmu` from svg:x/svg:y,
//   `extEmu` from svg:width/svg:height. A `draw:*` (or `dr3d:scene`) child of
//   a cell → anchored to that cell (a repeated cell, or a cell in a repeated
//   row, at its first column and row). Its svg:x/svg:y are SHEET-ABSOLUTE,
//   so `from`'s offset is svg:x − colLeft(col) (rows likewise) under the
//   read-time GridMetrics, clamped at 0 per axis and otherwise kept as read;
//   `table:end-cell-address` + `table:end-x`/`table:end-y` → `to`, a twoCell
//   anchor with `editAs: "twoCell"`; without it the anchor is `oneCell` with
//   `extEmu` from svg:width/svg:height.
// - Geometry. A `draw:line` (`draw:connector`, `draw:measure`) spans
//   svg:x1/y1–x2/y2; a `draw:g` is the union of its members' boxes; a
//   `draw:transform` made of `rotate` and `translate` steps (applied in
//   order, as LibreOffice writes and reads them) places the element's box
//   by its centre — `rotate(a)` is counter-clockwise radians, so
//   `rotationDeg` = −Σa in clockwise degrees; any other step (scale, skew,
//   matrix) leaves the transform unread (svg:x/svg:y, no rotation).
// - Order. `draw:z-index` orders the sheet's array (index 0 at the back); a
//   drawing without one goes in front of every drawing that has one, in
//   document order (`table:shapes` first, then the cells row by row). Ids
//   are minted "dr1", "dr2"… in that order.
// - What. `draw:frame` > `draw:object` → `chart` (`sourcePart` "Object
//   N/content.xml" from its xlink:href); > `draw:image` (the first one) →
//   `image` (`Pictures/…` → one MediaPart per package path, shared, mediaId
//   = the path, its type the manifest's media-type else the extension's,
//   its bytes deferred with registerPendingMedia; `fo:clip` in the frame's
//   graphic style → `crop`, as fractions of the picture's natural size at 96
//   dpi; an href outside the package → foreign "Linked picture"); >
//   `draw:object-ole` → foreign "Object" (its replacement picture the face);
//   > `draw:text-box` → a `rect` shape with the box's text.
//   `draw:custom-shape` → a shape by `draw:enhanced-geometry@draw:type`
//   (DRAW_TYPE_TO_GEOMETRY; `non-primitive` or unknown → `rect` with
//   `unsupportedGeometryXml`, the dashed placeholder), `draw:modifiers` →
//   adjustments (shapeCatalogue.ts's scale). `draw:rect`, `draw:ellipse`,
//   `draw:line` → `rect`, `ellipse`, `line` (`arrow` when its style ends in
//   a marker). `draw:g` → `group`, each member's box a fraction of the
//   group's (a nested group or a chart member → foreign). `draw:control` →
//   foreign "Form control", `dr3d:scene` → "3-D object", anything else →
//   "Object".
// - Style. The element's automatic graphic style (`draw:style-name`,
//   flattened with its parents and styles.xml's graphic default-style):
//   `draw:fill` none → no fill, else a solid `draw:fill-color` (a gradient,
//   hatch or bitmap paints as that colour; LibreOffice's #729fcf when none
//   is set); `draw:stroke` none → no outline, else `svg:stroke-color`
//   (#3465a4 when unset), `svg:stroke-width`, dash for "dash".
// - Text. The shape's (or text box's) `text:p`/`text:h` paragraphs →
//   DrawingText, each paragraph's alignment from its paragraph style, the
//   size, colour and weight from the first run (its span's text style, else
//   the paragraph's, else the graphic style's), the vertical anchor from
//   `draw:textarea-vertical-align` (default middle; top for a text box).
// - Alt text = the `svg:desc` child, name = `draw:name` ("" when absent).
//   `sourceXml` = the element verbatim.
//
// Writing (a DIRTY table — a clean one is its source verbatim): every
// drawing is placed by its behaviour — `fixed` in the table's
// `table:shapes`, `moveAndSize` and `move` in their `from` cell (tableXml.ts
// materialises the cell and splits a repeat run around it).
// - An imported drawing is its sourceXml through the retained-text pass,
//   PATCHED — nothing inside the element is regenerated. Each attribute
//   group is read back from the source the way the reader reads it, under
//   the write-time geometry and in the model's anchor cell, and rewritten
//   only when that read-back differs from the model by more than
//   PATCH_TOLERANCE_EMU (a tenth of a pixel: rounding never rewrites, a
//   one-pixel nudge always does) — and a source start before its anchor
//   cell's own start (read clamped: the cell moved under an edit, or the
//   source was inconsistent) counts as moved: the start position (svg:x/svg:y, or the
//   last `translate` of a draw:transform — appended when none — or a line's
//   svg:x1…y2, or a group's own translate), the size (svg:width/svg:height:
//   a oneCell or absolute anchor's extEmu; a twoCell's from→to box, only
//   once its anchor moved), the end cell (`table:end-cell-address` and
//   `table:end-x`/`-y`: a moveAndSize anchor's `to`; a move or fixed
//   anchor's recomputed point — drawings.ts writtenPoints, the source's
//   kept within a pixel), `draw:z-index`, `draw:name` and the `svg:desc`
//   child. A length is written in the unit the source used for it. The end
//   address keeps the source's spelling of the sheet name ($, quotes) while
//   it names the table's own sheet; otherwise — the sheet was renamed and
//   no pass rewrote it — the model's current name, quoted per ODF. A group
//   moves by its own translate; its size is not patched.
// - z-index: walking the array back to front, a drawing keeps its source
//   value while that is above every value written before it, else takes
//   one more than the last; a new drawing takes one more than the last; an
//   imported drawing without one keeps none while every drawing in front of
//   it lacks one too.
// - A drawing created here is written in the canonical form of
//   ods/drawingGolden.ts with a generated `de-gr{n}` graphic style (the
//   table writer's style registry, spliced into the envelope's automatic
//   styles like SPEC-122's `de-ce` styles). Its picture path is the media
//   part's (OdsMediaPaths re-paths anything outside `Pictures/`). A created
//   group, chart or foreign drawing has no wire form and is not written
//   (SPEC-121's cross-format copy drops them first); a created picture's
//   `crop` is not written.
//
// The reader upgrade (open.ts, SPEC-131's step) lifts the drawing elements
// out of a dirty stored sheet's cells' preservedXml — anchored to each
// cell's current position — and reads table:shapes from its retained
// sourceXml (upgradeOdsSheetDrawings).
import { decodeImageDimensions } from "../../core/imageDimensions";
import { DRAW_TYPE_TO_GEOMETRY, GEOMETRY_TO_DRAW_TYPE } from "../../core/drawingMl";
import { contentTypeForPartName, extensionForContentType, registerPendingMedia, type MediaPart } from "../../core/media";
import { DRAW_NAMESPACE, odfColorToCss, odfLengthToCm, SVG_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE } from "../../core/odfXml";
import { adjustmentToGuideValue, guideValueToAdjustment } from "../../core/shapeCatalogue";
import { LINE_GEOMETRIES, normalizeRotationDeg, shapeFillBaseColor, type ShapeFill, type ShapeGeometry, type ShapeOutline, type ShapeVisual } from "../../core/shapes";
import { DEFAULT_THEME, resolveColor, type DocumentTheme } from "../../core/theme";
import { CM_PER_INCH, POINTS_PER_INCH, cmToEmu, emuToCm } from "../../core/units";
import { XmlParseError } from "../../core/xml";
import { scanXml, type XmlScanAttribute } from "../../core/xmlScanner";
import { colToLetters } from "../address";
import { anchorBehaviour, drawingRectPx, EMU_PER_PX, writtenPoints } from "../drawings";
import { sheetNameNeedsQuotes } from "../formula/print";
import { buildGridMetrics, type GridMetrics } from "../geometry";
import { MAX_COLS, MAX_ROWS, type AnchorPoint, type Cell, type DrawingAnchor, type DrawingText, type Row, type Sheet, type SheetDrawing, type SheetDrawingContent } from "../model";
import { applySplices, decodeXmlText, escapeAttributeValue, escapeElementText, setAttribute, type Splice } from "../xmlSplice";
import { readManifest } from "./package";
import { parseOdsRangeAddress } from "./rangeAddress";
import { namespaceDeclarations, odsStyleBookFromSource, parseOdsPart, type OdsStyleBook } from "./styles";

const DR3D_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";

/** How far a read-back value may sit from the model's and still count as
 * unchanged: a tenth of a pixel (see the module header). */
export const PATCH_TOLERANCE_EMU = Math.round(EMU_PER_PX / 10);

/** LibreOffice's default shape fill and line colours (its graphic
 * default-style), used when a style sets none. */
const DEFAULT_FILL_COLOR = "#729fcf";
const DEFAULT_STROKE_COLOR = "#3465a4";

/** The canonical prefix for each namespace a drawing's geometry is read
 * from — what an attribute or element is matched by, whatever prefix the
 * producer bound (a standalone slice, whose prefixes are unbound, is
 * matched by the prefix as written). */
const CANONICAL_PREFIX: ReadonlyMap<string, string> = new Map([
  [DRAW_NAMESPACE, "draw"],
  [SVG_NAMESPACE, "svg"],
  [TABLE_NAMESPACE, "table"],
  [TEXT_NAMESPACE, "text"],
  [XLINK_NAMESPACE, "xlink"],
  [DR3D_NAMESPACE, "dr3d"],
]);

/** "prefix:local" with the canonical prefix of the name's namespace. */
function canonicalName(name: string, local: string, ns: string | undefined): string {
  const prefix = ns === undefined ? undefined : CANONICAL_PREFIX.get(ns);
  return prefix === undefined ? name : `${prefix}:${local}`;
}

function attrOf(attrs: readonly XmlScanAttribute[], qname: string): string | undefined {
  return attrs.find((attr) => canonicalName(attr.name, attr.local, attr.ns) === qname)?.value;
}

/** An ODF length → EMU (undefined when it is not one). */
function lengthEmu(value: string | undefined): number | undefined {
  const cm = odfLengthToCm(value);
  return cm === undefined ? undefined : cmToEmu(cm);
}

/** Whether an element (by its canonical name) is a drawing — a shape a
 * cell or `table:shapes` holds. */
export function isDrawingElementName(name: string): boolean {
  return (name.startsWith("draw:") && name !== "draw:glue-point") || name === "dr3d:scene";
}

/** Whether a scanned element (by namespace) is a drawing. */
export function isDrawingElement(ns: string | undefined, local: string): boolean {
  return (ns === DRAW_NAMESPACE && local !== "glue-point") || (ns === DR3D_NAMESPACE && local === "scene");
}

// --- geometry: the shared read-back ---------------------------------------------------------

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** One step of a draw:transform and where its argument text sits in the
 * attribute's (decoded) value. */
interface TransformStep {
  kind: "rotate" | "translate";
  values: number[];
  /** The raw argument texts (a translate's lengths keep their units). */
  args: string[];
  start: number;
  end: number;
}

/** A drawing element's geometry as its own text spells it — the read-back
 * the reader and the writer share, so "unchanged" means exactly "reads back
 * as the model holds it". Offsets index into the text it was read from. */
interface ElementGeometry {
  /** The element's canonical name ("draw:frame") and its name as written. */
  name: string;
  qname: string;
  attrs: XmlScanAttribute[];
  /** The unrotated box, sheet-absolute EMU (unrounded). */
  box: Box;
  /** How the position is spelt. */
  position: "xy" | "line" | "transform" | "group";
  rotationDeg?: number;
  /** The draw:transform's steps, when it is read (position "transform", or a group's own). */
  transform?: TransformStep[];
  /** The box before the transform: svg:x/svg:y (0 when absent) — or a group's member union. */
  local: Box;
  end?: { raw: string; col: number; row: number; sheet?: string; xEmu: number; yEmu: number };
  z?: number;
  drawName?: string;
  desc?: { start: number; end: number; textStart: number; textEnd: number; text: string };
  titleEnd?: number;
  /** Where the start tag ends and the close tag starts (−1: self-closing). */
  tagEnd: number;
  closeStart: number;
  /** A group's members (their geometry, offsets into the same text). */
  members: { geometry: ElementGeometry; start: number; end: number }[];
}

const TRANSFORM_STEP = /([A-Za-z]+)\s*\(([^)]*)\)/g;

/** A draw:transform's steps, or null when it holds anything but rotate and
 * translate (see the module header). */
function transformSteps(value: string): TransformStep[] | null {
  const steps: TransformStep[] = [];
  let rest = value;
  for (const match of value.matchAll(TRANSFORM_STEP)) {
    const kind = match[1]!.toLowerCase();
    const args = match[2]!.trim().split(/[\s,]+/).filter((arg) => arg !== "");
    const start = match.index;
    rest = rest.replace(match[0], "");
    if (kind === "rotate" && args.length === 1 && Number.isFinite(Number(args[0]))) {
      steps.push({ kind: "rotate", values: [Number(args[0])], args, start, end: start + match[0].length });
    } else if (kind === "translate" && (args.length === 1 || args.length === 2)) {
      const values = args.map((arg) => lengthEmu(arg) ?? (Number(arg) === 0 ? 0 : Number.NaN));
      if (values.some((each) => !Number.isFinite(each))) {
        return null;
      }
      steps.push({ kind: "translate", values: values.length === 1 ? [values[0]!, 0] : values, args, start, end: start + match[0].length });
    } else {
      return null;
    }
  }
  return rest.trim() === "" ? steps : null;
}

/** ODF's counter-clockwise rotation of a point by `a` radians (y down). */
function rotateCcw(x: number, y: number, a: number): [number, number] {
  return [x * Math.cos(a) + y * Math.sin(a), -x * Math.sin(a) + y * Math.cos(a)];
}

/** Where the steps take a point. */
function applySteps(steps: readonly TransformStep[], x: number, y: number): [number, number] {
  let at: [number, number] = [x, y];
  for (const step of steps) {
    at = step.kind === "rotate" ? rotateCcw(at[0], at[1], step.values[0]!) : [at[0] + step.values[0]!, at[1] + step.values[1]!];
  }
  return at;
}

/** The box a transform puts `local` at: its centre moved, its size kept. */
function transformedBox(local: Box, steps: readonly TransformStep[]): Box {
  const [cx, cy] = applySteps(steps, local.x + local.w / 2, local.y + local.h / 2);
  return { x: cx - local.w / 2, y: cy - local.h / 2, w: local.w, h: local.h };
}

const unionOf = (boxes: readonly Box[]): Box => {
  if (boxes.length === 0) {
    return { x: 0, y: 0, w: 0, h: 0 };
  }
  const left = Math.min(...boxes.map((box) => box.x));
  const top = Math.min(...boxes.map((box) => box.y));
  const right = Math.max(...boxes.map((box) => box.x + box.w));
  const bottom = Math.max(...boxes.map((box) => box.y + box.h));
  return { x: left, y: top, w: right - left, h: bottom - top };
};

/** The end cell of `table:end-cell-address`, or undefined. */
function endOf(attrs: readonly XmlScanAttribute[]): ElementGeometry["end"] {
  const raw = attrOf(attrs, "table:end-cell-address");
  const ref = raw === undefined ? null : parseOdsRangeAddress(raw.trim());
  if (!ref || ref.start.col === undefined || ref.start.row === undefined) {
    return undefined;
  }
  const end: NonNullable<ElementGeometry["end"]> = {
    raw: raw!,
    col: Math.min(ref.start.col, MAX_COLS - 1),
    row: Math.min(ref.start.row, MAX_ROWS - 1),
    xEmu: Math.max(0, lengthEmu(attrOf(attrs, "table:end-x")) ?? 0),
    yEmu: Math.max(0, lengthEmu(attrOf(attrs, "table:end-y")) ?? 0),
  };
  if (ref.sheet !== undefined) {
    end.sheet = ref.sheet;
  }
  return end;
}

/** Reads a drawing element's geometry from its text (`start` the index of
 * its start tag). Undefined when the text holds no element. */
function elementGeometryAt(xml: string, start: number): ElementGeometry | undefined {
  let geometry: ElementGeometry | undefined;
  let depth = 0;
  let childStart = -1;
  let childName = "";
  let descStart = -1;
  let descTextStart = -1;
  let descText = "";
  scanXml(xml.slice(start), (event) => {
    if (event.type === "text") {
      if (descStart >= 0) {
        descText += event.text;
      }
      return undefined;
    }
    if (event.type === "open") {
      const name = canonicalName(event.name, event.local, event.ns);
      if (depth === 0) {
        geometry = {
          name,
          qname: event.name,
          attrs: event.attrs,
          box: { x: 0, y: 0, w: 0, h: 0 },
          position: "xy",
          local: { x: 0, y: 0, w: 0, h: 0 },
          tagEnd: start + event.end,
          closeStart: -1,
          members: [],
        };
        if (event.selfClosing) {
          return "stop";
        }
        depth = 1;
        return undefined;
      }
      if (depth === 1) {
        if (name === "svg:desc") {
          if (event.selfClosing) {
            geometry!.desc = { start: start + event.start, end: start + event.end, textStart: start + event.end, textEnd: start + event.end, text: "" };
          } else {
            descStart = start + event.start;
            descTextStart = start + event.end;
            descText = "";
          }
        } else if (name === "svg:title" && event.selfClosing) {
          geometry!.titleEnd = start + event.end;
        }
        if (event.selfClosing) {
          if (geometry!.name === "draw:g" && isDrawingElementName(name)) {
            const member = elementGeometryAt(xml, start + event.start);
            if (member) {
              geometry!.members.push({ geometry: member, start: start + event.start, end: start + event.end });
            }
          }
        } else {
          childStart = start + event.start;
          childName = name;
        }
      }
      if (!event.selfClosing) {
        depth += 1;
      }
      return undefined;
    }
    depth -= 1;
    if (depth === 0) {
      geometry!.closeStart = start + event.start;
      return "stop";
    }
    if (depth === 1) {
      const end = start + event.end;
      if (descStart >= 0) {
        geometry!.desc = { start: descStart, end, textStart: descTextStart, textEnd: start + event.start, text: descText };
        descStart = -1;
      } else if (childName === "svg:title") {
        geometry!.titleEnd = end;
      } else if (geometry!.name === "draw:g" && isDrawingElementName(childName)) {
        const member = elementGeometryAt(xml, childStart);
        if (member) {
          geometry!.members.push({ geometry: member, start: childStart, end });
        }
      }
      childStart = -1;
    }
    return undefined;
  });
  if (!geometry) {
    return undefined;
  }
  placeGeometry(geometry);
  return geometry;
}

/** Fills in a scanned element's box, rotation, end cell, z-index and name. */
function placeGeometry(geometry: ElementGeometry): void {
  const attrs = geometry.attrs;
  const num = (qname: string): number => lengthEmu(attrOf(attrs, qname)) ?? 0;
  geometry.end = endOf(attrs);
  const z = Number(attrOf(attrs, "draw:z-index"));
  if (attrOf(attrs, "draw:z-index") !== undefined && Number.isInteger(z)) {
    geometry.z = z;
  }
  const drawName = attrOf(attrs, "draw:name");
  if (drawName !== undefined) {
    geometry.drawName = drawName;
  }
  const transform = attrOf(attrs, "draw:transform");
  const steps = transform === undefined ? null : transformSteps(transform);
  if (geometry.name === "draw:line" || geometry.name === "draw:connector" || geometry.name === "draw:measure") {
    const [x1, y1, x2, y2] = [num("svg:x1"), num("svg:y1"), num("svg:x2"), num("svg:y2")];
    geometry.position = "line";
    geometry.local = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
    geometry.box = geometry.local;
    return;
  }
  if (geometry.name === "draw:g") {
    geometry.position = "group";
    geometry.local = unionOf(geometry.members.map((member) => member.geometry.box));
  } else {
    geometry.local = { x: num("svg:x"), y: num("svg:y"), w: Math.max(0, num("svg:width")), h: Math.max(0, num("svg:height")) };
  }
  if (steps && steps.length > 0) {
    geometry.transform = steps;
    if (geometry.position === "xy") {
      geometry.position = "transform";
    }
    geometry.box = transformedBox(geometry.local, steps);
    const radians = steps.reduce((sum, step) => sum + (step.kind === "rotate" ? step.values[0]! : 0), 0);
    const degrees = normalizeRotationDeg((-radians * 180) / Math.PI);
    if (degrees !== 0) {
      geometry.rotationDeg = degrees;
    }
  } else {
    geometry.box = geometry.local;
  }
}

/** A drawing element's geometry from its own text. */
function elementGeometry(xml: string): ElementGeometry | undefined {
  const start = xml.indexOf("<");
  return start < 0 ? undefined : elementGeometryAt(xml, start);
}

/** The offset into a cell `startPx` from the sheet's edge of a sheet-absolute EMU position (clamped at 0). */
const offsetInto = (emu: number, startPx: number): number => Math.max(0, Math.round(emu - startPx * EMU_PER_PX));

/** The anchor the reader gives an element of `geometry` in `cell` (absent:
 * a `table:shapes` child) under `metrics` (see the module header). */
function anchorOf(geometry: ElementGeometry, cell: { row: number; col: number } | undefined, metrics: GridMetrics): DrawingAnchor {
  const { x, y, w, h } = geometry.box;
  const extEmu = { cx: Math.max(0, Math.round(w)), cy: Math.max(0, Math.round(h)) };
  if (!cell) {
    return { kind: "absolute", posEmu: { x: Math.round(x), y: Math.round(y) }, extEmu };
  }
  const from: AnchorPoint = { col: cell.col, colOffEmu: offsetInto(x, metrics.colLeft(cell.col)), row: cell.row, rowOffEmu: offsetInto(y, metrics.rowTop(cell.row)) };
  const end = geometry.end;
  if (end) {
    return { kind: "twoCell", editAs: "twoCell", from, to: { col: end.col, colOffEmu: end.xEmu, row: end.row, rowOffEmu: end.yEmu } };
  }
  return { kind: "oneCell", from, extEmu };
}

// --- media -----------------------------------------------------------------------------------

/** The package's pictures as the reader meets them: one MediaPart per
 * package path, created on first reference with its bytes deferred
 * (registerPendingMedia), its type the manifest's media-type else the
 * extension's. */
export class OdsMediaReader {
  private readonly parts: ReadonlyMap<string, Uint8Array>;
  private readonly types = new Map<string, string>();
  private readonly found = new Map<string, MediaPart>();
  private readonly sizes = new Map<string, { widthPx: number; heightPx: number } | null>();

  /** `parts`: the package entries; `known`: media parts already in the
   * model (an upgrade's), reused by path. */
  constructor(parts: ReadonlyMap<string, Uint8Array>, known: readonly MediaPart[] = []) {
    this.parts = parts;
    const manifest = parts.get("META-INF/manifest.xml");
    if (manifest) {
      try {
        for (const entry of readManifest(new TextDecoder().decode(manifest))) {
          if (entry.mediaType !== "") {
            this.types.set(entry.fullPath, entry.mediaType);
          }
        }
      } catch {
        // A malformed manifest types nothing: the extension does.
      }
    }
    for (const part of known) {
      this.found.set(part.mediaId, part);
    }
  }

  /** The media part at `path`, or undefined when the package has none. */
  part(path: string): MediaPart | undefined {
    let part = this.found.get(path);
    if (!part) {
      const bytes = this.parts.get(path);
      if (!bytes) {
        return undefined;
      }
      part = { mediaId: path, contentType: this.types.get(path) ?? contentTypeForPartName(path) ?? "application/octet-stream", data: "" };
      registerPendingMedia(part, bytes);
      this.found.set(path, part);
    }
    return part;
  }

  /** A picture's natural size in px (96 dpi), or null when unreadable. */
  naturalSize(path: string): { widthPx: number; heightPx: number } | null {
    if (!this.sizes.has(path)) {
      const bytes = this.parts.get(path);
      const part = this.part(path);
      this.sizes.set(path, bytes && part ? decodeImageDimensions(bytes, part.contentType) : null);
    }
    return this.sizes.get(path) ?? null;
  }

  /** Every part met (and every known one), in first-reference order. */
  media(): MediaPart[] {
    return [...this.found.values()];
  }
}

// --- reading -------------------------------------------------------------------------------

/** One drawing element as the reader's scan met it: its exact source slice
 * and, for an element inside a cell, that cell; absent for a `table:shapes`
 * child. */
export interface OdsDrawingSource {
  xml: string;
  cell?: { row: number; col: number };
}

/** The style lookups a drawing's content needs (OdsStyleBook). */
export type OdsDrawingStyles = Pick<OdsStyleBook, "graphicProperties" | "styleFont" | "paragraphAlign">;

/** What reading a sheet's drawings needs. */
export interface OdsDrawingReadContext {
  /** The document element's namespace declarations: a slice parses on its own under them. */
  declarations: string;
  styles: OdsDrawingStyles;
  media: OdsMediaReader;
}

const drawAttr = (el: Element, local: string): string | null => el.getAttributeNS(DRAW_NAMESPACE, local);

function elementChildren(el: Element): Element[] {
  return Array.from(el.children);
}

function firstChild(el: Element, ns: string, local: string): Element | undefined {
  return elementChildren(el).find((child) => child.namespaceURI === ns && child.localName === local);
}

/** A paragraph's text: text:s, text:tab and text:line-break expanded,
 * spans and links walked, anything else (a frame, a note) skipped. */
function paragraphText(el: Element): string {
  let text = "";
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === 3 || node.nodeType === 4) {
      text += node.nodeValue ?? "";
    } else if (node.nodeType === 1) {
      const child = node as Element;
      if (child.namespaceURI !== TEXT_NAMESPACE) {
        continue;
      }
      if (child.localName === "s") {
        const count = Number(child.getAttributeNS(TEXT_NAMESPACE, "c") ?? 1);
        text += " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 1000) : 1);
      } else if (child.localName === "tab") {
        text += "\t";
      } else if (child.localName === "line-break") {
        text += "\n";
      } else if (child.localName !== "note" && child.localName !== "annotation") {
        text += paragraphText(child);
      }
    }
  }
  return text;
}

/** The paragraphs (`text:p`, `text:h`, through lists) directly inside a container. */
function paragraphsOf(container: Element): Element[] {
  const out: Element[] = [];
  for (const child of elementChildren(container)) {
    if (child.namespaceURI !== TEXT_NAMESPACE) {
      continue;
    }
    if (child.localName === "p" || child.localName === "h") {
      out.push(child);
    } else if (child.localName === "list" || child.localName === "list-item") {
      out.push(...paragraphsOf(child));
    }
  }
  return out;
}

function firstSpan(el: Element): Element | undefined {
  for (const child of elementChildren(el)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "span") {
      return child;
    }
  }
  return undefined;
}

const ALIGN: Readonly<Record<string, "left" | "center" | "right">> = { start: "left", left: "left", center: "center", end: "right", right: "right" };

/** A container's paragraphs as DrawingText (see the module header), or undefined when they hold no text. */
function drawingText(container: Element, graphicStyle: string | null, ctx: OdsDrawingReadContext, defaultAnchor: DrawingText["anchor"]): DrawingText | undefined {
  const paragraphs = paragraphsOf(container);
  const read = paragraphs.map((p) => {
    const align = ALIGN[ctx.styles.paragraphAlign(p.getAttributeNS(TEXT_NAMESPACE, "style-name") ?? undefined) ?? ""];
    return align ? { text: paragraphText(p), align } : { text: paragraphText(p) };
  });
  if (read.every((p) => p.text === "")) {
    return undefined;
  }
  const vertical = ctx.styles.graphicProperties(graphicStyle ?? undefined).get("draw:textarea-vertical-align");
  const text: DrawingText = { paragraphs: read, anchor: vertical === "top" ? "top" : vertical === "bottom" ? "bottom" : vertical === "middle" || vertical === "justify" ? "middle" : defaultAnchor };
  const first = paragraphs.find((p) => paragraphText(p) !== "");
  const span = first ? firstSpan(first) : undefined;
  const font =
    ctx.styles.styleFont("text", span?.getAttributeNS(TEXT_NAMESPACE, "style-name") ?? undefined) ??
    ctx.styles.styleFont("paragraph", first?.getAttributeNS(TEXT_NAMESPACE, "style-name") ?? undefined) ??
    ctx.styles.styleFont("graphic", graphicStyle ?? undefined);
  if (font?.sizePt !== undefined) {
    text.sizePt = font.sizePt;
  }
  if (font?.color !== undefined) {
    text.color = font.color;
  }
  if (font?.bold) {
    text.bold = true;
  }
  return text;
}

/** A graphic style's fill and outline (see the module header). */
function styleVisual(props: ReadonlyMap<string, string>, line: boolean): Pick<ShapeVisual, "fill" | "outline"> {
  const visual: Pick<ShapeVisual, "fill" | "outline"> = {};
  const fill = props.get("draw:fill") ?? (line ? "none" : "solid");
  if (fill !== "none" && !line) {
    visual.fill = { kind: "solid", color: odfColorToCss(props.get("draw:fill-color")) ?? DEFAULT_FILL_COLOR };
  }
  const stroke = props.get("draw:stroke") ?? "solid";
  if (stroke !== "none") {
    const widthCm = odfLengthToCm(props.get("svg:stroke-width")) ?? 0;
    visual.outline = {
      color: odfColorToCss(props.get("svg:stroke-color")) ?? DEFAULT_STROKE_COLOR,
      widthPt: (widthCm / CM_PER_INCH) * POINTS_PER_INCH,
      dash: stroke === "dash" ? "dash" : "solid",
    };
  }
  return visual;
}

const LINK = /^(?:[a-z][a-z0-9+.-]*:|\/|\.\.\/)/i;

/** A package-relative href ("./Pictures/x.png" → "Pictures/x.png"), or null for a link outside the package. */
function packagePath(href: string | null): string | null {
  if (href === null || href.trim() === "") {
    return null;
  }
  const path = href.trim().replace(/^\.\//, "");
  return LINK.test(path) ? null : path;
}

/** `fo:clip="rect(top, right, bottom, left)"` → crop fractions of the picture's natural size (96 dpi). */
function cropOf(clip: string | undefined, naturalSize: () => { widthPx: number; heightPx: number } | null): { l: number; t: number; r: number; b: number } | undefined {
  const match = clip === undefined ? null : /^\s*rect\s*\(([^)]*)\)\s*$/.exec(clip);
  const sides = match ? match[1]!.split(/[\s,]+/).filter((each) => each !== "") : [];
  if (sides.length !== 4) {
    return undefined;
  }
  const [top, right, bottom, left] = sides.map((side) => (side === "auto" ? 0 : (odfLengthToCm(side) ?? 0)));
  if (top === 0 && right === 0 && bottom === 0 && left === 0) {
    return undefined;
  }
  const natural = naturalSize();
  if (!natural || natural.widthPx <= 0 || natural.heightPx <= 0) {
    return undefined;
  }
  const widthCm = (natural.widthPx / 96) * CM_PER_INCH;
  const heightCm = (natural.heightPx / 96) * CM_PER_INCH;
  const crop = { l: left! / widthCm, t: top! / heightCm, r: right! / widthCm, b: bottom! / heightCm };
  return crop.l === 0 && crop.t === 0 && crop.r === 0 && crop.b === 0 ? undefined : crop;
}

type FaceContent = Exclude<SheetDrawingContent, { kind: "group" | "chart" | "foreign" }>;

/** A `draw:image`'s content. */
function imageContent(image: Element, styleName: string | null, ctx: OdsDrawingReadContext): FaceContent | Extract<SheetDrawingContent, { kind: "foreign" }> {
  const path = packagePath(image.getAttributeNS(XLINK_NAMESPACE, "href"));
  if (path === null) {
    return { kind: "foreign", label: image.hasAttributeNS(XLINK_NAMESPACE, "href") ? "Linked picture" : "Object" };
  }
  ctx.media.part(path);
  const content: Extract<SheetDrawingContent, { kind: "image" }> = { kind: "image", mediaId: path };
  const crop = cropOf(ctx.styles.graphicProperties(styleName ?? undefined).get("fo:clip"), () => ctx.media.naturalSize(path));
  if (crop) {
    content.crop = crop;
  }
  return content;
}

/** The `Object N/content.xml` a frame's `draw:object` names. */
function objectPart(object: Element): string | undefined {
  const path = packagePath(object.getAttributeNS(XLINK_NAMESPACE, "href"));
  return path === null ? undefined : `${path.replace(/\/+$/, "")}/content.xml`;
}

/** A shape's content: its geometry, the graphic style's fill and outline, its text. */
function shapeContent(el: Element, geometry: ShapeGeometry, ctx: OdsDrawingReadContext, extra: Partial<ShapeVisual> = {}, textContainer: Element = el, defaultAnchor: DrawingText["anchor"] = "middle"): SheetDrawingContent {
  const styleName = drawAttr(el, "style-name");
  const props = ctx.styles.graphicProperties(styleName ?? undefined);
  const line = LINE_GEOMETRIES.has(geometry);
  const visual: ShapeVisual = { geometry, ...extra, ...styleVisual(props, line) };
  if (geometry === "line" && props.get("draw:marker-end") !== undefined) {
    visual.geometry = "arrow";
  }
  const text = line ? undefined : drawingText(textContainer, styleName, ctx, defaultAnchor);
  return text ? { kind: "shape", visual, text } : { kind: "shape", visual };
}

/** A `draw:custom-shape`'s geometry (see the module header). */
function customShapeContent(el: Element, ctx: OdsDrawingReadContext): SheetDrawingContent {
  const enhanced = firstChild(el, DRAW_NAMESPACE, "enhanced-geometry");
  const type = enhanced ? drawAttr(enhanced, "type") : null;
  const geometry = type === null ? undefined : DRAW_TYPE_TO_GEOMETRY[type];
  if (!geometry) {
    return shapeContent(el, "rect", ctx, { unsupportedGeometryXml: enhanced ? new XMLSerializer().serializeToString(enhanced) : "" });
  }
  const modifiers = enhanced ? drawAttr(enhanced, "modifiers") : null;
  const adjustments = (modifiers ?? "")
    .trim()
    .split(/\s+/)
    .filter((each) => each !== "")
    .map(Number)
    .filter((each) => Number.isFinite(each))
    .map(guideValueToAdjustment);
  return shapeContent(el, geometry, ctx, adjustments.length > 0 ? { adjustments } : {});
}

/** What a drawing element paints (see the module header). */
function contentOf(el: Element, geometry: ElementGeometry, ctx: OdsDrawingReadContext): SheetDrawingContent {
  if (el.namespaceURI === DR3D_NAMESPACE && el.localName === "scene") {
    return { kind: "foreign", label: "3-D object" };
  }
  if (el.namespaceURI !== DRAW_NAMESPACE) {
    return { kind: "foreign", label: "Object" };
  }
  switch (el.localName) {
    case "frame":
      return frameContent(el, ctx);
    case "custom-shape":
      return customShapeContent(el, ctx);
    case "rect":
      return shapeContent(el, "rect", ctx);
    case "ellipse":
      return shapeContent(el, "ellipse", ctx);
    case "line":
      return shapeContent(el, "line", ctx);
    case "g":
      return groupContent(el, geometry, ctx);
    case "control":
      return { kind: "foreign", label: "Form control" };
    default:
      return { kind: "foreign", label: "Object" };
  }
}

function frameContent(el: Element, ctx: OdsDrawingReadContext): SheetDrawingContent {
  const children = elementChildren(el).filter((child) => child.namespaceURI === DRAW_NAMESPACE);
  const styleName = drawAttr(el, "style-name");
  const object = children.find((child) => child.localName === "object");
  if (object) {
    const sourcePart = objectPart(object);
    return sourcePart === undefined ? { kind: "chart" } : { kind: "chart", sourcePart };
  }
  const image = children.find((child) => child.localName === "image");
  if (children.some((child) => child.localName === "object-ole")) {
    const face = image ? imageContent(image, styleName, ctx) : undefined;
    return face && face.kind === "image" ? { kind: "foreign", label: "Object", face } : { kind: "foreign", label: "Object" };
  }
  if (image) {
    return imageContent(image, styleName, ctx);
  }
  const textBox = children.find((child) => child.localName === "text-box");
  if (textBox) {
    return shapeContent(el, "rect", ctx, {}, textBox, "top");
  }
  return { kind: "foreign", label: "Object" };
}

type MemberContent = Extract<SheetDrawingContent, { kind: "group" }>["members"][number]["content"];

function groupContent(el: Element, geometry: ElementGeometry, ctx: OdsDrawingReadContext): SheetDrawingContent {
  const box = geometry.local;
  const elements = elementChildren(el).filter((child) => isDrawingElementName(canonicalName(child.tagName, child.localName, child.namespaceURI ?? undefined)));
  const members = geometry.members.map((member, index) => {
    const child = elements[index];
    const read = child ? contentOf(child, member.geometry, ctx) : ({ kind: "foreign", label: "Object" } as const);
    const content: MemberContent = read.kind === "group" ? { kind: "foreign", label: "Group" } : read.kind === "chart" ? { kind: "foreign", label: "Chart" } : read;
    const m = member.geometry.box;
    return {
      x: box.w > 0 ? (m.x - box.x) / box.w : 0,
      y: box.h > 0 ? (m.y - box.y) / box.h : 0,
      w: box.w > 0 ? m.w / box.w : 0,
      h: box.h > 0 ? m.h / box.h : 0,
      content,
    };
  });
  return { kind: "group", members };
}

/** The svg:desc child's text, or undefined. */
function descOf(el: Element): string | undefined {
  const desc = firstChild(el, SVG_NAMESPACE, "desc");
  const text = desc?.textContent ?? "";
  return text === "" ? undefined : text;
}

/** One element as a drawing (its id minted by the caller) and its z-index. */
function readDrawing(source: OdsDrawingSource, metrics: GridMetrics, ctx: OdsDrawingReadContext): { drawing: Omit<SheetDrawing, "id">; z?: number } | undefined {
  const geometry = elementGeometry(source.xml);
  if (!geometry) {
    return undefined;
  }
  let el: Element | null = null;
  try {
    el = parseOdsPart(`<ods-drawing ${ctx.declarations}>${source.xml}</ods-drawing>`).documentElement.firstElementChild;
  } catch (error) {
    if (!(error instanceof XmlParseError)) {
      throw error;
    }
  }
  const drawing: Omit<SheetDrawing, "id"> = {
    anchor: anchorOf(geometry, source.cell, metrics),
    name: geometry.drawName ?? "",
    sourceXml: source.xml,
    content: el ? contentOf(el, geometry, ctx) : { kind: "foreign", label: "Object" },
  };
  const altText = el ? descOf(el) : undefined;
  if (altText !== undefined) {
    drawing.altText = altText;
  }
  if (geometry.rotationDeg !== undefined) {
    drawing.rotationDeg = geometry.rotationDeg;
  }
  return geometry.z === undefined ? { drawing } : { drawing, z: geometry.z };
}

/** A drawing's place in the z-order: its z-index, else in front of all. */
const zKey = (each: { z?: number }): number => each.z ?? Number.POSITIVE_INFINITY;

/** A sheet's drawings from the elements the reader captured, in z-order
 * with ids minted (see the module header). `metrics` is the sheet's
 * read-time geometry. */
export function readOdsDrawings(sources: readonly OdsDrawingSource[], metrics: GridMetrics, ctx: OdsDrawingReadContext): SheetDrawing[] {
  const read = sources.flatMap((source, order) => {
    const one = readDrawing(source, metrics, ctx);
    return one ? [{ ...one, order }] : [];
  });
  return read
    .toSorted((a, b) => (zKey(a) === zKey(b) ? a.order - b.order : zKey(a) < zKey(b) ? -1 : 1))
    .map((each, index) => ({ id: `dr${index + 1}`, ...each.drawing }));
}

// --- writing ---------------------------------------------------------------------------------

/** What writing a dirty table's drawings needs (tableXml.ts builds it). */
export interface OdsDrawingWriteContext {
  /** The sheet's current name (end cells are on it). */
  sheetName: string;
  /** The sheet's write-time geometry. */
  metrics: GridMetrics;
  /** Generated graphic styles (`de-gr{n}`, ods/styles.ts OdsStyleRegistry). */
  registry: { graphicStyle(properties: string): string };
  theme?: DocumentTheme;
  /** The retained-text sheet-reference pass, for an imported drawing's source. */
  passRetained?: (text: string) => string;
  /** The package path a media id is written under (absent: the id itself). */
  mediaPath?: (mediaId: string) => string;
}

/** A dirty table's drawings as XML: the `table:shapes` element ("" when no
 * drawing is fixed) and the in-cell drawings by row, then column, each the
 * drawings anchored there in z-order. */
export interface OdsSheetDrawingsXml {
  shapes: string;
  cells: Map<number, Map<number, string>>;
}

/** The z-index each drawing is written with (undefined: leave the source's
 * absent one alone) — see the module header. */
function zIndexes(drawings: readonly SheetDrawing[], sourceZ: readonly (number | undefined)[]): (number | undefined)[] {
  // unnumbered[i]: drawing i and every one in front of it are imported without a z-index.
  const unnumbered: boolean[] = [];
  for (let i = drawings.length - 1; i >= 0; i -= 1) {
    unnumbered[i] = drawings[i]!.sourceXml !== undefined && sourceZ[i] === undefined && (i === drawings.length - 1 || unnumbered[i + 1]!);
  }
  let last = -1;
  return drawings.map((_drawing, index) => {
    const source = sourceZ[index];
    if (unnumbered[index]) {
      return undefined;
    }
    if (source !== undefined && source > last) {
      last = source;
      return source;
    }
    last += 1;
    return last;
  });
}

/** Writes a sheet's drawings (see the module header). */
export function odsSheetDrawingsXml(drawings: readonly SheetDrawing[], ctx: OdsDrawingWriteContext): OdsSheetDrawingsXml {
  const sources = drawings.map((drawing) => (drawing.sourceXml === undefined ? undefined : (ctx.passRetained ?? ((text: string): string => text))(drawing.sourceXml)));
  const geometries = sources.map((text) => (text === undefined ? undefined : elementGeometry(text)));
  const z = zIndexes(
    drawings,
    geometries.map((geometry) => geometry?.z),
  );
  let shapes = "";
  const cells = new Map<number, Map<number, string>>();
  drawings.forEach((drawing, index) => {
    const source = sources[index];
    const geometry = geometries[index];
    const xml =
      source !== undefined && geometry !== undefined ? patchedDrawing(source, geometry, drawing, z[index], ctx) : source !== undefined ? source : canonicalOdsDrawing(drawing, z[index] ?? 0, ctx);
    if (xml === "") {
      return;
    }
    const from = drawing.anchor.from;
    if (anchorBehaviour(drawing.anchor) === "fixed" || !from) {
      shapes += xml;
      return;
    }
    const row = Math.min(Math.max(0, from.row), MAX_ROWS - 1);
    const col = Math.min(Math.max(0, from.col), MAX_COLS - 1);
    let byCol = cells.get(row);
    if (!byCol) {
      byCol = new Map();
      cells.set(row, byCol);
    }
    byCol.set(col, (byCol.get(col) ?? "") + xml);
  });
  return { shapes: shapes === "" ? "" : `<table:shapes>${shapes}</table:shapes>`, cells };
}

/** The box a model anchor puts a drawing at, sheet-absolute EMU: the start
 * unclamped (offsets kept as read), the size the rendered rectangle's —
 * a move or fixed anchor's extEmu. */
function targetBox(anchor: DrawingAnchor, metrics: GridMetrics): Box {
  const rect = drawingRectPx(anchor, metrics);
  const behaviour = anchorBehaviour(anchor);
  let x = rect.left * EMU_PER_PX;
  let y = rect.top * EMU_PER_PX;
  if (behaviour !== "fixed" && anchor.from) {
    x = metrics.colLeft(anchor.from.col) * EMU_PER_PX + anchor.from.colOffEmu;
    y = metrics.rowTop(anchor.from.row) * EMU_PER_PX + anchor.from.rowOffEmu;
  } else if (anchor.posEmu) {
    x = anchor.posEmu.x;
    y = anchor.posEmu.y;
  }
  if (behaviour !== "moveAndSize" && anchor.extEmu) {
    return { x, y, w: anchor.extEmu.cx, h: anchor.extEmu.cy };
  }
  return { x, y, w: rect.width * EMU_PER_PX, h: rect.height * EMU_PER_PX };
}

const differs = (a: number, b: number): boolean => Math.abs(a - b) > PATCH_TOLERANCE_EMU;

/** `value` (cm) as a canonical length: centimetres to 4 decimals. */
export function odsCm(cm: number): string {
  const rounded = Number(cm.toFixed(4));
  return `${Object.is(rounded, -0) ? 0 : rounded}cm`;
}

const DIGITS_BY_UNIT: Readonly<Record<string, number>> = { cm: 4, mm: 3, in: 4, pt: 2, pc: 3 };
const CM_PER_UNIT: Readonly<Record<string, number>> = { cm: 1, mm: 0.1, in: CM_PER_INCH, pt: CM_PER_INCH / POINTS_PER_INCH, pc: (CM_PER_INCH / POINTS_PER_INCH) * 12 };

/** An EMU length written in the unit `like` (the source's value) uses — cm when it has none. */
function lengthLike(emu: number, like: string | undefined): string {
  const unit = like === undefined ? undefined : /([a-z]+)\s*$/i.exec(like.trim())?.[1]?.toLowerCase();
  const cm = emuToCm(emu);
  if (unit === undefined || CM_PER_UNIT[unit] === undefined || unit === "cm") {
    return odsCm(cm);
  }
  const value = Number((cm / CM_PER_UNIT[unit]!).toFixed(DIGITS_BY_UNIT[unit]!));
  return `${Object.is(value, -0) ? 0 : value}${unit}`;
}

/** A sheet name as an end cell writes it: bare, or quoted (`''` for a quote)
 * when a formula would quote it or it holds a `.` — LibreOffice's form, no `$`. */
function odsAddressSheet(name: string): string {
  return sheetNameNeedsQuotes(name) || name.includes(".") ? `'${name.replaceAll("'", "''")}'` : name;
}

const cellText = (col: number, row: number): string => `${colToLetters(col)}${row + 1}`;

/** The end-cell attribute value: the source's sheet spelling kept (and its
 * cell part's `$`s) while it names the sheet, else the model's name. */
function endAddress(point: AnchorPoint, source: ElementGeometry["end"], sheetName: string): string {
  if (source && (source.sheet === undefined || source.sheet.toUpperCase() === sheetName.toUpperCase())) {
    const raw = source.raw.trim();
    const dot = lastTopLevelDot(raw);
    if (dot >= 0) {
      const cell = raw.slice(dot + 1);
      const colAbs = cell.startsWith("$");
      const rowAbs = /[A-Za-z]\$/.test(cell);
      return `${raw.slice(0, dot + 1)}${colAbs ? "$" : ""}${colToLetters(point.col)}${rowAbs ? "$" : ""}${point.row + 1}`;
    }
  }
  return `${odsAddressSheet(sheetName)}.${cellText(point.col, point.row)}`;
}

/** The index of the `.` that separates an address's sheet from its cell (outside quotes), or −1. */
function lastTopLevelDot(text: string): number {
  let quoted = false;
  let found = -1;
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === "'") {
      quoted = !quoted;
    } else if (text[i] === "." && !quoted) {
      found = i;
    } else if (text[i] === ":" && !quoted) {
      break;
    }
  }
  return found;
}

/** An imported drawing's source text with its anchor patched (see the module header). */
function patchedDrawing(text: string, geometry: ElementGeometry, drawing: SheetDrawing, z: number | undefined, ctx: OdsDrawingWriteContext): string {
  const metrics = ctx.metrics;
  const anchor = drawing.anchor;
  const behaviour = anchorBehaviour(anchor);
  const start = text.indexOf("<");
  let tag = text.slice(start, geometry.tagEnd);
  const attr = (qname: string): string | undefined => attrOf(geometry.attrs, qname);
  const set = (qname: string, value: string | null): void => {
    tag = setAttribute(tag, qname, value);
  };
  const target = targetBox(anchor, metrics);
  // The start position, read back in the model's anchor cell.
  let movedX: boolean;
  let movedY: boolean;
  let readFrom: AnchorPoint | undefined;
  if (behaviour === "fixed" || !anchor.from) {
    movedX = differs(geometry.box.x, target.x);
    movedY = differs(geometry.box.y, target.y);
  } else {
    const left = metrics.colLeft(anchor.from.col) * EMU_PER_PX;
    const top = metrics.rowTop(anchor.from.row) * EMU_PER_PX;
    readFrom = { col: anchor.from.col, colOffEmu: offsetInto(geometry.box.x, metrics.colLeft(anchor.from.col)), row: anchor.from.row, rowOffEmu: offsetInto(geometry.box.y, metrics.rowTop(anchor.from.row)) };
    // A source position before the anchor cell's start reads back clamped: its cell moved.
    movedX = differs(readFrom.colOffEmu, anchor.from.colOffEmu) || geometry.box.x < left - PATCH_TOLERANCE_EMU;
    movedY = differs(readFrom.rowOffEmu, anchor.from.rowOffEmu) || geometry.box.y < top - PATCH_TOLERANCE_EMU;
  }
  // The end cell.
  let wantEnd: AnchorPoint | undefined;
  if (behaviour === "moveAndSize") {
    wantEnd = anchor.to ?? anchor.from;
  } else if (geometry.end) {
    const sourceTo: AnchorPoint = { col: geometry.end.col, colOffEmu: geometry.end.xEmu, row: geometry.end.row, rowOffEmu: geometry.end.yEmu };
    wantEnd = writtenPoints(anchor, metrics, { ...(readFrom ? { from: readFrom } : {}), to: sourceTo }).to;
  }
  let endMoved = false;
  if (wantEnd) {
    const end = geometry.end;
    const sheetMatches = end !== undefined && (end.sheet === undefined || end.sheet.toUpperCase() === ctx.sheetName.toUpperCase());
    if (!end || end.col !== wantEnd.col || end.row !== wantEnd.row || !sheetMatches) {
      set("table:end-cell-address", endAddress(wantEnd, end, ctx.sheetName));
      endMoved = true;
    }
    if (!end || attr("table:end-x") === undefined || differs(end.xEmu, wantEnd.colOffEmu)) {
      set("table:end-x", lengthLike(wantEnd.colOffEmu, attr("table:end-x")));
      endMoved = true;
    }
    if (!end || attr("table:end-y") === undefined || differs(end.yEmu, wantEnd.rowOffEmu)) {
      set("table:end-y", lengthLike(wantEnd.rowOffEmu, attr("table:end-y")));
      endMoved = true;
    }
  }
  // The size: an extEmu anchor's own; a twoCell box's only once the anchor moved.
  const sizeCounts = behaviour !== "moveAndSize" || movedX || movedY || endMoved;
  const sizedW = sizeCounts && differs(geometry.box.w, target.w);
  const sizedH = sizeCounts && differs(geometry.box.h, target.h);
  const w = sizedW ? target.w : geometry.box.w;
  const h = sizedH ? target.h : geometry.box.h;
  switch (geometry.position) {
    case "xy":
      if (movedX) {
        set("svg:x", lengthLike(target.x, attr("svg:x")));
      }
      if (movedY) {
        set("svg:y", lengthLike(target.y, attr("svg:y")));
      }
      if (sizedW) {
        set("svg:width", lengthLike(w, attr("svg:width")));
      }
      if (sizedH) {
        set("svg:height", lengthLike(h, attr("svg:height")));
      }
      break;
    case "transform":
      if (sizedW) {
        set("svg:width", lengthLike(w, attr("svg:width")));
      }
      if (sizedH) {
        set("svg:height", lengthLike(h, attr("svg:height")));
      }
      if (movedX || movedY || sizedW || sizedH) {
        const local = { ...geometry.local, w, h };
        const at = transformedBox(local, geometry.transform!);
        const want = { x: movedX ? target.x : geometry.box.x, y: movedY ? target.y : geometry.box.y };
        set("draw:transform", translatedTransform(attr("draw:transform")!, geometry.transform!, want.x - at.x, want.y - at.y));
      }
      break;
    case "line":
      if (movedX || movedY || sizedW || sizedH) {
        const box = { x: movedX ? target.x : geometry.box.x, y: movedY ? target.y : geometry.box.y, w, h };
        for (const [qx, qy] of [
          ["svg:x1", "svg:y1"],
          ["svg:x2", "svg:y2"],
        ] as const) {
          const x = lengthEmu(attr(qx)) ?? 0;
          const y = lengthEmu(attr(qy)) ?? 0;
          const nx = box.x + (geometry.box.w > 0 ? ((x - geometry.box.x) / geometry.box.w) * box.w : x - geometry.box.x);
          const ny = box.y + (geometry.box.h > 0 ? ((y - geometry.box.y) / geometry.box.h) * box.h : y - geometry.box.y);
          if (differs(nx, x)) {
            set(qx, lengthLike(nx, attr(qx)));
          }
          if (differs(ny, y)) {
            set(qy, lengthLike(ny, attr(qy)));
          }
        }
      }
      break;
    case "group":
      if (movedX || movedY) {
        const dx = movedX ? target.x - geometry.box.x : 0;
        const dy = movedY ? target.y - geometry.box.y : 0;
        set("draw:transform", translatedTransform(attr("draw:transform") ?? "", geometry.transform ?? [], dx, dy));
      }
      break;
  }
  if (z !== undefined && z !== geometry.z) {
    set("draw:z-index", String(z));
  }
  if (drawing.name !== (geometry.drawName ?? "")) {
    set("draw:name", drawing.name === "" ? null : drawing.name);
  }
  const altText = drawing.altText ?? "";
  if (geometry.closeStart < 0 && altText !== "") {
    // A self-closing element gains a body for its alt text.
    return `${tag.replace(/\s*\/>$/, ">")}<svg:desc>${escapeElementText(altText)}</svg:desc></${geometry.qname}>`;
  }
  return tag + bodyWithDesc(text, geometry, altText);
}

/** A draw:transform with its last translate moved by (dx, dy) EMU — one
 * appended when its last step is not a translate. */
function translatedTransform(value: string, steps: readonly TransformStep[], dx: number, dy: number): string {
  const last = steps.at(-1);
  if (!last || last.kind !== "translate") {
    return `${value.trim() === "" ? "" : `${value.trimEnd()} `}translate (${odsCm(emuToCm(dx))} ${odsCm(emuToCm(dy))})`;
  }
  const x = lengthLike(last.values[0]! + dx, last.args[0]);
  const y = lengthLike(last.values[1]! + dy, last.args[1] ?? last.args[0]);
  const head = value.slice(last.start, last.end).replace(/\(.*$/, "");
  return `${value.slice(0, last.start)}${head}(${x} ${y})${value.slice(last.end)}`;
}

/** The element's text after its start tag with its svg:desc set to
 * `altText` ("" removes it): the child replaced, removed, or inserted after
 * an svg:title — else last in a frame (after its content, ODF 1.2 §10.4.2)
 * and first in a shape (§9.2.21). */
function bodyWithDesc(text: string, geometry: ElementGeometry, altText: string): string {
  const rest = text.slice(geometry.tagEnd);
  if (altText === (geometry.desc?.text ?? "") || (geometry.closeStart < 0 && altText === "")) {
    return rest;
  }
  const at = (offset: number): number => offset - geometry.tagEnd;
  const element = altText === "" ? "" : `<svg:desc>${escapeElementText(altText)}</svg:desc>`;
  const splices: Splice[] = [];
  if (geometry.desc) {
    splices.push({ start: at(geometry.desc.start), end: at(geometry.desc.end), text: element });
  } else {
    const insertAt = geometry.titleEnd ?? (geometry.name === "draw:frame" ? geometry.closeStart : geometry.tagEnd);
    splices.push({ start: at(insertAt), end: at(insertAt), text: element });
  }
  return applySplices(rest, splices);
}

// --- canonical new drawings (ods/drawingGolden.ts) -------------------------------------------

/** A colour as a graphic style writes it (#rrggbb). */
const colorHex = (color: Parameters<typeof resolveColor>[0], theme: DocumentTheme): string => resolveColor(color, theme).toLowerCase();

/** The pt a stroke width is written in. */
const ptText = (pt: number): string => `${Number(pt.toFixed(4))}pt`;

/** A shape's graphic-properties attribute run (ods/drawingGolden.ts). */
export function odsShapeGraphicProperties(visual: ShapeVisual, theme: DocumentTheme = DEFAULT_THEME): string {
  const line = LINE_GEOMETRIES.has(visual.geometry);
  const base = line ? undefined : fillColor(visual.fill);
  let out = base === undefined ? ` draw:fill="none"` : ` draw:fill="solid" draw:fill-color="${colorHex(base, theme)}"`;
  const outline: ShapeOutline | undefined = visual.outline;
  out += outline
    ? ` draw:stroke="${outline.dash === "solid" ? "solid" : "dash"}" svg:stroke-color="${colorHex(outline.color, theme)}" svg:stroke-width="${ptText(outline.widthPt)}"`
    : ` draw:stroke="none"`;
  if (visual.geometry === "arrow") {
    out += ` draw:marker-end="Arrow" draw:marker-end-width="0.3cm"`;
  }
  return out;
}

function fillColor(fill: ShapeFill | undefined): Parameters<typeof resolveColor>[0] | undefined {
  if (!fill || fill.kind === "none") {
    return undefined;
  }
  return shapeFillBaseColor(fill);
}

/** A picture's graphic-properties attribute run: no line, no fill. */
export const ODS_IMAGE_GRAPHIC_PROPERTIES = ` draw:stroke="none" draw:fill="none"`;

/** The anchor attributes of a canonical drawing: the end cell (moveAndSize
 * only), then z-index, name and style. */
function canonicalHead(drawing: SheetDrawing, z: number, style: string, ctx: OdsDrawingWriteContext): string {
  let head = "";
  if (anchorBehaviour(drawing.anchor) === "moveAndSize") {
    const to = drawing.anchor.to ?? drawing.anchor.from ?? { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 };
    head +=
      ` table:end-cell-address="${escapeAttributeValue(`${odsAddressSheet(ctx.sheetName)}.${cellText(to.col, to.row)}`)}"` +
      ` table:end-x="${odsCm(emuToCm(to.colOffEmu))}" table:end-y="${odsCm(emuToCm(to.rowOffEmu))}"`;
  }
  head += ` draw:z-index="${z}"`;
  head += ` draw:name="${escapeAttributeValue(drawing.name)}"`;
  return `${head} draw:style-name="${style}"`;
}

const boxAttributes = (box: Box): string =>
  ` svg:width="${odsCm(emuToCm(box.w))}" svg:height="${odsCm(emuToCm(box.h))}" svg:x="${odsCm(emuToCm(box.x))}" svg:y="${odsCm(emuToCm(box.y))}"`;

const descXml = (altText: string | undefined): string => (altText === undefined || altText === "" ? "" : `<svg:desc>${escapeElementText(altText)}</svg:desc>`);

/** A shape's text as `text:p` paragraphs ("" for none). */
function textXml(text: DrawingText | undefined): string {
  return (text?.paragraphs ?? []).map((p) => (p.text === "" ? "<text:p/>" : `<text:p>${escapeElementText(p.text).replace(/\t/g, "<text:tab/>").replace(/\n/g, "<text:line-break/>")}</text:p>`)).join("");
}

/** A drawing created here in its canonical form (ods/drawingGolden.ts), or
 * "" for a kind with none (a created group, chart or foreign drawing). */
export function canonicalOdsDrawing(drawing: SheetDrawing, z: number, ctx: OdsDrawingWriteContext): string {
  const theme = ctx.theme ?? DEFAULT_THEME;
  const box = targetBox(drawing.anchor, ctx.metrics);
  const content = drawing.content;
  if (content.kind === "image") {
    const style = ctx.registry.graphicStyle(ODS_IMAGE_GRAPHIC_PROPERTIES);
    const href = ctx.mediaPath ? ctx.mediaPath(content.mediaId) : content.mediaId;
    return (
      `<draw:frame${canonicalHead(drawing, z, style, ctx)}${boxAttributes(box)}>` +
      `<draw:image xlink:href="${escapeAttributeValue(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${descXml(drawing.altText)}</draw:frame>`
    );
  }
  if (content.kind !== "shape") {
    return "";
  }
  const visual = content.visual;
  const style = ctx.registry.graphicStyle(odsShapeGraphicProperties(visual, theme));
  if (LINE_GEOMETRIES.has(visual.geometry)) {
    const line = ` svg:x1="${odsCm(emuToCm(box.x))}" svg:y1="${odsCm(emuToCm(box.y))}" svg:x2="${odsCm(emuToCm(box.x + box.w))}" svg:y2="${odsCm(emuToCm(box.y + box.h))}"`;
    const desc = descXml(drawing.altText);
    return desc === "" ? `<draw:line${canonicalHead(drawing, z, style, ctx)}${line}/>` : `<draw:line${canonicalHead(drawing, z, style, ctx)}${line}>${desc}</draw:line>`;
  }
  const modifiers = visual.adjustments && visual.adjustments.length > 0 ? ` draw:modifiers="${visual.adjustments.map(adjustmentToGuideValue).join(" ")}"` : "";
  const type = GEOMETRY_TO_DRAW_TYPE[visual.geometry] ?? "rectangle";
  return (
    `<draw:custom-shape${canonicalHead(drawing, z, style, ctx)}${boxAttributes(box)}>${descXml(drawing.altText)}${textXml(content.text)}` +
    `<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="${escapeAttributeValue(type)}"${modifiers}/></draw:custom-shape>`
  );
}

// --- media paths -----------------------------------------------------------------------------

/** The package path each media id is written under: a `Pictures/…` id is
 * its own path; anything else (a cross-format id) gets the next free
 * `Pictures/image{n}.{ext}` ("Media parts"). */
export class OdsMediaPaths {
  private readonly paths = new Map<string, string>();
  private next: number;

  constructor(taken: Iterable<string>) {
    let max = 0;
    for (const name of taken) {
      const match = /^Pictures\/image(\d+)\.[^/]*$/.exec(name);
      if (match) {
        max = Math.max(max, Number(match[1]));
      }
    }
    this.next = max + 1;
  }

  pathOf(mediaId: string, contentType: string): string {
    let path = this.paths.get(mediaId);
    if (path === undefined) {
      if (/^Pictures\/[^/]+$/.test(mediaId)) {
        path = mediaId;
      } else {
        path = `Pictures/image${this.next}.${extensionForContentType(contentType)}`;
        this.next += 1;
      }
      this.paths.set(mediaId, path);
    }
    return path;
  }
}

/** Every media id a drawing paints (its picture, a face's, a group member's). */
export function drawingMediaIds(drawing: SheetDrawing): string[] {
  const content = drawing.content;
  if (content.kind === "image") {
    return [content.mediaId];
  }
  if (content.kind === "foreign") {
    return content.face?.kind === "image" ? [content.face.mediaId] : [];
  }
  if (content.kind === "group") {
    return content.members.flatMap((member) => (member.content.kind === "image" ? [member.content.mediaId] : member.content.kind === "foreign" && member.content.face?.kind === "image" ? [member.content.face.mediaId] : []));
  }
  return [];
}

// --- the table:shapes splice ------------------------------------------------------------------

/** A table's leading children (between its start tag and its first column
 * or row) with `table:shapes` replaced by `shapes` ("" removes it): in the
 * old one's place, else after `office:forms` (or the last of the children
 * ODF puts before it), else first. `prefix` is the table's own prefix
 * ("table:"). Every other byte verbatim. */
export function leadingWithShapes(leading: string, shapes: string, prefix = "table:"): string {
  const spans: { name: string; start: number; end: number }[] = [];
  let depth = 0;
  let open = -1;
  let openName = "";
  try {
    scanXml(`<w>${leading}</w>`, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "open") {
        if (depth === 1) {
          if (event.selfClosing) {
            spans.push({ name: event.name, start: event.start - 3, end: event.end - 3 });
          } else {
            open = event.start - 3;
            openName = event.name;
          }
        }
        if (!event.selfClosing) {
          depth += 1;
        }
        return undefined;
      }
      depth -= 1;
      if (depth === 1 && open >= 0) {
        spans.push({ name: openName, start: open, end: event.end - 3 });
        open = -1;
      }
      return undefined;
    });
  } catch {
    return leading;
  }
  const old = spans.find((span) => span.name === `${prefix}shapes`);
  if (old) {
    // A table:shapes holding no drawing, with none to write, stays as it is.
    if (shapes === "" && !/<(?:draw|dr3d):/.test(leading.slice(old.start, old.end))) {
      return leading;
    }
    return leading.slice(0, old.start) + shapes + leading.slice(old.end);
  }
  if (shapes === "") {
    return leading;
  }
  const before = new Set([`${prefix}title`, `${prefix}desc`, `${prefix}table-source`, "office:dde-source", `${prefix}scenario`, "office:forms"]);
  const after = spans.filter((span) => before.has(span.name)).at(-1);
  const at = after ? after.end : 0;
  return leading.slice(0, at) + shapes + leading.slice(at);
}

// --- the reader upgrade ----------------------------------------------------------------------

/** What upgradeOdsSheetDrawings needs: the dirty stored sheet, the retained
 * content envelope's prefix (its root's namespace declarations and the
 * automatic graphic styles) and the source package's parts (`Pictures/…`,
 * `META-INF/manifest.xml`, styles.xml); `media` the model's media so far
 * (a part already there is reused by path). */
export interface OdsDrawingUpgradeInput {
  sheet: Sheet;
  prefix: string;
  parts: ReadonlyMap<string, Uint8Array>;
  media?: readonly MediaPart[];
}

/** The upgrade's result: the sheet's rows without the lifted elements
 * (copy-on-write — the same array when no cell held one; a cell left
 * holding nothing dropped, a row left empty too), its drawings in z-order,
 * and the model's media with the pictures they show added. */
export interface OdsDrawingUpgrade {
  rows: Row[];
  drawings: SheetDrawing[];
  media: MediaPart[];
}

/** The top-level elements of a cell's preservedXml, split into drawing
 * elements and the rest (verbatim, in order). */
function splitPreserved(xml: string): { drawings: string[]; rest: string } {
  const drawings: string[] = [];
  let rest = "";
  let depth = 0;
  let open = -1;
  let openName = "";
  let cursor = 0;
  const take = (name: string, start: number, end: number): void => {
    if (isDrawingElementName(name)) {
      rest += xml.slice(cursor, start);
      drawings.push(xml.slice(start, end));
      cursor = end;
    }
  };
  const wrapped = `<w>${xml}</w>`;
  try {
    scanXml(wrapped, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "open") {
        if (depth === 1) {
          if (event.selfClosing) {
            take(event.name, event.start - 3, event.end - 3);
          } else {
            open = event.start - 3;
            openName = event.name;
          }
        }
        if (!event.selfClosing) {
          depth += 1;
        }
        return undefined;
      }
      depth -= 1;
      if (depth === 1 && open >= 0) {
        take(openName, open, event.end - 3);
        open = -1;
      }
      return undefined;
    });
  } catch {
    return { drawings: [], rest: xml };
  }
  rest += xml.slice(cursor);
  return { drawings, rest };
}

/** Whether a cell holds anything once its preservedXml is gone. */
function cellHoldsSomething(cell: Cell): boolean {
  const { col: _col, value, ...rest } = cell;
  return value !== null || Object.keys(rest).length > 0;
}

/** A row holds something beyond its cells. */
function rowHoldsSomething(row: Row): boolean {
  return row.cells.length > 0 || row.heightPt !== undefined || row.hidden === true || row.filterHidden === true || row.styleId !== undefined || row.sourceAttrs !== undefined || row.sourceStyleName !== undefined || row.customHeight === true;
}

/** The `table:shapes` children of a table's retained source text (its direct child only). */
function retainedShapes(sourceXml: string | undefined): string[] {
  if (sourceXml === undefined) {
    return [];
  }
  const out: string[] = [];
  let depth = 0;
  let inShapes = false;
  let childStart = -1;
  try {
    scanXml(sourceXml, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "open") {
        const name = canonicalName(event.name, event.local, event.ns);
        if (depth === 1 && (name === "table:shapes" || event.local === "shapes") && !event.selfClosing) {
          inShapes = true;
        } else if (inShapes && depth === 2) {
          if (event.selfClosing) {
            out.push(sourceXml.slice(event.start, event.end));
          } else {
            childStart = event.start;
          }
        } else if (depth === 1 && (event.local === "table-column" || event.local === "table-row" || event.local === "table-rows" || event.local === "table-columns" || event.local === "table-header-rows" || event.local === "table-header-columns" || event.local === "table-row-group" || event.local === "table-column-group")) {
          return "stop";
        }
        if (!event.selfClosing) {
          depth += 1;
        }
        return undefined;
      }
      depth -= 1;
      if (inShapes && depth === 2 && childStart >= 0) {
        out.push(sourceXml.slice(childStart, event.end));
        childStart = -1;
      } else if (inShapes && depth === 1) {
        inShapes = false;
      }
      return undefined;
    });
  } catch {
    return out;
  }
  return out;
}

/** SPEC-131's reader-upgrade step for a dirty .ods sheet (see the module
 * header): the drawing elements its cells' preservedXml hold, anchored to
 * each cell's CURRENT position, plus its retained `table:shapes`, read as
 * the current reader reads them. */
export function upgradeOdsSheetDrawings(input: OdsDrawingUpgradeInput): OdsDrawingUpgrade {
  const { sheet } = input;
  const sources: OdsDrawingSource[] = retainedShapes(input.sheet.sourceXml).map((xml) => ({ xml }));
  let rows: Row[] | null = null;
  sheet.rows.forEach((row, at) => {
    if (!row.cells.some((cell) => cell.preservedXml !== undefined && /<(?:draw|dr3d):/.test(cell.preservedXml))) {
      return;
    }
    const cells: Cell[] = [];
    for (const cell of row.cells) {
      if (cell.preservedXml === undefined) {
        cells.push(cell);
        continue;
      }
      const split = splitPreserved(cell.preservedXml);
      if (split.drawings.length === 0) {
        cells.push(cell);
        continue;
      }
      for (const xml of split.drawings) {
        sources.push({ xml, cell: { row: row.index, col: cell.col } });
      }
      const next: Cell = { ...cell, preservedXml: split.rest };
      if (split.rest === "") {
        delete next.preservedXml;
      }
      if (cellHoldsSomething(next)) {
        cells.push(next);
      }
    }
    rows ??= sheet.rows.slice();
    rows[at] = { ...row, cells };
  });
  const prefix = input.prefix;
  const root = /<[\w.-]+:document-content\b[^>]*>/.exec(prefix)?.[0] ?? "";
  const declarations = namespaceDeclarations(
    [...root.matchAll(/\s(xmlns(?::[\w.-]+)?)\s*=\s*"([^"]*)"/g)].map((match) => ({ name: match[1]!, value: decodeXmlText(match[2]!) })),
  );
  const stylesBytes = input.parts.get("styles.xml");
  const styles = odsStyleBookFromSource(prefix, stylesBytes ? new TextDecoder().decode(stylesBytes) : undefined);
  const media = new OdsMediaReader(input.parts, input.media ?? []);
  const nextRows: Row[] = rows === null ? sheet.rows : (rows as Row[]).filter(rowHoldsSomething);
  const drawings = sources.length === 0 ? [] : readOdsDrawings(sources, buildGridMetrics({ ...sheet, rows: nextRows }), { declarations, styles, media });
  return { rows: nextRows, drawings, media: media.media() };
}
