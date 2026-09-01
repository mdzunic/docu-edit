// Reads/writes the ODF shape markup SPEC-051 models: `draw:custom-shape`
// (area geometries) and `draw:line` (Line/Arrow), the shape-flavoured
// siblings of drawing.ts's `draw:frame`/`draw:image` for pictures. Mirrors
// drawing.ts's own split (parse.../...Xml) so the two modules read side by
// side; this one never touches Image, drawing.ts never touches Shape.
import type { Block, ImageWrapMode, Paragraph, Shape } from "../documentModel";
import type {
  ArrowheadKind,
  ArrowheadSize,
  ShapeFill,
  ShapeGeometry,
  ShapeGradient,
  ShapeOutline,
  ShapeShadow,
} from "../core/shapes";
import { LINE_GEOMETRIES } from "../core/shapes";
import { adjustmentToGuideValue, guideValueToAdjustment } from "../core/shapeCatalogue";
import {
  GEOMETRY_TO_DRAW_TYPE,
  DRAW_TYPE_TO_GEOMETRY as DRAW_TYPE_TO_GEOMETRY_TABLE,
  PATTERN_PRESET_TO_ODF_HATCH,
  odfHatchToPatternPreset,
} from "../core/drawingMl";
import { enhancedPathString, enhancedPathViewBox, parseEnhancedPath } from "../core/enhancedPath";
import { escapeXmlAttribute, firstChildElement, parseXml, serializeElement } from "../core/xml";
import { cmToOdfLength, odfLengthToCm } from "./units";
import {
  DRAW_NAMESPACE,
  STYLE_NAMESPACE,
  SVG_NAMESPACE,
  TEXT_NAMESPACE,
  XLINK_NAMESPACE,
  drawAttr,
  styleAttr,
  svgAttr,
  textAttr,
} from "../core/odfXml";

// Our own extension namespace, used to round-trip groupOrdinal and rotation
// (ODF's real `draw:transform` is a general affine-matrix string; encoding
// our whole-degree rotation through it is unnecessary complexity when we
// only ever need to round-trip through our own reader/writer — see the
// implement summary). Unknown-namespaced attributes are ignored by
// LibreOffice/Word, so this never corrupts a file another app opens.
const DE_NAMESPACE = "urn:docu-edit:shape";

// Declared locally on every shape-related element rather than on the
// document root, so content.xml's bytes are unchanged for documents without
// shapes (the TABLE_XMLNS/FRAME_XMLNS pattern from writer.ts/drawing.ts).
// Needed on both the style:style (draw:fill/stroke, svg:stroke-*) and the
// draw:custom-shape/draw:line element itself (svg:width/height/x/y, etc).
const SHAPE_XMLNS = ` xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}"`;

// The ODF draw:type maps now live in core/drawingMl.ts (SPEC-079's two-phase
// move) so the `.odp` reader shares them; this module re-imports them (the
// reverse map aliased to its historical name), its behaviour unchanged.

// --- SPEC-073: shared drawing definitions (styles.xml <office:styles>) ------
// ODF expresses gradients, hatches, bitmap fills, arrow markers and dash
// patterns as NAMED definition elements a graphic style references by name.
// Every definition this writer mints is prefixed "de" so a later save can
// tell its own (regenerated) definitions from a foreign document's.

/** The definition elements the defs pre-pass manages inside styles.xml's
 * `<office:styles>` (all in the draw: namespace). */
export const ODF_DRAWING_DEF_LOCAL_NAMES: ReadonlySet<string> = new Set([
  "gradient",
  "hatch",
  "fill-image",
  "marker",
  "stroke-dash",
]);

/** Whether an element is one of the shared drawing definitions the writer
 * routes into styles.xml (see mergePreservedGraphicProps). */
export function isOdfDrawingDefElement(el: Element): boolean {
  return el.namespaceURI === DRAW_NAMESPACE && ODF_DRAWING_DEF_LOCAL_NAMES.has(el.localName);
}

// Definition elements declare their namespaces locally (the TABLE_XMLNS
// pattern): they are string-appended into a styles.xml whose root may not
// declare draw:/svg:/xlink:, and the preserved-styles patch parses each one
// stand-alone before importing it.
const DEF_DRAW_XMLNS = ` xmlns:draw="${DRAW_NAMESPACE}"`;
const DEF_SVG_XMLNS = ` xmlns:svg="${SVG_NAMESPACE}"`;
const DEF_XLINK_XMLNS = ` xmlns:xlink="${XLINK_NAMESPACE}"`;

/** The closed marker catalogue (SPEC-073 arrowheads on plain lines): one
 * fixed "deArrow*" name per ArrowheadKind, viewBox 0 0 10 10, tip at the
 * top — ODF rotates the marker along the line's direction itself. */
const MARKER_DEFS: Record<Exclude<ArrowheadKind, "none">, { name: string; d: string }> = {
  triangle: { name: "deArrowTriangle", d: "M0 10 L5 0 L10 10 Z" },
  open: { name: "deArrowOpen", d: "M0 10 L5 0 L10 10 L5 4 Z" },
  stealth: { name: "deArrowStealth", d: "M0 10 L5 0 L10 10 L5 7 Z" },
  diamond: { name: "deArrowDiamond", d: "M5 0 L10 5 L5 10 L0 5 Z" },
  oval: {
    name: "deArrowOval",
    d: "M5 0 C7.76 0 10 2.24 10 5 C10 7.76 7.76 10 5 10 C2.24 10 0 7.76 0 5 C0 2.24 2.24 0 5 0 Z",
  },
};
const MARKER_NAME_TO_ARROW: Record<string, ArrowheadKind> = Object.fromEntries(
  Object.entries(MARKER_DEFS).map(([kind, def]) => [def.name, kind as ArrowheadKind]),
);

/** Marker widths per ArrowheadSize — 0.3cm is SPEC-070's legacy arrow width,
 * kept as "medium" so old files keep their size. */
const ARROW_WIDTH_CM: Record<ArrowheadSize, string> = {
  small: "0.2cm",
  medium: "0.3cm",
  large: "0.45cm",
};

function arrowSizeFromWidth(width: string | null): ArrowheadSize {
  const cm = odfLengthToCm(width);
  if (cm !== undefined) {
    if (Math.abs(cm - 0.2) < 0.01) {
      return "small";
    }
    if (Math.abs(cm - 0.45) < 0.01) {
      return "large";
    }
  }
  return "medium";
}

/** The fixed dash-pattern names, one per non-solid model dash. The dot/dash
 * lengths are plausible renderings; only the NAME is contractual (the reader
 * maps it straight back to the model dash). */
const DASH_DEFS: Record<Exclude<ShapeOutline["dash"], "solid">, { name: string; body: string }> = {
  dash: { name: "deDash", body: ' draw:dots1="1" draw:dots1-length="0.2cm" draw:distance="0.15cm"' },
  dot: { name: "deDot", body: ' draw:dots1="1" draw:dots1-length="0.05cm" draw:distance="0.15cm"' },
  dashDot: {
    name: "deDashDot",
    body: ' draw:dots1="1" draw:dots1-length="0.2cm" draw:dots2="1" draw:dots2-length="0.05cm" draw:distance="0.15cm"',
  },
  longDash: { name: "deLongDash", body: ' draw:dots1="1" draw:dots1-length="0.4cm" draw:distance="0.15cm"' },
  longDashDot: {
    name: "deLongDashDot",
    body: ' draw:dots1="1" draw:dots1-length="0.4cm" draw:dots2="1" draw:dots2-length="0.05cm" draw:distance="0.15cm"',
  },
};
const DASH_NAME_TO_MODEL: Record<string, ShapeOutline["dash"]> = Object.fromEntries(
  Object.entries(DASH_DEFS).map(([dash, def]) => [def.name, dash as ShapeOutline["dash"]]),
);

/** Model gradient angle (degrees clockwise from "left to right", SPEC-073) →
 * ODF `draw:angle` (tenths of a degree, counter-clockwise from "bottom to
 * top"): angle10 = ((90 - angleDeg) mod 360) * 10. Bijective on whole
 * degrees; `odfGradientAngleToDeg` is the exact inverse. */
export function gradientAngleToOdf(angleDeg: number): number {
  return ((((90 - Math.round(angleDeg)) % 360) + 360) % 360) * 10;
}

function odfGradientAngleToDeg(angle10: number): number {
  return (((90 - angle10 / 10) % 360) + 360) % 360;
}

/** The shared-definition registry one save builds up: graphic styles claim
 * names through it while content.xml is generated, and the collected
 * definition elements land in styles.xml `<office:styles>` afterwards
 * (buildStylesXml on a fresh save, patchOdtShapeDefs on a preserved one).
 * Identical definitions dedupe to one name; preserved foreign definitions
 * (carried through `unsupportedStyleXml`) keep their own names verbatim. */
export interface OdfDrawingDefs {
  gradientNameFor(gradient: ShapeGradient): string;
  hatchNameFor(hatch: { style: "single" | "double"; angleDeg: number }, foreground: string): string;
  fillImageNameFor(href: string): string;
  markerNameFor(kind: Exclude<ArrowheadKind, "none">): string;
  strokeDashNameFor(dash: Exclude<ShapeOutline["dash"], "solid">): string;
  /** Routes a definition element preserved inside a block's
   * `unsupportedStyleXml` back into styles.xml (deduped by draw:name). */
  addPreservedDef(el: Element): void;
  /** Every collected definition element, serialised in first-use order. */
  defsXml(): string;
}

export function createOdfDrawingDefs(): OdfDrawingDefs {
  const defs: string[] = [];
  const claimedNames = new Set<string>();
  const byKey = new Map<string, string>();
  let gradCounter = 0;
  let hatchCounter = 0;
  let imageCounter = 0;
  const register = (key: string, mint: () => { name: string; xml: string }): string => {
    const existing = byKey.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const { name, xml } = mint();
    byKey.set(key, name);
    claimedNames.add(name);
    defs.push(xml);
    return name;
  };
  return {
    gradientNameFor(gradient: ShapeGradient): string {
      const style = gradient.type === "radial" ? "radial" : "linear";
      const start = gradient.stops[0]?.color ?? "#000000";
      const end = gradient.stops[gradient.stops.length - 1]?.color ?? start;
      // ODF's draw:gradient holds exactly two colours: intermediate stops
      // and per-stop alpha have no slot (the drop is reported centrally).
      const angleAttr = style === "linear" ? ` draw:angle="${gradientAngleToOdf(gradient.angleDeg ?? 0)}"` : "";
      const body =
        ` draw:style="${style}" draw:start-color="${escapeXmlAttribute(start)}"` +
        ` draw:end-color="${escapeXmlAttribute(end)}"` +
        ` draw:start-intensity="100%" draw:end-intensity="100%"${angleAttr}`;
      return register(`gradient|${body}`, () => {
        const name = `deGrad${++gradCounter}`;
        return { name, xml: `<draw:gradient${DEF_DRAW_XMLNS} draw:name="${name}"${body}/>` };
      });
    },
    hatchNameFor(hatch: { style: "single" | "double"; angleDeg: number }, foreground: string): string {
      const body =
        ` draw:style="${hatch.style}" draw:color="${escapeXmlAttribute(foreground)}"` +
        ` draw:rotation="${hatch.angleDeg * 10}" draw:distance="0.15cm"`;
      return register(`hatch|${body}`, () => {
        const name = `deHatch${++hatchCounter}`;
        return { name, xml: `<draw:hatch${DEF_DRAW_XMLNS} draw:name="${name}"${body}/>` };
      });
    },
    fillImageNameFor(href: string): string {
      return register(`fill-image|${href}`, () => {
        const name = `deFillImg${++imageCounter}`;
        return {
          name,
          xml:
            `<draw:fill-image${DEF_DRAW_XMLNS}${DEF_XLINK_XMLNS} draw:name="${name}"` +
            ` xlink:href="${escapeXmlAttribute(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`,
        };
      });
    },
    markerNameFor(kind: Exclude<ArrowheadKind, "none">): string {
      const def = MARKER_DEFS[kind];
      return register(`marker|${def.name}`, () => ({
        name: def.name,
        xml:
          `<draw:marker${DEF_DRAW_XMLNS}${DEF_SVG_XMLNS} draw:name="${def.name}"` +
          ` svg:viewBox="0 0 10 10" svg:d="${def.d}"/>`,
      }));
    },
    strokeDashNameFor(dash: Exclude<ShapeOutline["dash"], "solid">): string {
      const def = DASH_DEFS[dash];
      return register(`stroke-dash|${def.name}`, () => ({
        name: def.name,
        xml: `<draw:stroke-dash${DEF_DRAW_XMLNS} draw:name="${def.name}" draw:style="rect"${def.body}/>`,
      }));
    },
    addPreservedDef(el: Element): void {
      const name = el.getAttributeNS(DRAW_NAMESPACE, "name") ?? "";
      if (claimedNames.has(name)) {
        return; // A definition by this name is already scheduled.
      }
      register(`preserved|${el.localName}|${name}`, () => ({ name, xml: serializeElement(el) }));
    },
    defsXml(): string {
      return defs.join("");
    },
  };
}

/** Everything needed to parse a shape's text:p children into
 * `Shape.text.blocks` — injected by the caller (reader.ts) so this module
 * stays independent of the paragraph parser's own dependency graph. */
export interface OdfShapeTextContext {
  parseParagraph: (pEl: Element) => Paragraph;
  /** Resolves a picture fill's `xlink:href` (a package path like
   * "Pictures/image1.png") to a mediaId, registering the media part with the
   * reader — mirrors how odf/drawing.ts reads a `draw:image`. Absent (or
   * returning null) leaves the bitmap fill preserved verbatim. */
  resolveFillImage?: (href: string) => string | null;
}

function vAlignOf(el: Element): "top" | "middle" | "bottom" {
  const anchor = drawAttr(el, "text-vertical-align") ?? styleAttr(el, "vertical-align");
  if (anchor === "middle" || anchor === "center") {
    return "middle";
  }
  if (anchor === "bottom") {
    return "bottom";
  }
  return "top";
}

function parseShapeText(shapeEl: Element, ctx: OdfShapeTextContext): Shape["text"] | undefined {
  const textBox = firstChildElement(shapeEl, DRAW_NAMESPACE, "text-box");
  const container = textBox ?? shapeEl;
  const paragraphs = Array.from(container.children).filter(
    (child) => child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h"),
  );
  if (paragraphs.length === 0) {
    return undefined;
  }
  const blocks: Block[] = paragraphs.map((p) => ctx.parseParagraph(p));
  return { blocks, vAlign: vAlignOf(textBox ?? shapeEl) };
}

/** What one graphic style resolved to (SPEC-073): the modelled properties
 * plus the verbatim leftovers for `unsupportedStyleXml` — a synthetic
 * `style:graphic-properties` element carrying ONLY the attributes/children
 * this parse did not interpret (AC-35/AC-36), so a same-format save can
 * merge them back without ever conflicting with a modelled property. */
export interface ParsedGraphicStyle {
  fill?: ShapeFill;
  outline?: ShapeOutline;
  shadow?: ShapeShadow;
  wrap: ImageWrapMode;
  unsupportedStyleXml?: string;
}

export interface ParseGraphicStyleOptions {
  resolveFillImage?: (href: string) => string | null;
  /** "model": de-named `draw:marker-*` attrs become outline arrowheads
   * (shapes); "consume": marker attrs are known but regenerated from the
   * block's own fields, so they never land in the preserved set
   * (connectors, whose arrows ride in `de:` element attributes). */
  markers: "model" | "consume";
  /** Connectors have no fill in the model: `false` consumes only a
   * meaningless draw:fill/draw:fill-color pair (a connector has no area)
   * and leaves any richer fill vocabulary preserved. */
  parseFill: boolean;
}

const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

/** The consumed-attribute key: namespace + local name. */
function attrKey(ns: string | null, local: string): string {
  return `${ns ?? ""}|${local}`;
}

/** Snaps a floating-point length that is zero-up-to-rounding (cos(90°) is
 * 6.1e-17, not 0) to exactly 0 so it formats as "0cm". */
function snapZero(value: number): number {
  return Math.abs(value) < 1e-9 ? 0 : value;
}

/** Declares `xmlns:prefix` for every prefixed attribute/element about to be
 * placed on `el`, so serialising the (detached) element keeps the source
 * prefixes instead of the serializer's generated ns1/ns2. */
function declareNamespacesFor(el: Element, attrs: readonly Attr[], elements: readonly Element[]): void {
  const declared = new Set<string>([el.prefix ?? ""]);
  const declare = (prefix: string | null, ns: string | null) => {
    if (!ns || !prefix || declared.has(prefix)) {
      return;
    }
    declared.add(prefix);
    el.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, ns);
  };
  for (const attr of attrs) {
    declare(attr.prefix, attr.namespaceURI);
  }
  for (const child of elements) {
    declare(child.prefix, child.namespaceURI);
  }
}

/** Finds a named drawing definition (`draw:gradient` etc.) in content.xml or
 * styles.xml — definitions this writer emits live in styles.xml
 * `<office:styles>`, but LibreOffice may put them either place. */
function findDrawDef(
  localName: string,
  name: string,
  contentDoc: Document,
  stylesDoc: Document | undefined,
): Element | undefined {
  for (const doc of [contentDoc, stylesDoc]) {
    if (!doc) {
      continue;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(DRAW_NAMESPACE, localName))) {
      if (el.getAttributeNS(DRAW_NAMESPACE, "name") === name) {
        return el;
      }
    }
  }
  return undefined;
}

/** Parses a `draw:gradient` definition, or null when it uses vocabulary the
 * model cannot hold (axial/ellipsoid/square styles) — the caller then
 * preserves the referencing attributes and the definition verbatim. */
function parseGradientDef(def: Element): ShapeGradient | null {
  const style = drawAttr(def, "style");
  if (style !== "linear" && style !== "radial") {
    return null;
  }
  const start = drawAttr(def, "start-color");
  const end = drawAttr(def, "end-color");
  if (!start || !end) {
    return null;
  }
  const stops = [
    { position: 0, color: start.toLowerCase() },
    { position: 1, color: end.toLowerCase() },
  ];
  if (style === "radial") {
    return { type: "radial", stops };
  }
  const angleRaw = drawAttr(def, "angle");
  const angle10 = angleRaw !== null ? parseFloat(angleRaw) : 0;
  return {
    type: "linear",
    angleDeg: odfGradientAngleToDeg(Number.isFinite(angle10) ? angle10 : 0),
    stops,
  };
}

/** Resolves a drawing element's graphic style into the SPEC-073 model:
 * fill/outline/shadow/wrap from the attribute vocabulary this module's own
 * writer emits, everything else preserved verbatim (see ParsedGraphicStyle).
 * Shared by shapes and connectors. */
export function parseGraphicStyleProps(
  el: Element,
  contentDoc: Document,
  stylesDoc: Document | undefined,
  options: ParseGraphicStyleOptions,
): ParsedGraphicStyle {
  const styleName = drawAttr(el, "style-name");
  const find = (doc: Document | undefined) => {
    if (!doc || !styleName) {
      return undefined;
    }
    for (const styleEl of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (styleEl.getAttributeNS(STYLE_NAMESPACE, "name") === styleName) {
        return firstChildElement(styleEl, STYLE_NAMESPACE, "graphic-properties");
      }
    }
    return undefined;
  };
  const props = find(contentDoc) ?? find(stylesDoc);

  const result: ParsedGraphicStyle = { wrap: "inline" };
  if (props) {
    // Every attribute this parse interprets is recorded here; whatever is
    // left over (plus any child elements) becomes unsupportedStyleXml.
    const consumed = new Set<string>();
    const consume = (ns: string, local: string): void => {
      consumed.add(attrKey(ns, local));
    };
    const take = (ns: string, local: string): string | null => {
      consume(ns, local);
      return props.getAttributeNS(ns, local);
    };
    /** An opacity percentage attribute as a 0..1 alpha — 100% (opaque) maps
     * to `undefined`, matching the model's "absent = opaque" (D5). */
    const takeAlpha = (ns: string, local: string): number | undefined => {
      const raw = props.getAttributeNS(ns, local);
      if (raw === null) {
        return undefined;
      }
      const pct = parseFloat(raw);
      if (!Number.isFinite(pct)) {
        return undefined; // Malformed — leave it preserved.
      }
      consume(ns, local);
      return pct >= 100 ? undefined : Math.max(0, pct) / 100;
    };
    // Inlined into unsupportedStyleXml so a same-format save can route the
    // definition back into styles.xml (the defs pre-pass extracts them).
    const preservedDefs: Element[] = [];

    // --- fill -------------------------------------------------------------
    // draw:fill/draw:fill-color live in the draw: namespace; stroke details
    // reuse SVG's attribute names (svg: namespace, ODF's convention) —
    // matching graphicPropertiesXml's own emission, which is what a
    // same-format save must read back.
    const fillMode = drawAttr(props, "fill");
    if (!options.parseFill) {
      // A connector has no area: its style's draw:fill pair means nothing
      // and is simply regenerated (i.e. omitted) on save.
      consume(DRAW_NAMESPACE, "fill");
      consume(DRAW_NAMESPACE, "fill-color");
    } else if (fillMode === "solid") {
      consume(DRAW_NAMESPACE, "fill");
      const color = take(DRAW_NAMESPACE, "fill-color");
      if (color) {
        const fill: ShapeFill = { kind: "solid", color: color.toLowerCase() };
        const alpha = takeAlpha(DRAW_NAMESPACE, "opacity");
        if (alpha !== undefined) {
          fill.alpha = alpha;
        }
        result.fill = fill;
      }
    } else if (fillMode === null || fillMode === "none") {
      // Both the model's "no fill" forms write draw:fill="none"; reading it
      // back as NO fill (undefined, never {kind:"none"}) keeps text-box
      // detection and the shipped tests working.
      consume(DRAW_NAMESPACE, "fill");
    } else if (fillMode === "gradient") {
      const name = drawAttr(props, "fill-gradient-name");
      const def = name ? findDrawDef("gradient", name, contentDoc, stylesDoc) : undefined;
      const gradient = def ? parseGradientDef(def) : null;
      if (gradient) {
        consume(DRAW_NAMESPACE, "fill");
        consume(DRAW_NAMESPACE, "fill-gradient-name");
        result.fill = { kind: "gradient", gradient };
      } else if (def) {
        preservedDefs.push(def);
      }
    } else if (fillMode === "bitmap") {
      const name = drawAttr(props, "fill-image-name");
      const def = name ? findDrawDef("fill-image", name, contentDoc, stylesDoc) : undefined;
      const href = def ? def.getAttributeNS(XLINK_NAMESPACE, "href") : null;
      const mediaId = href && options.resolveFillImage ? options.resolveFillImage(href) : null;
      if (mediaId) {
        consume(DRAW_NAMESPACE, "fill");
        consume(DRAW_NAMESPACE, "fill-image-name");
        const repeat = take(STYLE_NAMESPACE, "repeat");
        const fill: ShapeFill = { kind: "picture", mediaId, mode: repeat === "repeat" ? "tile" : "stretch" };
        const alpha = takeAlpha(DRAW_NAMESPACE, "opacity");
        if (alpha !== undefined) {
          fill.alpha = alpha;
        }
        result.fill = fill;
      } else if (def) {
        preservedDefs.push(def);
      }
    } else if (fillMode === "hatch") {
      const name = drawAttr(props, "fill-hatch-name");
      const def = name ? findDrawDef("hatch", name, contentDoc, stylesDoc) : undefined;
      const hatchStyle = def ? drawAttr(def, "style") : null;
      const rotation = def ? parseFloat(drawAttr(def, "rotation") ?? "0") : NaN;
      const preset =
        (hatchStyle === "single" || hatchStyle === "double") && Number.isFinite(rotation)
          ? odfHatchToPatternPreset(hatchStyle, rotation / 10)
          : undefined;
      if (def && preset) {
        consume(DRAW_NAMESPACE, "fill");
        consume(DRAW_NAMESPACE, "fill-hatch-name");
        consume(DRAW_NAMESPACE, "fill-hatch-solid");
        const background = take(DRAW_NAMESPACE, "fill-color");
        result.fill = {
          kind: "pattern",
          preset,
          foreground: (drawAttr(def, "color") ?? "#000000").toLowerCase(),
          background: (background ?? "#ffffff").toLowerCase(),
        };
      } else if (def) {
        // A triple hatch / unmapped angle: the referencing attributes stay
        // preserved and the definition rides along for a same-format save.
        preservedDefs.push(def);
      }
    }
    // Any other draw:fill value stays entirely in the preserved set.

    // --- outline ----------------------------------------------------------
    const stroke = drawAttr(props, "stroke");
    if (stroke === "solid" || stroke === "dash") {
      consume(DRAW_NAMESPACE, "stroke");
      const color = take(SVG_NAMESPACE, "stroke-color") ?? "#000000";
      const widthCm = odfLengthToCm(take(SVG_NAMESPACE, "stroke-width")) ?? 0.03;
      let dash: ShapeOutline["dash"] = "solid";
      if (stroke === "dash") {
        const dashName = take(DRAW_NAMESPACE, "stroke-dash");
        dash = (dashName !== null ? DASH_NAME_TO_MODEL[dashName] : undefined) ?? "dash";
      }
      const outline: ShapeOutline = { color: color.toLowerCase(), widthPt: (widthCm / 2.54) * 72, dash };
      const alpha = takeAlpha(SVG_NAMESPACE, "stroke-opacity");
      if (alpha !== undefined) {
        outline.alpha = alpha;
      }
      const join = props.getAttributeNS(SVG_NAMESPACE, "stroke-linejoin");
      if (join === "round" || join === "bevel" || join === "miter") {
        consume(SVG_NAMESPACE, "stroke-linejoin");
        outline.join = join;
      }
      const cap = props.getAttributeNS(SVG_NAMESPACE, "stroke-linecap");
      if (cap === "round" || cap === "square") {
        consume(SVG_NAMESPACE, "stroke-linecap");
        outline.cap = cap;
      } else if (cap === "butt") {
        consume(SVG_NAMESPACE, "stroke-linecap"); // butt IS the model's flat default.
      }
      result.outline = outline;
    } else if (stroke === "none") {
      consume(DRAW_NAMESPACE, "stroke");
    }

    // --- arrow markers ----------------------------------------------------
    if (options.markers === "consume") {
      consume(DRAW_NAMESPACE, "marker-start");
      consume(DRAW_NAMESPACE, "marker-end");
      consume(DRAW_NAMESPACE, "marker-start-width");
      consume(DRAW_NAMESPACE, "marker-end-width");
    } else if (result.outline) {
      // Only OUR de-named markers are explicit model arrowheads (SPEC-073);
      // a foreign marker name stays preserved verbatim.
      const markerStart = drawAttr(props, "marker-start");
      if (markerStart !== null && MARKER_NAME_TO_ARROW[markerStart]) {
        consume(DRAW_NAMESPACE, "marker-start");
        result.outline.startArrow = {
          kind: MARKER_NAME_TO_ARROW[markerStart],
          size: arrowSizeFromWidth(take(DRAW_NAMESPACE, "marker-start-width")),
        };
      }
      const markerEnd = drawAttr(props, "marker-end");
      if (markerEnd !== null && MARKER_NAME_TO_ARROW[markerEnd]) {
        consume(DRAW_NAMESPACE, "marker-end");
        result.outline.endArrow = {
          kind: MARKER_NAME_TO_ARROW[markerEnd],
          size: arrowSizeFromWidth(take(DRAW_NAMESPACE, "marker-end-width")),
        };
      }
    }

    // --- shadow -----------------------------------------------------------
    if (drawAttr(props, "shadow") === "visible") {
      consume(DRAW_NAMESPACE, "shadow");
      const color = (take(DRAW_NAMESPACE, "shadow-color") ?? "#808080").toLowerCase();
      const opacityRaw = take(DRAW_NAMESPACE, "shadow-opacity");
      const opacityPct = opacityRaw !== null ? parseFloat(opacityRaw) : NaN;
      const x = odfLengthToCm(take(DRAW_NAMESPACE, "shadow-offset-x")) ?? 0;
      const y = odfLengthToCm(take(DRAW_NAMESPACE, "shadow-offset-y")) ?? 0;
      consume(DE_NAMESPACE, "shadow-blur-pt");
      const blurRaw = props.getAttributeNS(DE_NAMESPACE, "shadow-blur-pt");
      const distCm = Math.hypot(x, y);
      result.shadow = {
        color,
        alpha: Number.isFinite(opacityPct) ? Math.max(0, Math.min(100, opacityPct)) / 100 : 1,
        // Blur has no ODF slot — it round-trips through the de: extension.
        blurPt: blurRaw !== null && Number.isFinite(Number(blurRaw)) ? Number(blurRaw) : 0,
        distancePt: Math.round((distCm / 2.54) * 72 * 100) / 100,
        angleDeg:
          distCm > 0 ? ((Math.round((Math.atan2(y, x) * 180) / Math.PI) % 360) + 360) % 360 : 0,
      };
    }

    // --- wrap and the constant placement attrs ----------------------------
    const wrapAttr = styleAttr(props, "wrap");
    const runThrough = styleAttr(props, "run-through");
    if (wrapAttr === "parallel" || wrapAttr === "dynamic") {
      result.wrap = "square";
      consume(STYLE_NAMESPACE, "wrap");
    } else if (wrapAttr === "run-through" && runThrough === "background") {
      result.wrap = "behind";
      consume(STYLE_NAMESPACE, "wrap");
      consume(STYLE_NAMESPACE, "run-through");
    } else if (wrapAttr === "run-through") {
      result.wrap = "infront";
      consume(STYLE_NAMESPACE, "wrap");
      consume(STYLE_NAMESPACE, "run-through");
    } else if (wrapAttr === "none") {
      result.wrap = "topAndBottom";
      consume(STYLE_NAMESPACE, "wrap");
    }
    // Constants this writer always emits — consumed so a round trip of our
    // own output never mistakes them for foreign styling.
    consume(STYLE_NAMESPACE, "number-wrapped-paragraphs");
    consume(STYLE_NAMESPACE, "vertical-pos");
    consume(STYLE_NAMESPACE, "vertical-rel");
    consume(STYLE_NAMESPACE, "horizontal-pos");
    consume(STYLE_NAMESPACE, "horizontal-rel");

    // --- the preserved leftovers (AC-35) ----------------------------------
    const keepAttrs: Attr[] = [];
    for (let i = 0; i < props.attributes.length; i++) {
      const attr = props.attributes[i];
      if (attr.namespaceURI === XMLNS_NAMESPACE) {
        continue; // Serialisation re-declares whatever it needs.
      }
      if (consumed.has(attrKey(attr.namespaceURI, attr.localName))) {
        continue;
      }
      keepAttrs.push(attr);
    }
    const keepChildren = Array.from(props.children);
    if (keepAttrs.length > 0 || keepChildren.length > 0 || preservedDefs.length > 0) {
      const keep = props.ownerDocument.createElementNS(STYLE_NAMESPACE, "style:graphic-properties");
      // The synthetic element is detached, so the serializer would invent
      // ns1/ns2 prefixes — declare each source prefix explicitly instead,
      // keeping the preserved markup readable and byte-stable across trips.
      declareNamespacesFor(keep, keepAttrs, [...keepChildren, ...preservedDefs]);
      for (const attr of keepAttrs) {
        if (attr.namespaceURI) {
          keep.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
        } else {
          keep.setAttribute(attr.name, attr.value);
        }
      }
      for (const child of keepChildren) {
        keep.appendChild(child.cloneNode(true));
      }
      for (const def of preservedDefs) {
        keep.appendChild(def.cloneNode(true));
      }
      result.unsupportedStyleXml = serializeElement(keep);
    }
  }
  const anchorType = textAttr(el, "anchor-type");
  if (anchorType === null || anchorType === "as-char") {
    result.wrap = "inline";
  }
  return result;
}

export interface ParsedShape {
  shape: Shape;
}

/** Parses a `draw:custom-shape` or `draw:line` element into a Shape, or
 * null when its geometry/type is outside the catalogue this spec models —
 * the caller's SPEC-008 opaque fallback then preserves it verbatim. */
export function parseShapeElement(
  el: Element,
  contentDoc: Document,
  stylesDoc: Document | undefined,
  ctx: OdfShapeTextContext,
): ParsedShape | null {
  const isLine = el.namespaceURI === DRAW_NAMESPACE && el.localName === "line";
  const isCustom = el.namespaceURI === DRAW_NAMESPACE && el.localName === "custom-shape";
  if (!isLine && !isCustom) {
    return null;
  }

  const { fill, outline, shadow, wrap, unsupportedStyleXml } = parseGraphicStyleProps(el, contentDoc, stylesDoc, {
    resolveFillImage: ctx.resolveFillImage,
    markers: "model",
    parseFill: true,
  });
  const descEl = firstChildElement(el, SVG_NAMESPACE, "desc");
  const altText = descEl?.textContent || undefined;
  const groupOrdinalAttr = el.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;
  const rotationAttr = el.getAttributeNS(DE_NAMESPACE, "rotation");
  const rotationDeg = rotationAttr ? Number(rotationAttr) : undefined;

  let widthCm: number;
  let heightCm: number;
  let geometry: ShapeGeometry;
  let unsupportedGeometryXml: string | undefined;
  let path: Shape["path"];
  let pathSource: Shape["pathSource"];
  let adjustments: number[] = [];

  let lineOffsetXCm: number | undefined;
  let lineOffsetYCm: number | undefined;
  if (isLine) {
    const x1 = odfLengthToCm(svgAttr(el, "x1")) ?? 0;
    const y1 = odfLengthToCm(svgAttr(el, "y1")) ?? 0;
    const x2 = odfLengthToCm(svgAttr(el, "x2")) ?? 0;
    const y2 = odfLengthToCm(svgAttr(el, "y2")) ?? 0;
    widthCm = Math.abs(x2 - x1);
    heightCm = Math.abs(y2 - y1);
    // shapeElementXml always writes x1/y1 as the box origin (x2/y2 = origin
    // + width/height) — a draw:line has no separate svg:x/svg:y, unlike
    // draw:custom-shape below.
    lineOffsetXCm = x1;
    lineOffsetYCm = y1;
    // SPEC-070's legacy convention: an element-level draw:marker-end IS the
    // arrow geometry. SPEC-073's explicit arrowheads live on the graphic
    // STYLE as de-named markers instead, so an arrow-geometry line with an
    // explicit endArrow marks itself with de:arrow-geometry (the element
    // marker stays reserved for the legacy form).
    geometry =
      drawAttr(el, "marker-end") || el.getAttributeNS(DE_NAMESPACE, "arrow-geometry") ? "arrow" : "line";
  } else {
    const w = odfLengthToCm(svgAttr(el, "width"));
    const h = odfLengthToCm(svgAttr(el, "height"));
    if (w === undefined || h === undefined) {
      return null;
    }
    widthCm = w;
    heightCm = h;
    const enhancedGeometry = firstChildElement(el, DRAW_NAMESPACE, "enhanced-geometry");
    const drawType = enhancedGeometry ? drawAttr(enhancedGeometry, "type") : null;
    const mapped = drawType ? DRAW_TYPE_TO_GEOMETRY_TABLE[drawType] : undefined;
    geometry = mapped ?? "rect";
    if (!mapped) {
      // A non-primitive enhanced-geometry with an @draw:enhanced-path is a
      // path shape (SPEC-071 AC-5) — parse it and render its real outline,
      // unless it is formula-driven or uses an unmodelled command, in which
      // case it stays on SPEC-051's verbatim preservation path.
      const enhancedPath = enhancedGeometry ? drawAttr(enhancedGeometry, "enhanced-path") : null;
      const viewBox = enhancedGeometry ? svgAttr(enhancedGeometry, "viewBox") : null;
      const vb = viewBox ? viewBox.trim().split(/\s+/).map(Number) : null;
      const parsed =
        enhancedPath && vb && vb.length === 4 && Number.isFinite(vb[2]) && Number.isFinite(vb[3])
          ? parseEnhancedPath(enhancedPath, vb[2] - vb[0], vb[3] - vb[1])
          : null;
      if (parsed) {
        path = parsed.path;
        const smoothAttr = enhancedGeometry ? enhancedGeometry.getAttributeNS(DE_NAMESPACE, "smooth") : null;
        if (smoothAttr) {
          path.smooth = smoothAttr.split("").map((c) => c === "1");
        }
        if (parsed.converted && enhancedGeometry) {
          pathSource = { format: "odt" as const, xml: new XMLSerializer().serializeToString(enhancedGeometry) };
        }
      } else {
        unsupportedGeometryXml = enhancedGeometry ? new XMLSerializer().serializeToString(enhancedGeometry) : "";
      }
    } else if (enhancedGeometry) {
      // draw:modifiers is a space-separated list of adjustment values in the
      // shape's coordinate space, normalised to fractions (SPEC-069 AC-6b).
      const modifiers = drawAttr(enhancedGeometry, "modifiers");
      if (modifiers) {
        adjustments = modifiers
          .trim()
          .split(/\s+/)
          .map(Number)
          .filter((n) => !Number.isNaN(n))
          .map(guideValueToAdjustment);
      }
    }
  }

  const shape: Shape = { type: "shape", geometry, widthCm, heightCm, wrap };
  if (path) {
    shape.path = path;
    if (pathSource) {
      shape.pathSource = pathSource;
    }
  } else if (unsupportedGeometryXml !== undefined) {
    shape.unsupportedGeometryXml = unsupportedGeometryXml;
  }
  if (adjustments.length > 0) {
    shape.adjustments = adjustments;
  }
  if (fill) {
    shape.fill = fill;
  }
  if (outline) {
    shape.outline = outline;
  }
  if (shadow) {
    shape.shadow = shadow;
  }
  if (unsupportedStyleXml) {
    shape.unsupportedStyleXml = unsupportedStyleXml;
  }
  if (altText) {
    shape.altText = altText;
  }
  if (rotationDeg) {
    shape.rotationDeg = rotationDeg;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    shape.groupOrdinal = groupOrdinal;
  }
  if (wrap !== "inline") {
    shape.relativeToH = "column";
    shape.relativeToV = "paragraph";
    const offsetXCm = isLine ? lineOffsetXCm : odfLengthToCm(svgAttr(el, "x"));
    const offsetYCm = isLine ? lineOffsetYCm : odfLengthToCm(svgAttr(el, "y"));
    if (offsetXCm !== undefined) {
      shape.offsetXCm = offsetXCm;
    }
    if (offsetYCm !== undefined) {
      shape.offsetYCm = offsetYCm;
    }
  }
  if (!isLine) {
    const text = parseShapeText(el, ctx);
    if (text) {
      shape.text = text;
    }
  }
  return { shape };
}

/** What the graphic-style writers need beyond the block itself (SPEC-073):
 * the shared-definition registry, the media plan's href lookup for picture
 * fills, and whether preserved styling may be re-emitted (same-format saves
 * only — a cross-format model's `unsupportedStyleXml` is the OTHER format's
 * vocabulary and is reported as dropped instead). */
export interface OdfShapeStyleWriteContext {
  defs?: OdfDrawingDefs;
  hrefForMedia?: (mediaId: string) => string | undefined;
  /** Default true. */
  includePreserved?: boolean;
}

/** An opacity percentage attribute value ("40%") for a 0..1 alpha. */
function alphaPct(alpha: number): string {
  return `${Math.round(alpha * 100)}%`;
}

function fillStyleAttrs(shape: Shape, defs: OdfDrawingDefs, hrefForMedia?: (mediaId: string) => string | undefined): string {
  const fill = shape.fill;
  if (!fill || fill.kind === "none") {
    return ' draw:fill="none"';
  }
  const alphaAttr = "alpha" in fill && fill.alpha !== undefined && fill.alpha < 1 ? ` draw:opacity="${alphaPct(fill.alpha)}"` : "";
  if (fill.kind === "solid") {
    return ` draw:fill="solid" draw:fill-color="${escapeXmlAttribute(fill.color)}"${alphaAttr}`;
  }
  if (fill.kind === "gradient") {
    const name = defs.gradientNameFor(fill.gradient);
    return ` draw:fill="gradient" draw:fill-gradient-name="${escapeXmlAttribute(name)}"`;
  }
  if (fill.kind === "picture") {
    const href = hrefForMedia?.(fill.mediaId);
    if (!href) {
      return ' draw:fill="none"'; // Defensive: an unplanned media part paints nothing.
    }
    const name = defs.fillImageNameFor(href);
    return (
      ` draw:fill="bitmap" draw:fill-image-name="${escapeXmlAttribute(name)}"` +
      ` style:repeat="${fill.mode === "tile" ? "repeat" : "stretch"}"${alphaAttr}`
    );
  }
  // Pattern: only presets in the closed hatch table are expressible in ODF;
  // anything else flattens to the foreground colour (drop reported centrally).
  const hatch = PATTERN_PRESET_TO_ODF_HATCH[fill.preset];
  if (!hatch) {
    return ` draw:fill="solid" draw:fill-color="${escapeXmlAttribute(fill.foreground)}"`;
  }
  const name = defs.hatchNameFor(hatch, fill.foreground);
  return (
    ` draw:fill="hatch" draw:fill-hatch-name="${escapeXmlAttribute(name)}"` +
    ` draw:fill-hatch-solid="true" draw:fill-color="${escapeXmlAttribute(fill.background)}"`
  );
}

/** The stroke attribute run for an outline (SPEC-073) — shared by shape and
 * connector graphic styles so the two vocabularies cannot drift. The
 * `compound` member has no ODF vocabulary and writes as a single line (the
 * drop is reported centrally in opaque.ts). */
export function outlineStyleAttrs(outline: ShapeOutline, defs: OdfDrawingDefs): string {
  const widthCm = (outline.widthPt / 72) * 2.54;
  let attrs =
    ` draw:stroke="${outline.dash === "solid" ? "solid" : "dash"}"` +
    ` svg:stroke-color="${escapeXmlAttribute(outline.color)}"` +
    ` svg:stroke-width="${cmToOdfLength(widthCm)}"`;
  if (outline.dash !== "solid") {
    attrs += ` draw:stroke-dash="${defs.strokeDashNameFor(outline.dash)}"`;
  }
  if (outline.alpha !== undefined && outline.alpha < 1) {
    attrs += ` svg:stroke-opacity="${alphaPct(outline.alpha)}"`;
  }
  if (outline.join) {
    attrs += ` svg:stroke-linejoin="${outline.join}"`;
  }
  if (outline.cap && outline.cap !== "flat") {
    attrs += ` svg:stroke-linecap="${outline.cap}"`;
  }
  return attrs;
}

/** The de-named marker attribute run for a line shape's explicit outline
 * arrowheads (SPEC-073 — the legacy `arrow` geometry keeps its own
 * element-level `draw:marker-end="Arrow"` instead, see shapeElementXml). */
function shapeMarkerAttrs(outline: ShapeOutline | undefined, defs: OdfDrawingDefs): string {
  let attrs = "";
  if (outline?.startArrow && outline.startArrow.kind !== "none") {
    attrs +=
      ` draw:marker-start="${defs.markerNameFor(outline.startArrow.kind)}"` +
      ` draw:marker-start-width="${ARROW_WIDTH_CM[outline.startArrow.size]}"`;
  }
  if (outline?.endArrow && outline.endArrow.kind !== "none") {
    attrs +=
      ` draw:marker-end="${defs.markerNameFor(outline.endArrow.kind)}"` +
      ` draw:marker-end-width="${ARROW_WIDTH_CM[outline.endArrow.size]}"`;
  }
  return attrs;
}

/** The shadow attribute run (SPEC-073 AC-8): colour/opacity/offset in real
 * ODF vocabulary, blur in the de: extension (ODF has no blur slot) — the
 * caller declares xmlns:de. Offsets are the polar (distancePt, angleDeg)
 * pair resolved to cartesian centimetres. */
export function shadowStyleAttrs(shadow: ShapeShadow): string {
  const distCm = (shadow.distancePt / 72) * 2.54;
  const rad = (shadow.angleDeg * Math.PI) / 180;
  return (
    ` draw:shadow="visible" draw:shadow-color="${escapeXmlAttribute(shadow.color)}"` +
    ` draw:shadow-opacity="${alphaPct(shadow.alpha)}"` +
    ` draw:shadow-offset-x="${cmToOdfLength(snapZero(Math.cos(rad) * distCm))}"` +
    ` draw:shadow-offset-y="${cmToOdfLength(snapZero(Math.sin(rad) * distCm))}"` +
    ` de:shadow-blur-pt="${shadow.blurPt}"`
  );
}

/** The wrap attribute run shared by shape and connector graphic styles. */
export function wrapStyleAttrs(wrap: ImageWrapMode): string {
  if (wrap === "square") {
    return ' style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"';
  }
  if (wrap === "topAndBottom") {
    return ' style:wrap="none"';
  }
  if (wrap === "behind") {
    return ' style:wrap="run-through" style:run-through="background"';
  }
  if (wrap === "infront") {
    return ' style:wrap="run-through" style:run-through="foreground"';
  }
  return "";
}

const PLACEMENT_ATTRS =
  ' style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"';

/** Merges a block's preserved `unsupportedStyleXml` (a synthetic
 * style:graphic-properties element, see parseGraphicStyleProps) into the
 * generated element: preserved attributes join the generated ones (modelled
 * attributes win, though by construction the two sets are disjoint),
 * preserved child elements are re-appended, and any drawing DEFINITION
 * elements riding along are routed into the defs registry for styles.xml. */
export function mergePreservedGraphicProps(baseXml: string, preserved: Element, defs?: OdfDrawingDefs): string {
  let base: Element;
  try {
    const wrapped = parseXml(
      `<style:wrapper xmlns:style="${STYLE_NAMESPACE}" xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" xmlns:de="${DE_NAMESPACE}">${baseXml}</style:wrapper>`,
    );
    base = wrapped.documentElement.firstElementChild!;
  } catch {
    return baseXml;
  }
  const preservedAttrs = Array.from(preserved.attributes).filter((attr) => attr.namespaceURI !== XMLNS_NAMESPACE);
  const preservedChildren = Array.from(preserved.children).filter((child) => !isOdfDrawingDefElement(child));
  // serializeElement runs on the merged element alone: the wrapper's xmlns
  // declarations are out of scope then, so re-declare the preserved
  // prefixes on the element itself (see declareNamespacesFor).
  declareNamespacesFor(base, preservedAttrs, preservedChildren);
  for (const attr of preservedAttrs) {
    const exists = attr.namespaceURI
      ? base.hasAttributeNS(attr.namespaceURI, attr.localName)
      : base.hasAttribute(attr.name);
    if (exists) {
      continue;
    }
    if (attr.namespaceURI) {
      base.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
    } else {
      base.setAttribute(attr.name, attr.value);
    }
  }
  for (const child of Array.from(preserved.children)) {
    if (isOdfDrawingDefElement(child)) {
      defs?.addPreservedDef(child);
      continue;
    }
    base.appendChild(base.ownerDocument.importNode(child, true));
  }
  return serializeElement(base);
}

function parsePreservedStyle(xml: string | undefined): Element | undefined {
  if (!xml) {
    return undefined;
  }
  try {
    const el = parseXml(xml).documentElement;
    return el.namespaceURI === STYLE_NAMESPACE && el.localName === "graphic-properties" ? el : undefined;
  } catch {
    return undefined;
  }
}

function graphicPropertiesXml(shape: Shape, ctx: OdfShapeStyleWriteContext): string {
  const defs = ctx.defs ?? createOdfDrawingDefs();
  const preserved = ctx.includePreserved !== false ? parsePreservedStyle(shape.unsupportedStyleXml) : undefined;
  // A fill the reader could not model left its draw:fill attrs in the
  // preserved set (AC-35) — the writer must then not fight them with its own.
  const preservedHasFill = preserved !== undefined && preserved.hasAttributeNS(DRAW_NAMESPACE, "fill");
  const fillAttrs = preservedHasFill ? "" : fillStyleAttrs(shape, defs, ctx.hrefForMedia);
  const strokeAttrs = shape.outline ? outlineStyleAttrs(shape.outline, defs) : ' draw:stroke="none"';
  const markerAttrs = shapeMarkerAttrs(shape.outline, defs);
  const shadowAttrs = shape.shadow ? shadowStyleAttrs(shape.shadow) : "";
  const deXmlns = shadowAttrs ? ` xmlns:de="${DE_NAMESPACE}"` : "";
  const wrapAttrs = wrapStyleAttrs(shape.wrap);
  const base = `<style:graphic-properties${deXmlns}${fillAttrs}${strokeAttrs}${markerAttrs}${shadowAttrs}${wrapAttrs}${PLACEMENT_ATTRS}/>`;
  return preserved ? mergePreservedGraphicProps(base, preserved, defs) : base;
}

/** One generated graphic-family `style:style` for a shape — each shape gets
 * its own automatic style (no cross-shape dedup, unlike AutomaticStyles'
 * cached per-wrap-mode image styles) since fill/outline/wrap vary per shape. */
export function shapeGraphicStyleXml(name: string, shape: Shape, ctx: OdfShapeStyleWriteContext = {}): string {
  return `<style:style style:name="${escapeXmlAttribute(name)}"${SHAPE_XMLNS} style:family="graphic" style:parent-style-name="Graphics">${graphicPropertiesXml(shape, ctx)}</style:style>`;
}

export interface OdfShapeTextWriteContext {
  paragraphXml: (paragraph: Paragraph) => string;
}

function textBoxXml(shape: Shape, ctx: OdfShapeTextWriteContext): string {
  if (!shape.text) {
    return "";
  }
  const paragraphs = shape.text.blocks
    .map((block) => (block.type === "paragraph" ? ctx.paragraphXml(block) : ""))
    .join("");
  const vAlign =
    shape.text.vAlign === "middle" ? "middle" : shape.text.vAlign === "bottom" ? "bottom" : "top";
  return `<draw:text-box draw:text-vertical-align="${vAlign}">${paragraphs || "<text:p/>"}</draw:text-box>`;
}

/** Serialises a Shape into `draw:custom-shape` (area geometries) or
 * `draw:line` (Line/Arrow) — the shape-flavoured sibling of drawing.ts's
 * drawFrameXml. `styleName` is this shape's own generated graphic style
 * (see shapeGraphicStyleXml). */
export function shapeElementXml(shape: Shape, styleName: string, ctx: OdfShapeTextWriteContext, drawId?: string): string {
  const anchorType = shape.wrap === "inline" ? "as-char" : "paragraph";
  const descXml = shape.altText ? `<svg:desc>${escapeXmlAttribute(shape.altText)}</svg:desc>` : "";
  const groupAttr = shape.groupOrdinal !== undefined ? ` de:group="${shape.groupOrdinal}"` : "";
  const rotationAttr = shape.rotationDeg ? ` de:rotation="${shape.rotationDeg}"` : "";
  // SPEC-073: an arrow-geometry line WITH an explicit endArrow writes
  // de-named style markers instead of the legacy element marker, so the
  // geometry itself needs this extension attr to survive the round trip.
  const arrowGeometryAttr =
    shape.geometry === "arrow" && shape.outline?.endArrow ? ' de:arrow-geometry="1"' : "";
  const deXmlns = groupAttr || rotationAttr || arrowGeometryAttr ? ` xmlns:de="${DE_NAMESPACE}"` : "";
  // SPEC-070: a shape referenced by a connector's draw:start-shape/end-shape
  // carries the matching draw:id (also as xml:id, which ODF requires as the
  // canonical id) so the reference resolves on reopen.
  const idAttr = drawId
    ? ` xmlns:xml="http://www.w3.org/XML/1998/namespace" draw:id="${escapeXmlAttribute(drawId)}" xml:id="${escapeXmlAttribute(drawId)}"`
    : "";
  const xCm = shape.wrap !== "inline" ? shape.offsetXCm ?? 0 : undefined;
  const yCm = shape.wrap !== "inline" ? shape.offsetYCm ?? 0 : undefined;
  const posAttrs = xCm !== undefined && yCm !== undefined ? ` svg:x="${cmToOdfLength(xCm)}" svg:y="${cmToOdfLength(yCm)}"` : "";

  if (LINE_GEOMETRIES.has(shape.geometry) && !shape.unsupportedGeometryXml) {
    // The legacy arrow geometry (no explicit endArrow) keeps SPEC-070's
    // exact bytes; an explicit endArrow moved the marker into the graphic
    // style (see shapeMarkerAttrs) and rides on de:arrow-geometry instead.
    const markerAttr =
      shape.geometry === "arrow" && !shape.outline?.endArrow
        ? ' draw:marker-end="Arrow" draw:marker-end-width="0.3cm"'
        : "";
    return (
      `<draw:line${SHAPE_XMLNS} text:anchor-type="${anchorType}" draw:style-name="${escapeXmlAttribute(styleName)}"${idAttr}${deXmlns}${groupAttr}${rotationAttr}${arrowGeometryAttr}` +
      ` svg:x1="${cmToOdfLength(xCm ?? 0)}" svg:y1="${cmToOdfLength(yCm ?? 0)}"` +
      ` svg:x2="${cmToOdfLength((xCm ?? 0) + shape.widthCm)}" svg:y2="${cmToOdfLength((yCm ?? 0) + shape.heightCm)}"` +
      `${markerAttr}>` +
      descXml +
      "</draw:line>"
    );
  }

  const modifiersAttr =
    shape.adjustments && shape.adjustments.length > 0
      ? ` draw:modifiers="${shape.adjustments.map(adjustmentToGuideValue).join(" ")}"`
      : "";
  // A path shape writes as a non-primitive enhanced-geometry with a matching
  // @svg:viewBox on the shared 21600 grid (AC-27). A path reduced from arcs/
  // quadratics on read re-emits its verbatim source on a same-format save.
  const smoothAttr =
    shape.path?.smooth && shape.path.smooth.length > 0
      ? ` xmlns:de="${DE_NAMESPACE}" de:smooth="${shape.path.smooth.map((b) => (b ? "1" : "0")).join("")}"`
      : "";
  const geometryXml = shape.path
    ? shape.pathSource?.format === "odt"
      ? shape.pathSource.xml
      : `<draw:enhanced-geometry${smoothAttr} svg:viewBox="${enhancedPathViewBox()}" draw:type="non-primitive" draw:enhanced-path="${escapeXmlAttribute(enhancedPathString(shape.path))}"/>`
    : (shape.unsupportedGeometryXml ?? `<draw:enhanced-geometry draw:type="${GEOMETRY_TO_DRAW_TYPE[shape.geometry]}"${modifiersAttr}/>`);
  return (
    `<draw:custom-shape${SHAPE_XMLNS} text:anchor-type="${anchorType}" draw:style-name="${escapeXmlAttribute(styleName)}"${idAttr}${deXmlns}${groupAttr}${rotationAttr}` +
    `${posAttrs} svg:width="${cmToOdfLength(shape.widthCm)}" svg:height="${cmToOdfLength(shape.heightCm)}">` +
    descXml +
    textBoxXml(shape, ctx) +
    geometryXml +
    "</draw:custom-shape>"
  );
}
