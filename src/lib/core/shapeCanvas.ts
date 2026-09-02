// Canvas painting for a shape (SPEC-051 "Rendering and test surface"): only
// geometry, fill and outline are painted here — text is real DOM inside
// the ARIA proxy (see DocShapeBlock.vue), and the canvas never carries
// interactive/accessible state. Typed against the core ShapeVisual (the
// visual subset of the document model's Shape block), so any host — the
// page canvas, a slide canvas, the PDF/print DOM extractor — can paint one.
// Framework-free except for the CanvasRenderingContext2D parameter itself,
// which every caller supplies.
import {
  ARROWHEAD_SIZE_FACTORS,
  LINE_GEOMETRIES,
  shapeFillIsVisible,
  type ShapeArrowhead,
  type ShapeFill,
  type ShapeOutline,
  type ShapeShadow,
  type ShapeVisual,
} from "./shapes";
import { arrowheadOutline } from "./connectors";
import {
  lineSegment,
  rotatedBounds,
  roundedCalloutOutline,
  roundRectRadius,
  scalePolygon,
  shapePolygon,
  type Point,
} from "./shapeGeometry";
import { resolvedShapeAdjustments, shapeDefinition } from "./shapeCatalogue";
import type { ShapePath } from "./shapePath";

/** Host hooks the painter cannot supply itself (SPEC-073): a picture fill's
 * decoded bitmap. The painter is synchronous — an unresolved (still-loading)
 * picture paints nothing and the host repaints when the bitmap arrives. */
export interface ShapePaintOptions {
  pictures?: (mediaId: string) => CanvasImageSource | null;
}

/** `#rrggbb` + a 0..1 opacity as a canvas-ready rgba() string; alpha absent
 * or 1 keeps the hex untouched (identical bytes, cheapest path). */
function withAlpha(hex: string, alpha: number | undefined): string {
  if (alpha === undefined || alpha >= 1) {
    return hex;
  }
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, ${Math.max(0, alpha)})`;
}

function dashPattern(dash: ShapeOutline["dash"]): number[] {
  switch (dash) {
    case "dash":
      return [6, 4];
    case "dot":
      return [1, 3];
    case "dashDot":
      return [6, 3, 1, 3];
    case "longDash":
      return [10, 4];
    case "longDashDot":
      return [10, 3, 1, 3];
    default:
      return [];
  }
}

function applyOutline(ctx: CanvasRenderingContext2D, outline: ShapeOutline | undefined, widthPxPerPt: number): void {
  if (!outline) {
    return;
  }
  ctx.strokeStyle = withAlpha(outline.color, outline.alpha);
  ctx.lineWidth = Math.max(1, outline.widthPt * widthPxPerPt);
  ctx.setLineDash(dashPattern(outline.dash));
  ctx.lineJoin = outline.join ?? "miter";
  ctx.lineCap = outline.cap === "round" ? "round" : outline.cap === "square" ? "square" : "butt";
}

/** The linear/radial CanvasGradient for a gradient fill over a w x h box.
 * Built in the context's own (already devicePixelRatio * zoom scaled)
 * coordinate space, so the ramp is computed per device pixel rather than
 * scaled as a bitmap (SPEC-063 / AC-39). */
function gradientStyle(
  ctx: CanvasRenderingContext2D,
  fill: Extract<ShapeFill, { kind: "gradient" }>,
  widthPx: number,
  heightPx: number,
): CanvasGradient {
  const { gradient } = fill;
  const cx = widthPx / 2;
  const cy = heightPx / 2;
  let g: CanvasGradient;
  if (gradient.type === "radial") {
    g = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.hypot(widthPx, heightPx) / 2);
  } else {
    const rad = ((gradient.angleDeg ?? 0) * Math.PI) / 180;
    const dx = Math.cos(rad);
    const dy = Math.sin(rad);
    const half = (Math.abs(widthPx * dx) + Math.abs(heightPx * dy)) / 2;
    g = ctx.createLinearGradient(cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half);
  }
  for (const stop of gradient.stops) {
    g.addColorStop(Math.min(1, Math.max(0, stop.position)), withAlpha(stop.color, stop.alpha));
  }
  return g;
}

/** The hatch line segments (in the unit of `widthPx`/`heightPx`) for a
 * pattern preset — drawn as real vectors under the scaled transform, never a
 * bitmap tile. The closed preset set mirrors the format mapping tables in
 * the docx/odf shape modules; an unknown preset paints as `ltDnDiag`. */
function patternSegments(preset: string, w: number, h: number): [Point, Point][] {
  const STEP = 6;
  const lines: [Point, Point][] = [];
  const horizontal = () => {
    for (let y = STEP / 2; y < h; y += STEP) lines.push([{ x: 0, y }, { x: w, y }]);
  };
  const vertical = () => {
    for (let x = STEP / 2; x < w; x += STEP) lines.push([{ x, y: 0 }, { x, y: h }]);
  };
  const downDiag = () => {
    for (let c = -h; c < w; c += STEP) lines.push([{ x: c, y: 0 }, { x: c + h, y: h }]);
  };
  const upDiag = () => {
    for (let c = 0; c < w + h; c += STEP) lines.push([{ x: c, y: 0 }, { x: c - h, y: h }]);
  };
  if (preset === "ltHorz") horizontal();
  else if (preset === "ltVert") vertical();
  else if (preset === "ltUpDiag") upDiag();
  else if (preset === "smGrid") {
    horizontal();
    vertical();
  } else if (preset === "diagCross") {
    downDiag();
    upDiag();
  } else downDiag();
  return lines;
}

/** Fills the path currently open in `ctx` with a `ShapeFill` (SPEC-073):
 * solid/gradient set a fillStyle; picture clips and draws the bitmap;
 * pattern paints its background then strokes the hatch vectors clipped to
 * the path. No-ops for `none`/absent and for a picture the host has not
 * resolved yet. */
function fillCurrentPath(
  ctx: CanvasRenderingContext2D,
  fill: ShapeFill,
  widthPx: number,
  heightPx: number,
  opts: ShapePaintOptions,
  fillRule?: CanvasFillRule,
): void {
  if (fill.kind === "solid" || fill.kind === "gradient") {
    ctx.fillStyle =
      fill.kind === "solid" ? withAlpha(fill.color, fill.alpha) : gradientStyle(ctx, fill, widthPx, heightPx);
    if (fillRule) {
      ctx.fill(fillRule);
    } else {
      ctx.fill();
    }
    return;
  }
  if (fill.kind === "picture") {
    const image = opts.pictures?.(fill.mediaId);
    if (!image) {
      return;
    }
    ctx.save();
    if (fillRule) {
      ctx.clip(fillRule);
    } else {
      ctx.clip();
    }
    ctx.globalAlpha *= fill.alpha === undefined ? 1 : Math.max(0, Math.min(1, fill.alpha));
    if (fill.mode === "tile") {
      const pattern = ctx.createPattern(image, "repeat");
      if (pattern) {
        ctx.fillStyle = pattern;
        ctx.fillRect(0, 0, widthPx, heightPx);
      }
    } else {
      ctx.drawImage(image, 0, 0, widthPx, heightPx);
    }
    ctx.restore();
    return;
  }
  if (fill.kind === "pattern") {
    ctx.save();
    if (fillRule) {
      ctx.clip(fillRule);
    } else {
      ctx.clip();
    }
    ctx.fillStyle = fill.background;
    ctx.fillRect(0, 0, widthPx, heightPx);
    ctx.strokeStyle = fill.foreground;
    ctx.lineWidth = 1;
    ctx.setLineDash([]);
    ctx.beginPath();
    for (const [a, b] of patternSegments(fill.preset, widthPx, heightPx)) {
      ctx.moveTo(a.x, a.y);
      ctx.lineTo(b.x, b.y);
    }
    ctx.stroke();
    ctx.restore();
  }
}

/** Paints the outer-shadow pass (SPEC-073 AC-9) for the silhouette `trace`
 * draws: the shape is drawn far off to the left with the canvas shadow
 * offset bringing only its shadow back into view. Blur and offset are
 * multiplied by the context's current scale because canvas shadow geometry
 * lives in device space, unaffected by the transform — which is exactly what
 * computes the shadow in the devicePixelRatio * zoom scaled space rather
 * than scaling it as a bitmap (AC-39). Callers skip this entirely for
 * shapes with no shadow (AC-40). */
function paintShadowPass(
  ctx: CanvasRenderingContext2D,
  shadow: ShapeShadow,
  trace: () => void,
  fillVisible: boolean,
  outline: ShapeOutline | undefined,
): void {
  const m = ctx.getTransform();
  const scale = Math.hypot(m.a, m.b) || 1;
  const OFFSET = 100000;
  const rad = (shadow.angleDeg * Math.PI) / 180;
  const distPx = shadow.distancePt * PX_PER_PT * scale;
  ctx.save();
  ctx.shadowColor = withAlpha(shadow.color, shadow.alpha);
  ctx.shadowBlur = shadow.blurPt * PX_PER_PT * scale;
  // The off-screen displacement happens along the (possibly rotated) local
  // x axis; the device-space shadow offset walks back along that same axis
  // (m's unit x column), plus the shadow's own absolute-direction distance.
  ctx.shadowOffsetX = (m.a / scale) * OFFSET + Math.cos(rad) * distPx;
  ctx.shadowOffsetY = (m.b / scale) * OFFSET + Math.sin(rad) * distPx;
  ctx.translate(-OFFSET / scale, 0);
  trace();
  if (fillVisible) {
    ctx.fillStyle = "#000000";
    ctx.fill();
  }
  if (outline) {
    ctx.strokeStyle = "#000000";
    ctx.lineWidth = Math.max(1, outline.widthPt * PX_PER_PT);
    ctx.setLineDash(dashPattern(outline.dash));
    ctx.stroke();
  }
  ctx.restore();
}

/** Paints one outline arrowhead (SPEC-073 AC-14) at `tip`, pointing out
 * along `along`, reusing the connector head geometry so every head is one
 * shape everywhere. */
function paintOutlineArrowhead(
  ctx: CanvasRenderingContext2D,
  arrow: ShapeArrowhead | undefined,
  tip: Point,
  along: Point,
): void {
  if (!arrow || arrow.kind === "none") {
    return;
  }
  const head = arrowheadOutline(arrow.kind, tip, along, ctx.lineWidth * ARROWHEAD_SIZE_FACTORS[arrow.size]);
  if (!head) {
    return;
  }
  ctx.beginPath();
  head.points.forEach((p, i) => (i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)));
  if (head.filled) {
    ctx.closePath();
    ctx.fillStyle = ctx.strokeStyle;
    ctx.setLineDash([]);
    ctx.fill();
  } else {
    ctx.setLineDash([]);
    ctx.stroke();
  }
}

function polygonPath(ctx: CanvasRenderingContext2D, points: Point[]): void {
  ctx.beginPath();
  points.forEach((p, i) => (i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)));
  ctx.closePath();
}

/** Traces a `ShapePath`'s subpaths into `ctx` in real px, honouring each
 * subpath's `closed` flag — the source of the real freeform outline
 * (SPEC-071). Fill uses the non-zero winding rule the spec states. */
function shapePathTrace(ctx: CanvasRenderingContext2D, path: ShapePath, widthPx: number, heightPx: number): void {
  ctx.beginPath();
  for (const subpath of path.subpaths) {
    let started = false;
    for (const step of subpath.steps) {
      if (step.op === "move") {
        ctx.moveTo(step.x * widthPx, step.y * heightPx);
        started = true;
      } else if (step.op === "line" && started) {
        ctx.lineTo(step.x * widthPx, step.y * heightPx);
      } else if (step.op === "cubic" && started) {
        ctx.bezierCurveTo(
          step.x1 * widthPx,
          step.y1 * heightPx,
          step.x2 * widthPx,
          step.y2 * heightPx,
          step.x * widthPx,
          step.y * heightPx,
        );
      } else if (step.op === "close") {
        ctx.closePath();
      }
    }
  }
}

function roundRectPath(ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number, radiusFraction?: number): void {
  const base = radiusFraction === undefined ? roundRectRadius(widthPx, heightPx) : Math.min(widthPx, heightPx) * radiusFraction;
  const r = Math.min(base, widthPx / 2, heightPx / 2);
  ctx.beginPath();
  ctx.moveTo(r, 0);
  ctx.lineTo(widthPx - r, 0);
  ctx.arcTo(widthPx, 0, widthPx, r, r);
  ctx.lineTo(widthPx, heightPx - r);
  ctx.arcTo(widthPx, heightPx, widthPx - r, heightPx, r);
  ctx.lineTo(r, heightPx);
  ctx.arcTo(0, heightPx, 0, heightPx - r, r);
  ctx.lineTo(0, r);
  ctx.arcTo(0, 0, r, 0, r);
  ctx.closePath();
}

/** The rounded callout: a rounded-rectangle body whose bottom edge detours
 * into the tail. arcTo needs no radius maths of its own here — every step's
 * endpoint is already the arc's tangent point (roundedCalloutOutline). */
function roundedCalloutPath(ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number): void {
  const { start, steps, radius } = roundedCalloutOutline(widthPx, heightPx);
  ctx.beginPath();
  ctx.moveTo(start.x, start.y);
  for (const step of steps) {
    if (step.corner) {
      ctx.arcTo(step.corner.x, step.corner.y, step.to.x, step.to.y, radius);
    } else {
      ctx.lineTo(step.to.x, step.to.y);
    }
  }
  ctx.closePath();
}

function paintArrowhead(ctx: CanvasRenderingContext2D, from: Point, to: Point, sizePx: number): void {
  const angle = Math.atan2(to.y - from.y, to.x - from.x);
  const spread = Math.PI / 7;
  ctx.beginPath();
  ctx.moveTo(to.x, to.y);
  ctx.lineTo(to.x - sizePx * Math.cos(angle - spread), to.y - sizePx * Math.sin(angle - spread));
  ctx.lineTo(to.x - sizePx * Math.cos(angle + spread), to.y - sizePx * Math.sin(angle + spread));
  ctx.closePath();
  ctx.fillStyle = ctx.strokeStyle;
  ctx.fill();
}

/** 1pt in px at 96dpi — outline widths are stored in points. */
const PX_PER_PT = 96 / 72;

/** Sizes `canvas` for `shape` at `widthPx` x `heightPx` and paints it.
 *
 * The one place a shape's canvas is sized, so the page's own render and the
 * live resize/draw previews cannot size it differently — and so a rotated
 * shape gets the bigger surface its corners need (rotatedBounds) in every
 * one of them. The canvas is offset back by that padding, leaving the
 * shape's own box exactly where the caller placed it; a caller only has to
 * not clip its overflow.
 *
 * Takes the element rather than a context because the sizing IS the point;
 * paintShape below stays the framework- and DOM-free core it wraps. */
export function paintShapeIntoCanvas(
  canvas: HTMLCanvasElement,
  shape: ShapeVisual,
  widthPx: number,
  heightPx: number,
  zoomFraction = 1,
  opts: ShapePaintOptions = {},
): void {
  const bounds = rotatedBounds(widthPx, heightPx, shape.rotationDeg ?? 0);
  // A shadow spills past the shape's own (already rotation-padded) box by
  // its offset plus blur; grow the surface so it is not clipped (AC-9).
  const shadowPad = shape.shadow
    ? Math.ceil((shape.shadow.distancePt + 2 * shape.shadow.blurPt) * PX_PER_PT)
    : 0;
  const surfaceW = Math.max(1, Math.ceil(bounds.width) + 2 * shadowPad);
  const surfaceH = Math.max(1, Math.ceil(bounds.height) + 2 * shadowPad);
  const dpr = typeof window === "undefined" ? 1 : window.devicePixelRatio || 1;
  // SPEC-063 AC-20: the CSS size stays the model px so the .pages transform
  // scales it, but the backing store is multiplied by devicePixelRatio * zoom
  // so the extra device pixels the transform will stretch it across are
  // painted crisply rather than upscaled. zoom defaults to 1 (identity).
  const scale = dpr * (zoomFraction || 1);
  canvas.width = Math.max(1, Math.round(surfaceW * scale));
  canvas.height = Math.max(1, Math.round(surfaceH * scale));
  canvas.style.width = `${surfaceW}px`;
  canvas.style.height = `${surfaceH}px`;
  canvas.style.left = `${-bounds.padX - shadowPad}px`;
  canvas.style.top = `${-bounds.padY - shadowPad}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.clearRect(0, 0, surfaceW, surfaceH);
  ctx.translate(bounds.padX + shadowPad, bounds.padY + shadowPad);
  paintShape(ctx, shape, widthPx, heightPx, opts);
}

/** Paints a Shape's geometry/fill/outline into a `widthPx` x `heightPx`
 * canvas context — the context is expected to already be translated/scaled
 * to the shape's own top-left corner (0,0) in device pixels; rotation is
 * applied here, about the box centre. An unsupported geometry
 * (`unsupportedGeometryXml` set) paints as a dashed placeholder rectangle
 * at the shape's true bounds instead (AC-60). */
export function paintShape(
  ctx: CanvasRenderingContext2D,
  shape: ShapeVisual,
  widthPx: number,
  heightPx: number,
  opts: ShapePaintOptions = {},
): void {
  ctx.save();
  const cx = widthPx / 2;
  const cy = heightPx / 2;
  if (shape.rotationDeg) {
    ctx.translate(cx, cy);
    ctx.rotate((shape.rotationDeg * Math.PI) / 180);
    ctx.translate(-cx, -cy);
  }
  const fillVisible = shapeFillIsVisible(shape.fill);

  // A path shape paints its real outline (SPEC-071), taking precedence over
  // `geometry` and never falling to the dashed placeholder.
  if (shape.path) {
    const path = shape.path;
    const trace = () => shapePathTrace(ctx, path, widthPx, heightPx);
    if (shape.shadow) {
      paintShadowPass(ctx, shape.shadow, trace, fillVisible, shape.outline);
    }
    trace();
    if (shape.fill && fillVisible) {
      fillCurrentPath(ctx, shape.fill, widthPx, heightPx, opts, "nonzero");
    }
    if (shape.outline) {
      trace();
      applyOutline(ctx, shape.outline, PX_PER_PT);
      ctx.stroke();
    }
    ctx.restore();
    return;
  }

  if (shape.unsupportedGeometryXml !== undefined) {
    ctx.strokeStyle = "#888888";
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 4]);
    ctx.strokeRect(0.75, 0.75, Math.max(0, widthPx - 1.5), Math.max(0, heightPx - 1.5));
    ctx.restore();
    return;
  }

  if (LINE_GEOMETRIES.has(shape.geometry)) {
    if (!shape.outline) {
      ctx.restore();
      return;
    }
    const outline = shape.outline;
    const [a, b] = lineSegment(widthPx, heightPx);
    const trace = () => {
      ctx.beginPath();
      ctx.moveTo(a.x, a.y);
      ctx.lineTo(b.x, b.y);
    };
    if (shape.shadow) {
      paintShadowPass(ctx, shape.shadow, trace, false, outline);
    }
    applyOutline(ctx, outline, PX_PER_PT);
    trace();
    ctx.stroke();
    // SPEC-073 AC-14: explicit arrowheads on any line geometry. The `arrow`
    // geometry keeps its SPEC-070 built-in end triangle when the outline
    // carries no explicit end arrow.
    if (outline.startArrow) {
      paintOutlineArrowhead(ctx, outline.startArrow, a, { x: a.x - b.x, y: a.y - b.y });
    }
    if (outline.endArrow) {
      paintOutlineArrowhead(ctx, outline.endArrow, b, { x: b.x - a.x, y: b.y - a.y });
    } else if (shape.geometry === "arrow") {
      paintArrowhead(ctx, a, b, Math.max(8, ctx.lineWidth * 3));
    }
    ctx.restore();
    return;
  }

  const def = shapeDefinition(shape.geometry);
  const adjustments = resolvedShapeAdjustments(shape.geometry, shape.adjustments);
  let holePaths: Point[][] = [];

  const traceOutline = () => {
    if (shape.geometry === "ellipse") {
      ctx.beginPath();
      ctx.ellipse(cx, cy, Math.max(0, widthPx / 2), Math.max(0, heightPx / 2), 0, 0, Math.PI * 2);
    } else if (shape.geometry === "roundRect" || shape.geometry === "flowChartAlternateProcess") {
      roundRectPath(ctx, widthPx, heightPx, adjustments[0]);
    } else if (shape.geometry === "wedgeRoundRectCallout") {
      roundedCalloutPath(ctx, widthPx, heightPx);
    } else if (def) {
      polygonPath(ctx, scalePolygon(def.outline(adjustments), widthPx, heightPx));
      holePaths = (def.holes?.(adjustments) ?? []).map((hole) => scalePolygon(hole, widthPx, heightPx));
    } else {
      const unit = shapePolygon(shape.geometry);
      polygonPath(ctx, scalePolygon(unit ?? [], widthPx, heightPx));
    }
  };

  if (shape.shadow) {
    paintShadowPass(ctx, shape.shadow, traceOutline, fillVisible, shape.outline);
  }
  traceOutline();

  if (shape.fill && fillVisible) {
    if (holePaths.length > 0) {
      // Punch the holes (donut, frame) out of the fill with the even-odd
      // rule — the outline path is already open in `ctx`.
      for (const hole of holePaths) {
        ctx.moveTo(hole[0].x, hole[0].y);
        for (let i = 1; i < hole.length; i++) {
          ctx.lineTo(hole[i].x, hole[i].y);
        }
        ctx.closePath();
      }
      fillCurrentPath(ctx, shape.fill, widthPx, heightPx, opts, "evenodd");
    } else {
      fillCurrentPath(ctx, shape.fill, widthPx, heightPx, opts);
    }
  }
  if (shape.outline) {
    traceOutline();
    applyOutline(ctx, shape.outline, PX_PER_PT);
    ctx.stroke();
    for (const hole of holePaths) {
      ctx.beginPath();
      ctx.moveTo(hole[0].x, hole[0].y);
      for (let i = 1; i < hole.length; i++) {
        ctx.lineTo(hole[i].x, hole[i].y);
      }
      ctx.closePath();
      ctx.stroke();
    }
  }
  ctx.restore();
}
