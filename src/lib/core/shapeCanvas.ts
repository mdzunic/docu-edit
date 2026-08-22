// Canvas painting for a shape (SPEC-051 "Rendering and test surface"): only
// geometry, fill and outline are painted here — text is real DOM inside
// the ARIA proxy (see DocShapeBlock.vue), and the canvas never carries
// interactive/accessible state. Typed against the core ShapeVisual (the
// visual subset of the document model's Shape block), so any host — the
// page canvas, a slide canvas, the PDF/print DOM extractor — can paint one.
// Framework-free except for the CanvasRenderingContext2D parameter itself,
// which every caller supplies.
import {
  LINE_GEOMETRIES,
  type ShapeOutline,
  type ShapeVisual,
} from "./shapes";
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

function applyOutline(ctx: CanvasRenderingContext2D, outline: ShapeOutline | undefined, widthPxPerPt: number): void {
  if (!outline) {
    return;
  }
  ctx.strokeStyle = outline.color;
  ctx.lineWidth = Math.max(1, outline.widthPt * widthPxPerPt);
  ctx.setLineDash(outline.dash === "dash" ? [6, 4] : outline.dash === "dot" ? [1, 3] : []);
}

function polygonPath(ctx: CanvasRenderingContext2D, points: Point[]): void {
  ctx.beginPath();
  points.forEach((p, i) => (i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)));
  ctx.closePath();
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
): void {
  const bounds = rotatedBounds(widthPx, heightPx, shape.rotationDeg ?? 0);
  const surfaceW = Math.max(1, Math.ceil(bounds.width));
  const surfaceH = Math.max(1, Math.ceil(bounds.height));
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
  canvas.style.left = `${-bounds.padX}px`;
  canvas.style.top = `${-bounds.padY}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.clearRect(0, 0, surfaceW, surfaceH);
  ctx.translate(bounds.padX, bounds.padY);
  paintShape(ctx, shape, widthPx, heightPx);
}

/** Paints a Shape's geometry/fill/outline into a `widthPx` x `heightPx`
 * canvas context — the context is expected to already be translated/scaled
 * to the shape's own top-left corner (0,0) in device pixels; rotation is
 * applied here, about the box centre. An unsupported geometry
 * (`unsupportedGeometryXml` set) paints as a dashed placeholder rectangle
 * at the shape's true bounds instead (AC-60). */
export function paintShape(ctx: CanvasRenderingContext2D, shape: ShapeVisual, widthPx: number, heightPx: number): void {
  ctx.save();
  const cx = widthPx / 2;
  const cy = heightPx / 2;
  if (shape.rotationDeg) {
    ctx.translate(cx, cy);
    ctx.rotate((shape.rotationDeg * Math.PI) / 180);
    ctx.translate(-cx, -cy);
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
    const [a, b] = lineSegment(widthPx, heightPx);
    applyOutline(ctx, shape.outline, PX_PER_PT);
    ctx.beginPath();
    ctx.moveTo(a.x, a.y);
    ctx.lineTo(b.x, b.y);
    ctx.stroke();
    if (shape.geometry === "arrow") {
      paintArrowhead(ctx, a, b, Math.max(8, ctx.lineWidth * 3));
    }
    ctx.restore();
    return;
  }

  const def = shapeDefinition(shape.geometry);
  const adjustments = resolvedShapeAdjustments(shape.geometry, shape.adjustments);
  let holePaths: Point[][] = [];

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

  if (shape.fill) {
    ctx.fillStyle = shape.fill;
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
      ctx.fill("evenodd");
    } else {
      ctx.fill();
    }
  }
  if (shape.outline) {
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
