// Canvas painting for a Shape block (SPEC-051 "Rendering and test surface"):
// only geometry, fill and outline are painted here — text is real DOM inside
// the ARIA proxy (see DocShapeBlock.vue), and the canvas never carries
// interactive/accessible state. Framework-free except for the
// CanvasRenderingContext2D parameter itself, which every caller (the editor
// canvas, the PDF/print DOM extractor) supplies.
import {
  LINE_GEOMETRIES,
  type Shape,
  type ShapeOutline,
} from "./documentModel";
import { lineSegment, roundRectRadius, scalePolygon, shapePolygon, type Point } from "./shapeGeometry";

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

function roundRectPath(ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number): void {
  const r = Math.min(roundRectRadius(widthPx, heightPx), widthPx / 2, heightPx / 2);
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

/** Paints a Shape's geometry/fill/outline into a `widthPx` x `heightPx`
 * canvas context — the context is expected to already be translated/scaled
 * to the shape's own top-left corner (0,0) in device pixels; rotation is
 * applied here, about the box centre. An unsupported geometry
 * (`unsupportedGeometryXml` set) paints as a dashed placeholder rectangle
 * at the shape's true bounds instead (AC-60). */
export function paintShape(ctx: CanvasRenderingContext2D, shape: Shape, widthPx: number, heightPx: number): void {
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

  if (shape.geometry === "ellipse") {
    ctx.beginPath();
    ctx.ellipse(cx, cy, Math.max(0, widthPx / 2), Math.max(0, heightPx / 2), 0, 0, Math.PI * 2);
  } else if (shape.geometry === "roundRect" || shape.geometry === "wedgeRoundRectCallout") {
    if (shape.geometry === "roundRect") {
      roundRectPath(ctx, widthPx, heightPx);
    } else {
      const unit = shapePolygon(shape.geometry);
      polygonPath(ctx, scalePolygon(unit ?? [], widthPx, heightPx));
    }
  } else {
    const unit = shapePolygon(shape.geometry);
    polygonPath(ctx, scalePolygon(unit ?? [], widthPx, heightPx));
  }

  if (shape.fill) {
    ctx.fillStyle = shape.fill;
    ctx.fill();
  }
  if (shape.outline) {
    applyOutline(ctx, shape.outline, PX_PER_PT);
    ctx.stroke();
  }
  ctx.restore();
}
