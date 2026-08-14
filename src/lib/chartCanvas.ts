// Canvas painting for a Chart block (SPEC-052 "Rendering and test
// surface" — the same canvas + ARIA-proxy split SPEC-051 established for
// Shape): geometry, axes, gridlines, series and data labels are painted
// here; the proxy element (DocChartBlock.vue) carries the interactive/
// accessible state. Framework-free except for the CanvasRenderingContext2D
// parameter itself, which every caller (the editor canvas, the PDF/print DOM
// extractor) supplies — mirrors shapeCanvas.ts's own contract.
import {
  DEFAULT_CHART_COLORS,
  SINGLE_SERIES_CHART_TYPES,
  type Chart,
  type ChartSeries,
} from "./documentModel";

const AXIS_COLOR = "#444444";
const GRIDLINE_COLOR = "#dddddd";
const TEXT_COLOR = "#222222";
const LABEL_FONT = "10px sans-serif";
const TITLE_FONT = "bold 13px sans-serif";
const AXIS_TITLE_FONT = "11px sans-serif";

interface Layout {
  plotX: number;
  plotY: number;
  plotW: number;
  plotH: number;
}

function seriesToPlot(chart: Chart): ChartSeries[] {
  return SINGLE_SERIES_CHART_TYPES.has(chart.chartType) ? chart.series.slice(0, 1) : chart.series;
}

function valueExtent(chart: Chart, series: ChartSeries[]): { min: number; max: number } {
  const isStacked = chart.chartType === "columnStacked";
  let min = 0;
  let max = 0;
  if (isStacked) {
    for (let i = 0; i < chart.categories.length; i++) {
      let posSum = 0;
      let negSum = 0;
      for (const s of series) {
        const v = s.values[i] ?? 0;
        if (v >= 0) {
          posSum += v;
        } else {
          negSum += v;
        }
      }
      max = Math.max(max, posSum);
      min = Math.min(min, negSum);
    }
  } else {
    for (const s of series) {
      for (const v of s.values) {
        if (v === null) {
          continue;
        }
        max = Math.max(max, v);
        min = Math.min(min, v);
      }
    }
  }
  if (chart.options.valueAxisMin !== undefined) {
    min = chart.options.valueAxisMin;
  }
  if (chart.options.valueAxisMax !== undefined) {
    max = chart.options.valueAxisMax;
  }
  if (min === max) {
    max = min + 1;
  }
  return { min, max };
}

function paintLegend(
  ctx: CanvasRenderingContext2D,
  chart: Chart,
  series: ChartSeries[],
  widthPx: number,
  heightPx: number,
  titleH: number,
): { reservedLeft: number; reservedRight: number; reservedTop: number; reservedBottom: number } {
  const reserved = { reservedLeft: 0, reservedRight: 0, reservedTop: 0, reservedBottom: 0 };
  const pos = chart.options.legend;
  if (pos === "none" || series.length === 0) {
    return reserved;
  }
  const swatch = 8;
  const gap = 4;
  ctx.font = LABEL_FONT;
  if (pos === "right" || pos === "left") {
    const width = Math.min(
      widthPx * 0.35,
      Math.max(...series.map((s) => ctx.measureText(s.name).width)) + swatch + gap * 3,
    );
    if (pos === "right") {
      reserved.reservedRight = width;
    } else {
      reserved.reservedLeft = width;
    }
    const x = pos === "right" ? widthPx - width + gap : gap;
    let y = heightPx / 2 - (series.length * 16) / 2;
    for (const s of series) {
      ctx.fillStyle = s.color;
      ctx.fillRect(x, y, swatch, swatch);
      ctx.fillStyle = TEXT_COLOR;
      ctx.fillText(s.name, x + swatch + gap, y + swatch);
      y += 16;
    }
  } else {
    reserved.reservedBottom = pos === "bottom" ? 18 : 0;
    // The top legend paints below the title (if any) — computeLayout
    // reserves titleH + 18px together, so the painted position must stack
    // the same way or the two overlap (see the layout/paint split note on
    // paintChart below).
    reserved.reservedTop = pos === "top" ? 18 : 0;
    const y = pos === "top" ? titleH + 14 : heightPx - 8;
    let x = widthPx / 2 - (series.reduce((sum, s) => sum + ctx.measureText(s.name).width + swatch + gap * 3, 0)) / 2;
    for (const s of series) {
      ctx.fillStyle = s.color;
      ctx.fillRect(x, y - swatch, swatch, swatch);
      ctx.fillStyle = TEXT_COLOR;
      ctx.fillText(s.name, x + swatch + gap, y);
      x += ctx.measureText(s.name).width + swatch + gap * 3;
    }
  }
  return reserved;
}

function computeLayout(
  ctx: CanvasRenderingContext2D,
  chart: Chart,
  series: ChartSeries[],
  widthPx: number,
  heightPx: number,
): Layout {
  let top = 6;
  let bottom = 6;
  let left = 6;
  let right = 6;
  if (chart.options.title) {
    top += 20;
  }
  const legendReserve = paintLegendReserveOnly(ctx, chart, series, widthPx);
  top += legendReserve.reservedTop;
  bottom += legendReserve.reservedBottom;
  left += legendReserve.reservedLeft;
  right += legendReserve.reservedRight;

  const hasAxes = chart.chartType !== "pie" && chart.chartType !== "doughnut";
  if (hasAxes) {
    left += 34; // value axis labels
    bottom += 20; // category axis labels
    if (chart.options.valueAxisTitle) {
      left += 14;
    }
    if (chart.options.categoryAxisTitle) {
      bottom += 14;
    }
  }
  return {
    plotX: left,
    plotY: top,
    plotW: Math.max(1, widthPx - left - right),
    plotH: Math.max(1, heightPx - top - bottom),
  };
}

// Measures the legend's reserved space without painting — computeLayout
// needs the size before the plot area is known, paintChart paints it after.
function paintLegendReserveOnly(
  ctx: CanvasRenderingContext2D,
  chart: Chart,
  series: ChartSeries[],
  widthPx: number,
): { reservedLeft: number; reservedRight: number; reservedTop: number; reservedBottom: number } {
  const pos = chart.options.legend;
  if (pos === "none" || series.length === 0) {
    return { reservedLeft: 0, reservedRight: 0, reservedTop: 0, reservedBottom: 0 };
  }
  ctx.font = LABEL_FONT;
  if (pos === "right" || pos === "left") {
    const width = Math.min(
      widthPx * 0.35,
      Math.max(...series.map((s) => ctx.measureText(s.name).width)) + 20,
    );
    return pos === "right"
      ? { reservedLeft: 0, reservedRight: width, reservedTop: 0, reservedBottom: 0 }
      : { reservedLeft: width, reservedRight: 0, reservedTop: 0, reservedBottom: 0 };
  }
  return pos === "top"
    ? { reservedLeft: 0, reservedRight: 0, reservedTop: 18, reservedBottom: 0 }
    : { reservedLeft: 0, reservedRight: 0, reservedTop: 0, reservedBottom: 18 };
}

function paintAxes(
  ctx: CanvasRenderingContext2D,
  chart: Chart,
  layout: Layout,
  min: number,
  max: number,
): void {
  const { plotX, plotY, plotW, plotH } = layout;
  ctx.strokeStyle = AXIS_COLOR;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(plotX, plotY);
  ctx.lineTo(plotX, plotY + plotH);
  ctx.lineTo(plotX + plotW, plotY + plotH);
  ctx.stroke();

  const zeroY = plotY + plotH - ((0 - min) / (max - min)) * plotH;
  if (min < 0 && max > 0) {
    ctx.beginPath();
    ctx.moveTo(plotX, zeroY);
    ctx.lineTo(plotX + plotW, zeroY);
    ctx.stroke();
  }

  if (chart.options.gridlines) {
    ctx.strokeStyle = GRIDLINE_COLOR;
    const steps = 4;
    for (let i = 1; i <= steps; i++) {
      const y = plotY + (plotH / steps) * i;
      ctx.beginPath();
      ctx.moveTo(plotX, y);
      ctx.lineTo(plotX + plotW, y);
      ctx.stroke();
    }
  }

  ctx.fillStyle = TEXT_COLOR;
  ctx.font = LABEL_FONT;
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  const steps = 4;
  for (let i = 0; i <= steps; i++) {
    const v = min + ((max - min) * (steps - i)) / steps;
    const y = plotY + (plotH / steps) * i;
    ctx.fillText(formatAxisNumber(v), plotX - 4, y);
  }
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  chart.categories.forEach((cat, i) => {
    const x = plotX + ((i + 0.5) / Math.max(1, chart.categories.length)) * plotW;
    ctx.fillText(truncateLabel(cat, 10), x, plotY + plotH + 4);
  });

  if (chart.options.valueAxisTitle) {
    ctx.save();
    ctx.translate(10, plotY + plotH / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.font = AXIS_TITLE_FONT;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(chart.options.valueAxisTitle, 0, 0);
    ctx.restore();
  }
  if (chart.options.categoryAxisTitle) {
    ctx.font = AXIS_TITLE_FONT;
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText(chart.options.categoryAxisTitle, plotX + plotW / 2, plotY + plotH + 34);
  }
}

function formatAxisNumber(v: number): string {
  const rounded = Math.round(v * 100) / 100;
  return String(rounded);
}
function truncateLabel(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function paintDataLabel(ctx: CanvasRenderingContext2D, text: string, x: number, y: number): void {
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = LABEL_FONT;
  ctx.textAlign = "center";
  ctx.textBaseline = "bottom";
  ctx.fillText(text, x, y - 2);
}

function paintBarLike(ctx: CanvasRenderingContext2D, chart: Chart, series: ChartSeries[], layout: Layout, min: number, max: number): void {
  const { plotX, plotY, plotW, plotH } = layout;
  const n = chart.categories.length;
  if (n === 0) {
    return;
  }
  const categoryW = plotW / n;
  const zeroY = plotY + plotH - ((0 - min) / (max - min)) * plotH;
  const horizontal = chart.chartType === "barClustered";
  const stacked = chart.chartType === "columnStacked";

  for (let ci = 0; ci < n; ci++) {
    if (stacked) {
      let posY = zeroY;
      let negY = zeroY;
      const barW = categoryW * 0.6;
      const x = plotX + ci * categoryW + (categoryW - barW) / 2;
      for (const s of series) {
        const v = s.values[ci];
        if (v === null) {
          continue;
        }
        const h = (Math.abs(v) / (max - min)) * plotH;
        ctx.fillStyle = s.color;
        if (v >= 0) {
          ctx.fillRect(x, posY - h, barW, h);
          posY -= h;
        } else {
          ctx.fillRect(x, negY, barW, h);
          negY += h;
        }
        if (chart.options.dataLabels) {
          paintDataLabel(ctx, formatAxisNumber(v), x + barW / 2, v >= 0 ? posY + h : negY);
        }
      }
    } else if (horizontal) {
      const barH = (plotH / n) * 0.6;
      const y = plotY + ci * (plotH / n) + (plotH / n - barH) / 2;
      const groupW = plotW / n;
      series.forEach((s, si) => {
        const v = s.values[ci];
        if (v === null) {
          return;
        }
        const w = (Math.abs(v) / (max - min)) * plotW;
        const barSlice = barH / series.length;
        const barY = y + si * barSlice;
        ctx.fillStyle = s.color;
        const startX = v >= 0 ? plotX : plotX - w;
        ctx.fillRect(startX, barY, w, barSlice);
        if (chart.options.dataLabels) {
          ctx.save();
          ctx.textAlign = "left";
          ctx.textBaseline = "middle";
          ctx.fillStyle = TEXT_COLOR;
          ctx.font = LABEL_FONT;
          ctx.fillText(formatAxisNumber(v), startX + w + 2, barY + barSlice / 2);
          ctx.restore();
        }
      });
      void groupW;
    } else {
      const barW = (categoryW * 0.6) / series.length;
      const groupStartX = plotX + ci * categoryW + (categoryW - barW * series.length) / 2;
      series.forEach((s, si) => {
        const v = s.values[ci];
        if (v === null) {
          return;
        }
        const h = (Math.abs(v) / (max - min)) * plotH;
        const x = groupStartX + si * barW;
        const y = v >= 0 ? zeroY - h : zeroY;
        ctx.fillStyle = s.color;
        ctx.fillRect(x, y, barW, h);
        if (chart.options.dataLabels) {
          paintDataLabel(ctx, formatAxisNumber(v), x + barW / 2, y);
        }
      });
    }
  }
}

function pointsFor(chart: Chart, s: ChartSeries, layout: Layout, min: number, max: number): (null | { x: number; y: number })[] {
  const { plotX, plotY, plotW, plotH } = layout;
  const n = chart.categories.length;
  return s.values.map((v, i) => {
    if (v === null) {
      return null;
    }
    const x = n <= 1 ? plotX + plotW / 2 : plotX + (i / (n - 1)) * plotW;
    const y = plotY + plotH - ((v - min) / (max - min)) * plotH;
    return { x, y };
  });
}

function paintLineOrArea(ctx: CanvasRenderingContext2D, chart: Chart, series: ChartSeries[], layout: Layout, min: number, max: number): void {
  const isArea = chart.chartType === "area";
  const showMarkers = chart.chartType === "lineMarkers";
  for (const s of series) {
    const points = pointsFor(chart, s, layout, min, max);
    if (isArea) {
      const zeroY = layout.plotY + layout.plotH - ((0 - min) / (max - min)) * layout.plotH;
      ctx.beginPath();
      let started = false;
      let lastX = layout.plotX;
      for (const p of points) {
        if (!p) {
          continue;
        }
        if (!started) {
          ctx.moveTo(p.x, zeroY);
          started = true;
        }
        ctx.lineTo(p.x, p.y);
        lastX = p.x;
      }
      if (started) {
        ctx.lineTo(lastX, zeroY);
        ctx.closePath();
        ctx.fillStyle = `${s.color}80`;
        ctx.fill();
      }
    }
    ctx.strokeStyle = s.color;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    let drawing = false;
    for (const p of points) {
      if (!p) {
        drawing = false;
        continue;
      }
      if (!drawing) {
        ctx.moveTo(p.x, p.y);
        drawing = true;
      } else {
        ctx.lineTo(p.x, p.y);
      }
    }
    ctx.stroke();
    if (showMarkers) {
      ctx.fillStyle = s.color;
      for (const p of points) {
        if (!p) {
          continue;
        }
        ctx.beginPath();
        ctx.arc(p.x, p.y, 2.5, 0, Math.PI * 2);
        ctx.fill();
      }
    }
    if (chart.options.dataLabels) {
      for (let i = 0; i < points.length; i++) {
        const p = points[i];
        const v = s.values[i];
        if (p && v !== null) {
          paintDataLabel(ctx, formatAxisNumber(v), p.x, p.y);
        }
      }
    }
  }
}

function paintScatter(ctx: CanvasRenderingContext2D, chart: Chart, series: ChartSeries[], layout: Layout, min: number, max: number): void {
  const { plotX, plotY, plotW, plotH } = layout;
  const xs = chart.categories.map((c) => Number(c)).map((n) => (Number.isFinite(n) ? n : NaN));
  const validXs = xs.filter((n) => Number.isFinite(n));
  const xMin = validXs.length > 0 ? Math.min(...validXs) : 0;
  const xMax = validXs.length > 0 ? Math.max(...validXs) : 1;
  for (const s of series) {
    ctx.fillStyle = s.color;
    s.values.forEach((v, i) => {
      const xv = Number.isFinite(xs[i]) ? xs[i] : i;
      if (v === null) {
        return;
      }
      const x = plotX + ((xv - xMin) / Math.max(1e-9, xMax - xMin)) * plotW;
      const y = plotY + plotH - ((v - min) / (max - min)) * plotH;
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();
      if (chart.options.dataLabels) {
        paintDataLabel(ctx, formatAxisNumber(v), x, y);
      }
    });
  }
}

function paintPieLike(ctx: CanvasRenderingContext2D, chart: Chart, series: ChartSeries[], layout: Layout): void {
  const s = series[0];
  if (!s) {
    return;
  }
  const { plotX, plotY, plotW, plotH } = layout;
  const cx = plotX + plotW / 2;
  const cy = plotY + plotH / 2;
  const r = Math.max(1, Math.min(plotW, plotH) / 2 - 4);
  const total = s.values.reduce((sum: number, v) => sum + (v !== null ? Math.abs(v) : 0), 0) || 1;
  let angle = -Math.PI / 2;
  const colors = chart.categories.map((_, i) => DEFAULT_CHART_COLORS[i % DEFAULT_CHART_COLORS.length]);
  s.values.forEach((v, i) => {
    if (v === null || v === 0) {
      return;
    }
    const slice = (Math.abs(v) / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = colors[i];
    ctx.fill();
    if (chart.options.dataLabels) {
      const mid = angle + slice / 2;
      const lx = cx + Math.cos(mid) * r * 0.65;
      const ly = cy + Math.sin(mid) * r * 0.65;
      ctx.fillStyle = "#ffffff";
      ctx.font = LABEL_FONT;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(formatAxisNumber(v), lx, ly);
    }
    angle += slice;
  });
  if (chart.chartType === "doughnut") {
    ctx.fillStyle = "#ffffff";
    ctx.beginPath();
    ctx.arc(cx, cy, r * 0.5, 0, Math.PI * 2);
    ctx.fill();
  }
}

/** Paints a Chart's title/legend/axes/series/data-labels into a `widthPx` x
 * `heightPx` canvas context — the context is expected to already be
 * translated/scaled to the chart's own top-left corner (0,0) in device
 * pixels (mirrors shapeCanvas.ts's paintShape). An unsupported chart type
 * (`unsupportedChartXml` set) paints as a dashed placeholder rectangle at
 * the chart's true bounds instead, exactly like an unsupported shape
 * geometry. */
export function paintChart(ctx: CanvasRenderingContext2D, chart: Chart, widthPx: number, heightPx: number): void {
  ctx.save();
  ctx.clearRect(0, 0, widthPx, heightPx);

  if (chart.unsupportedChartXml !== undefined) {
    ctx.strokeStyle = "#888888";
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 4]);
    ctx.strokeRect(0.75, 0.75, Math.max(0, widthPx - 1.5), Math.max(0, heightPx - 1.5));
    ctx.setLineDash([]);
    ctx.fillStyle = "#666666";
    ctx.font = LABEL_FONT;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("Chart not supported", widthPx / 2, heightPx / 2);
    ctx.restore();
    return;
  }

  const series = seriesToPlot(chart);
  // Painted top-to-bottom in the same order computeLayout reserves space
  // for (title, then a top legend if any) so the two never overlap — see
  // paintLegend's own "top" branch, which takes this same titleH.
  let titleH = 0;
  if (chart.options.title) {
    ctx.fillStyle = TEXT_COLOR;
    ctx.font = TITLE_FONT;
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(chart.options.title, widthPx / 2, 4);
    titleH = 18;
  }

  const layout = computeLayout(ctx, chart, series, widthPx, heightPx);
  const { min, max } = valueExtent(chart, series);
  const hasAxes = chart.chartType !== "pie" && chart.chartType !== "doughnut";
  if (hasAxes) {
    paintAxes(ctx, chart, layout, min, max);
    if (chart.chartType === "columnClustered" || chart.chartType === "columnStacked" || chart.chartType === "barClustered") {
      paintBarLike(ctx, chart, series, layout, min, max);
    } else if (chart.chartType === "line" || chart.chartType === "lineMarkers" || chart.chartType === "area") {
      paintLineOrArea(ctx, chart, series, layout, min, max);
    } else if (chart.chartType === "scatter") {
      paintScatter(ctx, chart, series, layout, min, max);
    }
  } else {
    paintPieLike(ctx, chart, series, layout);
  }

  paintLegend(ctx, chart, series, widthPx, heightPx, titleH);
  ctx.restore();
}
