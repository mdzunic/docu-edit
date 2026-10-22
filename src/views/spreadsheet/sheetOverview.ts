// The overview canvas's painter (SPEC-137 "The grid's ceiling"): when the
// viewport alone would pass the grid's ceiling, SheetOverviewCanvas paints the
// sheet instead of DOM cells — from the same band items the paper binds
// (useSheetCells' overviewBands: the viewport's stored cells, merges,
// resolved borders and gridlines, every one through the cells' own style
// resolution, cellDisplay and formatValue), placed by GridMetrics.
//
// - The canvas covers the grid's content box (headers included) inside the
//   zoom's scale transform; its backing store is the box × zoom × the device
//   pixel ratio, so one canvas pixel is one device pixel at any zoom.
// - The paper keeps SheetCellLayer's literal colours (PAPER) in every app
//   theme, and the dark theme's dim-page treatment reaches it as the DOM
//   paper's computed `brightness()` laid over it as black; the headers follow
//   the app theme (tokens read from the canvas's computed style).
// - Text whose rendered size (px × zoom) is under GRID_GREEK_BELOW_PX is
//   greeked: a 1 px bar in its colour, as wide as the cells' own width
//   estimate (estimateTextWidth), aligned as the text would be.
// - Repaints at most once per animation frame: one reactive effect tracks
//   everything the frame reads (scroll, zoom, data, styles) and schedules the
//   next frame instead of painting at once; the theme's attributes on the
//   root schedule one too. jsdom has no 2D context: getContext's null paints
//   nothing.
import { effect, stop } from "vue";
import { fontStack } from "../../lib/core/fontSubstitution";
import { estimateTextWidth, maxDigitWidthPx } from "../../lib/spreadsheet/geometry";
import type { BandItem, CellItem } from "./sheetCellItems";
import { drawsText, type AxisWindow } from "./sheetRenderPlan";
import { COLUMN_HEADER_HEIGHT_PX as CH, ROW_HEADER_WIDTH_PX as RH } from "./useSheetViewport";

export const GRID_GREEK_BELOW_PX = 6;

/** The paper's literal colours (SheetCellLayer's `.sheet-paper`). */
export const PAPER = { background: "#ffffff", gridline: "#d4d4d4", text: "#000000" } as const;

const DEFAULT_SIZE_PT = 11;
const DEFAULT_FAMILY = "Calibri, Carlito, sans-serif";
const HEADER_FONT_PX = 11;

/** What a frame paints: the bands and the content box's geometry (unscaled px). */
export interface OverviewFrame {
  bands: BandItem[];
  zoom: number;
  width: number;
  height: number;
  frozenW: number;
  viewCols: number;
  viewRows: number;
  gridLines: boolean;
  rows: AxisWindow;
  cols: AxisWindow;
}

interface Chrome {
  surface: string;
  text: string;
  line: string;
  font: string;
}

/** A band's region: its bands-relative x and width, its cells and lines. */
function regions(frame: OverviewFrame): { frozen: boolean; x: number; width: number }[] {
  const out = frame.frozenW > 0 ? [{ frozen: true, x: 0, width: frame.frozenW }] : [];
  return [...out, { frozen: false, x: frame.frozenW, width: Math.max(0, frame.viewCols - frame.frozenW) }];
}

function textOf(cell: CellItem): string {
  return cell.runs ? cell.runs.map((run) => run.text).join("") : cell.fill ? cell.fill.before + cell.fill.after : cell.text;
}

/** Paints one cell's fill, data bar and text at paper (x, y). */
function paintCell(ctx: CanvasRenderingContext2D, cell: CellItem, x: number, y: number, k: number, zoom: number, dpr: number): void {
  const d = (v: number): number => Math.round(v * k);
  const fill = cell.boxStyle["background-color"];
  if (fill) {
    ctx.fillStyle = fill;
    ctx.fillRect(d(x), d(y), d(x + cell.width) - d(x), d(y + cell.height) - d(y));
  } else if (cell.merged) {
    ctx.fillStyle = PAPER.background;
    ctx.fillRect(d(x), d(y), d(x + cell.width - 1) - d(x), d(y + cell.height - 1) - d(y));
  }
  const bar = cell.cf?.dataBar;
  if (bar && !cell.paintOnly) {
    ctx.fillStyle = bar.color;
    const inner = cell.width - 4;
    ctx.fillRect(d(x + 2 + (inner * bar.startPct) / 100), d(y + 2), Math.max(1, d((inner * bar.widthPct) / 100)), d(cell.height - 4));
  }
  if (!drawsText(cell)) {
    return;
  }
  const text = textOf(cell);
  const style = cell.textStyle;
  const sizePt = Number.parseFloat(style["font-size"] ?? "") || DEFAULT_SIZE_PT;
  const sizePx = (sizePt * 4) / 3;
  const left = x + (Number.parseFloat(style.left ?? "") || 0);
  const width = Number.parseFloat(style.width ?? "") || cell.width;
  const align = cell.boxStyle["text-align"] ?? "left";
  const valign = cell.boxStyle["justify-content"] ?? "flex-end";
  ctx.fillStyle = style.color ?? PAPER.text;
  if (sizePx * zoom < GRID_GREEK_BELOW_PX) {
    const family = style["font-family"] ?? "Calibri";
    const barWidth = Math.min(width - 4, estimateTextWidth(text, maxDigitWidthPx(family, sizePt)));
    if (barWidth <= 0 || text.trim() === "") {
      return;
    }
    const barX = align === "right" ? left + width - 2 - barWidth : align === "center" ? left + (width - barWidth) / 2 : left + 2;
    const barY = valign === "flex-start" ? y + 2 + sizePx / 2 : valign === "center" ? y + cell.height / 2 : y + cell.height - 2 - sizePx * 0.4;
    ctx.fillRect(d(barX), d(barY), Math.max(1, d(barWidth)), Math.max(1, Math.round(dpr)));
    return;
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(d(left), d(y), d(width), d(cell.height));
  ctx.clip();
  const family = style["font-family"] ? fontStack(style["font-family"]) : DEFAULT_FAMILY;
  ctx.font = `${style["font-style"] === "italic" ? "italic " : ""}${style["font-weight"] === "bold" ? "bold " : ""}${sizePx * k}px ${family}`;
  ctx.textAlign = align === "right" ? "right" : align === "center" ? "center" : "left";
  ctx.textBaseline = valign === "flex-start" ? "top" : valign === "center" ? "middle" : "bottom";
  const tx = align === "right" ? left + width - 2 : align === "center" ? left + width / 2 : left + 2;
  const ty = valign === "flex-start" ? y + 1 : valign === "center" ? y + cell.height / 2 : y + cell.height - 1;
  ctx.fillText(text, d(tx), d(ty));
  ctx.restore();
}

/** Paints the headers: the strips, each item's boundary and (when legible)
 * its label, in the app theme's colours. */
function paintHeaders(ctx: CanvasRenderingContext2D, frame: OverviewFrame, k: number, chrome: Chrome): void {
  const d = (v: number): number => Math.round(v * k);
  ctx.fillStyle = chrome.surface;
  ctx.fillRect(0, 0, d(frame.width), d(CH));
  ctx.fillRect(0, 0, d(RH), d(frame.height));
  const legible = HEADER_FONT_PX * frame.zoom >= GRID_GREEK_BELOW_PX;
  ctx.font = `${HEADER_FONT_PX * k}px ${chrome.font}`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  const axis = (window: AxisWindow, frozenSize: number, horizontal: boolean): void => {
    for (const [items, base] of [[window.frozen, 0], [window.scrolled, frozenSize]] as const) {
      for (const item of items) {
        if (item.size <= 0) {
          continue;
        }
        const at = (horizontal ? RH : CH) + base + item.start;
        ctx.fillStyle = chrome.line;
        if (horizontal) {
          ctx.fillRect(d(at + item.size) - 1, 0, 1, d(CH));
        } else {
          ctx.fillRect(0, d(at + item.size) - 1, d(RH), 1);
        }
        if (legible) {
          ctx.fillStyle = chrome.text;
          const label = horizontal ? colLabel(item.index) : String(item.index + 1);
          ctx.fillText(label, d(horizontal ? at + item.size / 2 : RH / 2), d(horizontal ? CH / 2 : at + item.size / 2));
        }
      }
    }
  };
  ctx.save();
  ctx.beginPath();
  ctx.rect(d(RH), 0, d(frame.width - RH), d(CH));
  ctx.clip();
  axis(frame.cols, frame.frozenW, true);
  ctx.restore();
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, d(CH), d(RH), d(frame.height - CH));
  ctx.clip();
  const frozenH = frame.bands.find((band) => band.kind === "main")?.top ?? 0;
  axis(frame.rows, frozenH, false);
  ctx.restore();
}

function colLabel(index: number): string {
  let letters = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

/** Paints a frame into `ctx` (its canvas sized width × zoom × dpr). */
export function paintOverview(ctx: CanvasRenderingContext2D, frame: OverviewFrame, view: { dpr: number; dim: number; chrome: Chrome }): void {
  const k = frame.zoom * view.dpr;
  const d = (v: number): number => Math.round(v * k);
  const hair = Math.max(1, Math.round(k));
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, d(frame.width), d(frame.height));
  ctx.fillStyle = PAPER.background;
  ctx.fillRect(d(RH), d(CH), d(frame.viewCols), d(frame.viewRows));
  for (const band of frame.bands) {
    for (const region of regions(frame)) {
      const x0 = RH + region.x;
      const y0 = CH + band.top;
      ctx.save();
      ctx.beginPath();
      ctx.rect(d(x0), d(y0), d(region.width), d(band.height));
      ctx.clip();
      if (frame.gridLines) {
        ctx.fillStyle = PAPER.gridline;
        for (const line of band.gridRows) {
          ctx.fillRect(d(x0), d(y0 + line.at), d(region.width), hair);
        }
        for (const line of region.frozen ? band.gridColsFrozen : band.gridColsScrolled) {
          ctx.fillRect(d(x0 + line.at), d(y0), hair, d(band.height));
        }
      }
      for (const row of band.rows) {
        for (const cell of region.frozen ? row.frozen : row.scrolled) {
          paintCell(ctx, cell, x0 + cell.left, y0 + row.top + cell.top, k, frame.zoom, view.dpr);
        }
      }
      for (const border of region.frozen ? band.bordersFrozen : band.bordersScrolled) {
        const horizontal = border.style["border-top"] !== undefined;
        const [width = "1px", , color = PAPER.text] = (border.style["border-top"] ?? border.style["border-left"] ?? "").split(" ");
        const thick = Math.max(1, Math.round(Number.parseFloat(width) * k));
        const bx = x0 + (Number.parseFloat(border.style.left ?? "") || 0);
        const by = y0 + (Number.parseFloat(border.style.top ?? "") || 0);
        const length = Number.parseFloat((horizontal ? border.style.width : border.style.height) ?? "") || 0;
        ctx.globalAlpha = border.style.opacity ? Number(border.style.opacity) : 1;
        ctx.fillStyle = color;
        ctx.fillRect(d(bx), d(by), horizontal ? d(length) : thick, horizontal ? thick : d(length));
        ctx.globalAlpha = 1;
      }
      ctx.restore();
    }
  }
  if (view.dim < 1) {
    ctx.fillStyle = `rgba(0, 0, 0, ${1 - view.dim})`;
    ctx.fillRect(d(RH), d(CH), d(frame.viewCols), d(frame.viewRows));
  }
  paintHeaders(ctx, frame, k, view.chrome);
}

/** The factor of a computed `filter: brightness(x)`, 1 for anything else. */
function brightnessOf(filter: string): number {
  const match = /brightness\(([\d.]+)\)/.exec(filter);
  return match ? Math.min(1, Number(match[1])) : 1;
}

/** The next animation frame (a 16 ms timer where there is none). */
function later(run: () => void): number {
  return typeof requestAnimationFrame === "function" ? requestAnimationFrame(run) : Number(setTimeout(run, 16));
}

/** Narrows SheetOverviewCanvas's runtime-typed `frame` prop. */
export function asOverviewFrame(value: unknown): () => OverviewFrame {
  return value as () => OverviewFrame;
}

/** Paints `frame()` into `canvas` now and again at most once per animation
 * frame whenever what it read changes; returns the disposer. */
export function mountOverviewPainter(canvas: HTMLCanvasElement, frame: () => OverviewFrame): () => void {
  let pending: number | null = null;
  const draw = (): void => {
    const input = frame();
    const dpr = typeof window !== "undefined" && window.devicePixelRatio > 0 ? window.devicePixelRatio : 1;
    const width = Math.max(1, Math.round(input.width * input.zoom * dpr));
    const height = Math.max(1, Math.round(input.height * input.zoom * dpr));
    if (canvas.width !== width || canvas.height !== height) {
      canvas.width = width;
      canvas.height = height;
    }
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      return;
    }
    const css = getComputedStyle(canvas);
    const token = (name: string, fallback: string): string => css.getPropertyValue(name).trim() || fallback;
    const paper = canvas.parentElement?.querySelector(".sheet-paper");
    paintOverview(ctx, input, {
      dpr,
      dim: paper ? brightnessOf(getComputedStyle(paper).filter) : 1,
      chrome: { surface: token("--surface-sunken", "#f3f3f3"), text: token("--text-secondary", "#555555"), line: token("--border-subtle", "#dddddd"), font: token("--font-ui", "sans-serif") },
    });
  };
  const runner = effect(draw, {
    scheduler: () => {
      pending ??= later(() => {
        pending = null;
        runner();
      });
    },
  });
  const observer = typeof MutationObserver === "function" ? new MutationObserver(() => runner.effect.scheduler?.()) : null;
  observer?.observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme", "data-dim-page"] });
  return () => {
    stop(runner);
    observer?.disconnect();
    if (pending !== null) {
      if (typeof cancelAnimationFrame === "function") {
        cancelAnimationFrame(pending);
      }
      clearTimeout(pending);
    }
  };
}
