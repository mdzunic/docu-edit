// The view's text measurer for row autofit (SPEC-122 "Row autofit"):
// lib/spreadsheet/autofit.ts decides how tall a row's content is through a
// CellTextMeasurer, and this is the browser one — CanvasRenderingContext2D
// .measureText for widths, the line height from fontBoundingBoxAscent +
// fontBoundingBoxDescent (1.2 × the font size where the browser has no font
// bounding box, and everywhere when there is no canvas, as in jsdom). The
// fonts autofit hands over are already resolved family names (autofit
// resolves the workbook's theme fonts), so nothing here reads a theme.
import type { CellTextMeasurer, MeasuredFont } from "../../lib/spreadsheet/autofit";

const PX_PER_PT = 96 / 72;
/** The line height of a font with no measurable bounding box, × its size. */
export const FALLBACK_LINE_HEIGHT = 1.2;
/** The average character width of an unmeasurable font, × its size. */
const FALLBACK_CHAR_WIDTH = 0.55;

/** The canvas `font` shorthand of a measured font. */
export function canvasFont(font: MeasuredFont): string {
  const family = /^[\w-]+$/.test(font.family) ? font.family : `"${font.family.replaceAll('"', "")}"`;
  return `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${font.sizePt}pt ${family}, sans-serif`;
}

/** The browser's measurer (see the module header). */
export function createCanvasCellMeasurer(): CellTextMeasurer {
  let context: CanvasRenderingContext2D | null | undefined;
  const lineHeights = new Map<string, number>();

  function canvas(font: MeasuredFont): CanvasRenderingContext2D | null {
    if (context === undefined) {
      context = typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
    }
    if (context) {
      context.font = canvasFont(font);
    }
    return context;
  }

  return {
    widthPx(text: string, font: MeasuredFont): number {
      const ctx = canvas(font);
      const measured = ctx?.measureText(text).width;
      return measured !== undefined && Number.isFinite(measured) ? measured : text.length * font.sizePt * PX_PER_PT * FALLBACK_CHAR_WIDTH;
    },
    lineHeightPx(font: MeasuredFont): number {
      const key = canvasFont(font);
      const known = lineHeights.get(key);
      if (known !== undefined) {
        return known;
      }
      const metrics = canvas(font)?.measureText("Hg");
      const box = (metrics?.fontBoundingBoxAscent ?? Number.NaN) + (metrics?.fontBoundingBoxDescent ?? Number.NaN);
      const height = Number.isFinite(box) && box > 0 ? box : FALLBACK_LINE_HEIGHT * font.sizePt * PX_PER_PT;
      lineHeights.set(key, height);
      return height;
    },
  };
}
