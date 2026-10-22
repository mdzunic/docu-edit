import { afterEach, describe, expect, it, vi } from "vitest";
import type { PdfPageInput } from "./pdfBuild";
import { paintPdfPage, rasterisePdfPage, type RasterContext } from "./pdfRaster";

/** A 2D context stand-in that records every call with the state it ran in. */
function recordingContext() {
  const calls: string[] = [];
  const state = {
    fillStyle: "" as RasterContext["fillStyle"],
    strokeStyle: "" as RasterContext["strokeStyle"],
    lineWidth: 1,
    font: "",
    lineCap: "butt" as CanvasLineCap,
    lineJoin: "miter" as CanvasLineJoin,
  };
  const ctx = {
    ...state,
    save: () => calls.push("save"),
    restore: () => calls.push("restore"),
    fillRect: (x: number, y: number, w: number, h: number) => calls.push(`fillRect ${ctx.fillStyle} ${x} ${y} ${w} ${h}`),
    strokeRect: (x: number, y: number, w: number, h: number) =>
      calls.push(`strokeRect ${ctx.strokeStyle} ${ctx.lineWidth} ${x} ${y} ${w} ${h}`),
    fillText: (text: string, x: number, y: number) => calls.push(`fillText ${ctx.font} ${ctx.fillStyle} ${text} ${x} ${y}`),
    drawImage: (_image: unknown, x: number, y: number, w: number, h: number) => calls.push(`drawImage ${x} ${y} ${w} ${h}`),
    beginPath: () => calls.push("beginPath"),
    moveTo: (x: number, y: number) => calls.push(`moveTo ${x} ${y}`),
    lineTo: (x: number, y: number) => calls.push(`lineTo ${x} ${y}`),
    bezierCurveTo: () => calls.push("bezierCurveTo"),
    closePath: () => calls.push("closePath"),
    fill: () => calls.push(`fill ${ctx.fillStyle}`),
    stroke: () => calls.push(`stroke ${ctx.strokeStyle} ${ctx.lineWidth}`),
    setLineDash: () => undefined,
  };
  return { ctx: ctx as unknown as RasterContext, calls };
}

function page(overrides: Partial<PdfPageInput> = {}): PdfPageInput {
  return { widthPt: 100, heightPt: 50, items: [], links: [], headings: [], images: [], ...overrides };
}

describe("paintPdfPage (SPEC-104)", () => {
  it("paints white paper, then rects, images, text and on-top layers in PDF order", () => {
    const { ctx, calls } = recordingContext();
    const image = {} as CanvasImageSource;
    paintPdfPage(
      ctx,
      page({
        rects: [
          { xPt: 0, yPt: 0, widthPt: 100, heightPt: 50, fill: [0, 0, 1] },
          { xPt: 10, yPt: 10, widthPt: 20, heightPt: 10, stroke: [0, 0, 0] },
          { xPt: 1, yPt: 1, widthPt: 2, heightPt: 2, fill: "#ff0000", onTop: true },
        ],
        images: [{ src: "a", xPt: 5, yPt: 5, widthPt: 10, heightPt: 10 }, { src: "missing", xPt: 0, yPt: 0, widthPt: 1, heightPt: 1 }],
        items: [{ text: "Hi", xPt: 10, yPt: 20, sizePt: 12, face: "calibri-bold", color: [1, 0, 0] }],
      }),
      2,
      new Map([["a", image]]),
    );
    expect(calls).toEqual([
      "save",
      "fillRect #ffffff 0 0 200 100",
      "fillRect rgb(0, 0, 255) 0 0 200 100",
      "strokeRect rgb(0, 0, 0) 1.5 20 20 40 20",
      "drawImage 10 10 20 20",
      'fillText bold 24px "Carlito", sans-serif rgb(255, 0, 0) Hi 20 40',
      "fillRect rgb(255, 0, 0) 2 2 4 4",
      "restore",
    ]);
  });

  it("paints vector paths with their fill and stroke", () => {
    const { ctx, calls } = recordingContext();
    paintPdfPage(
      ctx,
      page({
        paths: [
          {
            subpaths: [{ start: { xPt: 0, yPt: 0 }, segments: [{ toXPt: 10, toYPt: 0 }], closed: true }],
            fill: [0, 1, 0],
            stroke: { color: [0, 0, 0], widthPt: 1 },
          },
        ],
      }),
      1,
      new Map(),
    );
    expect(calls).toContain("moveTo 0 0");
    expect(calls).toContain("lineTo 10 0");
    expect(calls).toContain("fill rgb(0, 255, 0)");
    expect(calls).toContain("stroke rgb(0, 0, 0) 1");
  });
});

describe("rasterisePdfPage (SPEC-104)", () => {
  afterEach(() => vi.restoreAllMocks());

  it("sizes the canvas to the absolute width it is given and releases it after (AC-37)", async () => {
    const canvases: HTMLCanvasElement[] = [];
    const sizes: string[] = [];
    const create = document.createElement.bind(document);
    vi.spyOn(document, "createElement").mockImplementation((tag: string) => {
      const el = create(tag);
      if (tag === "canvas") {
        const canvas = el as HTMLCanvasElement;
        canvases.push(canvas);
        canvas.getContext = (() => {
          sizes.push(`${canvas.width}x${canvas.height}`);
          return recordingContext().ctx;
        }) as unknown as HTMLCanvasElement["getContext"];
        // jsdom's Blob has no arrayBuffer(); a PNG-signature stand-in does.
        const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
        canvas.toBlob = (callback: BlobCallback) => callback({ arrayBuffer: async () => png.buffer } as Blob);
      }
      return el;
    });
    const bytes = await rasterisePdfPage(page({ widthPt: 960, heightPt: 540 }), 2560);
    expect(sizes).toEqual(["2560x1440"]);
    expect(Array.from(bytes)).toEqual([0x89, 0x50, 0x4e, 0x47]);
    expect(canvases[0].width).toBe(0);
  });
});
