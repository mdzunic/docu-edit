import { describe, it, expect, beforeEach } from "vitest";
import {
  createDocument,
  sectionContentWidthCm,
  type Chart,
  type Image,
  type Paragraph,
  type Shape,
  type Table,
} from "./documentModel";
import {
  buildThumbnail,
  clearThumbnailCache,
  getCachedThumbnail,
  renderThumbnail,
  THUMBNAIL_MAX_LINES,
} from "./libraryThumbnail";

function paragraph(text: string, headingLevel?: number): Paragraph {
  return {
    type: "paragraph",
    runs: [{ text, formatting: {} }],
    ...(headingLevel === undefined ? {} : { headingLevel }),
  };
}

function image(mediaId: string, widthCm: number, heightCm: number): Image {
  return {
    type: "image",
    mediaId,
    widthCm,
    heightCm,
    wrap: "inline",
    naturalWidthPx: 100,
    naturalHeightPx: 100,
  };
}

beforeEach(() => {
  clearThumbnailCache();
});

describe("buildThumbnail (SPEC-067)", () => {
  it("extracts the leading paragraphs with their heading levels", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [
      paragraph("The Title", 1),
      paragraph("Body one"),
      paragraph("Body two"),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    expect(thumb.lines).toEqual([
      { text: "The Title", level: 1 },
      { text: "Body one", level: 0 },
      { text: "Body two", level: 0 },
    ]);
  });

  it("skips empty paragraphs and caps the line count", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [
      paragraph("   "),
      ...Array.from({ length: THUMBNAIL_MAX_LINES + 5 }, (_unused, i) =>
        paragraph(`line ${i}`),
      ),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.lines).toHaveLength(THUMBNAIL_MAX_LINES);
    expect(thumb.lines[0]).toEqual({ text: "line 0", level: 0 });
  });

  it("includes images as lines sized the way the page sizes them", () => {
    const model = createDocument("doc");
    const contentWidthCm = sectionContentWidthCm(model, 0);
    model.media = [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }];
    model.sections[0].blocks = [
      paragraph("Title", 1),
      image("m1", contentWidthCm / 2, contentWidthCm / 4),
      paragraph("After"),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    expect(thumb.lines).toEqual([
      { text: "Title", level: 1 },
      {
        text: "",
        level: 0,
        image: {
          src: "data:image/png;base64,AAAA",
          widthFraction: 0.5,
          aspectRatio: 2,
        },
      },
      { text: "After", level: 0 },
    ]);
  });

  it("previews an image-only document instead of falling back to the badge", () => {
    const model = createDocument("doc");
    model.media = [{ mediaId: "m1", contentType: "image/jpeg", data: "BBBB" }];
    model.sections[0].blocks = [image("m1", 10, 5)];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    expect(thumb.lines).toHaveLength(1);
    expect(thumb.lines[0]?.image?.src).toBe("data:image/jpeg;base64,BBBB");
  });

  it("skips an image whose media part is missing or empty, without failing", () => {
    const model = createDocument("doc");
    model.media = [{ mediaId: "present", contentType: "image/png", data: "" }];
    model.sections[0].blocks = [
      image("absent", 10, 5),
      image("present", 10, 5),
      paragraph("still here"),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    expect(thumb.lines).toEqual([{ text: "still here", level: 0 }]);
  });

  it("clamps an over-wide image to the full content width", () => {
    const model = createDocument("doc");
    model.media = [{ mediaId: "m1", contentType: "image/png", data: "CCCC" }];
    model.sections[0].blocks = [image("m1", 100, 50)];
    const thumb = buildThumbnail(model);
    expect(thumb.lines[0]?.image?.widthFraction).toBe(1);
  });

  it("a page-filling image consumes the line budget like the space it takes", () => {
    const model = createDocument("doc");
    const pageHeightCm = model.defaultStyles.pageSize.heightCm;
    model.media = [{ mediaId: "m1", contentType: "image/png", data: "DDDD" }];
    model.sections[0].blocks = [
      image("m1", 10, pageHeightCm),
      ...Array.from({ length: 5 }, (_unused, i) => paragraph(`after ${i}`)),
    ];
    const thumb = buildThumbnail(model);
    // The full-page image alone fills the preview; the trailing paragraphs
    // (which could never fit under it on the miniature) are not extracted.
    expect(thumb.lines).toHaveLength(1);
    expect(thumb.lines[0]?.image).toBeDefined();
  });

  it("includes tables as a capped text grid with the table's proportions", () => {
    const model = createDocument("doc");
    const table: Table = {
      type: "table",
      columnWidthsCm: [6, 3],
      rows: [
        {
          isHeader: true,
          cells: [
            { blocks: [paragraph("Region")] },
            { blocks: [paragraph("Sales")] },
          ],
        },
        {
          cells: [
            { blocks: [paragraph("North")] },
            { blocks: [paragraph("120")] },
          ],
        },
      ],
    };
    model.sections[0].blocks = [paragraph("Report"), table];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    const line = thumb.lines[1];
    expect(line?.table).toEqual({
      rows: [
        ["Region", "Sales"],
        ["North", "120"],
      ],
      headerRows: 1,
      columnFractions: [6 / 9, 3 / 9],
      widthFraction: Math.min(1, 9 / sectionContentWidthCm(model, 0)),
    });
  });

  it("caps a large table to the miniature's rows and columns", () => {
    const model = createDocument("doc");
    const table: Table = {
      type: "table",
      columnWidthsCm: Array.from({ length: 9 }, () => 2),
      rows: Array.from({ length: 10 }, (_unused, r) => ({
        cells: Array.from({ length: 9 }, (_alsoUnused, c) => ({
          blocks: [paragraph(`r${r}c${c}`)],
        })),
      })),
    };
    model.sections[0].blocks = [table];
    const thumb = buildThumbnail(model);
    const line = thumb.lines[0];
    expect(line?.table?.rows).toHaveLength(6);
    expect(line?.table?.rows[0]).toHaveLength(6);
  });

  it("includes charts as a deep-cloned block sized like the page sizes it", () => {
    const model = createDocument("doc");
    const chart: Chart = {
      type: "chart",
      chartType: "bar",
      categories: ["A", "B"],
      series: [{ name: "S1", values: [1, 2], color: "#2e6fd8" }],
      options: {},
      widthCm: sectionContentWidthCm(model, 0) / 2,
      heightCm: 5,
      wrap: "inline",
    } as unknown as Chart;
    model.sections[0].blocks = [chart];
    const thumb = buildThumbnail(model);
    const line = thumb.lines[0];
    expect(line?.chart?.widthFraction).toBeCloseTo(0.5, 5);
    expect(line?.chart?.chart).toEqual(chart);
    // A deep clone, never an alias of the live model's block.
    expect(line?.chart?.chart).not.toBe(chart);
  });

  it("includes a text box as its painted geometry plus its text lines", () => {
    const model = createDocument("doc");
    const shape: Shape = {
      type: "shape",
      geometry: "rect",
      widthCm: sectionContentWidthCm(model, 0) / 2,
      heightCm: 4,
      wrap: "square",
      text: {
        vAlign: "middle",
        blocks: [paragraph("Callout title"), paragraph("Second line")],
      },
    };
    model.sections[0].blocks = [shape];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    const line = thumb.lines[0]?.shape;
    expect(line?.shape).toEqual({ geometry: "rect" });
    expect(line?.textLines).toEqual(["Callout title", "Second line"]);
    expect(line?.vAlign).toBe("middle");
    expect(line?.widthFraction).toBeCloseTo(0.5, 5);
    expect(line?.widthPx).toBeGreaterThan(0);
    expect(line?.heightPx).toBeGreaterThan(0);
  });

  it("carries a shape's paint properties but strips preserved geometry XML", () => {
    const model = createDocument("doc");
    const shape: Shape = {
      type: "shape",
      geometry: "ellipse",
      widthCm: 4,
      heightCm: 4,
      wrap: "inline",
      fill: "#ff8800",
      outline: { color: "#123456", widthPt: 2, dash: "solid" },
      rotationDeg: 30,
      unsupportedGeometryXml: "<a:custGeom>big preserved xml</a:custGeom>",
    };
    model.sections[0].blocks = [shape];
    const thumb = buildThumbnail(model);
    const line = thumb.lines[0]?.shape;
    expect(line?.shape).toEqual({
      geometry: "ellipse",
      fill: "#ff8800",
      outline: { color: "#123456", widthPt: 2, dash: "solid" },
      rotationDeg: 30,
      // Presence alone drives the dashed placeholder; the XML never enters
      // the thumbnail cache.
      unsupportedGeometryXml: "",
    });
  });

  it("clamps a line shape's zero cross axis so its ratio stays finite", () => {
    const model = createDocument("doc");
    const shape: Shape = {
      type: "shape",
      geometry: "line",
      widthCm: 6,
      heightCm: 0,
      wrap: "square",
      outline: { color: "#000000", widthPt: 1, dash: "solid" },
    };
    model.sections[0].blocks = [shape];
    const thumb = buildThumbnail(model);
    const line = thumb.lines[0]?.shape;
    expect(line).toBeDefined();
    expect(Number.isFinite(line?.aspectRatio)).toBe(true);
    expect(line?.heightPx).toBeGreaterThan(0);
  });

  it("never throws on a malformed model — it reports failure instead", () => {
    const broken = { sections: [{ blocks: [{ type: "paragraph" }] }] } as never;
    expect(() => buildThumbnail(broken)).not.toThrow();
  });
});

describe("thumbnail cache (SPEC-067)", () => {
  it("caches per id + modified time and re-renders when the timestamp changes", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [paragraph("first")];

    const first = renderThumbnail("doc", 1000, model);
    expect(getCachedThumbnail("doc", 1000)).toBe(first);

    // Same id + timestamp returns the cached instance.
    expect(renderThumbnail("doc", 1000, model)).toBe(first);

    // A newer edit invalidates the old cache entry.
    expect(getCachedThumbnail("doc", 2000)).toBeUndefined();
    model.sections[0].blocks = [paragraph("edited")];
    const second = renderThumbnail("doc", 2000, model);
    expect(second).not.toBe(first);
    expect(second.lines[0]?.text).toBe("edited");
    expect(getCachedThumbnail("doc", 1000)).toBeUndefined();
  });
});
