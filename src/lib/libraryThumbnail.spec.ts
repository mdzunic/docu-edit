import { describe, it, expect, beforeEach } from "vitest";
import {
  createDocument,
  sectionContentWidthCm,
  sectionPageSize,
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
  THUMBNAIL_DIAGRAM_MAX_ITEMS,
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

  // SPEC-070: without this a flowchart previewed as a column of loose boxes
  // with nothing joining them — the connectors were skipped outright, and the
  // boxes were stacked one per line instead of placed where they sit.
  it("lays a run of drawings and connectors out at their real relative positions", () => {
    const model = createDocument("doc");
    const box = (offsetXCm: number, offsetYCm: number): Shape => ({
      type: "shape",
      geometry: "rect",
      widthCm: 4,
      heightCm: 2,
      wrap: "square",
      offsetXCm,
      offsetYCm,
    });
    model.sections[0].blocks = [
      box(0, 0),
      box(10, 6),
      {
        type: "connector",
        kind: "straight",
        widthCm: 6,
        heightCm: 6,
        wrap: "square",
        offsetXCm: 4,
        offsetYCm: 1,
        start: { xCm: 0, yCm: 0 },
        end: { xCm: 6, yCm: 6 },
        // Right-mid of the first box to left-mid of the second.
        startAttachment: { target: { block: 0 }, site: 1 },
        endAttachment: { target: { block: 1 }, site: 3 },
        outline: { color: "#c0392b", widthPt: 2, dash: "dash" },
        endArrow: "triangle",
      },
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    // One diagram line, not three stacked ones.
    expect(thumb.lines).toHaveLength(1);
    const diagram = thumb.lines[0]?.diagram;
    expect(diagram).toBeDefined();

    // Both boxes are placed, and the second sits right of and below the
    // first — exactly as their offsets put them on the page.
    expect(diagram!.items).toHaveLength(2);
    const [first, second] = diagram!.items;
    expect(first.leftFraction).toBeCloseTo(0, 5);
    expect(first.topFraction).toBeCloseTo(0, 5);
    expect(second.leftFraction).toBeGreaterThan(first.leftFraction);
    expect(second.topFraction).toBeGreaterThan(first.topFraction);
    // Sized in the same frame: a 4cm box in a 14cm-wide diagram.
    expect(first.widthFraction).toBeCloseTo(4 / 14, 5);
    for (const item of diagram!.items) {
      expect(item.leftFraction + item.widthFraction).toBeLessThanOrEqual(1.0001);
      expect(item.topFraction + item.heightFraction).toBeLessThanOrEqual(1.0001);
    }

    // The connector is routed in the diagram's own frame, carrying its
    // formatting and one arrowhead.
    expect(diagram!.connectors).toHaveLength(1);
    const [paint] = diagram!.connectors;
    expect(paint.color).toBe("#c0392b");
    expect(paint.dash).toBe("6 4");
    expect(paint.heads).toHaveLength(1);
    const coords = [paint.d, ...paint.heads.map((head) => head.d)]
      .join(" ")
      .match(/-?\d+(?:\.\d+)?/g)!
      .map(Number);
    const xs = coords.filter((_, index) => index % 2 === 0);
    const ys = coords.filter((_, index) => index % 2 === 1);
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(0);
    expect(Math.max(...xs)).toBeLessThanOrEqual(diagram!.widthPx);
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(0);
    expect(Math.max(...ys)).toBeLessThanOrEqual(diagram!.heightPx);
  });

  it("indents a diagram to its own place across the content width", () => {
    const model = createDocument("doc");
    const contentWidthCm = sectionContentWidthCm(model, 0);
    const shape = (offsetXCm: number): Shape => ({
      type: "shape",
      geometry: "rect",
      widthCm: 2,
      heightCm: 2,
      wrap: "square",
      offsetXCm,
      offsetYCm: 0,
    });
    model.sections[0].blocks = [shape(contentWidthCm / 2), shape(contentWidthCm / 2 + 3)];
    const diagram = buildThumbnail(model).lines[0]?.diagram;
    expect(diagram).toBeDefined();
    expect(diagram!.leftFraction).toBeCloseTo(0.5, 2);
    expect(diagram!.leftFraction + diagram!.widthFraction).toBeLessThanOrEqual(1.0001);
  });

  it("keeps a lone drawing on its own line — there is no relative geometry to show", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [
      { type: "shape", geometry: "rect", widthCm: 4, heightCm: 2, wrap: "square", offsetXCm: 5 },
      paragraph("After the shape"),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.lines[0]?.shape).toBeDefined();
    expect(thumb.lines[0]?.diagram).toBeUndefined();
    expect(thumb.lines[1]?.text).toBe("After the shape");
  });

  it("does not let the empty anchor paragraphs between drawings split a diagram", () => {
    const model = createDocument("doc");
    const shape = (offsetXCm: number): Shape => ({
      type: "shape",
      geometry: "rect",
      widthCm: 3,
      heightCm: 2,
      wrap: "square",
      offsetXCm,
      offsetYCm: 0,
    });
    model.sections[0].blocks = [shape(0), paragraph(""), shape(6), paragraph("   "), shape(12)];
    const thumb = buildThumbnail(model);
    expect(thumb.lines).toHaveLength(1);
    expect(thumb.lines[0]?.diagram?.items).toHaveLength(3);
  });

  it("closes a run at THUMBNAIL_DIAGRAM_MAX_ITEMS so one page cannot paint unbounded drawings", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = Array.from({ length: THUMBNAIL_DIAGRAM_MAX_ITEMS + 4 }, (_, index) => ({
      type: "shape" as const,
      geometry: "rect" as const,
      widthCm: 0.4,
      heightCm: 0.4,
      wrap: "square" as const,
      offsetXCm: index * 0.5,
      offsetYCm: 0,
    }));
    const diagrams = buildThumbnail(model).lines.filter((line) => line.diagram);
    expect(diagrams.length).toBeGreaterThan(1);
    for (const line of diagrams) {
      expect(line.diagram!.items.length).toBeLessThanOrEqual(THUMBNAIL_DIAGRAM_MAX_ITEMS);
    }
  });

  it("leaves inline drawings and images in flow — the page does not position those by offset", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [
      { type: "shape", geometry: "rect", widthCm: 4, heightCm: 2, wrap: "inline", offsetXCm: 9 },
      { type: "shape", geometry: "ellipse", widthCm: 4, heightCm: 2, wrap: "inline", offsetXCm: 0 },
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.lines).toHaveLength(2);
    expect(thumb.lines.every((line) => line.shape !== undefined)).toBe(true);
  });

  it("scales a tall elbow connector down so it cannot fill the miniature", () => {
    const model = createDocument("doc");
    const contentWidthCm = sectionContentWidthCm(model, 0);
    model.sections[0].blocks = [
      {
        type: "connector",
        kind: "elbow",
        widthCm: 2,
        heightCm: 20,
        wrap: "square",
        offsetXCm: 0,
        offsetYCm: 0,
        // A long vertical drop: the painted box is tall and nearly empty.
        start: { xCm: 0, yCm: 0 },
        end: { xCm: 2, yCm: 20 },
        endArrow: "triangle",
      },
    ];
    const line = buildThumbnail(model).lines[0]?.connector;
    expect(line).toBeDefined();
    // Aspect ratio untouched (no distortion)…
    expect(line!.aspectRatio).toBeCloseTo(line!.widthPx / line!.heightPx, 5);
    // …but the box it is shown at takes at most a quarter of the page height.
    const shownHeightCm =
      ((line!.widthFraction * contentWidthCm) / line!.widthPx) * line!.heightPx;
    expect(shownHeightCm / sectionPageSize(model, 0).heightCm).toBeLessThanOrEqual(0.2501);
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
