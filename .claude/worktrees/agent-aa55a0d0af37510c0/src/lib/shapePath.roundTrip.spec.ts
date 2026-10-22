// SPEC-071 AC-34/AC-28: a document holding a Word-style freeform (open
// polyline), a Writer-style closed enhanced-path shape, a closed curve, an
// open scribble and a shape mixing smooth and corner vertices opens, renders
// as real outlines, and saves+reopens equalStructural in both .docx and
// .odt — and a cross-format save carries the path with no DropRecord.
import { describe, it, expect } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { firstStructuralDifference } from "./equality";
import { createDocument, type DocumentModel, type Shape } from "./documentModel";
import { snapPathToGrid, type ShapePath } from "./core/shapePath";

function pathShape(path: ShapePath, overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 4,
    heightCm: 3,
    wrap: "square",
    offsetXCm: 1,
    offsetYCm: 1,
    relativeToH: "column",
    relativeToV: "paragraph",
    outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
    path: snapPathToGrid(path),
    ...overrides,
  };
}

// An open polyline (Word freeform): outline only, no fill.
const freeform: ShapePath = {
  subpaths: [
    { steps: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 0.5, y: 0.8 }, { op: "line", x: 1, y: 0.2 }], closed: false },
  ],
};

// A closed quadrilateral (Writer enhanced-path), filled.
const closedPoly: ShapePath = {
  subpaths: [
    {
      steps: [
        { op: "move", x: 0, y: 0 },
        { op: "line", x: 1, y: 0 },
        { op: "line", x: 1, y: 1 },
        { op: "line", x: 0, y: 1 },
        { op: "close" },
      ],
      closed: true,
    },
  ],
};

// A closed curve, every vertex smooth.
const closedCurve: ShapePath = {
  subpaths: [
    {
      steps: [
        { op: "move", x: 0, y: 0.5 },
        { op: "cubic", x1: 0, y1: 0, x2: 1, y2: 0, x: 1, y: 0.5 },
        { op: "cubic", x1: 1, y1: 1, x2: 0, y2: 1, x: 0, y: 0.5 },
        { op: "close" },
      ],
      closed: true,
    },
  ],
  smooth: [true, true],
};

// An open scribble (cubics), outline only.
const scribble: ShapePath = {
  subpaths: [
    {
      steps: [
        { op: "move", x: 0, y: 0.2 },
        { op: "cubic", x1: 0.2, y1: 0, x2: 0.4, y2: 0.6, x: 0.6, y: 0.4 },
        { op: "cubic", x1: 0.7, y1: 0.3, x2: 0.9, y2: 0.9, x: 1, y: 0.7 },
      ],
      closed: false,
    },
  ],
  smooth: [true, false, true],
};

// A mix of smooth and corner vertices.
const mixed: ShapePath = {
  subpaths: [
    {
      steps: [
        { op: "move", x: 0, y: 0 },
        { op: "cubic", x1: 0.25, y1: 0, x2: 0.75, y2: 0, x: 1, y: 0.5 },
        { op: "line", x: 0.5, y: 1 },
        { op: "close" },
      ],
      closed: true,
    },
  ],
  smooth: [false, true, false],
};

function pathModel(): DocumentModel {
  const model = createDocument("id-shape-paths");
  model.sections[0].blocks = [
    { type: "paragraph", runs: [{ text: "Paths", formatting: {} }] },
    pathShape(freeform),
    pathShape(closedPoly, { fill: { kind: "solid", color: "#4472c4" } }),
    pathShape(closedCurve, { fill: { kind: "solid", color: "#70ad47" } }),
    pathShape(scribble),
    pathShape(mixed, { fill: { kind: "solid", color: "#ed7d31" } }),
    { type: "paragraph", runs: [{ text: "End", formatting: {} }] },
  ];
  return model;
}

describe("path shapes round-trip (SPEC-071 AC-34)", () => {
  it("opens, saves and reopens equalStructural through .docx", async () => {
    const opened = await parseDocx((await writeDocx(pathModel())).bytes);
    const reopened = await parseDocx((await writeDocx({ ...pathModel(), sections: opened.sections })).bytes);
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff, diff ?? "").toBeNull();
    const shapes = opened.sections[0].blocks.filter((b) => b.type === "shape") as Shape[];
    expect(shapes).toHaveLength(5);
    // Every shape carries a modelled path and nothing preserved opaquely.
    expect(shapes.every((s) => s.path && !s.unsupportedGeometryXml)).toBe(true);
    // The closed curve's smoothness survived.
    expect(shapes[2].path!.smooth).toEqual([true, true]);
  });

  it("opens, saves and reopens equalStructural through .odt", async () => {
    const opened = await parseOdt((await writeOdt(pathModel())).bytes);
    const reopened = await parseOdt((await writeOdt({ ...pathModel(), sections: opened.sections })).bytes);
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff, diff ?? "").toBeNull();
    const shapes = opened.sections[0].blocks.filter((b) => b.type === "shape") as Shape[];
    expect(shapes.every((s) => s.path && !s.unsupportedGeometryXml)).toBe(true);
    expect(shapes[4].path!.smooth).toEqual([false, true, false]);
  });

  it("never carries both path and unsupportedGeometryXml (AC-3)", async () => {
    const docx = (await parseDocx((await writeDocx(pathModel())).bytes)).sections[0].blocks;
    const odt = (await parseOdt((await writeOdt(pathModel())).bytes)).sections[0].blocks;
    for (const block of [...docx, ...odt]) {
      if (block.type === "shape") {
        expect(!!(block.path && block.unsupportedGeometryXml)).toBe(false);
      }
    }
  });

  it("carries the path cross-format with no DropRecord (AC-28)", async () => {
    // .docx -> model -> .odt
    const fromDocx = await parseDocx((await writeDocx(pathModel())).bytes);
    const toOdt = await writeOdt({ ...pathModel(), sections: fromDocx.sections });
    expect(toOdt.dropped.filter((d) => d.kind === "shapeGeometry")).toHaveLength(0);
    // .odt -> model -> .docx
    const fromOdt = await parseOdt((await writeOdt(pathModel())).bytes);
    const toDocx = await writeDocx({ ...pathModel(), sections: fromOdt.sections });
    expect(toDocx.dropped.filter((d) => d.kind === "shapeGeometry")).toHaveLength(0);
    // The reopened .odt shapes still carry the paths.
    const reopened = await parseOdt(toOdt.bytes);
    const shapes = reopened.sections[0].blocks.filter((b) => b.type === "shape") as Shape[];
    expect(shapes.every((s) => s.path)).toBe(true);
  });
});
