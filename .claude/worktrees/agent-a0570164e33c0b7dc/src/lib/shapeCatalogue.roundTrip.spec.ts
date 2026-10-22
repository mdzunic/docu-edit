// SPEC-069 AC-14: a document holding an expanded-catalogue flowchart box, a
// banner, a chevron, a cylinder and an adjusted callout opens, saves and
// reopens equalStructural in both .docx and .odt — the fixpoint check every
// round-trip suite uses (open once, write+reopen, compare the two reads).
import { describe, it, expect } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { firstStructuralDifference } from "./equality";
import { createDocument, type DocumentModel, type Shape } from "./documentModel";

function shape(overrides: Partial<Shape>): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 3,
    heightCm: 2,
    wrap: "square",
    offsetXCm: 1,
    offsetYCm: 1,
    relativeToH: "column",
    relativeToV: "paragraph",
    fill: { kind: "solid", color: "#4472c4" },
    outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
    ...overrides,
  };
}

/** A flowchart process, an up ribbon banner, a chevron, a cylinder and a
 * callout tuned away from its default tail (AC-14a). */
function catalogueModel(): DocumentModel {
  const model = createDocument("id-shape-catalogue");
  model.sections[0].blocks = [
    { type: "paragraph", runs: [{ text: "Diagram", formatting: {} }] },
    shape({ geometry: "flowChartProcess" }),
    shape({ geometry: "ribbon2" }),
    shape({ geometry: "chevron", adjustments: [0.4] }),
    shape({ geometry: "can", adjustments: [0.2] }),
    shape({ geometry: "wedgeRectCallout", adjustments: [0.7, 0.95] }),
    { type: "paragraph", runs: [{ text: "End", formatting: {} }] },
  ];
  return model;
}

describe("expanded shape catalogue round-trip (SPEC-069 AC-14)", () => {
  it("opens, saves and reopens equalStructural through .docx", async () => {
    const opened = await parseDocx((await writeDocx(catalogueModel())).bytes);
    const reopened = await parseDocx(
      (await writeDocx({ ...catalogueModel(), sections: opened.sections })).bytes,
    );
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff, diff ?? "").toBeNull();
    // The catalogue geometries and their adjustments survived the first open.
    const shapes = opened.sections[0].blocks.filter((b) => b.type === "shape") as Shape[];
    expect(shapes.map((s) => s.geometry)).toEqual([
      "flowChartProcess",
      "ribbon2",
      "chevron",
      "can",
      "wedgeRectCallout",
    ]);
    expect(shapes.find((s) => s.geometry === "chevron")!.adjustments![0]).toBeCloseTo(0.4, 3);
    expect(shapes.find((s) => s.geometry === "wedgeRectCallout")!.adjustments).toHaveLength(2);
  });

  it("opens, saves and reopens equalStructural through .odt", async () => {
    const opened = await parseOdt((await writeOdt(catalogueModel())).bytes);
    const reopened = await parseOdt(
      (await writeOdt({ ...catalogueModel(), sections: opened.sections })).bytes,
    );
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff, diff ?? "").toBeNull();
    const shapes = opened.sections[0].blocks.filter((b) => b.type === "shape") as Shape[];
    expect(shapes.map((s) => s.geometry)).toEqual([
      "flowChartProcess",
      "ribbon2",
      "chevron",
      "can",
      "wedgeRectCallout",
    ]);
    expect(shapes.find((s) => s.geometry === "can")!.adjustments![0]).toBeCloseTo(0.2, 3);
  });
});
