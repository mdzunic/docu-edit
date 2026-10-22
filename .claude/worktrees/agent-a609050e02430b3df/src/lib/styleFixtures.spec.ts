// SPEC-073 AC-38: documents holding a gradient-filled shape, a
// picture-filled shape, a shadowed callout, a semi-transparent shape over
// text and a double-headed arrow open, save and reopen `equalStructural`
// through both formats. Mirrors connectorFixtures.spec.ts's
// writeDocx/parseDocx + writeOdt/parseOdt pattern — the fixture is built
// with this product's own writer, then re-read, so the whole package (media
// parts, style definitions, graphic styles) is exercised, not just the
// shape markup.
import { describe, it, expect } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { equalStructural, firstStructuralDifference } from "./equality";
import {
  createDocument,
  isShape,
  type DocumentModel,
  type Paragraph,
  type Shape,
} from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function shape(overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 4,
    heightCm: 2,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 1,
    offsetYCm: 1,
    ...overrides,
  };
}

// A 1x1 red PNG.
const PNG_1PX =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function styledModel(format: "docx" | "odt"): DocumentModel {
  const model = createDocument("styled-fixture");
  model.format = format;
  // The part-path mediaId a reopened document holds (readers name media by
  // package part), so save → reopen is an identity on the fill's mediaId.
  const mediaId = format === "docx" ? "word/media/image1.png" : "Pictures/image1.png";
  model.media = [{ mediaId, contentType: "image/png", data: PNG_1PX }];
  model.sections = [
    {
      blocks: [
        paragraph("Shape styling fixtures"),
        // A gradient-filled shape (two stops — both formats express it).
        shape({
          fill: {
            kind: "gradient",
            gradient: {
              type: "linear",
              angleDeg: 90,
              stops: [
                { position: 0, color: "#4472c4" },
                { position: 1, color: "#8faadc" },
              ],
            },
          },
          outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
        }),
        // A picture-filled shape referencing the shared media part (AC-4).
        shape({ offsetXCm: 6, fill: { kind: "picture", mediaId, mode: "stretch" } }),
        // A shadowed callout.
        shape({
          geometry: "wedgeRectCallout",
          offsetYCm: 5,
          fill: { kind: "solid", color: "#dae3f3" },
          outline: { color: "#4472c4", widthPt: 1.5, dash: "solid" },
          shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 },
        }),
        // A semi-transparent shape over text (AC-7's model half).
        shape({
          offsetYCm: 0.2,
          wrap: "infront",
          fill: { kind: "solid", color: "#70ad47", alpha: 0.5 },
          outline: { color: "#2e5395", widthPt: 1, dash: "solid", alpha: 0.5 },
        }),
        // A double-headed arrow: the arrow geometry plus an explicit start
        // arrowhead (the implicit SPEC-070 end triangle stays implicit).
        shape({
          geometry: "arrow",
          heightCm: 0,
          offsetYCm: 8,
          outline: {
            color: "#2e5395",
            widthPt: 2,
            dash: "solid",
            startArrow: { kind: "triangle", size: "medium" },
          },
        }),
        paragraph("After the drawings"),
      ],
    },
  ];
  if (format === "docx") {
    // The .docx anchor always carries relativeHeight, which reads back as
    // zOrder 0 — bake it in so save → reopen is an identity there. The .odt
    // writer/reader carry no z-order for an unstacked shape.
    for (const block of model.sections[0].blocks) {
      if (isShape(block)) {
        block.zOrder = 0;
      }
    }
  }
  return model;
}

const FIXED_MODIFIED = new Date("2026-01-01T00:00:00Z");

function reopenedModel(parsedSections: DocumentModel["sections"], media: DocumentModel["media"], format: "docx" | "odt"): DocumentModel {
  const model = createDocument("styled-fixture");
  model.format = format;
  model.sections = parsedSections;
  model.media = media;
  return model;
}

describe("SPEC-073 AC-38 — styled fixtures round-trip equalStructural (.docx)", () => {
  it("save → reopen preserves every styled shape structurally", async () => {
    const source = styledModel("docx");
    const { bytes, dropped } = await writeDocx(source, { modified: FIXED_MODIFIED });
    expect(dropped).toEqual([]);
    const parsed = await parseDocx(bytes);
    const reopened = reopenedModel(parsed.sections, parsed.media, "docx");
    const sourceShapes = source.sections[0].blocks.filter((b) => isShape(b));
    const reopenedShapes = reopened.sections[0].blocks.filter((b) => isShape(b));
    expect(reopenedShapes).toHaveLength(sourceShapes.length);
    for (let i = 0; i < sourceShapes.length; i++) {
      expect(
        equalStructural(sourceShapes[i], reopenedShapes[i]),
        `shape ${i}: ${JSON.stringify(firstStructuralDifference(sourceShapes[i], reopenedShapes[i]))}`,
      ).toBe(true);
    }
    // The picture fill's media part survives as one part (AC-4).
    expect(parsed.media.filter((part) => part.contentType === "image/png")).toHaveLength(1);
  });

  it("a second save of the reopened document is stable (no drift)", async () => {
    const source = styledModel("docx");
    const first = await writeDocx(source, { modified: FIXED_MODIFIED });
    const parsed = await parseDocx(first.bytes);
    const reopened = reopenedModel(parsed.sections, parsed.media, "docx");
    const second = await writeDocx(reopened, { modified: FIXED_MODIFIED });
    const reparsed = await parseDocx(second.bytes);
    const a = parsed.sections[0].blocks.filter((b) => isShape(b));
    const b = reparsed.sections[0].blocks.filter((c) => isShape(c));
    expect(b).toHaveLength(a.length);
    for (let i = 0; i < a.length; i++) {
      expect(equalStructural(a[i], b[i]), `shape ${i}`).toBe(true);
    }
  });
});

describe("SPEC-073 AC-38 — styled fixtures round-trip equalStructural (.odt)", () => {
  it("save → reopen preserves every styled shape structurally", async () => {
    const source = styledModel("odt");
    const { bytes, dropped } = await writeOdt(source);
    expect(dropped).toEqual([]);
    const parsed = await parseOdt(bytes);
    const reopened = reopenedModel(parsed.sections, parsed.media, "odt");
    const sourceShapes = source.sections[0].blocks.filter((b) => isShape(b));
    const reopenedShapes = reopened.sections[0].blocks.filter((b) => isShape(b));
    expect(reopenedShapes).toHaveLength(sourceShapes.length);
    for (let i = 0; i < sourceShapes.length; i++) {
      expect(
        equalStructural(sourceShapes[i], reopenedShapes[i]),
        `shape ${i}: ${JSON.stringify(firstStructuralDifference(sourceShapes[i], reopenedShapes[i]))}`,
      ).toBe(true);
    }
  });
});
