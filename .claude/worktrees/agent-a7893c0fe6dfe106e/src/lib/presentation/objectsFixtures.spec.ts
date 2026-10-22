// SPEC-096: the spec096-objects fixture pair parses to the expected object
// model in BOTH formats — the three unaligned shapes, the 37°-rotated image,
// the tuned roundRect, the verbatim group with a NON-identity chOff/chExt
// mapping (children land in slide space through a ×2 scale), and the rotated
// verbatim group that pins the confirm-to-release path.
import { describe, expect, it } from "vitest";
import {
  buildSpec096OdpFixtureBytes,
  buildSpec096PptxFixtureBytes,
  SPEC096_ADJUSTMENT,
  SPEC096_GROUP_BOX,
  SPEC096_IMAGE_BOX,
  SPEC096_IMAGE_ROTATION,
  SPEC096_ROTATED_GROUP_BOX,
  SPEC096_ROTATED_GROUP_DEG,
  SPEC096_ROUND_RECT_BOX,
  SPEC096_SHAPE_BOXES,
} from "./fixtures/build";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import type { GroupFrame, ImageFrame, ShapeFrame, SlideFrame } from "./model";

function geomOf(frame: SlideFrame): number[] {
  return [frame.xCm, frame.yCm, frame.widthCm, frame.heightCm];
}

function expectGeomClose(frame: SlideFrame, expected: readonly number[]): void {
  geomOf(frame).forEach((v, i) => expect(Math.abs(v - expected[i])).toBeLessThan(0.02));
}

describe("spec096-objects pptx fixture", () => {
  it("parses the shapes, rotated image and tuned roundRect", async () => {
    const parsed = await parsePptx(await buildSpec096PptxFixtureBytes());
    const frames = parsed.slides[0].frames;
    const rects = frames.filter((f): f is ShapeFrame => f.kind === "shape" && f.geometry === "rect");
    SPEC096_SHAPE_BOXES.forEach((box, i) => expectGeomClose(rects[i], box));
    const image = frames.find((f): f is ImageFrame => f.kind === "image")!;
    expectGeomClose(image, SPEC096_IMAGE_BOX);
    expect(image.rotationDeg).toBe(SPEC096_IMAGE_ROTATION);
    const round = frames.find((f): f is ShapeFrame => f.kind === "shape" && f.geometry === "roundRect")!;
    expectGeomClose(round, SPEC096_ROUND_RECT_BOX);
    expect(round.adjustments).toHaveLength(1);
    expect(Math.abs(round.adjustments![0] - SPEC096_ADJUSTMENT)).toBeLessThan(1e-5);
  });

  it("maps the NON-identity chOff/chExt group's child into slide space", async () => {
    const parsed = await parsePptx(await buildSpec096PptxFixtureBytes());
    const groups = parsed.slides[0].frames.filter((f): f is GroupFrame => f.kind === "group");
    const nonIdentity = groups[0];
    expect(nonIdentity.xml).toBeTruthy();
    expectGeomClose(nonIdentity, SPEC096_GROUP_BOX);
    // The child was authored at (0,0)–(gw/4 × gh/2) in a HALF-scale child
    // space, so slide space doubles it: (20, 12, 2, 3).
    const child = nonIdentity.children[0];
    expectGeomClose(child, [SPEC096_GROUP_BOX[0], SPEC096_GROUP_BOX[1], SPEC096_GROUP_BOX[2] / 2, SPEC096_GROUP_BOX[3]]);
  });

  it("keeps the rotated group verbatim with its rotation modelled", async () => {
    const parsed = await parsePptx(await buildSpec096PptxFixtureBytes());
    const groups = parsed.slides[0].frames.filter((f): f is GroupFrame => f.kind === "group");
    const rotated = groups[1];
    expect(rotated.xml).toBeTruthy();
    expect(rotated.rotationDeg).toBe(SPEC096_ROTATED_GROUP_DEG);
    expectGeomClose(rotated, SPEC096_ROTATED_GROUP_BOX);
  });
});

describe("spec096-objects odp fixture", () => {
  it("parses the shapes, rotated image and tuned roundRect (draw:modifiers)", async () => {
    const parsed = await parseOdp(await buildSpec096OdpFixtureBytes());
    const frames = parsed.slides[0].frames;
    const rects = frames.filter((f): f is ShapeFrame => f.kind === "shape" && f.geometry === "rect");
    SPEC096_SHAPE_BOXES.forEach((box, i) => expectGeomClose(rects[i], box));
    const image = frames.find((f): f is ImageFrame => f.kind === "image")!;
    expect(image.rotationDeg).toBe(SPEC096_IMAGE_ROTATION);
    const round = frames.find((f): f is ShapeFrame => f.kind === "shape" && f.geometry === "roundRect")!;
    expect(round.adjustments).toHaveLength(1);
    expect(Math.abs(round.adjustments![0] - SPEC096_ADJUSTMENT)).toBeLessThan(1e-5);
  });

  it("keeps the draw:g group verbatim, bounds from the children's union", async () => {
    const parsed = await parseOdp(await buildSpec096OdpFixtureBytes());
    const group = parsed.slides[0].frames.find((f): f is GroupFrame => f.kind === "group")!;
    expect(group.xml).toBeTruthy();
    expect(group.children).toHaveLength(2);
    expectGeomClose(group, SPEC096_GROUP_BOX);
  });
});
