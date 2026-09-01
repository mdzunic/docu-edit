// SPEC-096 round trips over the spec096-objects fixtures, naming the
// relations: a no-edit download stays byte-identical part-for-part under
// `equalBytes` (compared after unzip, never container bytes); rotate +
// adjust + group + align → save → reopen reproduces rotationDeg exactly,
// adjustments within 1e-5, geometry within ±0.01cm and the group's children
// order (equalStructural over the modelled subset — its IGNORED_KEYS drops
// frame ids and XML_STRING_KEYS holds `xml`, so a verbatim GroupFrame.xml
// compares canonically); writing the reopened model twice yields identical
// slide-part bytes. `equalSemantic` is typed `(model: DocumentModel)` and is
// NOT usable on a PresentationModel — no semantic relation is claimed.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { equalStructural, firstStructuralDifference, equalSemantic } from "../equality";
import { unzip } from "../core/zip";
import { buildSpec096OdpFixtureBytes, buildSpec096PptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { alignRects } from "../core/arrange";
import {
  frameRotatedBox,
  groupFrames,
  setFrameAdjustments,
  setFramePosition,
  setFrameRotation,
} from "./frames";
import type { GroupFrame, ImageFrame, PresentationModel, ShapeFrame, SlideFrame } from "./model";

async function openFixture(format: "pptx" | "odp"): Promise<{ model: PresentationModel; source: Map<string, Uint8Array> }> {
  const bytes = format === "pptx" ? await buildSpec096PptxFixtureBytes() : await buildSpec096OdpFixtureBytes();
  const model = await readPresentationBytes(bytes, format, "Objects fixture");
  return { model, source: await unzip(bytes) };
}

async function write(model: PresentationModel): Promise<Map<string, Uint8Array>> {
  const { bytes } = model.format === "pptx" ? await writePptx(model) : await writeOdp(model);
  return unzip(bytes);
}

function r3(v: number): number {
  return Math.round(v * 1000) / 1000;
}

/** The modelled subset a SPEC-096 edit promises to reproduce: kind, geometry
 * name, 3dp geometry, rotation, adjustments and (recursively) children —
 * never ids or verbatim xml. */
function projection(frame: SlideFrame): Record<string, unknown> {
  const out: Record<string, unknown> = {
    kind: frame.kind,
    box: [r3(frame.xCm), r3(frame.yCm), r3(frame.widthCm), r3(frame.heightCm)],
  };
  if (frame.rotationDeg) out.rotationDeg = frame.rotationDeg;
  if (frame.kind === "shape") {
    out.geometry = frame.geometry;
    if (frame.adjustments) out.adjustments = frame.adjustments.map((a) => Math.round(a * 1e5));
  }
  if (frame.kind === "group") {
    out.children = frame.children.map(projection);
  }
  return out;
}

for (const format of ["pptx", "odp"] as const) {
  describe(`spec096-objects ${format}: no-edit round trip`, () => {
    it("stays byte-identical part-for-part under equalBytes", async () => {
      const { model, source } = await openFixture(format);
      const written = await write(model);
      expect([...written.keys()]).toEqual([...source.keys()]);
      for (const [name, data] of source) {
        expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
      }
    });
  });

  describe(`spec096-objects ${format}: rotate + adjust + group + align → save → reopen`, () => {
    async function editedAndReopened(): Promise<{ edited: PresentationModel; reopened: PresentationModel }> {
      const { model } = await openFixture(format);
      const frames = model.slides[0].frames;
      const rects = frames.filter((f): f is ShapeFrame => f.kind === "shape" && f.geometry === "rect");
      const image = frames.find((f): f is ImageFrame => f.kind === "image")!;
      const round = frames.find((f): f is ShapeFrame => f.kind === "shape" && f.geometry === "roundRect")!;
      // Rotate the image to 123°, tune the roundRect, align the three rects
      // left (the same rect algebra the Align menu runs), group two of them.
      setFrameRotation(model, 0, image.id, 123);
      setFrameAdjustments(model, 0, round.id, [0.33]);
      const aligned = alignRects(rects.map(frameRotatedBox), "left", "selection", {
        xCm: 0,
        yCm: 0,
        widthCm: model.slideWidthCm,
        heightCm: model.slideHeightCm,
      });
      rects.forEach((rect, i) => setFramePosition(model, 0, rect.id, aligned[i].xCm, aligned[i].yCm));
      groupFrames(model, 0, [rects[0].id, rects[1].id]);
      const { bytes } = model.format === "pptx" ? await writePptx(model) : await writeOdp(model);
      const reopened = await readPresentationBytes(bytes, format, "Objects fixture");
      return { edited: model, reopened };
    }

    it("reproduces rotation exactly, adjustments within 1e-5, geometry within ±0.01cm, child order", async () => {
      const { edited, reopened } = await editedAndReopened();
      const image = reopened.slides[0].frames.find((f): f is ImageFrame => f.kind === "image")!;
      expect(image.rotationDeg).toBe(123);
      const round = reopened.slides[0].frames.find(
        (f): f is ShapeFrame => f.kind === "shape" && f.geometry === "roundRect",
      )!;
      expect(Math.abs(round.adjustments![0] - 0.33)).toBeLessThan(1e-5);
      // The AUTHORED group (its two children were aligned to x=2) — the pptx
      // fixture also carries two verbatim groups elsewhere on the slide.
      const group = reopened.slides[0].frames.find(
        (f): f is GroupFrame => f.kind === "group" && f.children.length === 2 && Math.abs(f.children[0].xCm - 2) < 0.05,
      )!;
      expect(group.children).toHaveLength(2);
      const editedGroup = edited.slides[0].frames.find((f): f is GroupFrame => f.kind === "group" && !f.xml)!;
      editedGroup.children.forEach((child, i) => {
        const back = group.children[i];
        expect(Math.abs(back.xCm - child.xCm)).toBeLessThan(0.01);
        expect(Math.abs(back.yCm - child.yCm)).toBeLessThan(0.01);
        expect(Math.abs(back.widthCm - child.widthCm)).toBeLessThan(0.01);
        expect(Math.abs(back.heightCm - child.heightCm)).toBeLessThan(0.01);
      });
      // The aligned rects' left edges came back exactly: the authored
      // group's two children plus the one still-top-level rect.
      const topLevelRects = reopened.slides[0].frames.filter(
        (f): f is ShapeFrame => f.kind === "shape" && f.geometry === "rect",
      );
      for (const rect of [...group.children, ...topLevelRects]) {
        expect(Math.abs(rect.xCm - 2)).toBeLessThan(0.01);
      }
    });

    it("compares equal under equalStructural over the modelled subset", async () => {
      const { edited, reopened } = await editedAndReopened();
      // pptx keeps the two verbatim groups distinct from the authored one;
      // compare the projection frame-by-frame.
      const a = edited.slides[0].frames.map(projection);
      const b = reopened.slides[0].frames.map(projection);
      expect(firstStructuralDifference(a, b)).toBeNull();
      expect(equalStructural(a, b)).toBe(true);
    });

    it("compares a verbatim GroupFrame.xml canonically (XML_STRING_KEYS) and drops ids (IGNORED_KEYS)", async () => {
      const { reopened } = await editedAndReopened();
      const clone = structuredClone(reopened);
      // Re-mint every frame id and reformat a verbatim group's xml with
      // shuffled attribute order — equalStructural must not care.
      for (const frame of clone.slides[0].frames) {
        frame.id = `re-${frame.id}`;
        if (frame.kind === "group" && frame.xml) {
          frame.xml = frame.xml.replace(/<a:off x="([^"]*)" y="([^"]*)"\/>/, '<a:off y="$2" x="$1"/>');
        }
      }
      expect(equalStructural(reopened.slides, clone.slides)).toBe(true);
      // equalSemantic is typed (model: DocumentModel) — not usable here, so
      // no semantic relation is claimed for presentations.
      // @ts-expect-error a PresentationModel is not a DocumentModel
      void (() => equalSemantic(reopened, clone));
    });

    it("writes the reopened model twice with identical slide-part bytes", async () => {
      const { reopened } = await editedAndReopened();
      const first = await write(reopened);
      const second = await write(structuredClone(reopened));
      const slideParts = [...first.keys()].filter(
        (name) => name.startsWith("ppt/slides/slide") || name === "content.xml",
      );
      expect(slideParts.length).toBeGreaterThan(0);
      for (const name of slideParts) {
        expect(equalBytes(first.get(name)!, second.get(name)!), `part ${name} must be write-stable`).toBe(true);
      }
    });
  });
}
