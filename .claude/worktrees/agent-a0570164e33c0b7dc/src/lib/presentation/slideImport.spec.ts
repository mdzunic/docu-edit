// SPEC-099 — import repair: id re-minting (group children included),
// animation retargeting and drop-with-frame, stripSlideParts leaving each
// caller's dirty handling intact, media de-duplication, rescale, and both
// theme/layout options.
import { describe, expect, it } from "vitest";
import type { ForeignFrame, GroupFrame, PresentationModel, Slide, SlideFrame, TextFrame } from "./model";
import { readPresentationBytes } from "./open";
import { buildPptxFixtureBytes } from "./fixtures/build";
import { addSlide, createBlankPresentation, stripRetention } from "./slides";
import { addSection, sectionsOf } from "./sections";
import { carriesPackageReference, importSlides } from "./slideImport";

const PNG_A = { contentType: "image/png", data: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==" };

function blank(n = 1): PresentationModel {
  const model = createBlankPresentation();
  for (let i = 1; i < n; i++) {
    addSlide(model, i - 1);
  }
  delete model.structureDirty;
  return model;
}

function allIds(frames: SlideFrame[]): string[] {
  return frames.flatMap((frame) => (frame.kind === "group" ? [frame.id, ...allIds(frame.children)] : [frame.id]));
}

function textFrame(id: string, text: string): TextFrame {
  return { id, kind: "text", xCm: 1, yCm: 1, widthCm: 10, heightCm: 2, paragraphs: [{ runs: [{ text }] }] };
}

function foreign(id: string, xml: string): ForeignFrame {
  return { id, kind: "foreign", xml, label: "Diagram", xCm: 1, yCm: 4, widthCm: 4, heightCm: 4 };
}

/** A hand-made source deck: one slide with a text frame, a foreign frame
 * carrying a package reference, and animations on both. */
function sourceWithForeign(): PresentationModel {
  const model = blank();
  const slide: Slide = {
    id: "s1",
    frames: [textFrame("t1", "Kept"), foreign("f1", `<p:graphicFrame><a:graphic r:id="rId7"/></p:graphicFrame>`)],
    animations: [
      { frameId: "t1", effect: "fadeIn", trigger: "click" },
      { frameId: "f1", effect: "appear", trigger: "click" },
    ],
    transition: { type: "fade" },
    partName: "ppt/slides/slide9.xml",
    preserved: { "p:clrMapOvr": "<p:clrMapOvr/>" },
  };
  model.slides = [slide];
  return model;
}

describe("importSlides — ids and retention", () => {
  it("re-mints the slide and every frame id, group children included, and rebuilds from the model", async () => {
    const source = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Source");
    const target = blank();
    const original = source.slides[3];
    const report = importSlides(target, source, [3]);
    expect(report.slides).toBe(1);
    const imported = target.slides[1];
    expect(imported.id).not.toBe(original.id);
    const before = allIds(original.frames);
    const after = allIds(imported.frames);
    expect(after.length).toBeGreaterThan(before.length - 1);
    for (const id of after) {
      expect(before).not.toContain(id);
    }
    expect(imported.frames.some((frame) => frame.kind === "group" && frame.children.length === 2)).toBe(true);
    expect(imported.partName).toBeUndefined();
    expect(imported.sourceXml).toBeUndefined();
    expect(imported.preserved).toBeUndefined();
    expect(imported.dirty).toBe(true);
    expect(imported.notesDirty).toBe(true);
    expect(target.structureDirty).toBe(true);
    // The source deck is untouched.
    expect(source.slides[3]).toBe(original);
    expect(original.partName).toBe("ppt/slides/slide4.xml");
  });

  it("stripRetention keeps its own dirty handling: dirty deleted, notesDirty left alone", () => {
    const model = blank();
    model.slides[0].dirty = true;
    model.slides[0].notesDirty = true;
    model.slides[0].partName = "ppt/slides/slide1.xml";
    const stripped = stripRetention(model);
    expect(stripped.slides[0].dirty).toBeUndefined();
    expect(stripped.slides[0].notesDirty).toBe(true);
    expect(stripped.slides[0].partName).toBeUndefined();
  });
});

describe("importSlides — frames that cannot be re-homed", () => {
  it("drops a foreign frame with a package reference, names it, and takes its animation along", () => {
    const target = blank();
    const report = importSlides(target, sourceWithForeign(), [0], { frameName: (frame) => `named:${frame.kind}` });
    const imported = target.slides[1];
    expect(imported.frames.map((frame) => frame.kind)).toEqual(["text"]);
    expect(report.droppedFrames).toBe(1);
    expect(report.droppedFrameNames).toEqual(["named:foreign"]);
    expect(report.lostAnimations).toBe(1);
    // The surviving animation follows the re-minted text frame.
    expect(imported.animations).toEqual([{ frameId: imported.frames[0].id, effect: "fadeIn", trigger: "click" }]);
    // A modelled transition survives.
    expect(imported.transition).toEqual({ type: "fade" });
  });

  it("drops a group whose xml carries a non-image reference but rebuilds one holding only plain images", () => {
    const image = { id: "i1", kind: "image" as const, mediaId: "ppt/media/image1.png", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 };
    const plain: GroupFrame = {
      id: "g1",
      kind: "group",
      children: [image, textFrame("t2", "Caption")],
      xml: `<p:grpSp><p:pic><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic></p:grpSp>`,
      xCm: 0,
      yCm: 0,
      widthCm: 10,
      heightCm: 3,
    };
    const linked: GroupFrame = { ...structuredClone(plain), id: "g2", xml: `<p:grpSp><a:hlinkClick r:id="rId4"/></p:grpSp>` };
    linked.children = linked.children.map((child, i) => ({ ...child, id: `c${i}` }));
    const source = blank();
    source.media = [{ mediaId: "ppt/media/image1.png", ...PNG_A }];
    source.slides = [{ id: "s", frames: [plain, linked] }];
    const target = blank();
    const report = importSlides(target, source, [0]);
    const frames = target.slides[1].frames;
    expect(frames).toHaveLength(1);
    const group = frames[0] as GroupFrame;
    expect(group.kind).toBe("group");
    expect(group.xml).toBeUndefined();
    expect(report.droppedFrameNames).toEqual(["Group"]);
  });

  it("drops a foreign frame from the other format and counts preserved-only timing as lost", () => {
    const source = sourceWithForeign();
    source.format = "odp";
    source.slides[0].frames = [textFrame("t1", "Kept"), foreign("f2", "<draw:frame/>")];
    source.slides[0].animations = undefined;
    source.slides[0].transition = undefined;
    source.slides[0].preserved = { anim: "<anim:par/>", "draw:transition": "1" };
    const report = importSlides(blank(), source, [0]);
    expect(report.droppedFrames).toBe(1);
    expect(report.lostTransitions).toBe(1);
    expect(report.lostAnimations).toBe(1);
  });

  it("recognises package references in verbatim xml", () => {
    expect(carriesPackageReference(`<x r:embed="rId1"/>`)).toBe(true);
    expect(carriesPackageReference(`<x r:link="rId1"/>`)).toBe(true);
    expect(carriesPackageReference(`<draw:image xlink:href="Pictures/a.png"/>`)).toBe(true);
    expect(carriesPackageReference(`<x xlink:href="https://example.com"/>`)).toBe(false);
    expect(carriesPackageReference(`<a:blip r:embed="rId1"/>`, true)).toBe(false);
    expect(carriesPackageReference(`<draw:frame><draw:image xlink:href="Pictures/a.png"/></draw:frame>`, true)).toBe(false);
  });
});

describe("importSlides — media", () => {
  it("re-homes an image into the target's directory under a fresh id, replacing target.media", async () => {
    const source = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Source");
    const target = blank();
    target.format = "odp";
    const before = target.media;
    const report = importSlides(target, source, [3]);
    expect(report.mediaReused).toBe(0);
    expect(target.media).not.toBe(before);
    expect(target.media).toHaveLength(1);
    expect(target.media![0].mediaId).toBe("Pictures/image1.png");
    const image = target.slides[1].frames.find((frame) => frame.kind === "image");
    expect(image && image.kind === "image" && image.mediaId).toBe("Pictures/image1.png");
  });

  it("reuses an identical part the target already holds", async () => {
    const source = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Source");
    const target = blank();
    const part = source.media![0];
    target.media = [{ mediaId: "ppt/media/existing.png", contentType: part.contentType, data: part.data }];
    const before = target.media;
    const report = importSlides(target, source, [3, 3]);
    expect(report.mediaReused).toBe(2);
    expect(target.media).toBe(before);
    for (const slide of target.slides.slice(1)) {
      const image = slide.frames.find((frame) => frame.kind === "image");
      expect(image && image.kind === "image" && image.mediaId).toBe("ppt/media/existing.png");
    }
  });
});

describe("importSlides — size, theme and layout", () => {
  it("rescales on a slide-size mismatch and says so", () => {
    const source = blank();
    source.slides[0].frames = [textFrame("t", "Wide")];
    const target = blank();
    target.slideWidthCm = 25.4;
    target.slideHeightCm = 19.05;
    const report = importSlides(target, source, [0]);
    expect(report.rescaled).toBe(true);
    const frame = target.slides[1].frames[0];
    expect(frame.widthCm).toBeCloseTo(10 * (25.4 / source.slideWidthCm), 2);
    expect(importSlides(blank(), source, [0]).rescaled).toBe(false);
  });

  it("use destination keeps theme references and maps layoutKey by kind", () => {
    const source = blank();
    const frame = textFrame("t", "Themed");
    frame.paragraphs[0].runs[0].color = { theme: "accent1" };
    frame.paragraphs[0].runs[0].fontFamily = { theme: "major" };
    source.slides[0] = { id: "s", frames: [frame], layoutKey: "src:title" };
    source.layouts = [{ key: "src:title", kind: "title", name: "Title", slots: [] }];
    const target = blank();
    target.layouts = [
      { key: "dst:blank", kind: "blank", name: "Blank", slots: [] },
      { key: "dst:title", kind: "title", name: "Title Slide", slots: [] },
    ];
    importSlides(target, source, [0], { formatting: "destination" });
    const imported = target.slides[1];
    expect(imported.layoutKey).toBe("dst:title");
    const run = (imported.frames[0] as TextFrame).paragraphs[0].runs[0];
    expect(run.color).toEqual({ theme: "accent1" });
    expect(run.fontFamily).toEqual({ theme: "major" });
    // No layout of that kind in the target → dropped.
    target.layouts = [{ key: "dst:blank", kind: "blank", name: "Blank", slots: [] }];
    importSlides(target, source, [0]);
    expect(target.slides[2].layoutKey).toBeUndefined();
  });

  it("keep source resolves every reference through the source theme and drops layoutKey", () => {
    const source = blank();
    source.theme = {
      name: "Source",
      colors: {
        dk1: "#000000",
        lt1: "#ffffff",
        dk2: "#111111",
        lt2: "#eeeeee",
        accent1: "#ff0000",
        accent2: "#00ff00",
        accent3: "#0000ff",
        accent4: "#ffff00",
        accent5: "#00ffff",
        accent6: "#ff00ff",
        hlink: "#0000ee",
        folHlink: "#551a8b",
      },
      fonts: { major: "Georgia", minor: "Verdana" },
    };
    const frame = textFrame("t", "Themed");
    frame.paragraphs[0].runs[0].color = { theme: "accent1" };
    frame.paragraphs[0].runs[0].fontFamily = { theme: "minor" };
    frame.fill = { type: "solid", color: { theme: "accent3" } };
    source.slides[0] = { id: "s", frames: [frame], layoutKey: "src:title" };
    const target = blank();
    importSlides(target, source, [0], { formatting: "source" });
    const imported = target.slides[1];
    expect(imported.layoutKey).toBeUndefined();
    const out = imported.frames[0] as TextFrame;
    expect(out.paragraphs[0].runs[0].color).toBe("#ff0000");
    expect(out.paragraphs[0].runs[0].fontFamily).toBe("Verdana");
    expect(out.fill).toEqual({ type: "solid", color: "#0000ff" });
  });
});

describe("importSlides — sections", () => {
  it("takes the insertion point's section and leaves the deck normalised", () => {
    const target = blank(4);
    addSection(target, 0, "Intro");
    const body = addSection(target, 2, "Body");
    const report = importSlides(target, sourceWithForeign(), [0, 0], { at: 3 });
    expect(report.slides).toBe(2);
    expect(target.slides[3].sectionKey).toBe(body);
    expect(target.slides[4].sectionKey).toBe(body);
    expect(sectionsOf(target).map((run) => [run.section.name, run.startIndex, run.endIndex])).toEqual([
      ["Intro", 0, 2],
      ["Body", 2, 6],
    ]);
  });
});
