// slideOutline derivation (AC-4) and frame-kind classification (AC-6), tested
// both on hand-built slides (the pure derivation) and on the in-repo fixtures
// (the readers' classification tables).
import { describe, expect, it } from "vitest";
import { paragraphText, slideFrameKinds, slideOutline } from "./slideModel";
import type { Slide, SlideFrame, SlideParagraph } from "./model";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { buildOdpFixtureBytes, buildPptxFixtureBytes, FIXTURE_FRAME_KINDS } from "./fixtures/build";

let nextId = 0;
function frameId(): string {
  return `f${nextId++}`;
}

function para(...texts: string[]): SlideParagraph {
  return { runs: texts.map((text) => ({ text })) };
}

function textFrame(role: "title" | "body" | undefined, paragraphs: SlideParagraph[]): SlideFrame {
  const frame: SlideFrame = { id: frameId(), kind: "text", xCm: 0, yCm: 0, widthCm: 5, heightCm: 2, paragraphs };
  if (role) {
    frame.role = role;
  }
  return frame;
}

function slide(frames: SlideFrame[]): Slide {
  return { id: frameId(), frames };
}

describe("slideOutline", () => {
  it("takes the title from the first title-role frame, lines from the rest in order", () => {
    const s = slide([
      textFrame("title", [para("My title")]),
      textFrame("body", [para("first"), para("second")]),
    ]);
    expect(slideOutline(s)).toEqual({ title: "My title", lines: ["first", "second"] });
  });

  it("descends into groups for outline lines (grouped text is slide text)", () => {
    const grouped: SlideFrame = {
      id: frameId(),
      kind: "group",
      xCm: 0,
      yCm: 0,
      widthCm: 5,
      heightCm: 2,
      xml: "<g/>",
      children: [textFrame(undefined, [para("grouped note")])],
    };
    const s = slide([textFrame(undefined, [para("plain")]), grouped]);
    expect(slideOutline(s)).toEqual({ lines: ["plain", "grouped note"] });
  });

  it("omits the title when no title frame is present, and drops empty runs", () => {
    const s = slide([textFrame("body", [para(""), para("kept")])]);
    expect(slideOutline(s)).toEqual({ lines: ["kept"] });
  });

  it("ignores image and foreign frames", () => {
    const s = slide([
      textFrame("title", [para("T")]),
      { id: frameId(), kind: "image", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, mediaId: "m" },
      { id: frameId(), kind: "foreign", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, xml: "<x/>", label: "Table" },
    ]);
    expect(slideOutline(s)).toEqual({ title: "T", lines: [] });
  });

  it("joins a multi-run paragraph via paragraphText", () => {
    expect(paragraphText({ runs: [{ text: "Hello " }, { text: "world" }] })).toBe("Hello world");
  });
});

describe("frame kind classification (both formats)", () => {
  it("classifies each fixture slide's frames identically across pptx and odp", async () => {
    const pptx = (await parsePptx(await buildPptxFixtureBytes())).slides;
    const odp = (await parseOdp(await buildOdpFixtureBytes())).slides;
    expect(pptx.map((s) => slideFrameKinds(s))).toEqual(FIXTURE_FRAME_KINDS);
    expect(odp.map((s) => slideFrameKinds(s))).toEqual(FIXTURE_FRAME_KINDS);
  });
});
