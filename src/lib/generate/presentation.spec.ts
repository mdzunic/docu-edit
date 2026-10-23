// SPEC-106 — PresentationBuilder: every method's model output and validation
// errors (path text included, `sizePt` among the rejected run keys), the
// placeholder frames coming from slides.ts's own exports, LengthInput
// conversion, the fontSizePt → sizePt run mapping with its drop count, the
// output methods' delegation to SPEC-080's writers, toDeck purity, and the
// product boundary the deck path keeps (AC-21).
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it, vi } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import type { ImageFrame, TextFrame } from "../presentation/model";
import { writeOdp } from "../presentation/odp/writer";
import { writePptx } from "../presentation/pptx/writer";
import { createBlankPresentation, r3, titleBodyFrames, titleSlideFrames } from "../presentation/slides";
import { DocumentGenerationError } from "./errors";
import { PresentationBuilder } from "./presentation";

// 1×1 and 4×2 PNGs.
const PNG_1X1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const PNG_4X2 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";
const MODIFIED = new Date(Date.UTC(2026, 8, 30, 12, 0, 0));

function withoutId<T extends { id: string }>(frame: T): Omit<T, "id"> {
  const { id: _id, ...rest } = frame;
  return rest;
}

function geometry(frame: { xCm: number; yCm: number; widthCm: number; heightCm: number }): number[] {
  return [frame.xCm, frame.yCm, frame.widthCm, frame.heightCm];
}

function textFrames(builder: PresentationBuilder, slide = 0): TextFrame[] {
  return builder.toDeck().slides[slide].frames.filter((frame): frame is TextFrame => frame.kind === "text");
}

function expectGenerationError(fn: () => unknown, message: string | RegExp): void {
  expect(fn).toThrow(DocumentGenerationError);
  expect(fn).toThrow(message);
}

describe("new PresentationBuilder(options)", () => {
  it("starts from createBlankPresentation()'s model with no first slide", () => {
    const blank = createBlankPresentation();
    const deck = new PresentationBuilder().toDeck();
    expect(deck.slides).toEqual([]);
    expect(deck.kind).toBe("presentation");
    expect(deck.format).toBe(blank.format);
    expect(deck.title).toBe(blank.title);
    expect([deck.slideWidthCm, deck.slideHeightCm]).toEqual([blank.slideWidthCm, blank.slideHeightCm]);
    expect(deck.opaque).toBeUndefined();
  });

  it("takes a title, a format and a named size", () => {
    const deck = new PresentationBuilder({ title: "Q3 review", format: "odp", size: "4:3" }).toDeck();
    expect(deck.title).toBe("Q3 review");
    expect(deck.format).toBe("odp");
    // SPEC-090's 4:3, restated: 25.4 × 19.05 cm.
    expect([deck.slideWidthCm, deck.slideHeightCm]).toEqual([25.4, 19.05]);
    expect(new PresentationBuilder({ size: "16:9" }).toDeck().slideWidthCm).toBe(33.867);
  });

  it("takes an explicit size as LengthInputs, converted to cm", () => {
    const deck = new PresentationBuilder({ size: { width: { in: 10 }, height: { pt: 540 } } }).toDeck();
    expect([deck.slideWidthCm, deck.slideHeightCm]).toEqual([25.4, 19.05]);
  });

  it("rejects unknown options, formats and sizes by path", () => {
    expectGenerationError(() => new PresentationBuilder({ bogus: 1 } as never), "options.bogus: unknown property");
    expectGenerationError(() => new PresentationBuilder({ format: "docx" } as never), /^options\.format: expected one of "pptx", "odp"/);
    expectGenerationError(() => new PresentationBuilder({ size: "16:10" } as never), /^options\.size: expected one of "16:9", "4:3"/);
    expectGenerationError(() => new PresentationBuilder({ size: { width: { cm: 1 }, height: { cm: 10 } } }), /^options\.size\.width: expected a slide dimension/);
    expectGenerationError(() => new PresentationBuilder({ size: { width: 20, height: { cm: 10 } } } as never), /^options\.size\.width: expected a length/);
    expectGenerationError(() => new PresentationBuilder({ title: 7 } as never), "options.title: expected a string");
  });
});

describe("every add*/set* method returns this (AC-2)", () => {
  it("chains", () => {
    const builder = new PresentationBuilder();
    const rect = { x: { cm: 1 }, y: { cm: 1 }, width: { cm: 4 }, height: { cm: 2 } };
    expect(builder.addTitleSlide("A")).toBe(builder);
    expect(builder.addBulletSlide("B", ["x"])).toBe(builder);
    expect(builder.addImageSlide("C", { data: PNG_1X1, contentType: "image/png" })).toBe(builder);
    expect(builder.addBlankSlide()).toBe(builder);
    expect(builder.addBullets(["y"])).toBe(builder);
    expect(builder.addTextBox("z", rect)).toBe(builder);
    expect(builder.addImage({ data: PNG_1X1, contentType: "image/png", x: { cm: 1 }, y: { cm: 1 }, width: { cm: 2 } })).toBe(builder);
    expect(builder.setNotes(["n"])).toBe(builder);
    expect(builder.setBackground("#123456")).toBe(builder);
    expect(builder.toDeck().slides).toHaveLength(4);
  });
});

describe("adding slides by intent (AC-4, AC-5)", () => {
  it("addTitleSlide lands slides.ts's titleSlideFrames pair with the text filled", () => {
    const builder = new PresentationBuilder().addTitleSlide("Q3 review", "Prepared 30 September");
    const [title, subtitle] = textFrames(builder);
    const [expectedTitle, expectedSubtitle] = titleSlideFrames(33.867, 19.05);
    expect(withoutId({ ...title, paragraphs: [] })).toEqual(withoutId(expectedTitle));
    expect(withoutId({ ...subtitle, paragraphs: [] })).toEqual(withoutId(expectedSubtitle));
    expect(title.paragraphs).toEqual([{ runs: [{ text: "Q3 review" }] }]);
    expect(subtitle.paragraphs).toEqual([{ runs: [{ text: "Prepared 30 September" }] }]);
  });

  it("addTitleSlide without a subtitle keeps the empty subtitle placeholder, as a blank deck's first slide does", () => {
    const [, subtitle] = textFrames(new PresentationBuilder().addTitleSlide("Only a title"));
    expect(subtitle.role).toBe("subtitle");
    expect(subtitle.paragraphs).toEqual([]);
  });

  it("addBulletSlide lands slides.ts's titleBodyFrames pair at the deck's size", () => {
    const builder = new PresentationBuilder({ size: "4:3" }).addBulletSlide("Highlights", ["One"]);
    const [title, body] = textFrames(builder);
    const [expectedTitle, expectedBody] = titleBodyFrames(25.4, 19.05);
    expect(geometry(title)).toEqual(geometry(expectedTitle));
    expect(withoutId({ ...body, paragraphs: [] })).toEqual(withoutId(expectedBody));
    expect(body.paragraphs).toEqual([{ runs: [{ text: "One" }], bullet: "disc" }]);
  });

  it("nests a nested array one level deeper, clamped at MAX_INDENT_LEVEL", () => {
    const deep: unknown[] = ["L8"];
    let nested: unknown[] = deep;
    for (let i = 0; i < 9; i++) {
      nested = [nested];
    }
    const builder = new PresentationBuilder().addBulletSlide("Highlights", [
      "Revenue up 12%",
      ["Europe up 20%", "APAC flat", ["Germany"]],
      "Costs flat",
      nested as never,
    ]);
    const [, body] = textFrames(builder);
    expect(body.paragraphs.map((p) => [p.runs[0].text, p.level ?? 0])).toEqual([
      ["Revenue up 12%", 0],
      ["Europe up 20%", 1],
      ["APAC flat", 1],
      ["Germany", 2],
      ["Costs flat", 0],
      ["L8", 8],
    ]);
  });

  it("takes explicit { text, level } and { runs, level } rows, and numbered style", () => {
    const builder = new PresentationBuilder().addBulletSlide(
      "Steps",
      [{ text: "First", level: 0 }, { runs: ["Second ", { text: "bold", bold: true }], level: 2 }, { text: "Third" }],
      { style: "number" },
    );
    const [, body] = textFrames(builder);
    expect(body.paragraphs).toEqual([
      { runs: [{ text: "First" }], bullet: "number" },
      { runs: [{ text: "Second " }, { text: "bold", bold: true }], bullet: "number", level: 2 },
      { runs: [{ text: "Third" }], bullet: "number" },
    ]);
  });

  it("adds no paragraph for a row with no text (neither reader keeps an empty paragraph)", () => {
    const [, body] = textFrames(new PresentationBuilder().addBulletSlide("T", ["a", "", { text: "" }, "b"]));
    expect(body.paragraphs.map((p) => p.runs[0].text)).toEqual(["a", "b"]);
  });

  it("names the offending bullet path", () => {
    const builder = new PresentationBuilder().addTitleSlide("a").addTitleSlide("b");
    expectGenerationError(() => builder.addBulletSlide("c", ["ok", 42 as never]), /^slides\[2\]\.bullets\[1\]: expected string or run array/);
    expectGenerationError(() => builder.addBulletSlide("c", ["ok", [null as never]]), /^slides\[2\]\.bullets\[1\]\[0\]: expected string or run array/);
    expectGenerationError(() => builder.addBulletSlide("c", [{ text: "x", level: 9 }]), "slides[2].bullets[0].level: expected an integer between 0 and 8");
    expectGenerationError(() => builder.addBulletSlide("c", [{ text: "x", bold: true } as never]), "slides[2].bullets[0].bold: unknown property");
    expectGenerationError(() => builder.addBulletSlide("c", "x" as never), "slides[2].bullets: expected an array");
    expectGenerationError(() => builder.addBulletSlide("c", ["x"], { style: "roman" } as never), /^slides\[2\]\.style: expected one of/);
    expectGenerationError(() => builder.addBulletSlide(5 as never, ["x"]), "slides[2].title: expected string or run array");
  });

  it("addImageSlide fits the image into the body rect at its own aspect, centred", () => {
    const builder = new PresentationBuilder().addImageSlide("Pipeline", { data: PNG_4X2, contentType: "image/png" });
    const deck = builder.toDeck();
    const [title, image] = deck.slides[0].frames;
    const [expectedTitle, body] = titleBodyFrames(33.867, 19.05);
    expect(geometry(title)).toEqual(geometry(expectedTitle));
    expect(image.kind).toBe("image");
    // 2:1 inside the 28.448 × 11.43 body: height-bound, centred horizontally.
    const width = r3(11.43 * 2);
    expect(geometry(image)).toEqual([r3(body.xCm + (body.widthCm - width) / 2), body.yCm, width, 11.43]);
    expect(deck.media).toEqual([{ mediaId: "ppt/media/image1.png", contentType: "image/png", data: PNG_4X2 }]);
  });

  it("addImageSlide takes a width or height (the other side from the aspect), shrunk to fit", () => {
    const sized = new PresentationBuilder()
      .addImageSlide("A", { data: PNG_4X2, contentType: "image/png", width: { cm: 10 } })
      .addImageSlide("B", { data: PNG_4X2, contentType: "image/png", height: { cm: 4 } })
      .addImageSlide("C", { data: PNG_4X2, contentType: "image/png", width: { cm: 100 }, height: { cm: 10 } })
      .toDeck();
    const size = (i: number) => [sized.slides[i].frames[1].widthCm, sized.slides[i].frames[1].heightCm];
    expect(size(0)).toEqual([10, 5]);
    expect(size(1)).toEqual([8, 4]);
    // 100 × 10 cm overflows the 28.448 cm-wide body: shrunk at its aspect.
    expect(size(2)).toEqual([28.448, 2.845]);
  });

  it("mints media through insertImageFrame's {dir}image{n}.{ext} rule, per format", () => {
    const pptx = new PresentationBuilder()
      .addImageSlide("A", { data: PNG_1X1, contentType: "image/png" })
      .addImageSlide("B", { data: PNG_1X1, contentType: "image/png" })
      .toDeck();
    expect(pptx.media!.map((part) => part.mediaId)).toEqual(["ppt/media/image1.png", "ppt/media/image2.png"]);
    expect((pptx.slides[1].frames[1] as ImageFrame).mediaId).toBe("ppt/media/image2.png");
    const odp = new PresentationBuilder({ format: "odp" }).addImageSlide("A", { data: PNG_1X1, contentType: "image/png" }).toDeck();
    expect(odp.media![0].mediaId).toBe("Pictures/image1.png");
  });

  it("validates the image by path before landing anything", () => {
    const builder = new PresentationBuilder();
    expectGenerationError(() => builder.addImageSlide("A", { data: PNG_1X1, contentType: "image/webp" } as never), /^slides\[0\]\.image\.contentType: expected one of/);
    expectGenerationError(() => builder.addImageSlide("A", { data: "%%%", contentType: "image/png" }), "slides[0].image.data: not valid base64");
    expectGenerationError(() => builder.addImageSlide("A", { data: btoa("not a png"), contentType: "image/png" }), "slides[0].image.data: could not decode image dimensions for image/png");
    expectGenerationError(() => builder.addImageSlide("A", { data: PNG_1X1, contentType: "image/png", x: { cm: 1 } } as never), "slides[0].image.x: unknown property");
    expectGenerationError(() => builder.addImageSlide("A", { data: PNG_1X1, contentType: "image/png", altText: 3 } as never), "slides[0].image.altText: expected a string");
    expect(builder.toDeck().slides).toEqual([]);
    expect(builder.toDeck().media).toBeUndefined();
  });

  it("addBlankSlide appends a slide with no frames", () => {
    expect(new PresentationBuilder().addBlankSlide().toDeck().slides[0].frames).toEqual([]);
  });
});

describe("per-slide content methods act on the most recent slide (AC-6)", () => {
  it("throw when no slide has been added", () => {
    const builder = new PresentationBuilder();
    const rect = { x: { cm: 1 }, y: { cm: 1 }, width: { cm: 4 }, height: { cm: 2 } };
    expectGenerationError(() => builder.addBullets(["x"]), /^addBullets: no slide to act on/);
    expectGenerationError(() => builder.addTextBox("x", rect), /^addTextBox: no slide to act on/);
    expectGenerationError(() => builder.addImage({ data: PNG_1X1, contentType: "image/png", x: { cm: 0 }, y: { cm: 0 }, width: { cm: 1 } }), /^addImage: no slide to act on/);
    expectGenerationError(() => builder.setNotes(["x"]), /^setNotes: no slide to act on/);
    expectGenerationError(() => builder.setBackground("#000000"), /^setBackground: no slide to act on/);
  });

  it("addBullets appends to the body, or adds one at addSlide's geometry", () => {
    const builder = new PresentationBuilder()
      .addBulletSlide("A", ["one"])
      .addBullets(["two", ["three"]], { style: "number" })
      .addTitleSlide("B")
      .addBullets(["four"])
      .addBlankSlide()
      .addBullets(["five"]);
    const deck = builder.toDeck();
    const body = (i: number) => deck.slides[i].frames.find((f): f is TextFrame => f.kind === "text" && f.role === "body")!;
    expect(body(0).paragraphs).toEqual([
      { runs: [{ text: "one" }], bullet: "disc" },
      { runs: [{ text: "two" }], bullet: "number" },
      { runs: [{ text: "three" }], bullet: "number", level: 1 },
    ]);
    const expectedBody = titleBodyFrames(33.867, 19.05)[1];
    expect(geometry(body(1))).toEqual(geometry(expectedBody));
    expect(deck.slides[1].frames.map((f) => (f as TextFrame).role)).toEqual(["title", "subtitle", "body"]);
    expect(deck.slides[2].frames).toHaveLength(1);
    expect(body(2).paragraphs[0].runs[0].text).toBe("five");
    expectGenerationError(() => builder.addBullets([7 as never]), /^slides\[2\]\.bullets\[0\]: expected string or run array/);
  });

  it("addTextBox lands a free text box, every rect field through lengthToEmu → emuToCm → r3", () => {
    const deck = new PresentationBuilder()
      .addBlankSlide()
      .addTextBox("in cm", { x: { cm: 1.23456 }, y: { cm: 2 }, width: { cm: 10 }, height: { cm: 2.5 } })
      .addTextBox([{ text: "in other units", italic: true }], { x: { pt: 72 }, y: { in: 1 }, width: { emu: 3600000 }, height: { pt: 36 } })
      .toDeck();
    const [first, second] = deck.slides[0].frames as TextFrame[];
    expect(first.role).toBeUndefined();
    expect(geometry(first)).toEqual([1.235, 2, 10, 2.5]);
    expect(first.paragraphs).toEqual([{ runs: [{ text: "in cm" }] }]);
    expect(geometry(second)).toEqual([2.54, 2.54, 10, 1.27]);
    expect(second.paragraphs).toEqual([{ runs: [{ text: "in other units", italic: true }] }]);
  });

  it("addTextBox names a bad rect field by path", () => {
    const builder = new PresentationBuilder().addBlankSlide().addTextBox("a", { x: { cm: 1 }, y: { cm: 1 }, width: { cm: 1 }, height: { cm: 1 } });
    expectGenerationError(() => builder.addTextBox("b", { x: 3, y: { cm: 1 }, width: { cm: 1 }, height: { cm: 1 } } as never), /^slides\[0\]\.textBoxes\[1\]\.x: expected a length/);
    expectGenerationError(() => builder.addTextBox("b", { x: { cm: 1 }, y: { cm: -1 }, width: { cm: 1 }, height: { cm: 1 } }), "slides[0].textBoxes[1].y.cm: expected a non-negative number");
    expectGenerationError(() => builder.addTextBox("b", { x: { cm: 1 }, y: { cm: 1 }, width: { cm: 0 }, height: { cm: 1 } }), "slides[0].textBoxes[1].width: expected a positive length");
    expectGenerationError(() => builder.addTextBox("b", { x: { cm: 1 }, y: { cm: 1 }, width: { cm: 1 }, height: { cm: 1 }, z: 1 } as never), "slides[0].textBoxes[1].z: unknown property");
    expect(builder.toDeck().slides[0].frames).toHaveLength(1);
  });

  it("addImage places an explicit rect, height from the aspect, clamped into the slide", () => {
    const deck = new PresentationBuilder()
      .addBlankSlide()
      .addImage({ data: PNG_4X2, contentType: "image/png", x: { cm: 1 }, y: { cm: 2 }, width: { cm: 6 } })
      .addImage({ data: PNG_4X2, contentType: "image/png", x: { cm: 30 }, y: { cm: 18 }, width: { cm: 8 }, height: { cm: 3 } })
      .addImage({ data: PNG_4X2, contentType: "image/png", x: { cm: 0 }, y: { cm: 0 }, width: { cm: 100 } })
      .toDeck();
    const [placed, moved, shrunk] = deck.slides[0].frames;
    expect(geometry(placed)).toEqual([1, 2, 6, 3]);
    // Moved fully onto the 33.867 × 19.05 cm slide.
    expect(geometry(moved)).toEqual([25.867, 16.05, 8, 3]);
    // 100 × 50 cm shrunk at its aspect to the slide's width.
    expect(geometry(shrunk)).toEqual([0, 0, 33.867, 16.934]);
    expect(deck.media!.map((part) => part.mediaId)).toEqual(["ppt/media/image1.png", "ppt/media/image2.png", "ppt/media/image3.png"]);
  });

  it("addImage names a bad field by path", () => {
    const builder = new PresentationBuilder().addBlankSlide();
    expectGenerationError(() => builder.addImage({ data: PNG_1X1, contentType: "image/png", x: { cm: 1 }, y: { cm: 1 } } as never), /^slides\[0\]\.images\[0\]\.width: expected a length/);
    expectGenerationError(() => builder.addImage({ data: 5, contentType: "image/png", x: { cm: 1 }, y: { cm: 1 }, width: { cm: 1 } } as never), "slides[0].images[0].data: expected base64 image bytes");
  });

  it("setNotes strips trailing empties, never stores [], and replaces earlier notes", () => {
    const builder = new PresentationBuilder().addTitleSlide("A").setNotes(["Lead with Europe.", "", "Then APAC.", "", ""]);
    expect(builder.toDeck().slides[0].notes).toEqual(["Lead with Europe.", "", "Then APAC."]);
    builder.setNotes(["first line\nsecond line"]);
    expect(builder.toDeck().slides[0].notes).toEqual(["first line", "second line"]);
    builder.setNotes(["", ""]);
    expect("notes" in builder.toDeck().slides[0]).toBe(false);
    builder.setNotes([]);
    expect("notes" in builder.toDeck().slides[0]).toBe(false);
    expectGenerationError(() => builder.setNotes(["ok", 2 as never]), "slides[0].notes[1]: expected a string");
    expectGenerationError(() => builder.setNotes("text" as never), "slides[0].notes: expected an array");
  });

  it("setBackground stores a lowercase solid fill", () => {
    const builder = new PresentationBuilder().addBlankSlide().setBackground("#AABBCC");
    expect(builder.toDeck().slides[0].background).toEqual({ type: "solid", color: "#aabbcc" });
    expectGenerationError(() => builder.setBackground("blue"), 'slides[0].background: expected a "#rrggbb" colour, got "blue"');
  });
});

describe("text fields take SPEC-045's RunSpec (AC-8)", () => {
  it("maps fontSizePt → sizePt and keeps bold/italic/underline/font/colour", () => {
    const [title] = textFrames(
      new PresentationBuilder().addTitleSlide([
        "plain ",
        { text: "styled", bold: true, italic: true, underline: true, fontFamily: "Georgia", fontSizePt: 40, color: "#FF0000" },
        { text: " not bold", bold: false },
      ]),
    );
    expect(title.paragraphs[0].runs).toEqual([
      { text: "plain " },
      { text: "styled", bold: true, italic: true, underline: true, fontFamily: "Georgia", sizePt: 40, color: "#ff0000" },
      { text: " not bold" },
    ]);
  });

  it("rejects the model-side sizePt as an unknown run key", () => {
    const builder = new PresentationBuilder();
    expectGenerationError(() => builder.addTitleSlide([{ text: "big", sizePt: 40 } as never]), "slides[0].title[0].sizePt: unknown property");
    expectGenerationError(() => builder.addTitleSlide("ok", [{ text: "x", fontSizePt: -1 }]), "slides[0].subtitle[0].fontSizePt: expected a positive number");
    expect(builder.toDeck().slides).toEqual([]);
  });

  it("drops strikethrough and highlight, which SlideRun has no field for, with a count", () => {
    const builder = new PresentationBuilder()
      .addBulletSlide("A", [{ runs: [{ text: "struck", strikethrough: true }, { text: "lit", highlight: "#ffff00" }] }, { runs: [{ text: "also", strikethrough: true }] }])
      .addTitleSlide([{ text: "fine", strikethrough: false }]);
    const [, body] = textFrames(builder);
    expect(body.paragraphs[0].runs).toEqual([{ text: "struck" }, { text: "lit" }]);
    expect(builder.dropped()).toEqual([
      { kind: "runFormatting", location: "slide 1", reason: "strikethrough dropped from 2 runs: a slide run has no strikethrough field" },
      { kind: "runFormatting", location: "slide 1", reason: "highlight dropped from 1 run: a slide run has no highlight field" },
    ]);
  });

  it("reports an image's alt text, which a slide image frame cannot keep", () => {
    const builder = new PresentationBuilder().addImageSlide("A", { data: PNG_1X1, contentType: "image/png", altText: "A red dot" });
    expect(builder.dropped()).toEqual([
      { kind: "media", location: "slide 1", reason: "image alt text not kept: a slide image frame has no alt-text field" },
    ]);
  });
});

describe("output (AC-9, AC-10)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function sample(format: "pptx" | "odp" = "pptx"): PresentationBuilder {
    return new PresentationBuilder({ title: "Q3 review", format })
      .addTitleSlide("Q3 review", "Prepared 30 September")
      .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%"]])
      .setNotes(["Lead with Europe."])
      .addImageSlide("Pipeline", { data: PNG_4X2, contentType: "image/png" })
      .setBackground("#336699");
  }

  it("toDeck is a pure, plain-JSON copy with no retained package", () => {
    const builder = sample();
    const deck = builder.toDeck();
    expect(JSON.parse(JSON.stringify(deck))).toEqual(deck);
    expect(deck.opaque).toBeUndefined();
    expect(deck.structureDirty).toBeUndefined();
    for (const slide of deck.slides) {
      expect(slide.partName).toBeUndefined();
      expect(slide.sourceXml).toBeUndefined();
      expect(slide.preserved).toBeUndefined();
      expect(slide.dirty).toBeUndefined();
    }
    deck.slides.pop();
    (deck.slides[0].frames[0] as TextFrame).paragraphs = [];
    expect(builder.toDeck().slides).toHaveLength(3);
    expect(builder.toDeck()).toEqual(builder.toDeck());
    expect(builder.toDeck()).not.toBe(builder.toDeck());
  });

  it("toBytes delegates to writePptx/writeOdp: the same parts, byte for byte", async () => {
    const builder = sample();
    for (const format of ["pptx", "odp"] as const) {
      const ours = await unzip(await builder.toBytes(format, { modified: MODIFIED }));
      const deck = { ...builder.toDeck(), format };
      const writer = format === "pptx" ? await writePptx(deck, { modified: MODIFIED }) : await writeOdp(deck, { modified: MODIFIED });
      const theirs = await unzip(writer.bytes);
      expect([...ours.keys()]).toEqual([...theirs.keys()]);
      for (const [name, data] of theirs) {
        expect(equalBytes(ours.get(name)!, data), name).toBe(true);
      }
    }
  });

  it("toBytes defaults to the deck's own format", async () => {
    expect((await unzip(await sample("pptx").toBytes())).has("ppt/presentation.xml")).toBe(true);
    const odp = await unzip(await sample("odp").toBytes());
    expect(new TextDecoder().decode(odp.get("mimetype"))).toBe("application/vnd.oasis.opendocument.presentation");
  });

  it("options.modified overrides the stamped timestamp", async () => {
    const parts = await unzip(await sample().toBytes("pptx", { modified: MODIFIED }));
    expect(new TextDecoder().decode(parts.get("docProps/core.xml"))).toContain("2026-09-30T12:00:00Z");
    const meta = await unzip(await sample().toBytes("odp", { modified: MODIFIED }));
    expect(new TextDecoder().decode(meta.get("meta.xml"))).toContain("2026-09-30T12:00:00");
  });

  it("validates format and options before writing", async () => {
    const builder = sample();
    await expect(builder.toBytes("docx" as never)).rejects.toThrow(/^format: expected one of "pptx", "odp", got "docx"/);
    await expect(builder.toBytes("pptx", { modified: "yesterday" } as never)).rejects.toThrow("options.modified: expected a valid Date");
    await expect(builder.toBytes("pptx", { now: new Date() } as never)).rejects.toThrow("options.now: unknown property");
  });

  it("toBlob carries the format's MIME type", async () => {
    expect((await sample().toBlob()).type).toBe("application/vnd.openxmlformats-officedocument.presentationml.presentation");
    expect((await sample().toBlob("odp")).type).toBe("application/vnd.oasis.opendocument.presentation");
  });

  it("toBuffer returns a Buffer, and throws a clear error where Buffer is not global", async () => {
    const buffer = await sample().toBuffer("odp", { modified: MODIFIED });
    expect(Buffer.isBuffer(buffer)).toBe(true);
    vi.stubGlobal("Buffer", undefined);
    await expect(sample().toBuffer()).rejects.toThrow(DocumentGenerationError);
    await expect(sample().toBuffer()).rejects.toThrow(/Buffer is not available in this environment/);
  });

  it("fileName delegates to pptxFileName/odpFileName", () => {
    expect(sample().fileName()).toBe("Q3 review.pptx");
    expect(sample().fileName("odp")).toBe("Q3 review.odp");
    expect(sample("odp").fileName()).toBe("Q3 review.odp");
    expect(new PresentationBuilder({ title: "  " }).fileName()).toBe("Untitled presentation.pptx");
    expect(() => sample().fileName("docx" as never)).toThrow(DocumentGenerationError);
  });
});

// AC-21: the deck path imports lib/presentation and lib/core only (plus its
// generate/ siblings) — never documentModel.ts or a .docx/.odt/.md layer —
// and slideMarkdown.ts reaches the text side through exactly one export and
// one type. Followed transitively over RUNTIME imports (`import type` is
// erased); the two products' own boundary specs cover the reverse direction.
function runtimeImports(file: string): string[] {
  const source = readFileSync(file, "utf8");
  const specifiers: string[] = [];
  for (const match of source.matchAll(/(?:^|\n)\s*(import|export)\s+([^;]*?)\s+from\s*"([^"]+)"/g)) {
    const clause = match[2];
    if (/^type\s/.test(clause)) {
      continue;
    }
    // `import { type A, b }` loads the module; `import { type A }` does not.
    const names = /^\{([^}]*)\}$/.exec(clause.trim());
    if (names && names[1].split(",").every((name) => name.trim() === "" || /^type\s/.test(name.trim()))) {
      continue;
    }
    specifiers.push(match[3]);
  }
  for (const match of source.matchAll(/import\(\s*"([^"]+)"\s*\)/g)) {
    specifiers.push(match[1]);
  }
  return specifiers.filter((s) => s.startsWith(".")).map((s) => resolve(dirname(file), `${s}.ts`));
}

describe("the deck path's product boundary (AC-21)", () => {
  const generateDir = dirname(fileURLToPath(import.meta.url));
  const libDir = resolve(generateDir, "..");

  function reachable(entry: string): Set<string> {
    const seen = new Set<string>();
    const queue = [resolve(generateDir, entry)];
    while (queue.length > 0) {
      const file = queue.pop()!;
      if (!seen.has(file)) {
        seen.add(file);
        queue.push(...runtimeImports(file));
      }
    }
    return seen;
  }

  const textSide = (file: string): boolean =>
    file === join(libDir, "documentModel.ts") || ["docx", "odf", "md"].some((dir) => file.startsWith(join(libDir, dir) + "/"));

  for (const entry of ["presentation.ts", "deckDescription.ts", "deckDrops.ts", "readDeck.ts", "deckHtml.ts"]) {
    it(`${entry} never reaches documentModel.ts or the .docx/.odt/.md layers`, () => {
      const reached = [...reachable(entry)];
      expect(reached.filter(textSide)).toEqual([]);
      expect(reached.some((file) => file.startsWith(join(libDir, "presentation") + "/"))).toBe(entry !== "deckDrops.ts");
    });
  }

  it("slideMarkdown.ts touches the text product through parseInlineRuns and the Run type only", () => {
    const source = readFileSync(join(generateDir, "slideMarkdown.ts"), "utf8");
    const textImports = [...source.matchAll(/import\s+([^;]*?)\s+from\s*"(\.\.\/(?:documentModel|md\/[^"]+|docx\/[^"]+|odf\/[^"]+))"/g)].map(
      (match) => `${match[1]} from ${match[2]}`,
    );
    expect(textImports.toSorted()).toEqual(["type { Run } from ../documentModel", "{ parseInlineRuns } from ../md/reader"]);
  });

  it("the two builders never import each other; index.ts re-exports both", () => {
    expect([...reachable("presentation.ts")]).not.toContain(join(generateDir, "builder.ts"));
    expect([...reachable("builder.ts")]).not.toContain(join(generateDir, "presentation.ts"));
    const index = readFileSync(join(generateDir, "index.ts"), "utf8");
    expect(index).toContain('export { DocumentBuilder } from "./builder";');
    expect(index).toContain('export { PresentationBuilder } from "./presentation";');
  });
});
