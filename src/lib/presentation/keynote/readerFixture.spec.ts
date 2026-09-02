// SPEC-087 AC-50/AC-51: the content contract against the REAL Keynote file
// e2e/fixtures/mini-deck.key, read via Node fs. That fixture can only be
// produced by a human running scripts/make-keynote-fixture.applescript on a
// Mac with Keynote (CI cannot run Keynote; a synthesised .key would only
// test our own assumptions — reader.spec.ts already does that). While the
// fixture is absent this suite FAILS LOUDLY by design (AC-51): a vacuously
// green run would let the three format facts the spec marked "confirm
// against the fixture" (the notes ref at slide field 27, the slide-node
// child list, image recognition by data reference) go unconfirmed forever.
import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { parseKeynote, type ParsedKeynote } from "./reader";
import { titleFromFileName } from "../../openDocument";
import type { ImageFrame, TextFrame } from "../model";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PATH = join(here, "..", "..", "..", "..", "e2e", "fixtures", "mini-deck.key");
const IMAGE_PATH = join(here, "..", "..", "..", "..", "e2e", "fixtures", "fixture-image.png");

export const FIXTURE_MISSING =
  "e2e/fixtures/mini-deck.key is missing — run scripts/make-keynote-fixture.applescript on a Mac with Keynote and commit the result (SPEC-087's prerequisite). The Keynote-import content contract cannot be verified without it.";

const fixturePresent = existsSync(FIXTURE_PATH);

it("the Keynote fixture is present (AC-51)", () => {
  expect(fixturePresent, FIXTURE_MISSING).toBe(true);
});

async function parseFixture(): Promise<ParsedKeynote> {
  const bytes = new Uint8Array(readFileSync(FIXTURE_PATH));
  return parseKeynote(await unzip(bytes));
}

function frameTexts(frame: TextFrame): string[] {
  return frame.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(""));
}

/** Keynote points → cm for the fixture's 1920×1080 show (AC-22's rule). */
const POINT_TO_CM = 33.867 / 1920;

describe.runIf(fixturePresent)("the mini-deck.key content contract (AC-13..AC-23, AC-33/AC-34, AC-50)", () => {
  it("imports two slides into a from-scratch pptx-format model (AC-13..AC-17)", async () => {
    const parsed = await parseFixture();
    expect(parsed.slides).toHaveLength(2);
    expect(parsed.opaque).toEqual([]);
    for (const slide of parsed.slides) {
      expect(slide.partName).toBeUndefined();
      expect(slide.sourceXml).toBeUndefined();
    }
    // The funnel opens this as format "pptx" titled from the file name.
    expect(titleFromFileName("mini-deck.key")).toBe("mini-deck");
  });

  it("slide 1 carries the title and the docu-edit text frame (AC-18)", async () => {
    const parsed = await parseFixture();
    const textFrames = parsed.slides[0].frames.filter((frame): frame is TextFrame => frame.kind === "text");
    const title = textFrames.find((frame) => frame.role === "title");
    expect(title, "slide 1 should have a title placeholder frame").toBeDefined();
    expect(frameTexts(title as TextFrame)).toEqual(["Keynote fixture deck"]);
    expect(
      textFrames.some((frame) => frameTexts(frame).some((text) => text.includes("Imported by docu-edit"))),
    ).toBe(true);
  });

  it("slide 2 carries the title and the two disc-bulleted body paragraphs (AC-19, AC-28/AC-29)", async () => {
    const parsed = await parseFixture();
    const textFrames = parsed.slides[1].frames.filter((frame): frame is TextFrame => frame.kind === "text");
    const title = textFrames.find((frame) => frame.role === "title");
    expect(frameTexts(title as TextFrame)).toEqual(["Second slide"]);
    expect(title?.paragraphs[0].bullet).toBeUndefined();

    const body = textFrames.find((frame) => frame.role === "body");
    expect(body, "slide 2 should have a body placeholder frame").toBeDefined();
    expect(frameTexts(body as TextFrame)).toEqual(["First bullet point", "Second bullet point"]);
    for (const paragraph of (body as TextFrame).paragraphs) {
      expect(paragraph.bullet).toBe("disc");
    }
  });

  it("slide 2's image frame carries the fixture PNG verbatim under SPEC-082's minted id (AC-20, AC-32..AC-34)", async () => {
    const parsed = await parseFixture();
    const images = parsed.slides[1].frames.filter((frame): frame is ImageFrame => frame.kind === "image");
    expect(images).toHaveLength(1);
    expect(images[0].mediaId).toMatch(/^ppt\/media\/image\d+\.(png|jpg|jpeg|gif)$/);

    const part = parsed.media.find((media) => media.mediaId === images[0].mediaId);
    expect(part).toBeDefined();
    const expected = Buffer.from(readFileSync(IMAGE_PATH)).toString("base64");
    expect(part?.data).toBe(expected);
    expect(part?.contentType).toBe("image/png");
  });

  it("slide 2 carries the speaker note; slide 1 has none (AC-21, AC-30/AC-31)", async () => {
    const parsed = await parseFixture();
    expect(parsed.slides[0].notes).toBeUndefined();
    expect(parsed.slides[1].notes).toEqual(["A speaker note for the importer"]);
  });

  it("the 16:9 deck scales to 33.867×19.05cm and the image lands within 0.05cm of its known geometry (AC-22/AC-23)", async () => {
    const parsed = await parseFixture();
    expect(parsed.slideWidthCm).toBeCloseTo(33.867, 3);
    expect(parsed.slideHeightCm).toBeCloseTo(19.05, 2);

    // The AppleScript places the image at {1300, 620} with width 300pt; the
    // source PNG is square, so height ≈ 300pt too.
    const image = parsed.slides[1].frames.find((frame) => frame.kind === "image") as ImageFrame;
    expect(Math.abs(image.xCm - 1300 * POINT_TO_CM)).toBeLessThan(0.05);
    expect(Math.abs(image.yCm - 620 * POINT_TO_CM)).toBeLessThan(0.05);
    expect(Math.abs(image.widthCm - 300 * POINT_TO_CM)).toBeLessThan(0.05);
  });
});
