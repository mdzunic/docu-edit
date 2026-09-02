// SPEC-087: the importer against synthetic packages built by fixtures.ts —
// detection predicates, both container layouts, the import mapping (roles,
// styling, geometry scale, notes, images), drop counting, the rejection
// messages, cycle termination and the 50-slide stress graph. The contract
// against a REAL Keynote-written file lives in readerFixture.spec.ts; these
// synthetic cases prove the decoder's mechanics, not Apple's bytes.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import {
  isIwaPackage,
  isLegacyKeynote,
  KeynoteOpenError,
  keynoteImportNotice,
  KEYNOTE_UNREADABLE_MESSAGE,
  NOT_KEYNOTE_MESSAGE,
  parseKeynote,
  type ParsedKeynote,
} from "./reader";
import {
  buildKeynoteFromArchives,
  buildStressKeynote,
  buildSyntheticKeynote,
  FIXTURE_PNG_BASE64,
  float32Field,
  iwaFileOf,
  lengthField,
  message,
  reference,
  stringField,
  SYNTHETIC_IMAGE_GEOMETRY,
  varintField,
  type ArchiveInput,
} from "./fixtures";
import type { ImageFrame, TextFrame } from "../model";

async function parseBuilt(bytes: Uint8Array): Promise<ParsedKeynote> {
  return parseKeynote(await unzip(bytes));
}

function textOf(frame: TextFrame): string[] {
  return frame.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(""));
}

const POINT_TO_CM = 33.867 / 1920;

describe("detection predicates", () => {
  it("recognises the flat layout by Index/Document.iwa (AC-5)", async () => {
    expect(isIwaPackage(await unzip(await buildSyntheticKeynote()))).toBe(true);
  });

  it("recognises the nested layout by Index.zip (AC-6)", async () => {
    expect(isIwaPackage(await unzip(await buildSyntheticKeynote({ nested: true })))).toBe(true);
  });

  it("does not match an ordinary ZIP", () => {
    expect(isIwaPackage(new Map([["word/document.xml", new Uint8Array(1)]]))).toBe(false);
  });

  it("recognises iWork '05–'09 packages (AC-8)", () => {
    expect(isLegacyKeynote(new Map([["index.apxl", new Uint8Array(1)]]))).toBe(true);
    expect(isLegacyKeynote(new Map([["index.apxl.gz", new Uint8Array(1)]]))).toBe(true);
    expect(isLegacyKeynote(new Map([["Index/Document.iwa", new Uint8Array(1)]]))).toBe(false);
  });
});

describe("parseKeynote — the synthetic deck's content", () => {
  it("imports two slides with the contract's text, roles and styling", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote());
    expect(parsed.slides).toHaveLength(2);
    expect(parsed.dropped).toEqual({ objects: 0, effects: 0 });

    const [slide1, slide2] = parsed.slides;
    const title1 = slide1.frames[0] as TextFrame;
    expect(title1.kind).toBe("text");
    expect(title1.role).toBe("title");
    expect(textOf(title1)).toEqual(["Keynote fixture deck"]);
    // Title styling: 36pt bold #000000, no bullet (AC-26/AC-29).
    expect(title1.paragraphs[0].runs[0]).toMatchObject({ sizePt: 36, bold: true, color: "#000000" });
    expect(title1.paragraphs[0].bullet).toBeUndefined();

    const body1 = slide1.frames[1] as TextFrame;
    expect(body1.role).toBe("body");
    expect(textOf(body1)).toEqual(["Imported by docu-edit"]);

    const title2 = slide2.frames[0] as TextFrame;
    expect(textOf(title2)).toEqual(["Second slide"]);
    const body2 = slide2.frames[1] as TextFrame;
    // Body: split on \n, 18pt, disc bullets (AC-24/AC-26/AC-28).
    expect(textOf(body2)).toEqual(["First bullet point", "Second bullet point"]);
    for (const paragraph of body2.paragraphs) {
      expect(paragraph.bullet).toBe("disc");
      expect(paragraph.runs[0].sizePt).toBe(18);
    }
  });

  it("imports the same contract from the nested Index.zip layout (AC-6)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote({ nested: true }));
    expect(parsed.slides).toHaveLength(2);
    expect(textOf(parsed.slides[0].frames[0] as TextFrame)).toEqual(["Keynote fixture deck"]);
    expect(parsed.slides[1].frames.some((frame) => frame.kind === "image")).toBe(true);
  });

  it("scales 16:9 Keynote points to 33.867×19.05cm and frame geometry within 0.05cm (AC-22/AC-23)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote());
    expect(parsed.slideWidthCm).toBeCloseTo(33.867, 3);
    expect(parsed.slideHeightCm).toBeCloseTo(19.05, 2);

    const image = parsed.slides[1].frames.find((frame) => frame.kind === "image") as ImageFrame;
    expect(Math.abs(image.xCm - SYNTHETIC_IMAGE_GEOMETRY.x * POINT_TO_CM)).toBeLessThan(0.05);
    expect(Math.abs(image.yCm - SYNTHETIC_IMAGE_GEOMETRY.y * POINT_TO_CM)).toBeLessThan(0.05);
    expect(Math.abs(image.widthCm - SYNTHETIC_IMAGE_GEOMETRY.w * POINT_TO_CM)).toBeLessThan(0.05);
    expect(Math.abs(image.heightCm - SYNTHETIC_IMAGE_GEOMETRY.h * POINT_TO_CM)).toBeLessThan(0.05);
  });

  it("scales a 4:3 deck to 25.4cm wide", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote({ size: { w: 1024, h: 768 } }));
    expect(parsed.slideWidthCm).toBeCloseTo(25.4, 3);
    expect(parsed.slideHeightCm).toBeCloseTo(19.05, 2);
  });

  it("imports the image with SPEC-082's minted id and the verbatim Data/ bytes (AC-32..AC-34)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote());
    const image = parsed.slides[1].frames.find((frame) => frame.kind === "image") as ImageFrame;
    expect(image.mediaId).toBe("ppt/media/image1.png");
    expect(parsed.media).toHaveLength(1);
    expect(parsed.media[0]).toMatchObject({ mediaId: "ppt/media/image1.png", contentType: "image/png" });
    expect(parsed.media[0].data).toBe(FIXTURE_PNG_BASE64);
  });

  it("imports speaker notes into Slide.notes and leaves note-less slides without the field (AC-30/AC-31)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote());
    expect(parsed.slides[0].notes).toBeUndefined();
    expect(parsed.slides[1].notes).toEqual(["A speaker note for the importer"]);
  });

  it("returns a from-scratch model: no opaque, no retention fields (AC-14/AC-15)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote());
    expect(parsed.opaque).toEqual([]);
    for (const slide of parsed.slides) {
      expect(slide.partName).toBeUndefined();
      expect(slide.sourceXml).toBeUndefined();
    }
  });
});

describe("parseKeynote — drops and the import notice", () => {
  it("drops and counts a drawable that is neither text nor image (AC-36)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote({ withForeignDrawable: true }));
    expect(parsed.dropped).toEqual({ objects: 1, effects: 0 });
  });

  it("counts a slide transition as a dropped effect (AC-36)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote({ withTransition: true }));
    expect(parsed.dropped).toEqual({ objects: 0, effects: 1 });
  });

  it("drops and counts an image whose data cannot be resolved (AC-35)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote({ unresolvableImage: true }));
    expect(parsed.dropped).toEqual({ objects: 1, effects: 0 });
    expect(parsed.media).toEqual([]);
    expect(parsed.slides[1].frames.some((frame) => frame.kind === "image")).toBe(false);
  });

  it("words the notice with plural, singular and zero forms (AC-37/AC-38)", () => {
    expect(keynoteImportNotice({ objects: 0, effects: 0 })).toBeNull();
    expect(keynoteImportNotice({ objects: 2, effects: 3 })).toBe(
      "Imported from Keynote. 2 objects and 3 effects couldn't be imported.",
    );
    expect(keynoteImportNotice({ objects: 1, effects: 0 })).toBe(
      "Imported from Keynote. 1 object couldn't be imported.",
    );
    expect(keynoteImportNotice({ objects: 0, effects: 1 })).toBe(
      "Imported from Keynote. 1 effect couldn't be imported.",
    );
  });
});

describe("parseKeynote — text handling on a bespoke graph", () => {
  /** A one-slide deck whose single non-placeholder drawable holds `text`. */
  function deckWithText(text: string): ArchiveInput[] {
    return [
      { id: 1, type: 1, payload: message(reference(2, 3)) },
      {
        id: 3,
        type: 2,
        payload: message(
          lengthField(3, message(reference(2, 10))),
          lengthField(4, message(float32Field(1, 1920), float32Field(2, 1080))),
        ),
      },
      { id: 10, type: 4, payload: message(reference(1, 11)) },
      { id: 11, type: 4, payload: message(reference(2, 21)) },
      { id: 21, type: 5, payload: message(lengthField(7, varintField(1, 31))) },
      { id: 31, type: 7, payload: message(lengthField(1, message(reference(2, 41)))) },
      { id: 41, type: 2001, payload: message(stringField(3, text)) },
    ];
  }

  it("splits paragraphs on \\n and U+2029 (AC-24)", async () => {
    const parsed = await parseBuilt(await buildKeynoteFromArchives(deckWithText("one\ntwo three")));
    const frame = parsed.slides[0].frames[0] as TextFrame;
    expect(textOf(frame)).toEqual(["one", "two", "three"]);
  });

  it("drops U+FFFC and zero-width characters (AC-25)", async () => {
    const parsed = await parseBuilt(
      await buildKeynoteFromArchives(deckWithText("a￼b​c‍d﻿e")),
    );
    expect(textOf(parsed.slides[0].frames[0] as TextFrame)).toEqual(["abcde"]);
  });

  it("gives a non-placeholder text drawable the 'other' role at 20pt with no bullet (AC-26)", async () => {
    const parsed = await parseBuilt(await buildKeynoteFromArchives(deckWithText("floating box")));
    const frame = parsed.slides[0].frames[0] as TextFrame;
    expect(frame.role).toBe("other");
    expect(frame.paragraphs[0].runs[0].sizePt).toBe(20);
    expect(frame.paragraphs[0].bullet).toBeUndefined();
  });

  it("gives a drawable without geometry the role's default band", async () => {
    const parsed = await parseBuilt(await buildKeynoteFromArchives(deckWithText("no geometry here")));
    const frame = parsed.slides[0].frames[0] as TextFrame;
    // other → the centred half-width 4:3 box.
    expect(frame.widthCm).toBeCloseTo(33.867 / 2, 2);
    expect(frame.heightCm).toBeCloseTo((33.867 / 2) * 0.75, 2);
    expect(frame.xCm).toBeCloseTo(33.867 / 4, 2);
  });

  it("imports an empty text storage as nothing, uncounted", async () => {
    const parsed = await parseBuilt(await buildKeynoteFromArchives(deckWithText("​")));
    expect(parsed.slides[0].frames).toEqual([]);
    expect(parsed.dropped).toEqual({ objects: 0, effects: 0 });
  });
});

describe("parseKeynote — rejections (AC-9..AC-12, AC-42)", () => {
  it("rejects a package whose root document archive is not Keynote's (AC-9)", async () => {
    const foreign: ArchiveInput[] = [{ id: 1, type: 10000, payload: message(varintField(1, 1)) }];
    const bytes = await buildKeynoteFromArchives(foreign);
    await expect(parseBuilt(bytes)).rejects.toThrow(NOT_KEYNOTE_MESSAGE);
    await expect(parseBuilt(bytes)).rejects.toBeInstanceOf(KeynoteOpenError);
  });

  it("rejects an undecodable object graph (AC-10)", async () => {
    const { zip } = await import("../../core/zipWriter");
    const bytes = await zip([{ name: "Index/Document.iwa", data: Uint8Array.from([0x42, 0x42, 0x42]) }]);
    await expect(parseBuilt(bytes)).rejects.toThrow(NOT_KEYNOTE_MESSAGE);
  });

  it("rejects a truncated IWA file with the typed error (AC-42)", async () => {
    const { document, metadata } = await import("./fixtures").then((f) => f.syntheticDeckArchives());
    const whole = iwaFileOf(document);
    const { zip } = await import("../../core/zipWriter");
    const bytes = await zip([
      { name: "Index/Document.iwa", data: whole.subarray(0, whole.length - 10) },
      { name: "Index/Metadata.iwa", data: iwaFileOf(metadata) },
    ]);
    const rejection = expect(parseBuilt(bytes)).rejects;
    await rejection.toBeInstanceOf(KeynoteOpenError);
  });

  it("rejects a deck resolving to zero slides as unreadable (AC-11)", async () => {
    const empty: ArchiveInput[] = [
      { id: 1, type: 1, payload: message(reference(2, 3)) },
      {
        id: 3,
        type: 2,
        payload: message(
          lengthField(3, message(reference(2, 10))),
          lengthField(4, message(float32Field(1, 1920), float32Field(2, 1080))),
        ),
      },
      { id: 10, type: 4, payload: message() },
    ];
    await expect(parseBuilt(await buildKeynoteFromArchives(empty))).rejects.toThrow(KEYNOTE_UNREADABLE_MESSAGE);
  });
});

describe("parseKeynote — defensive walks (AC-41, AC-43, AC-44)", () => {
  it("terminates on a reference cycle in the slide tree (AC-43)", async () => {
    const parsed = await parseBuilt(await buildSyntheticKeynote({ withCycle: true }));
    expect(parsed.slides).toHaveLength(2);
  });

  it("decodes a 50-slide stress graph iteratively (AC-44)", async () => {
    const parsed = await parseBuilt(await buildStressKeynote(50));
    expect(parsed.slides).toHaveLength(50);
    expect(textOf(parsed.slides[49].frames[0] as TextFrame)).toEqual(["Slide 50"]);
  });
});
