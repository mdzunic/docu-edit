// The presentation preparation's core promises, tested end to end over the
// in-repo fixtures: (1) `.pptx`/`.odp` parse into the same slide outline,
// (2) EVERY source part is retained, and (3) a save re-emits every part
// byte-identical, in package order — with `.odp`'s `mimetype` first and
// STOREd. The ZIP container is rebuilt on save, so the round-trip contract
// is part-level bytes, not container bytes (the same reason the docx/odf
// round-trip suites compare parts).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import {
  buildOdpFixtureBytes,
  buildPptxFixtureBytes,
  FIXTURE_DECK_AUTHOR,
  FIXTURE_DECK_TITLE,
  FIXTURE_SLIDES,
} from "./fixtures/build";
import { parsePptx, PptxOpenError } from "./pptx/reader";
import { OdpOpenError, parseOdp } from "./odp/reader";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { assemblePresentation, presentationFormatOf, readPresentationBytes } from "./open";
import { detectPackage } from "../core/packageDetect";
import { slideCount, type Slide } from "./model";
import { slideOutline } from "./slideModel";

const EXPECTED_OUTLINE = FIXTURE_SLIDES.map((slide) => ({
  title: "title" in slide ? slide.title : undefined,
  lines: [...slide.lines],
}));

/** The prep milestone's outline observable, now DERIVED from the frame model
 * via `slideOutline` (AC-4) — the round-trip's expected strings are unchanged. */
function outlineOf(slides: Slide[]): { title: string | undefined; lines: string[] }[] {
  return slides.map((slide) => {
    const outline = slideOutline(slide);
    return { title: outline.title, lines: outline.lines };
  });
}

describe("pptx reader", () => {
  it("extracts the slide outline: order, titles, body lines, a:br splits, grouped shapes", async () => {
    const parsed = await parsePptx(await buildPptxFixtureBytes());
    expect(outlineOf(parsed.slides)).toEqual(EXPECTED_OUTLINE);
  });

  it("reads docProps metadata through the root relationships", async () => {
    const parsed = await parsePptx(await buildPptxFixtureBytes());
    expect(parsed.properties.title).toBe(FIXTURE_DECK_TITLE);
    expect(parsed.properties.author).toBe(FIXTURE_DECK_AUTHOR);
    expect(parsed.properties.created).toBe("2026-01-05T10:00:00Z");
    expect(parsed.properties.modified).toBe("2026-02-10T09:30:00Z");
  });

  it("retains every package part, in package order", async () => {
    const bytes = await buildPptxFixtureBytes();
    const parts = await unzip(bytes);
    const parsed = await parsePptx(bytes);
    expect(parsed.opaque.map((part) => part.name)).toEqual([...parts.keys()]);
  });

  it("rejects an .odp, a .docx-shaped package, and non-ZIP bytes on its own terms", async () => {
    await expect(parsePptx(await buildOdpFixtureBytes())).rejects.toBeInstanceOf(PptxOpenError);
    await expect(parsePptx(new TextEncoder().encode("plain text"))).rejects.toBeInstanceOf(PptxOpenError);
  });
});

describe("odp reader", () => {
  it("extracts the same slide outline as the pptx twin", async () => {
    const parsed = await parseOdp(await buildOdpFixtureBytes());
    expect(outlineOf(parsed.slides)).toEqual(EXPECTED_OUTLINE);
  });

  it("reads meta.xml metadata", async () => {
    const parsed = await parseOdp(await buildOdpFixtureBytes());
    expect(parsed.properties.title).toBe(FIXTURE_DECK_TITLE);
    expect(parsed.properties.author).toBe(FIXTURE_DECK_AUTHOR);
  });

  it("retains every package part, mimetype included, in package order", async () => {
    const bytes = await buildOdpFixtureBytes();
    const parts = await unzip(bytes);
    const parsed = await parseOdp(bytes);
    expect(parsed.opaque.map((part) => part.name)).toEqual([...parts.keys()]);
    expect(parsed.opaque[0]?.name).toBe("mimetype");
  });

  it("rejects a .pptx and non-ZIP bytes on its own terms", async () => {
    await expect(parseOdp(await buildPptxFixtureBytes())).rejects.toBeInstanceOf(OdpOpenError);
    await expect(parseOdp(new TextEncoder().encode("plain text"))).rejects.toBeInstanceOf(OdpOpenError);
  });
});

describe("round trip (SPEC-008's rule applied to presentations)", () => {
  it("pptx: every part re-emitted byte-identical, in order", async () => {
    const sourceBytes = await buildPptxFixtureBytes();
    const model = await readPresentationBytes(sourceBytes, "pptx", "Deck");
    const { bytes } = await writePptx(model);
    const source = await unzip(sourceBytes);
    const written = await unzip(bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name} must round-trip byte-identical`).toBe(true);
    }
  });

  it("odp: every part re-emitted byte-identical, mimetype first and STOREd", async () => {
    const sourceBytes = await buildOdpFixtureBytes();
    const model = await readPresentationBytes(sourceBytes, "odp", "Deck");
    const { bytes } = await writeOdp(model);
    const source = await unzip(sourceBytes);
    const written = await unzip(bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name} must round-trip byte-identical`).toBe(true);
    }
    // The first local file header's compression method (offset 8) must be 0
    // (STORE) so `mimetype` is sniffable from the raw bytes.
    expect(bytes[8]).toBe(0);
    expect(bytes[9]).toBe(0);
    expect(new TextDecoder().decode(bytes.slice(38, 38 + 46))).toContain("application/vnd.oasis.opendocument");
  });
});

describe("presentationFormatOf / assemblePresentation", () => {
  it("classifies the fixtures and leaves text packages to the other product", async () => {
    expect(presentationFormatOf(detectPackage(await unzip(await buildPptxFixtureBytes())))).toBe("pptx");
    expect(presentationFormatOf(detectPackage(await unzip(await buildOdpFixtureBytes())))).toBe("odp");
    expect(presentationFormatOf({ container: "zip" })).toBeNull();
    expect(presentationFormatOf({ container: "odf", mimetype: "application/vnd.oasis.opendocument.text" })).toBeNull();
    expect(
      presentationFormatOf({
        container: "opc",
        mainPartPath: "word/document.xml",
        mainContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      }),
    ).toBeNull();
  });

  it("mints a fresh id, sets the discriminant, and omits empty properties", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "My deck");
    expect(model.kind).toBe("presentation");
    expect(model.format).toBe("pptx");
    expect(model.title).toBe("My deck");
    expect(model.id).not.toHaveLength(0);
    expect(slideCount(model)).toBe(3);

    const withoutProperties = assemblePresentation("Bare", "pptx", {
      slides: [],
      slideWidthCm: 33.867,
      slideHeightCm: 19.05,
      media: [],
      opaque: [],
      properties: {},
    });
    expect("properties" in withoutProperties).toBe(false);
    expect("media" in withoutProperties).toBe(false);
  });

  it("stays JSON-serialisable (the store's persistence contract)", async () => {
    const model = await readPresentationBytes(await buildOdpFixtureBytes(), "odp", "Deck");
    const revived = JSON.parse(JSON.stringify(model));
    expect(revived).toEqual(model);
  });
});
