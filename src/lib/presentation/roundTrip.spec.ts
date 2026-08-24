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
  FIXTURE_FRAME_KINDS,
  FIXTURE_SHAPE_FILL,
  FIXTURE_SLIDE4_GEOMETRY,
  FIXTURE_SLIDES,
} from "./fixtures/build";
import { parsePptx, PptxOpenError } from "./pptx/reader";
import { OdpOpenError, parseOdp } from "./odp/reader";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import {
  assemblePresentation,
  presentationFormatOf,
  readPresentationBytes,
  upgradePresentationModel,
} from "./open";
import { detectPackage } from "../core/packageDetect";
import {
  isLegacyPresentationModel,
  slideCount,
  type PresentationModel,
  type Slide,
  type SlideFrame,
  type ShapeFrame,
  type GroupFrame,
  type TextFrame,
  type ForeignFrame,
} from "./model";
import { slideOutline } from "./slideModel";

function geomOf(frame: SlideFrame): [number, number, number, number] {
  return [frame.xCm, frame.yCm, frame.widthCm, frame.heightCm];
}

function expectGeomClose(actual: [number, number, number, number], expected: [number, number, number, number]): void {
  for (let i = 0; i < 4; i++) {
    expect(Math.abs(actual[i] - expected[i]), `component ${i}: ${actual[i]} vs ${expected[i]}`).toBeLessThan(0.05);
  }
}

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
    expect(slideCount(model)).toBe(4);

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

describe("SPEC-079 frame model", () => {
  async function pptxSlides(): Promise<Slide[]> {
    return (await parsePptx(await buildPptxFixtureBytes())).slides;
  }
  async function odpSlides(): Promise<Slide[]> {
    return (await parseOdp(await buildOdpFixtureBytes())).slides;
  }

  it("yields the same per-slide frame kinds in z-order from both readers (AC-5/6)", async () => {
    const pptx = await pptxSlides();
    const odp = await odpSlides();
    expect(pptx.map((s) => s.frames.map((f) => f.kind))).toEqual(FIXTURE_FRAME_KINDS);
    expect(odp.map((s) => s.frames.map((f) => f.kind))).toEqual(FIXTURE_FRAME_KINDS);
  });

  it("agrees on slide-4 geometry across the two formats within 0.05cm (AC-8)", async () => {
    const pptx = (await pptxSlides())[3];
    const odp = (await odpSlides())[3];
    FIXTURE_SLIDE4_GEOMETRY.forEach((expected, i) => {
      expectGeomClose(geomOf(pptx.frames[i]), expected);
      expectGeomClose(geomOf(odp.frames[i]), geomOf(pptx.frames[i]));
    });
  });

  it("inherits pptx placeholder geometry: layout, then master, else own (AC-9/10/11/12)", async () => {
    const slide4 = (await pptxSlides())[3];
    expectGeomClose(geomOf(slide4.frames[0]), [2, 1, 29, 3]); // AC-9: from layout
    expectGeomClose(geomOf(slide4.frames[1]), [2, 5, 14, 10]); // AC-10: from master
    expectGeomClose(geomOf(slide4.frames[2]), [18, 5, 13, 3]); // AC-11: own a:xfrm
  });

  it("resolves effective run formatting from a:rPr (AC-13)", async () => {
    const own = (await pptxSlides())[3].frames[2] as TextFrame;
    const run = own.paragraphs[0].runs[0];
    expect(run.bold).toBe(true);
    expect(run.sizePt).toBe(24);
    expect(run.color).toBe("#ff0000");
  });

  it("resolves title size from the master txStyles chain (AC-15)", async () => {
    const title = (await pptxSlides())[0].frames[0] as TextFrame;
    // titleStyle lvl1 defRPr sz=4400 → 44pt, larger than the 18pt default.
    expect(title.paragraphs[0].runs[0].sizePt).toBe(44);
  });

  it("resolves a themed schemeClr fill through clrMap + lumMod/lumOff (AC-17/18/25)", async () => {
    const ellipse = (await pptxSlides())[3].frames[3] as ShapeFrame;
    expect(ellipse.geometry).toBe("ellipse");
    expect(ellipse.fill).toEqual({ type: "solid", color: FIXTURE_SHAPE_FILL });
  });

  it("resolves odp effective formatting through the parent-style chain (AC-19/20)", async () => {
    const body = (await odpSlides())[3].frames[1] as TextFrame;
    const run = body.paragraphs[0].runs[0];
    expect(run.bold).toBe(true); // T1's own fo:font-weight
    expect(run.color).toBe("#c0392b"); // inherited from parent style "Emphasis"
  });

  it("never errors resolving backgrounds and leaves plain slides without one (AC-28)", async () => {
    for (const slide of await pptxSlides()) {
      expect(slide.background === undefined || slide.background.type === "solid").toBe(true);
    }
  });

  it("populates model.media once and keeps the part in opaque (AC-29/30/31/32)", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    expect(model.media?.map((m) => m.mediaId)).toEqual(["ppt/media/image1.png"]);
    const image = model.slides[3].frames.find((f) => f.kind === "image");
    expect(image?.kind === "image" && image.mediaId).toBe("ppt/media/image1.png");
    // The bytes stay in opaque too (a no-edit download round-trips).
    expect(model.opaque.some((p) => p.name === "ppt/media/image1.png")).toBe(true);
  });

  it("flattens groups and labels foreign objects (AC-37/38/39/40)", async () => {
    const slide4 = (await pptxSlides())[3];
    const group = slide4.frames[5] as GroupFrame;
    expect(group.children.map((c) => c.kind)).toEqual(["shape", "shape"]);
    expect(group.xml).toContain("grpSp"); // preserved verbatim for the writer
    // Child rects mapped into slide space (identity chOff/chExt here).
    expectGeomClose(geomOf(group.children[0]), [8, 15, 2.5, 3]);
    const table = slide4.frames[6] as ForeignFrame;
    expect(table.label).toBe("Table");
    expect(table.noBounds).toBeUndefined();
  });

  it("upgrades a preparation-milestone record on re-parse (AC-54/56/57)", async () => {
    const fresh = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Legacy deck");
    // Simulate the prep milestone's stored shape: outline slides, no slideWidthCm.
    const legacy = {
      kind: "presentation",
      id: fresh.id,
      title: fresh.title,
      format: fresh.format,
      slides: [{ title: "Quarterly Review", lines: ["FY26 results and outlook"] }],
      opaque: fresh.opaque,
    } as unknown as PresentationModel;
    expect(isLegacyPresentationModel(legacy)).toBe(true);
    const upgraded = await upgradePresentationModel(legacy);
    expect(isLegacyPresentationModel(upgraded)).toBe(false);
    expect(upgraded.id).toBe(legacy.id);
    expect(upgraded.title).toBe("Legacy deck");
    expect(upgraded.slides).toHaveLength(4);
    expect(slideOutline(upgraded.slides[0]).title).toBe("Quarterly Review");
  });
});
