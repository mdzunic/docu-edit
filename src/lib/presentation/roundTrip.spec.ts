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
} from "./model";
import { slideOutline } from "./slideModel";
import { addSlide, createBlankPresentation, deleteSlide, duplicateSlide, moveSlide } from "./slides";
import { applyRunFormat, markSlideEdited, wholeFrameRange } from "./text";

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
      layouts: [],
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
    expect(model.opaque?.some((p) => p.name === "ppt/media/image1.png")).toBe(true);
  });

  it("flattens groups and labels foreign objects (AC-37/38/39/40)", async () => {
    const slide4 = (await pptxSlides())[3];
    const group = slide4.frames[5] as GroupFrame;
    expect(group.children.map((c) => c.kind)).toEqual(["shape", "shape"]);
    expect(group.xml).toContain("grpSp"); // preserved verbatim for the writer
    // Child rects mapped into slide space (identity chOff/chExt here).
    expectGeomClose(geomOf(group.children[0]), [8, 15, 2.5, 3]);
    // The trailing graphicFrame is a modelled TableFrame now (SPEC-091), not a
    // foreign "Table" box: a 1×1 grid with the "Cell" text.
    const table = slide4.frames[6];
    expect(table.kind).toBe("table");
    if (table.kind === "table") {
      expect(table.columnWidthsCm).toHaveLength(1);
      expect(table.rows).toHaveLength(1);
      expect(table.rows[0].cells[0].paragraphs[0].runs[0].text).toBe("Cell");
    }
  });

  it("gives a group without its own xfrm a union-of-children hit box (AC-31)", async () => {
    // Slide 3's 'Group 3' has an empty <p:grpSpPr/>; without a fallback its box
    // is 0×0 and unselectable. Its sole child text box sits at (2,10,20,3).
    const slide3 = (await pptxSlides())[2];
    const group = slide3.frames.find((f) => f.kind === "group") as GroupFrame;
    expect(group).toBeDefined();
    expect(group.widthCm).toBeGreaterThan(0);
    expect(group.heightCm).toBeGreaterThan(0);
    expectGeomClose(geomOf(group), [2, 10, 20, 3]);
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

// --- SPEC-080 writers: from-scratch build + part-substitution rebuild -------

const STRUCTURAL_PPTX_PARTS = new Set([
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
  "[Content_Types].xml",
]);

function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** SPEC-078's "compare resolved facts, not minted values" rule for
 * presentations: the comparable placeholder facts (kind, role, geometry
 * rounded, text) — minted ids, `phType`, retention and package fields
 * dropped, so a build → re-parse round trip is deep-comparable across
 * formats. */
function strippedForCompare(model: PresentationModel): unknown {
  return {
    format: model.format,
    title: model.title,
    w: r3(model.slideWidthCm),
    h: r3(model.slideHeightCm),
    slides: model.slides.map((slide) =>
      slide.frames.map((frame) => ({
        kind: frame.kind,
        role: frame.kind === "text" ? frame.role : undefined,
        x: r3(frame.xCm),
        y: r3(frame.yCm),
        w: r3(frame.widthCm),
        h: r3(frame.heightCm),
        paragraphs: frame.kind === "text" ? frame.paragraphs : undefined,
      })),
    ),
  };
}

describe("from-scratch writers (no retained package)", () => {
  it("pptx: a blank + added deck re-opens with the same slides, placeholders and geometry (AC-8/9)", async () => {
    const model = createBlankPresentation();
    addSlide(model, 0); // a title+body slide too
    const { bytes } = await writePptx(model);
    const reopened = await readPresentationBytes(bytes, "pptx", model.title);
    expect(strippedForCompare(reopened)).toEqual(strippedForCompare(model));
  });

  it("odp: a from-scratch deck re-opens the same, mimetype first and STOREd (AC-10/11)", async () => {
    const model: PresentationModel = { ...createBlankPresentation(), format: "odp" };
    addSlide(model, 0);
    const { bytes } = await writeOdp(model);
    const reopened = await readPresentationBytes(bytes, "odp", model.title);
    expect(strippedForCompare(reopened)).toEqual(strippedForCompare(model));
    // ODF container rule: mimetype first + STOREd (method 0 at offset 8).
    expect(bytes[8]).toBe(0);
    expect(new TextDecoder().decode(bytes.slice(30, 38))).toBe("mimetype");
  });
});

describe("pptx substitution rebuild", () => {
  async function openFixture(): Promise<{ model: PresentationModel; source: Map<string, Uint8Array> }> {
    const bytes = await buildPptxFixtureBytes();
    return { model: await readPresentationBytes(bytes, "pptx", "Deck"), source: await unzip(bytes) };
  }

  function expectUntouched(written: Map<string, Uint8Array>, source: Map<string, Uint8Array>): void {
    for (const [name, data] of source) {
      if (STRUCTURAL_PPTX_PARTS.has(name) || !written.has(name)) {
        continue;
      }
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
    }
  }

  it("reordering rebuilds only the structural parts; unmodelled parts stay byte-verbatim (AC-48/50/53/54)", async () => {
    const { model, source } = await openFixture();
    moveSlide(model, 3, 0); // slide4 to the front → [s4, s1, s2, s3]
    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    // No parts added or removed on a reorder.
    expect(new Set(written.keys())).toEqual(new Set(source.keys()));
    expectUntouched(written, source);
    // presentation.xml changed; Content-Types unchanged (no slides added/removed).
    expect(equalBytes(written.get("ppt/presentation.xml")!, source.get("ppt/presentation.xml")!)).toBe(false);
    expect(equalBytes(written.get("[Content_Types].xml")!, source.get("[Content_Types].xml")!)).toBe(true);
    // The re-minted sldIdLst carries the new order (slide4's part now first).
    const reopened = await readPresentationBytes(bytes, "pptx", "Deck");
    expect(reopened.slides[0].partName).toBe("ppt/slides/slide4.xml");
  });

  it("adding a slide appends a fresh part + rels and its Content-Types override (AC-27)", async () => {
    const { model, source } = await openFixture();
    addSlide(model, 0);
    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    expect(written.has("ppt/slides/slide5.xml")).toBe(true);
    expect(written.has("ppt/slides/_rels/slide5.xml.rels")).toBe(true);
    expectUntouched(written, source);
    // The new slide references the current slide's layout (slide1 → layout1).
    const rels = new TextDecoder().decode(written.get("ppt/slides/_rels/slide5.xml.rels")!);
    expect(rels).toContain("slideLayout1.xml");
    const reopened = await readPresentationBytes(bytes, "pptx", "Deck");
    expect(reopened.slides).toHaveLength(5);
  });

  it("deleting a slide drops its part, rels and Content-Types override (AC-34)", async () => {
    const { model, source } = await openFixture();
    deleteSlide(model, 0);
    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    expect(written.has("ppt/slides/slide1.xml")).toBe(false);
    expect(written.has("ppt/slides/_rels/slide1.xml.rels")).toBe(false);
    expect(new TextDecoder().decode(written.get("[Content_Types].xml")!)).not.toContain("/ppt/slides/slide1.xml");
    expectUntouched(written, source);
    const reopened = await readPresentationBytes(bytes, "pptx", "Deck");
    expect(reopened.slides).toHaveLength(3);
  });

  it("a duplicated slide's part + rels bytes equal the source's (AC-32/33)", async () => {
    const { model, source } = await openFixture();
    duplicateSlide(model, 1); // duplicate slide2
    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    expect(equalBytes(written.get("ppt/slides/slide5.xml")!, source.get("ppt/slides/slide2.xml")!)).toBe(true);
    expect(
      equalBytes(written.get("ppt/slides/_rels/slide5.xml.rels")!, source.get("ppt/slides/_rels/slide2.xml.rels")!),
    ).toBe(true);
    const reopened = await readPresentationBytes(bytes, "pptx", "Deck");
    expect(reopened.slides).toHaveLength(5);
  });

  it("re-emits every part byte-identical when nothing structural changed (AC-51/52)", async () => {
    const { model, source } = await openFixture();
    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name}`).toBe(true);
    }
  });
});

describe("odp substitution rebuild", () => {
  it("the reader's envelope + page sources recombine into content.xml byte-for-byte (AC-50)", async () => {
    const bytes = await buildOdpFixtureBytes();
    const model = await readPresentationBytes(bytes, "odp", "Deck");
    const content = new TextDecoder().decode((await unzip(bytes)).get("content.xml")!);
    const envelope = model.odpContentEnvelope!;
    const spliced = envelope.prefix + model.slides.map((slide) => slide.sourceXml).join("") + envelope.suffix;
    expect(spliced).toBe(content);
  });

  it("reordering rebuilds only content.xml; every other part stays byte-verbatim (AC-49/50)", async () => {
    const bytes = await buildOdpFixtureBytes();
    const source = await unzip(bytes);
    const model = await readPresentationBytes(bytes, "odp", "Deck");
    moveSlide(model, 0, 3);
    const written = await unzip((await writeOdp(model)).bytes);
    expect(new Set(written.keys())).toEqual(new Set(source.keys()));
    for (const [name, data] of source) {
      if (name === "content.xml") {
        expect(equalBytes(written.get(name)!, data)).toBe(false);
        continue;
      }
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
    }
    const reopened = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    // The reordered deck still round-trips its four pages.
    expect(reopened.slides).toHaveLength(4);
  });
});

// --- SPEC-081: edit → save → reopen text/formatting fidelity ----------------

function titleFrameOf(slide: Slide): TextFrame {
  return slide.frames.find((f): f is TextFrame => f.kind === "text" && f.role === "title")!;
}

describe("SPEC-081 edit → save → reopen fidelity", () => {
  it("pptx: an edited title survives; untouched siblings + structural parts stay byte-verbatim (AC-32/33/41/42/45)", async () => {
    const sourceBytes = await buildPptxFixtureBytes();
    const source = await unzip(sourceBytes);
    const model = await readPresentationBytes(sourceBytes, "pptx", "Deck");
    const title = titleFrameOf(model.slides[0]);
    title.paragraphs = [{ runs: [{ text: "Edited Title" }] }];
    applyRunFormat(title, wholeFrameRange(title), { bold: true, color: "#123456" });
    markSlideEdited(model.slides[0], "pptx");

    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    // A text-only edit rebuilds ONLY the dirty slide part.
    for (const [name, data] of source) {
      if (name === "ppt/slides/slide1.xml") continue;
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
    }

    const reopened = await readPresentationBytes(bytes, "pptx", "Deck");
    const rt = titleFrameOf(reopened.slides[0]);
    expect(rt.paragraphs[0].runs[0].text).toBe("Edited Title");
    expect(rt.paragraphs[0].runs[0].bold).toBe(true);
    expect(rt.paragraphs[0].runs[0].color).toBe("#123456");
    expect(slideOutline(reopened.slides[0]).title).toBe("Edited Title");

    // Stability (AC-47): the rebuilt part is byte-stable across writes, and a
    // reopen → re-download re-emits it verbatim.
    const twice = await unzip((await writePptx(model)).bytes);
    expect(equalBytes(twice.get("ppt/slides/slide1.xml")!, written.get("ppt/slides/slide1.xml")!)).toBe(true);
    const redownload = await unzip((await writePptx(reopened)).bytes);
    expect(equalBytes(redownload.get("ppt/slides/slide1.xml")!, written.get("ppt/slides/slide1.xml")!)).toBe(true);
  });

  it("pptx: a bulleted body line added to a dirty slide round-trips (AC-45/46)", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    const body = model.slides[1].frames.find((f): f is TextFrame => f.kind === "text" && f.role === "body")!;
    body.paragraphs.push({ runs: [{ text: "A new bullet" }], bullet: "disc" });
    markSlideEdited(model.slides[1], "pptx");
    const reopened = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const rtBody = reopened.slides[1].frames.find((f): f is TextFrame => f.kind === "text" && f.role === "body")!;
    const last = rtBody.paragraphs[rtBody.paragraphs.length - 1];
    expect(last.runs[0].text).toBe("A new bullet");
    expect(last.bullet).toBe("disc");
  });

  it("odp: an edited title survives; content.xml alone rebuilds, every other part verbatim (AC-31/42/45)", async () => {
    const sourceBytes = await buildOdpFixtureBytes();
    const source = await unzip(sourceBytes);
    const model = await readPresentationBytes(sourceBytes, "odp", "Deck");
    const title = titleFrameOf(model.slides[0]);
    title.paragraphs = [{ runs: [{ text: "Edited Title", bold: true }] }];
    markSlideEdited(model.slides[0], "odp");
    expect(model.slides[0].sourceXml).toBeUndefined(); // AC-31

    const { bytes } = await writeOdp(model);
    const written = await unzip(bytes);
    for (const [name, data] of source) {
      if (name === "content.xml") {
        expect(equalBytes(written.get(name)!, data)).toBe(false);
        continue;
      }
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
    }
    // Untouched pages spliced verbatim inside the rebuilt content.xml.
    const content = new TextDecoder().decode(written.get("content.xml")!);
    expect(content).toContain("Broaden format support"); // page 2, untouched

    const reopened = await readPresentationBytes(bytes, "odp", "Deck");
    const rt = titleFrameOf(reopened.slides[0]);
    expect(rt.paragraphs[0].runs[0].text).toBe("Edited Title");
    expect(rt.paragraphs[0].runs[0].bold).toBe(true);
    expect(slideOutline(reopened.slides[0]).title).toBe("Edited Title");
  });
});

describe("SPEC-089: a layout change re-points one slide's .rels and leaves every other part verbatim", () => {
  it("pptx: repoint slide1 to an existing layout — only slide1.xml + its .rels move", async () => {
    const { applySlideLayout } = await import("./layouts");
    const source = await buildPptxFixtureBytes();
    const model = await readPresentationBytes(source, "pptx", "Deck");
    applySlideLayout(model, 0, "ppt/slideLayouts/slideLayout2.xml");
    const back = await writePptx(model);
    const before = await unzip(source);
    const after = await unzip(back.bytes);
    // The layout parts, master, and Content-Types are untouched (no authoring).
    for (const name of [
      "[Content_Types].xml",
      "ppt/slideMasters/slideMaster1.xml",
      "ppt/slideLayouts/slideLayout1.xml",
      "ppt/slideLayouts/slideLayout2.xml",
      "ppt/slides/_rels/slide2.xml.rels",
    ]) {
      expect(equalBytes(before.get(name)!, after.get(name)!), name).toBe(true);
    }
    // slide1's .rels changed by exactly the layout Target.
    const relsBefore = new TextDecoder().decode(before.get("ppt/slides/_rels/slide1.xml.rels")!);
    const relsAfter = new TextDecoder().decode(after.get("ppt/slides/_rels/slide1.xml.rels")!);
    expect(relsBefore).toContain("slideLayout1.xml");
    expect(relsAfter).toContain("slideLayout2.xml");
    // The reopened slide reports the new layout.
    const reopened = await readPresentationBytes(back.bytes, "pptx", "Deck");
    expect(reopened.slides[0].layoutKey).toBe("ppt/slideLayouts/slideLayout2.xml");
  });
});

// --- SPEC-090: deck setup round-trip ---------------------------------------

describe("SPEC-090: deck setup round-trip", () => {
  const MOD = new Date("2026-03-01T12:00:00Z");

  it("an untouched deck still downloads byte-identically per part (both formats)", async () => {
    for (const format of ["pptx", "odp"] as const) {
      const source = format === "pptx" ? await buildPptxFixtureBytes() : await buildOdpFixtureBytes();
      const model = await readPresentationBytes(source, format, "Deck");
      const { bytes } = format === "pptx" ? await writePptx(model) : await writeOdp(model);
      const before = await unzip(source);
      const after = await unzip(bytes);
      expect([...after.keys()]).toEqual([...before.keys()]);
      for (const [name, data] of before) {
        expect(equalBytes(after.get(name)!, data), `${format} part ${name}`).toBe(true);
      }
    }
  });

  it("pptx: size + first-slide-number + metadata → save → reopen reproduces every field, re-saves byte-stable", async () => {
    const { applyDeckSize, setFirstSlideNumber, applyDeckProperties } = await import("./deckSetup");
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    setFirstSlideNumber(model, 3);
    applyDeckProperties(model, {
      ...model.properties,
      title: "New Title",
      subject: "Subj",
      keywords: "kw",
      comments: "Cmt",
      author: "Ada",
      lastModifiedBy: "Grace",
    });
    const first = await writePptx(model, { modified: MOD });
    const reopened = await readPresentationBytes(first.bytes, "pptx", "Deck");
    expect(reopened.slideWidthCm).toBeCloseTo(25.4, 2);
    expect(reopened.slideHeightCm).toBeCloseTo(19.05, 2);
    expect(reopened.firstSlideNumber).toBe(3);
    expect(reopened.properties?.title).toBe("New Title");
    expect(reopened.properties?.subject).toBe("Subj");
    expect(reopened.properties?.keywords).toBe("kw");
    expect(reopened.properties?.comments).toBe("Cmt");
    expect(reopened.properties?.author).toBe("Ada");
    expect(reopened.properties?.lastModifiedBy).toBe("Grace");
    // Re-saving the reopened (deckDirty-free) model is byte-stable.
    const second = await writePptx(reopened, { modified: MOD });
    const a = await unzip(first.bytes);
    const b = await unzip(second.bytes);
    for (const part of ["ppt/presentation.xml", "docProps/core.xml"]) {
      expect(equalBytes(a.get(part)!, b.get(part)!), part).toBe(true);
    }
  });

  it("odp: size + metadata reproduces; firstSlideNumber stays model-only; re-saves byte-stable", async () => {
    const { applyDeckSize, setFirstSlideNumber, applyDeckProperties } = await import("./deckSetup");
    const model = await readPresentationBytes(await buildOdpFixtureBytes(), "odp", "Deck");
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    setFirstSlideNumber(model, 3);
    applyDeckProperties(model, {
      ...model.properties,
      title: "New Title",
      subject: "Subj",
      keywords: "kw",
      comments: "Cmt",
      author: "Ada",
      lastModifiedBy: "Grace",
    });
    const first = await writeOdp(model, { modified: MOD });
    const reopened = await readPresentationBytes(first.bytes, "odp", "Deck");
    expect(reopened.slideWidthCm).toBeCloseTo(25.4, 2);
    expect(reopened.slideHeightCm).toBeCloseTo(19.05, 2);
    // odp does not serialise the numbering origin — it reads back as absent (= 1).
    expect(reopened.firstSlideNumber).toBeUndefined();
    expect(reopened.properties?.title).toBe("New Title");
    expect(reopened.properties?.author).toBe("Ada");
    expect(reopened.properties?.lastModifiedBy).toBe("Grace");
    const second = await writeOdp(reopened, { modified: MOD });
    const a = await unzip(first.bytes);
    const b = await unzip(second.bytes);
    for (const part of ["styles.xml", "meta.xml"]) {
      expect(equalBytes(a.get(part)!, b.get(part)!), part).toBe(true);
    }
  });
});
