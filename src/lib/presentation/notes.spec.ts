// SPEC-083 speaker notes: reader extraction (both formats, exact strings,
// trailing-empty stripping, notes-less slides), the pptx notes-slide package
// surgery (first-notes creation, in-place rebuild, removal-with-no-orphans),
// and the edit → save → reopen round trip in both formats.
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import { buildOdpFixtureBytes, buildPptxFixtureBytes, FIXTURE_NOTES } from "./fixtures/build";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { serializeNotesSlideXml } from "./pptx/serialize";
import { markSlideEdited } from "./text";
import type { PresentationModel, ParsedPresentation } from "./model";

const decoder = new TextDecoder("utf-8");

async function pptxParts(bytes: Uint8Array): Promise<Map<string, string>> {
  const parts = await unzip(bytes);
  const out = new Map<string, string>();
  for (const [name, data] of parts) {
    out.set(name, decoder.decode(data));
  }
  return out;
}

function modelFrom(parsed: ParsedPresentation, format: "pptx" | "odp"): PresentationModel {
  return {
    kind: "presentation",
    id: "deck",
    title: "Deck",
    format,
    slides: parsed.slides,
    slideWidthCm: parsed.slideWidthCm,
    slideHeightCm: parsed.slideHeightCm,
    opaque: parsed.opaque,
    media: parsed.media,
    odpContentEnvelope: parsed.odpContentEnvelope,
  };
}

describe("notes reader", () => {
  it("pptx: reads slide 1's notes-slide body, flattening formatting and stripping the trailing empty (AC-2/4)", async () => {
    const { slides } = await parsePptx(await buildPptxFixtureBytes());
    expect(slides[0].notes).toEqual(FIXTURE_NOTES);
    // Slides without a notes-slide part have the key absent, never [] (AC-1).
    expect(slides[1].notes).toBeUndefined();
    expect(slides[2].notes).toBeUndefined();
    expect(slides[3].notes).toBeUndefined();
  });

  it("odp: reads the presentation:notes frame, ignoring the page-thumbnail, stripping the trailing empty (AC-3/4)", async () => {
    const { slides } = await parseOdp(await buildOdpFixtureBytes());
    expect(slides[0].notes).toEqual(FIXTURE_NOTES);
    expect(slides[1].notes).toBeUndefined();
  });
});

describe("serializeNotesSlideXml", () => {
  it("emits the body placeholder with one a:p per entry and a sldImg placeholder", () => {
    const xml = serializeNotesSlideXml(["Line one", "Line two"]);
    expect(xml).toContain(`<p:ph type="sldImg"/>`);
    expect(xml).toContain(`<p:ph type="body" idx="1"/>`);
    expect(xml).toContain(`<a:t>Line one</a:t>`);
    expect(xml).toContain(`<a:t>Line two</a:t>`);
  });
});

describe("pptx notes package surgery", () => {
  it("rebuilds an existing notes part in place; the slide part stays byte-verbatim (AC-15)", async () => {
    const original = await buildPptxFixtureBytes();
    const before = await pptxParts(original);
    const model = modelFrom(await parsePptx(original), "pptx");
    model.slides[0].notes = ["Reworded opener"];
    model.slides[0].notesDirty = true;
    const { bytes } = await writePptx(model);
    const after = await pptxParts(bytes);
    // The notes part was rebuilt from the model…
    expect(after.get("ppt/notesSlides/notesSlide1.xml")).toContain(`<a:t>Reworded opener</a:t>`);
    // …but the slide part and the notes part's rels stay byte-verbatim.
    expect(after.get("ppt/slides/slide1.xml")).toBe(before.get("ppt/slides/slide1.xml"));
    expect(after.get("ppt/notesSlides/_rels/notesSlide1.xml.rels")).toBe(
      before.get("ppt/notesSlides/_rels/notesSlide1.xml.rels"),
    );
  });

  it("creates a notes-slide part, rels, Content-Types override and a notes master on first notes (AC-17..22)", async () => {
    const original = await buildPptxFixtureBytes();
    const before = await pptxParts(original);
    const model = modelFrom(await parsePptx(original), "pptx");
    // Slide 2 has no notes part in the fixture, and the fixture has no notes master.
    model.slides[1].notes = ["Brand new note"];
    model.slides[1].notesDirty = true;
    const { bytes } = await writePptx(model);
    const after = await pptxParts(bytes);

    // A uniquely-named notes-slide part + its rels (→ slide, → notes master).
    expect(after.has("ppt/notesSlides/notesSlide2.xml")).toBe(true);
    const notesRels = after.get("ppt/notesSlides/_rels/notesSlide2.xml.rels")!;
    expect(notesRels).toContain("relationships/slide");
    expect(notesRels).toContain("relationships/notesMaster");
    // The slide's rels gained the notesSlide relationship.
    expect(after.get("ppt/slides/_rels/slide2.xml.rels")).toContain("relationships/notesSlide");
    // Content-Types overrides for the notes-slide part and the notes master.
    const ct = after.get("[Content_Types].xml")!;
    expect(ct).toContain("/ppt/notesSlides/notesSlide2.xml");
    expect(ct).toContain("/ppt/notesMasters/notesMaster1.xml");
    // The minimal notes master + its rels.
    expect(after.has("ppt/notesMasters/notesMaster1.xml")).toBe(true);
    expect(after.has("ppt/notesMasters/_rels/notesMaster1.xml.rels")).toBe(true);
    // presentation.xml gained p:notesMasterIdLst and its rels the notesMaster.
    expect(after.get("ppt/presentation.xml")).toContain("<p:notesMasterIdLst>");
    expect(after.get("ppt/_rels/presentation.xml.rels")).toContain("relationships/notesMaster");
    // The slide part itself is untouched.
    expect(after.get("ppt/slides/slide2.xml")).toBe(before.get("ppt/slides/slide2.xml"));
  });

  it("omits the notes-slide part, its rels and relationship when notes are deleted (AC-23/24)", async () => {
    const original = await buildPptxFixtureBytes();
    const model = modelFrom(await parsePptx(original), "pptx");
    delete model.slides[0].notes;
    model.slides[0].notesDirty = true;
    const { bytes } = await writePptx(model);
    const after = await pptxParts(bytes);
    expect(after.has("ppt/notesSlides/notesSlide1.xml")).toBe(false);
    expect(after.has("ppt/notesSlides/_rels/notesSlide1.xml.rels")).toBe(false);
    expect(after.get("ppt/slides/_rels/slide1.xml.rels")).not.toContain("relationships/notesSlide");
    expect(after.get("[Content_Types].xml")).not.toContain("/ppt/notesSlides/notesSlide1.xml");
  });
});

describe("notes round trip (AC-26)", () => {
  it("pptx: edit → save → reopen shows the edited notes", async () => {
    const model = modelFrom(await parsePptx(await buildPptxFixtureBytes()), "pptx");
    model.slides[0].notes = ["Edited note A", "Edited note B"];
    model.slides[0].notesDirty = true;
    const { bytes } = await writePptx(model);
    const reopened = await parsePptx(bytes);
    expect(reopened.slides[0].notes).toEqual(["Edited note A", "Edited note B"]);
  });

  it("odp: edit → save → reopen shows the edited notes", async () => {
    const model = modelFrom(await parseOdp(await buildOdpFixtureBytes()), "odp");
    model.slides[0].notes = ["Edited odp note"];
    markSlideEdited(model.slides[0], "odp");
    const { bytes } = await writeOdp(model);
    const reopened = await parseOdp(bytes);
    expect(reopened.slides[0].notes).toEqual(["Edited odp note"]);
  });

  it("pptx: a notes edit alongside a structural change round-trips both (AC-15/26)", async () => {
    const { addSlide } = await import("./slides");
    const model = modelFrom(await parsePptx(await buildPptxFixtureBytes()), "pptx");
    model.slides[0].notes = ["Structural + notes"];
    model.slides[0].notesDirty = true;
    addSlide(model, 0); // sets structureDirty → the substitution path
    const { bytes } = await writePptx(model);
    const reopened = await parsePptx(bytes);
    expect(reopened.slides.length).toBe(5);
    expect(reopened.slides[0].notes).toEqual(["Structural + notes"]);
  });

  it("a deck whose notes were never touched re-emits its notes parts byte-verbatim (AC-27)", async () => {
    const original = await buildPptxFixtureBytes();
    const before = await pptxParts(original);
    const model = modelFrom(await parsePptx(original), "pptx");
    const { bytes } = await writePptx(model); // no edits
    const after = await pptxParts(bytes);
    expect(after.get("ppt/notesSlides/notesSlide1.xml")).toBe(before.get("ppt/notesSlides/notesSlide1.xml"));
  });
});
