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

  it("pptx: notes typed on a newly-added slide round-trip (AC-17/26)", async () => {
    const { addSlide } = await import("./slides");
    const model = modelFrom(await parsePptx(await buildPptxFixtureBytes()), "pptx");
    const at = addSlide(model, model.slides.length - 1); // new last slide
    model.slides[at].notes = ["Note on the new slide"];
    model.slides[at].notesDirty = true;
    const { bytes } = await writePptx(model);
    const after = await pptxParts(bytes);
    // A notes-slide part + Content-Types override were minted for the new slide.
    const notesParts = [...after.keys()].filter((n) => /notesSlides\/notesSlide\d+\.xml$/.test(n));
    expect(notesParts.length).toBe(2); // fixture's slide1 + the new slide
    const reopened = await parsePptx(bytes);
    expect(reopened.slides.length).toBe(5);
    expect(reopened.slides[at].notes).toEqual(["Note on the new slide"]);
  });

  it("pptx: a duplicated slide's edited notes get their own part, not the source's (AC-26)", async () => {
    const { duplicateSlide } = await import("./slides");
    const model = modelFrom(await parsePptx(await buildPptxFixtureBytes()), "pptx");
    // Slide 1 has notes in the fixture; duplicate it and edit the copy's notes.
    const at = duplicateSlide(model, 0);
    model.slides[at].notes = ["Only on the copy"];
    model.slides[at].notesDirty = true;
    const { bytes } = await writePptx(model);
    const after = await pptxParts(bytes);
    const reopened = await parsePptx(bytes);
    // Source keeps its own notes; the copy carries the edited notes — separate parts.
    expect(reopened.slides[0].notes).toEqual(FIXTURE_NOTES);
    expect(reopened.slides[at].notes).toEqual(["Only on the copy"]);
    const notesParts = [...after.keys()].filter((n) => /notesSlides\/notesSlide\d+\.xml$/.test(n));
    expect(notesParts.length).toBe(2);
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

// --- SPEC-093 notes non-body placeholder delta (AC-92/93/123/135) ----------
import { preserveNotesFieldPlaceholders } from "./pptx/serialize";

describe("notes non-body placeholders (SPEC-093 AC-92/93)", () => {
  const NOTES_SOURCE =
    `<p:notes><p:cSld><p:spTree>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>old</a:t></a:r></a:p></p:txBody></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Notes footer</a:t></a:r></a:p></p:txBody></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide Number Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:fld id="{x}" type="slidenum"><a:t>7</a:t></a:fld></a:p></p:txBody></p:sp>` +
    `</p:spTree></p:cSld></p:notes>`;

  it("re-emits the source's hdr/ftr/dt/sldNum placeholders, refreshing a slidenum a:t", () => {
    const nonBody = preserveNotesFieldPlaceholders(NOTES_SOURCE, "12", "14 May 2026");
    // Footer preserved verbatim (a plain run stays), slide number refreshed to 12.
    expect(nonBody).toContain(`type="ftr"`);
    expect(nonBody).toContain(`<a:t>Notes footer</a:t>`);
    expect(nonBody).toContain(`type="slidenum"><a:t>12</a:t></a:fld>`);
    // The body/sldImg placeholders are NOT re-emitted here (the scaffold owns them).
    expect(nonBody).not.toContain(`type="body"`);
    expect(nonBody).not.toContain(`type="sldImg"`);
  });

  it("the notes-slide serialiser splices the preserved placeholders after the body", () => {
    const nonBody = preserveNotesFieldPlaceholders(NOTES_SOURCE, "3", "14 May 2026");
    const xml = serializeNotesSlideXml(["My note"], nonBody);
    expect(xml).toContain(`<a:t>My note</a:t>`);
    expect(xml).toContain(`type="ftr"`);
    expect(xml).toContain(`type="slidenum"><a:t>3</a:t></a:fld>`);
    // Non-body placeholders come after the body scaffold, before the spTree close.
    expect(xml.indexOf(`type="body"`)).toBeLessThan(xml.indexOf(`type="ftr"`));
  });
});
