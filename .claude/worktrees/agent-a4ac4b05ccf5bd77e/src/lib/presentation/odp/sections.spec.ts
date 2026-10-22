// SPEC-099 — the odp section mapping: the custom-property value both ways,
// clamp/extend at read, the sibling-`meta:user-defined` preserve path, a
// byte-verbatim meta.xml when sections did not change, and membership
// recomputed after a reorder.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { equalBytes } from "../../core/equality";
import { decodedPartBytes } from "../../core/packageParts";
import type { PresentationModel } from "../model";
import { readPresentationBytes } from "../open";
import { addSlide, createBlankPresentation, moveSlide } from "../slides";
import { addSection, renameSection } from "../sections";
import { writeOdp } from "./writer";
import { decodeSectionsValue, encodeSectionsValue, SECTIONS_PROPERTY_NAME } from "./sections";

const decoder = new TextDecoder();

function deck(n: number): PresentationModel {
  const model = createBlankPresentation();
  model.format = "odp";
  for (let i = 1; i < n; i++) {
    addSlide(model, i - 1);
  }
  delete model.structureDirty;
  return model;
}

/** A sectioned odp deck read back from bytes (so it carries a retained
 * package), with one unrelated custom property beside the sections. */
async function sectionedOdp(): Promise<PresentationModel> {
  const model = deck(5);
  addSection(model, 0, "Intro");
  addSection(model, 2, "Body");
  model.properties = { customProperties: [{ name: "Reviewer", type: "text", value: "Ada" }] };
  return readPresentationBytes((await writeOdp(model)).bytes, "odp");
}

async function metaOf(bytes: Uint8Array): Promise<string> {
  return decoder.decode((await unzip(bytes)).get("meta.xml")!);
}

describe("the sections property value", () => {
  it("encodes compact JSON with run lengths", () => {
    const model = deck(3);
    addSection(model, 0, "Intro");
    addSection(model, 2, "Second");
    expect(encodeSectionsValue(model)).toBe(
      `{"v":1,"sections":[{"key":"de-sec1","name":"Intro","n":2},{"key":"de-sec2","name":"Second","n":1}]}`,
    );
    expect(encodeSectionsValue(deck(2))).toBeUndefined();
  });

  it("decodes back onto the slides", () => {
    const slides = deck(3).slides;
    const sections = decodeSectionsValue(
      `{"v":1,"sections":[{"key":"de-sec1","name":"Intro","n":2},{"key":"de-sec2","name":"Second","n":1}]}`,
      slides,
    );
    expect(sections).toEqual([
      { key: "de-sec1", name: "Intro" },
      { key: "de-sec2", name: "Second" },
    ]);
    expect(slides.map((slide) => slide.sectionKey)).toEqual(["de-sec1", "de-sec1", "de-sec2"]);
  });

  it("clamps a total overrunning the deck and extends the last run over a shortfall", () => {
    const over = deck(3).slides;
    decodeSectionsValue(`{"v":1,"sections":[{"key":"de-sec1","name":"A","n":2},{"key":"de-sec2","name":"B","n":9}]}`, over);
    expect(over.map((slide) => slide.sectionKey)).toEqual(["de-sec1", "de-sec1", "de-sec2"]);
    const short = deck(4).slides;
    decodeSectionsValue(`{"v":1,"sections":[{"key":"de-sec1","name":"A","n":1},{"key":"de-sec2","name":"B","n":1}]}`, short);
    expect(short.map((slide) => slide.sectionKey)).toEqual(["de-sec1", "de-sec2", "de-sec2", "de-sec2"]);
  });

  it("rejects anything but a v1 payload and re-mints unsafe or duplicate keys", () => {
    const slides = deck(2).slides;
    expect(decodeSectionsValue("not json", slides)).toBeUndefined();
    expect(decodeSectionsValue(`{"v":2,"sections":[]}`, slides)).toBeUndefined();
    expect(decodeSectionsValue(`{"v":1,"sections":[]}`, slides)).toBeUndefined();
    expect(decodeSectionsValue(`{"v":1,"sections":[{"key":"a","name":"A","n":-1}]}`, slides)).toBeUndefined();
    const sections = decodeSectionsValue(
      `{"v":1,"sections":[{"key":"<x>","name":"A","n":1},{"key":"de-sec1","name":"  ","n":1}]}`,
      slides,
    );
    expect(sections?.map((s) => s.key)).toEqual(["de-sec2", "de-sec1"]);
    expect(sections?.[1].name).toBe("Default Section");
  });
});

describe("writeOdp with sections", () => {
  it("reads sections back and keeps them out of the custom-property list", async () => {
    const model = await sectionedOdp();
    expect(model.sections?.map((s) => s.name)).toEqual(["Intro", "Body"]);
    expect(model.slides.map((slide) => slide.sectionKey)).toEqual(["de-sec1", "de-sec1", "de-sec2", "de-sec2", "de-sec2"]);
    expect(model.properties?.customProperties).toEqual([{ name: "Reviewer", type: "text", value: "Ada" }]);
  });

  it("re-emits meta.xml byte-verbatim while the sections are unchanged — even on a structural save", async () => {
    const model = await sectionedOdp();
    const original = decodedPartBytes(model.opaque!).get("meta.xml")!;
    model.structureDirty = true; // e.g. a slide edit elsewhere
    model.slides[4].dirty = true;
    const saved = (await unzip((await writeOdp(model)).bytes)).get("meta.xml")!;
    expect(equalBytes(saved, original)).toBe(true);
  });

  it("a section rename rebuilds meta.xml through buildMetaXml, the unrelated property surviving", async () => {
    const model = await sectionedOdp();
    renameSection(model, "de-sec2", "Results");
    const meta = await metaOf((await writeOdp(model)).bytes);
    expect(meta).toContain(`meta:name="Reviewer" meta:value-type="string">Ada<`);
    expect(meta).toContain(`meta:name="${SECTIONS_PROPERTY_NAME}" meta:value-type="string">`);
    expect(meta).toContain("Results");
    const reread = await readPresentationBytes((await writeOdp(model)).bytes, "odp");
    expect(reread.sections).toEqual(model.sections);
  });

  it("recomputes run lengths after a reorder", async () => {
    const model = await sectionedOdp();
    moveSlide(model, 4, 0); // a Body slide to the front → joins Intro
    const reread = await readPresentationBytes((await writeOdp(model)).bytes, "odp");
    expect(reread.slides.map((slide) => slide.sectionKey)).toEqual(model.slides.map((slide) => slide.sectionKey));
    expect(reread.slides.map((slide) => slide.sectionKey)).toEqual(["de-sec1", "de-sec1", "de-sec1", "de-sec2", "de-sec2"]);
  });
});
