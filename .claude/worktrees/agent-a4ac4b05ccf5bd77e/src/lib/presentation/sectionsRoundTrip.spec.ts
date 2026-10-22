// SPEC-099 — the round-trip AC's deck and the spec099-sections fixtures:
// a deck with three sections, a pasted slide and two reused slides → save →
// reopen reproduces sections, every sectionKey, slide order, retargeted
// animations and remapped mediaIds under equalStructural over the modelled
// subset (no semantic relation is claimed — equalSemantic is typed for the
// text model); a section-only edit re-emits only the structural parts; a
// no-edit download stays byte-identical part-for-part; stripRetention carries
// sections into the other format's mapping. roundTrip.spec.ts is untouched.
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import { equalBytes } from "../core/equality";
import { equalStructural } from "../equality";
import type { PresentationModel, Slide } from "./model";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { slideOutline } from "./slideModel";
import { addSection, renameSection, sectionsOf } from "./sections";
import { stripRetention } from "./slides";
import { importSlides } from "./slideImport";
import { buildSlideClipboardPayload, parseSlideClipboardPayload, slideClipboardFlavours } from "./slideClipboard";
import {
  buildPptxFixtureBytes,
  buildSpec099OdpFixtureBytes,
  buildSpec099PptxFixtureBytes,
  SPEC099_ANIMATED_SLIDE,
  SPEC099_IMAGE_SLIDE,
  SPEC099_ODP_EXTRA_PROPERTY,
  SPEC099_PPTX_EXTRA_EXT,
  SPEC099_SECTIONS,
  SPEC099_TITLES,
} from "./fixtures/build";

const decoder = new TextDecoder();
const INSTANT = new Date("2026-10-02T10:00:00Z");

async function write(model: PresentationModel): Promise<Uint8Array> {
  return model.format === "odp"
    ? (await writeOdp(model, { modified: INSTANT, now: INSTANT })).bytes
    : (await writePptx(model, { modified: INSTANT, now: INSTANT })).bytes;
}

async function reopen(model: PresentationModel): Promise<PresentationModel> {
  return readPresentationBytes(await write(model), model.format, model.title);
}

function titles(model: PresentationModel): (string | undefined)[] {
  return model.slides.map((slide) => slideOutline(slide).title);
}

function runs(model: PresentationModel): [string, string[]][] {
  return sectionsOf(model).map((run) => [
    run.section.name,
    model.slides.slice(run.startIndex, run.endIndex).map((slide) => slideOutline(slide).title ?? ""),
  ]);
}

function frameIndex(slide: Slide, id: string): number {
  return slide.frames.findIndex((frame) => frame.id === id);
}

/** Each section's name and run length. */
function runShape(deck: PresentationModel): [string, number][] {
  return sectionsOf(deck).map((run) => [run.section.name, run.endIndex - run.startIndex]);
}

/** The modelled subset compared after a round trip: the sections, and per
 * slide its section key, outline, animations (by target frame position) and
 * image media ids. */
function modelledSubset(model: PresentationModel): unknown {
  return {
    sections: model.sections,
    slides: model.slides.map((slide) => ({
      sectionKey: slide.sectionKey,
      outline: slideOutline(slide),
      animations: (slide.animations ?? []).map((animation) => ({ ...animation, frameId: frameIndex(slide, animation.frameId) })),
      media: slide.frames.flatMap((frame) => (frame.kind === "image" ? [frame.mediaId] : [])),
    })),
  };
}

describe("the spec099-sections fixtures (AC-16)", () => {
  for (const format of ["pptx", "odp"] as const) {
    it(`${format}: six slides in two named sections, an image and an animation`, async () => {
      const bytes = format === "pptx" ? await buildSpec099PptxFixtureBytes() : await buildSpec099OdpFixtureBytes();
      const model = await readPresentationBytes(bytes, format, "Sections fixture");
      expect(titles(model)).toEqual(SPEC099_TITLES);
      expect(runs(model)).toEqual(SPEC099_SECTIONS.map((section) => [section.name, section.titles]));
      expect(model.slides[SPEC099_IMAGE_SLIDE].frames.some((frame) => frame.kind === "image")).toBe(true);
      expect(model.slides[SPEC099_ANIMATED_SLIDE].animations).toHaveLength(1);
    });
  }

  it("pptx: the unrelated p:ext rides beside the section list", async () => {
    const parts = await unzip(await buildSpec099PptxFixtureBytes());
    const xml = decoder.decode(parts.get("ppt/presentation.xml")!);
    expect(xml).toContain(`<p:extLst>${SPEC099_PPTX_EXTRA_EXT}<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">`);
  });

  it("odp: the unrelated custom property is modelled; the sections one is not", async () => {
    const model = await readPresentationBytes(await buildSpec099OdpFixtureBytes(), "odp");
    expect(model.properties?.customProperties).toEqual([SPEC099_ODP_EXTRA_PROPERTY]);
  });
});

describe("no-edit and section-only saves (AC-15)", () => {
  for (const format of ["pptx", "odp"] as const) {
    it(`${format}: a no-edit download is byte-identical part-for-part`, async () => {
      const bytes = format === "pptx" ? await buildSpec099PptxFixtureBytes() : await buildSpec099OdpFixtureBytes();
      const model = await readPresentationBytes(bytes, format);
      const before = await unzip(bytes);
      const after = await unzip(await write(model));
      expect([...after.keys()]).toEqual([...before.keys()]);
      for (const [name, data] of before) {
        expect(equalBytes(after.get(name)!, data), name).toBe(true);
      }
    });
  }

  it("pptx: a section-only edit re-emits presentation.xml (+ .rels, [Content_Types].xml) and nothing else", async () => {
    const bytes = await buildSpec099PptxFixtureBytes();
    const model = await readPresentationBytes(bytes, "pptx");
    renameSection(model, model.sections![1].key, "Deep dive");
    const before = await unzip(bytes);
    const after = await unzip(await write(model));
    const structural = new Set(["ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"]);
    expect([...after.keys()]).toEqual([...before.keys()]);
    for (const [name, data] of before) {
      if (!structural.has(name)) {
        expect(equalBytes(after.get(name)!, data), name).toBe(true);
      }
    }
    const presentation = decoder.decode(after.get("ppt/presentation.xml")!);
    expect(presentation).toContain(`name="Deep dive"`);
    // The unrelated p:ext survives byte-verbatim beside the rewritten list.
    expect(presentation).toContain(`<p:extLst>${SPEC099_PPTX_EXTRA_EXT}<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">`);
  });

  it("odp: a section-only edit re-emits meta.xml, every other part byte-identical", async () => {
    const bytes = await buildSpec099OdpFixtureBytes();
    const model = await readPresentationBytes(bytes, "odp");
    renameSection(model, model.sections![0].key, "Kickoff");
    const before = await unzip(bytes);
    const after = await unzip(await write(model));
    expect([...after.keys()]).toEqual([...before.keys()]);
    for (const [name, data] of before) {
      if (name !== "meta.xml") {
        expect(equalBytes(after.get(name)!, data), name).toBe(true);
      }
    }
    const meta = decoder.decode(after.get("meta.xml")!);
    expect(meta).toContain("Kickoff");
    expect(meta).toContain(`meta:name="${SPEC099_ODP_EXTRA_PROPERTY.name}"`);
    expect((await readPresentationBytes(await write(model), "odp")).sections?.[0].name).toBe("Kickoff");
  });
});

describe("the round-trip deck (AC-15)", () => {
  for (const format of ["pptx", "odp"] as const) {
    it(`${format}: three sections, a pasted slide and two reused slides survive save → reopen`, async () => {
      const bytes = format === "pptx" ? await buildSpec099PptxFixtureBytes() : await buildSpec099OdpFixtureBytes();
      const model = await readPresentationBytes(bytes, format, "Round trip");
      // A third section over the last slide.
      addSection(model, 5, "Appendix");
      // Paste the animated slide (through the clipboard's wire form) after slide 2.
      const { payload } = slideClipboardFlavours(model, [SPEC099_ANIMATED_SLIDE]);
      const pasted = parseSlideClipboardPayload(payload!)!;
      importSlides(model, pasted, [0], { at: 2 });
      // Reuse two mini-deck slides (one carrying the image) at the end.
      const mini = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
      const report = importSlides(model, mini, [0, 3], { at: model.slides.length });
      expect(report.slides).toBe(2);
      expect(model.sections).toHaveLength(3);

      const reopened = await reopen(model);
      expect(equalStructural(modelledSubset(reopened), modelledSubset(model))).toBe(true);
      expect(titles(reopened)).toEqual(titles(model));
      // The pasted copy kept its animation, retargeted onto its own frame.
      const copy = reopened.slides[2];
      expect(copy.animations).toHaveLength(1);
      expect(copy.frames.some((frame) => frame.id === copy.animations![0].frameId)).toBe(true);
      // The reused image resolves to a part the reopened deck holds.
      const image = reopened.slides[reopened.slides.length - 1].frames.find((frame) => frame.kind === "image");
      expect(image?.kind).toBe("image");
      if (image?.kind === "image") {
        expect(reopened.media?.some((part) => part.mediaId === image.mediaId)).toBe(true);
      }
    });
  }

  it("stripRetention carries sections into the other format's mapping", async () => {
    const model = await readPresentationBytes(await buildSpec099PptxFixtureBytes(), "pptx");
    const copy = stripRetention(model);
    copy.format = "odp";
    const asOdp = await readPresentationBytes(await write(copy), "odp");
    expect(runs(asOdp)).toEqual(runs(model));
    expect(asOdp.sections?.map((section) => section.key)).toEqual(model.sections?.map((section) => section.key));
    // And back: the section runs (names and lengths) survive the second hop.
    const back = stripRetention(asOdp);
    back.format = "pptx";
    const asPptx = await readPresentationBytes(await write(back), "pptx");
    expect(runShape(asPptx)).toEqual(runShape(model));
    expect(asPptx.slides.map((slide) => slide.sectionKey)).toEqual(model.slides.map((slide) => slide.sectionKey));
  });

  it("a copy's payload keeps only the referenced media", async () => {
    const model = await readPresentationBytes(await buildSpec099PptxFixtureBytes(), "pptx");
    expect(buildSlideClipboardPayload(model, [SPEC099_IMAGE_SLIDE]).media).toHaveLength(1);
    expect(buildSlideClipboardPayload(model, [0]).media).toHaveLength(0);
  });
});
