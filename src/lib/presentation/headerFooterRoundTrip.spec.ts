// SPEC-093 round trip, both formats (AC-81..85): the header/footer fixtures
// build → reopen → the deck values + every Slide.fields entry survive
// (equalStructural over resolved facts, geometry rounded per SPEC-078's rule);
// a no-edit re-save is byte-identical; a header/footer edit rebuilds only the
// field-carrying slide parts, leaving every other part equalBytes-identical;
// and the pptx datetime* delta is pinned explicitly.
import { describe, expect, it } from "vitest";
import { equalStructural } from "../equality";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { defaultFieldRect } from "./headerFooter";
import { markSlideEdited } from "./text";
import {
  EXPECTED_HEADER_FOOTER,
  EXPECTED_HEADER_FOOTER_OVERRIDE,
  buildOdpHeaderFooterFixtureBytes,
  buildPptxHeaderFooterFixtureBytes,
} from "./fixtures/build";
import type { PresentationModel, SlideFieldKind, SlideFieldPlaceholder } from "./model";

const r3 = (n: number): number => Math.round(n * 1000) / 1000;

/** A field entry's comparable facts (geometry rounded, minted nothing), for
 * equalStructural across a build → re-parse round trip. */
function fact(field: SlideFieldPlaceholder): unknown {
  return {
    kind: field.kind,
    x: r3(field.xCm),
    y: r3(field.yCm),
    w: r3(field.widthCm),
    h: r3(field.heightCm),
    align: field.align,
    text: field.text,
    wireType: field.wireType,
    wireStyleName: field.wireStyleName,
  };
}

function facts(model: PresentationModel, slideIndex: number): unknown[] {
  return (model.slides[slideIndex].fields ?? []).map(fact);
}

/** The expected geometry for a kind at the fixture's deck size, rounded. */
function rect(kind: SlideFieldKind, w: number, h: number): { x: number; y: number; w: number; h: number } {
  const r = defaultFieldRect(kind, w, h);
  return { x: r3(r.xCm), y: r3(r.yCm), w: r3(r.widthCm), h: r3(r.heightCm) };
}

describe("SPEC-093 header/footer round trip", () => {
  describe("pptx", () => {
    it("reopens with the deck values + every field entry (equalStructural)", async () => {
      const bytes = await buildPptxHeaderFooterFixtureBytes();
      const m = await readPresentationBytes(bytes, "pptx");
      const w = m.slideWidthCm;
      const h = m.slideHeightCm;
      expect(equalStructural(m.headerFooter, EXPECTED_HEADER_FOOTER)).toBe(true);
      expect(m.firstSlideNumber).toBeUndefined(); // counts from 1
      // Slide 0 (title) suppressed: no field placeholders emitted, none read.
      expect(m.slides[0].fields).toBeUndefined();
      // Slide 1: all three, default look/geometry.
      expect(equalStructural(facts(m, 1), [
        { kind: "footer", ...rect("footer", w, h), align: "center", text: undefined, wireType: undefined, wireStyleName: undefined },
        { kind: "slideNumber", ...rect("slideNumber", w, h), align: "right", text: undefined, wireType: undefined, wireStyleName: undefined },
        { kind: "date", ...rect("date", w, h), align: undefined, text: undefined, wireType: undefined, wireStyleName: undefined },
      ])).toBe(true);
      // Slide 2: per-slide footer override.
      const footer2 = m.slides[2].fields!.find((f) => f.kind === "footer");
      expect(footer2?.text).toBe(EXPECTED_HEADER_FOOTER_OVERRIDE);
      // Slide 3: out-of-subset date preserved verbatim, cached text kept.
      const date3 = m.slides[3].fields!.find((f) => f.kind === "date");
      expect(date3?.wireType).toBe("datetime13");
      expect(date3?.text).toBe("14.5.2026");
    });

    it("a no-edit re-save is byte-identical part-for-part (AC-83)", async () => {
      const bytes = await buildPptxHeaderFooterFixtureBytes();
      const m = await readPresentationBytes(bytes, "pptx");
      const resaved = (await writePptx(m)).bytes;
      const before = await unzip(bytes);
      const after = await unzip(resaved);
      expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
      for (const [name, data] of before) {
        expect(equalBytes(after.get(name)!, data), `part ${name}`).toBe(true);
      }
    });

    it("a header/footer edit rebuilds only field-carrying parts; the rest stay equalBytes (AC-82)", async () => {
      const bytes = await buildPptxHeaderFooterFixtureBytes();
      const m = await readPresentationBytes(bytes, "pptx");
      // Mark a deck-wide header/footer change without altering the values.
      m.headerFooterDirty = true;
      for (const slide of m.slides) {
        if (slide.fields) markSlideEdited(slide, m.format);
      }
      const resaved = (await writePptx(m)).bytes;
      const before = await unzip(bytes);
      const after = await unzip(resaved);
      const fieldParts = new Set(
        [...before.keys()].filter((name) => /ppt\/slides\/slide[234]\.xml$/.test(name)),
      );
      for (const [name, data] of before) {
        if (fieldParts.has(name) || name === "ppt/presentation.xml") continue;
        expect(equalBytes(after.get(name)!, data), `part ${name} unchanged`).toBe(true);
      }
    });

    it("pins the sanctioned datetime* delta (AC-84): an auto date writes datetime3", async () => {
      const bytes = await buildPptxHeaderFooterFixtureBytes();
      const parts = await unzip(bytes);
      const decode = (u: Uint8Array): string => new TextDecoder().decode(u);
      const slide1 = decode(parts.get("ppt/slides/slide2.xml")!); // slide index 1
      expect(slide1).toContain(`type="datetime3"`); // d MMMM yyyy ⇄ datetime3
      // presentation.xml carries the title-slide suppression attribute.
      expect(decode(parts.get("ppt/presentation.xml")!)).toContain(`showSpecialPlsOnTitleSld="0"`);
    });
  });

  describe("odp", () => {
    it("reopens with the deck values + every field entry (equalStructural)", async () => {
      const bytes = await buildOdpHeaderFooterFixtureBytes();
      const m = await readPresentationBytes(bytes, "odp");
      const w = m.slideWidthCm;
      const h = m.slideHeightCm;
      // ODF has no showSpecialPlsOnTitleSld equivalent — the title slide's fields
      // are simply off, so its deck value does not survive.
      expect(equalStructural(m.headerFooter, { footerText: "Confidential" })).toBe(true);
      expect(m.slides[0].fields).toBeUndefined();
      expect(equalStructural(facts(m, 1), [
        { kind: "footer", ...rect("footer", w, h), align: "center", text: undefined, wireType: undefined, wireStyleName: undefined },
        { kind: "slideNumber", ...rect("slideNumber", w, h), align: "right", text: undefined, wireType: undefined, wireStyleName: undefined },
        { kind: "date", ...rect("date", w, h), align: undefined, text: undefined, wireType: undefined, wireStyleName: undefined },
      ])).toBe(true);
      expect(m.slides[2].fields!.find((f) => f.kind === "footer")?.text).toBe(EXPECTED_HEADER_FOOTER_OVERRIDE);
      // Out-of-subset date: foreign data style kept, ODF caches nothing.
      const date3 = m.slides[3].fields!.find((f) => f.kind === "date");
      expect(date3?.wireStyleName).toBe("D1");
      expect(date3?.text).toBeUndefined();
    });

    it("a no-edit re-save is byte-identical part-for-part (AC-83)", async () => {
      const bytes = await buildOdpHeaderFooterFixtureBytes();
      const m = await readPresentationBytes(bytes, "odp");
      const resaved = (await writeOdp(m)).bytes;
      const before = await unzip(bytes);
      const after = await unzip(resaved);
      for (const [name, data] of before) {
        expect(equalBytes(after.get(name)!, data), `part ${name}`).toBe(true);
      }
    });
  });
});
