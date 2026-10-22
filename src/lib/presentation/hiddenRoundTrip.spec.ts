// SPEC-104 — `Slide.hidden` and its two wire mappings, round-tripped through
// the real readers and writers: pptx `p:sld/@show="0"` (a single-attribute
// splice inside `preserved["@root"]`), odp `presentation:visibility="hidden"`
// on the page's drawing-page style (pinned golden-first, odp/visibilityGolden).
// Presentation round trips assert per-part `equalBytes`, `equalStructural`
// over the `slides` array and explicit field-level `toEqual` — the
// DocumentModel-typed `equalSemantic`/`resolve` do not apply to a deck.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { base64ToBytes } from "../core/packageParts";
import { unzip } from "../core/zip";
import { equalStructural } from "../equality";
import {
  buildOdpHiddenFixtureBytes,
  buildPptxHiddenFixtureBytes,
  EXPECTED_HIDDEN_FLAGS,
  SPEC104_NOTES,
  SPEC104_TITLES,
} from "./fixtures/build";
import type { PresentationFormat, PresentationModel } from "./model";
import { readPresentationBytes } from "./open";
import { writeOdp } from "./odp/writer";
import { writePptx } from "./pptx/writer";
import { rootAttrsHidden, spliceShowAttr } from "./pptx/showAttr";
import {
  GOLDEN_ODP_HIDDEN_PAGE_ATTR,
  GOLDEN_ODP_HIDDEN_STYLE,
  IMPRESS_ODP_HIDDEN_STYLE,
} from "./odp/visibilityGolden";
import { slideOutline } from "./slideModel";
import { createBlankPresentation, setSlideHidden } from "./slides";

const FORMATS: { format: PresentationFormat; build: () => Promise<Uint8Array> }[] = [
  { format: "pptx", build: buildPptxHiddenFixtureBytes },
  { format: "odp", build: buildOdpHiddenFixtureBytes },
];

function write(model: PresentationModel): Promise<{ bytes: Uint8Array }> {
  return model.format === "odp" ? writeOdp(model) : writePptx(model);
}

function hiddenFlags(model: PresentationModel): boolean[] {
  return model.slides.map((slide) => slide.hidden === true);
}

/** The slide-bearing part of a package: each pptx slide part, or odp's
 * content.xml (every page lives in it). */
function slideParts(format: PresentationFormat, parts: Map<string, Uint8Array>): string[] {
  return format === "odp"
    ? ["content.xml"]
    : [...parts.keys()].filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
}

describe.each(FORMATS)("$format hidden slides (SPEC-104)", ({ format, build }) => {
  it("reads the fixture's hidden flags exactly — only `true` is ever stored", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    expect(hiddenFlags(model)).toEqual(EXPECTED_HIDDEN_FLAGS);
    expect(model.slides.map((slide) => slideOutline(slide).title)).toEqual(SPEC104_TITLES);
    expect(model.slides[0].notes).toEqual(SPEC104_NOTES);
    for (const slide of model.slides) {
      expect("hidden" in slide ? slide.hidden : true).toBe(true);
    }
  });

  it("the reader sets no dirty flag and mutates no opaque part (AC-4)", async () => {
    const bytes = await build();
    const source = await unzip(bytes);
    const model = await readPresentationBytes(bytes, format, "Deck");
    expect(model.structureDirty).toBeUndefined();
    for (const slide of model.slides) {
      expect(slide.dirty).toBeUndefined();
      expect(slide.notesDirty).toBeUndefined();
    }
    for (const part of model.opaque ?? []) {
      expect(equalBytes(base64ToBytes(part.data), source.get(part.name)!), `opaque ${part.name}`).toBe(true);
    }
  });

  it("an untouched deck downloads byte-identically, part by part (AC-38)", async () => {
    const bytes = await build();
    const source = await unzip(bytes);
    const model = await readPresentationBytes(bytes, format, "Deck");
    const written = await unzip((await write(model)).bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name}`).toBe(true);
    }
  });

  it("hide → save → reopen reproduces hidden exactly, and a re-save is byte-stable for the slide part (AC-39)", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    setSlideHidden(model, 0, true);
    setSlideHidden(model, 1, false);
    expect(hiddenFlags(model)).toEqual([true, false, false]);
    const saved = (await write(model)).bytes;
    const reopened = await readPresentationBytes(saved, format, "Deck");
    expect(hiddenFlags(reopened)).toEqual([true, false, false]);
    expect("hidden" in reopened.slides[1]).toBe(false);
    // Nothing else about the slides moved.
    const strip = (m: PresentationModel) =>
      m.slides.map((slide) => ({ frames: slide.frames, notes: slide.notes, background: slide.background }));
    expect(equalStructural(strip(reopened), strip(model))).toBe(true);
    expect(reopened.slides[0].notes).toEqual(SPEC104_NOTES);

    const resaved = await unzip((await write(reopened)).bytes);
    const first = await unzip(saved);
    for (const name of slideParts(format, first)) {
      expect(equalBytes(resaved.get(name)!, first.get(name)!), `slide part ${name}`).toBe(true);
    }
  });

  it("toggling records the edit on the slide alone — the deck structure is untouched", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    setSlideHidden(model, 2, true);
    expect(model.slides[2].dirty).toBe(true);
    expect(model.structureDirty).toBeUndefined();
    // Setting the state it already has is a no-op, not an edit.
    setSlideHidden(model, 0, false);
    expect(model.slides[0].dirty).toBeUndefined();
  });
});

describe("pptx p:sld/@show splice (SPEC-104)", () => {
  const ROOT = ` xmlns:a="urn:a" xmlns:p="urn:p" showMasterSp="0"`;

  it("reads show=\"0\" and show=\"false\" as hidden, anything else as shown", () => {
    expect(rootAttrsHidden(`${ROOT} show="0"`)).toBe(true);
    expect(rootAttrsHidden(`${ROOT} show="false"`)).toBe(true);
    expect(rootAttrsHidden(`${ROOT} show="1"`)).toBe(false);
    expect(rootAttrsHidden(ROOT)).toBe(false); // showMasterSp is not show
    expect(rootAttrsHidden(undefined)).toBe(false);
  });

  it("appends show=\"0\" after the last attribute, and removes it with its leading space", () => {
    const hidden = spliceShowAttr(ROOT, true);
    expect(hidden).toBe(`${ROOT} show="0"`);
    expect(spliceShowAttr(hidden, false)).toBe(ROOT);
  });

  it("is idempotent both ways and never moves another attribute", () => {
    const mid = ` xmlns:a="urn:a" show="false" showMasterSp="0"`;
    expect(spliceShowAttr(mid, true)).toBe(mid);
    expect(spliceShowAttr(spliceShowAttr(ROOT, true), true)).toBe(`${ROOT} show="0"`);
    expect(spliceShowAttr(ROOT, false)).toBe(ROOT);
    expect(spliceShowAttr(mid, false)).toBe(` xmlns:a="urn:a" showMasterSp="0"`);
    // A shown value is replaced in place, never moved.
    expect(spliceShowAttr(` xmlns:a="urn:a" show="1" showMasterSp="0"`, true)).toBe(
      ` xmlns:a="urn:a" show="0" showMasterSp="0"`,
    );
    expect(spliceShowAttr(` show="1"`, false)).toBe(` show="1"`);
  });

  it("a rebuilt hidden slide keeps every other root attribute byte-for-byte", async () => {
    const model = await readPresentationBytes(await buildPptxHiddenFixtureBytes(), "pptx", "Deck");
    const rootBefore = model.slides[1].preserved?.["@root"] ?? "";
    expect(rootBefore.endsWith(` show="0"`)).toBe(true);
    // Edit the hidden slide in some other way: it rebuilds, still hidden.
    model.slides[1].dirty = true;
    const written = await unzip((await writePptx(model)).bytes);
    const xml = new TextDecoder().decode(written.get(model.slides[1].partName!)!);
    expect(xml).toContain(`<p:sld${rootBefore}>`);
  });
});

describe("odp presentation:visibility golden (SPEC-104)", () => {
  it("the serialiser emits the canonical hidden drawing-page style and reference", async () => {
    const model = createBlankPresentation();
    model.format = "odp";
    model.slides[0].hidden = true;
    const content = new TextDecoder().decode((await unzip((await writeOdp(model)).bytes)).get("content.xml")!);
    expect(content).toContain(GOLDEN_ODP_HIDDEN_STYLE);
    expect(content).toMatch(new RegExp(`<draw:page [^>]*${GOLDEN_ODP_HIDDEN_PAGE_ATTR}>`));
  });

  it("the reader accepts the golden and Impress's own phrasing", async () => {
    const bytes = await buildOdpHiddenFixtureBytes();
    const parts = await unzip(bytes);
    const content = new TextDecoder().decode(parts.get("content.xml")!);
    expect(content).toContain(GOLDEN_ODP_HIDDEN_STYLE);
    const impress = content
      .replace(GOLDEN_ODP_HIDDEN_STYLE, IMPRESS_ODP_HIDDEN_STYLE)
      .replace(`draw:style-name="de-DP1"`, `draw:style-name="dp2"`);
    parts.set("content.xml", new TextEncoder().encode(impress));
    const { zip } = await import("../core/zipWriter");
    const rezipped = await zip([...parts].map(([name, data]) => ({ name, data, store: name === "mimetype" })));
    const model = await readPresentationBytes(rezipped, "odp", "Deck");
    expect(hiddenFlags(model)).toEqual(EXPECTED_HIDDEN_FLAGS);
  });

  it("showing a hidden page drops its visibility and keeps its other attributes", async () => {
    const model = await readPresentationBytes(await buildOdpHiddenFixtureBytes(), "odp", "Deck");
    setSlideHidden(model, 1, false);
    const bytes = (await writeOdp(model)).bytes;
    const content = new TextDecoder().decode((await unzip(bytes)).get("content.xml")!);
    expect(hiddenFlags(await readPresentationBytes(bytes, "odp", "Deck"))).toEqual([false, false, false]);
    // The page's style stated nothing but its visibility, so the shown page
    // references no drawing-page style at all; its other attributes stay.
    const pageTag = /<draw:page draw:name="page_2"[^>]*>/.exec(content)?.[0] ?? "";
    expect(pageTag).toBe(`<draw:page draw:name="page_2" draw:master-page-name="Default" presentation:presentation-page-layout-name="de-AL2">`);
  });

  it("a visibility-only change keeps a transition the editor does not model", async () => {
    const model = await readPresentationBytes(await buildOdpHiddenFixtureBytes(), "odp", "Deck");
    // Give page 3 an out-of-subset transition in its own drawing-page style.
    const exotic = `smil:type="exoticWipe" smil:subtype="spiral" presentation:transition-speed="slow"`;
    const envelope = model.odpContentEnvelope!;
    envelope.prefix = envelope.prefix.replace(
      "</office:automatic-styles>",
      `<style:style style:name="dpX" style:family="drawing-page"><style:drawing-page-properties ${exotic}/></style:style></office:automatic-styles>`,
    );
    const page = model.slides[2];
    page.preserved = { ...page.preserved, "@page": `${page.preserved?.["@page"] ?? ""} draw:style-name="dpX"` };
    setSlideHidden(model, 2, true);
    const content = new TextDecoder().decode((await unzip((await writeOdp(model)).bytes)).get("content.xml")!);
    expect(content).toContain(`<style:drawing-page-properties ${exotic} presentation:visibility="hidden"/>`);
  });
});
