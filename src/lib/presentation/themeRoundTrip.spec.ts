// SPEC-094 AC: the mixed themed deck in both formats — save → reopen
// reproduces `model.theme` and every reference exactly in pptx
// (`equalStructural` over the theme and the frames' modelled subset); odp
// reproduces the untransformed refs and the theme while transformed refs
// arrive as their literals (the one sanctioned odp delta, pinned here); a
// no-edit download of either stays byte-identical part-for-part
// (`equalBytes`, `core/equality.ts`, per part after `unzip`).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { equalStructural } from "../equality";
import { unzip } from "../core/zip";
import { DEFAULT_THEME, resolveColor } from "../core/theme";
import {
  buildSpec094OdpFixtureBytes,
  buildSpec094PptxFixtureBytes,
  SPEC094_LITERAL_RUN_COLOR,
  SPEC094_ODP_EXTENSION_BLOCK,
  SPEC094_REFS,
} from "./fixtures/build";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { applyTheme } from "./theme";
import { markSlideEdited } from "./text";
import type { PresentationModel, ShapeFrame, Slide, TextFrame } from "./model";
import { BUILT_IN_THEMES } from "../core/theme";

async function openPptx(): Promise<PresentationModel> {
  const parsed = await parsePptx(await buildSpec094PptxFixtureBytes());
  return { kind: "presentation", id: "d", title: "Theme fixture", format: "pptx", ...parsed };
}

async function openOdp(): Promise<PresentationModel> {
  const parsed = await parseOdp(await buildSpec094OdpFixtureBytes());
  return { kind: "presentation", id: "d", title: "Theme fixture", format: "odp", ...parsed };
}

/** The modelled subset the round trip pins: theme scheme + the mixed slide's
 * colour/font-bearing values. */
function subset(model: PresentationModel) {
  const mixed = model.slides[0];
  const [title, body, rect, ellipse, literal] = mixed.frames as [TextFrame, TextFrame, ShapeFrame, ShapeFrame, TextFrame];
  return {
    theme: model.theme ? { name: model.theme.name, colors: model.theme.colors, fonts: model.theme.fonts } : undefined,
    titleColor: title.paragraphs[0].runs[0].color,
    titleFont: title.paragraphs[0].runs[0].fontFamily,
    bodyFont: body.paragraphs[0].runs[0].fontFamily,
    rectFill: rect.fill,
    ellipseFill: ellipse.fill,
    literalColor: literal.paragraphs[0].runs[0].color,
  };
}

function dirtyAll(model: PresentationModel): void {
  for (const slide of model.slides) {
    markSlideEdited(slide, model.format);
  }
}

describe("pptx theme round trip (AC-50)", () => {
  it("reads the mixed deck's theme and every ref", async () => {
    const model = await openPptx();
    expect(model.theme?.name).toBe("Office");
    expect(model.theme?.colors).toEqual(DEFAULT_THEME.colors);
    expect(model.theme?.fonts).toEqual(DEFAULT_THEME.fonts);
    expect(model.theme?.partName).toBe("ppt/theme/theme1.xml");
    const s = subset(model);
    expect(s.titleColor).toEqual(SPEC094_REFS.titleColor);
    expect(s.titleFont).toEqual(SPEC094_REFS.titleFont);
    expect(s.bodyFont).toEqual(SPEC094_REFS.bodyFont);
    expect(s.rectFill).toEqual({ type: "solid", color: SPEC094_REFS.accentFill });
    expect(s.ellipseFill).toEqual({ type: "solid", color: SPEC094_REFS.lighterFill });
    expect(s.literalColor).toBe(SPEC094_LITERAL_RUN_COLOR);
  });

  it("save (rebuilt slides) → reopen reproduces the theme and every ref exactly", async () => {
    const model = await openPptx();
    dirtyAll(model);
    const reopened = await parsePptx((await writePptx(model)).bytes);
    expect(equalStructural(subset(model), subset({ ...model, ...reopened }))).toBe(true);
    // The preserve-path slide keeps its (unparseable) p:clrMapOvr content —
    // the reader's serialiser may add xmlns attributes, so pin the payload.
    expect(reopened.slides[1].preserved?.["p:clrMapOvr"]).toContain('<p:extLst uri="urn:example:unparsed"/>');
  });

  it("applying a theme keeps every non-theme part equalBytes (AC-29)", async () => {
    const model = await openPptx();
    const source = await unzip(await buildSpec094PptxFixtureBytes());
    applyTheme(model, BUILT_IN_THEMES[1]);
    expect(model.slides.every((slide: Slide) => slide.dirty === undefined)).toBe(true); // AC-28
    const written = await unzip((await writePptx(model)).bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, bytes] of source) {
      if (name === model.theme?.partName) {
        expect(equalBytes(bytes, written.get(name)!)).toBe(false); // the theme moved
      } else {
        expect(equalBytes(bytes, written.get(name)!), `${name} must stay byte-identical`).toBe(true);
      }
    }
    // And the rewritten theme part still reads back as the applied scheme.
    const reopened = await parsePptx((await writePptx(model)).bytes);
    expect(reopened.theme?.colors).toEqual(BUILT_IN_THEMES[1].colors);
    expect(reopened.theme?.fonts).toEqual(BUILT_IN_THEMES[1].fonts);
    // Refs still track; the literal run is untouched (AC-38).
    const s = subset({ ...model, ...reopened });
    expect(s.rectFill).toEqual({ type: "solid", color: SPEC094_REFS.accentFill });
    expect(s.literalColor).toBe(SPEC094_LITERAL_RUN_COLOR);
  });
});

describe("odp theme round trip (AC-51)", () => {
  it("reads the theme and the untransformed refs; transformed refs arrive as literals", async () => {
    const model = await openOdp();
    expect(model.theme?.name).toBe("Office");
    expect(model.theme?.colors).toEqual(DEFAULT_THEME.colors);
    expect(model.theme?.partName).toBe("styles.xml");
    const s = subset(model);
    // Untransformed refs survive the parent-style link.
    expect(s.rectFill).toEqual({ type: "solid", color: SPEC094_REFS.accentFill });
    expect(s.titleColor).toEqual(SPEC094_REFS.titleColor);
    // The one sanctioned delta: a transformed ref arrives as its literal…
    expect(s.ellipseFill).toEqual({
      type: "solid",
      color: resolveColor(SPEC094_REFS.lighterFill, DEFAULT_THEME),
    });
    // …and theme fonts as family names.
    expect(s.bodyFont).toBe(DEFAULT_THEME.fonts.minor);
    expect(s.literalColor).toBe(SPEC094_LITERAL_RUN_COLOR);
  });

  it("save (rebuilt pages) → reopen is a fixed point over the modelled subset", async () => {
    const model = await openOdp();
    dirtyAll(model);
    const reopened = await parseOdp((await writeOdp(model)).bytes);
    expect(equalStructural(subset(model), subset({ ...model, ...reopened }))).toBe(true);
    // The unrecognised extension theme block rides styles.xml verbatim.
    const parts = await unzip((await writeOdp(model)).bytes);
    expect(new TextDecoder().decode(parts.get("styles.xml")!)).toContain(SPEC094_ODP_EXTENSION_BLOCK);
  });

  it("a theme change re-serialises only ref-carrying pages; other parts stay equalBytes (AC-33)", async () => {
    const model = await openOdp();
    const source = await unzip(await buildSpec094OdpFixtureBytes());
    applyTheme(model, BUILT_IN_THEMES[1]);
    // Only the mixed slide carries refs — the preserve-path page stays verbatim.
    expect(model.slides[0].dirty).toBe(true);
    expect(model.slides[1].dirty).toBeUndefined();
    expect(model.slides[1].sourceXml).toBeDefined();
    const written = await unzip((await writeOdp(model)).bytes);
    for (const [name, bytes] of source) {
      if (name === "content.xml" || name === "styles.xml") {
        continue; // the two parts a theme change may rewrite
      }
      expect(equalBytes(bytes, written.get(name)!), `${name} must stay byte-identical`).toBe(true);
    }
    // The verbatim page's source markup is spliced back unchanged.
    expect(new TextDecoder().decode(written.get("content.xml")!)).toContain(model.slides[1].sourceXml!);
    // The extension block survives the theme splice.
    expect(new TextDecoder().decode(written.get("styles.xml")!)).toContain(SPEC094_ODP_EXTENSION_BLOCK);
  });
});

describe("no-edit download (AC-52)", () => {
  it("pptx stays byte-identical part-for-part", async () => {
    const model = await openPptx();
    const source = await unzip(await buildSpec094PptxFixtureBytes());
    const written = await unzip((await writePptx(model)).bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, bytes] of source) {
      expect(equalBytes(bytes, written.get(name)!), `${name} must stay byte-identical`).toBe(true);
    }
  });

  it("odp stays byte-identical part-for-part", async () => {
    const model = await openOdp();
    const source = await unzip(await buildSpec094OdpFixtureBytes());
    const written = await unzip((await writeOdp(model)).bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, bytes] of source) {
      expect(equalBytes(bytes, written.get(name)!), `${name} must stay byte-identical`).toBe(true);
    }
  });
});
