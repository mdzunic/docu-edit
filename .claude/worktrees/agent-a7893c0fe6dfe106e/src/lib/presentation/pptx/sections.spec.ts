// SPEC-099 — the pptx section wire mapping: golden equality both ways, the
// in-place splice, the sibling-`p:ext` preserve path and membership
// recomputed after a reorder.
import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import { unzip } from "../../core/zip";
import type { PresentationModel } from "../model";
import { readPresentationBytes } from "../open";
import { createBlankPresentation, addSlide, moveSlide } from "../slides";
import { addSection, renameSection, removeSection } from "../sections";
import { writePptx } from "./writer";
import { GOLDEN_SECTIONS, SECTIONS_GOLDEN } from "./sectionsGolden";
import {
  P14_NAMESPACE,
  readSections,
  sectionKeyFromWireId,
  sectionsExtXml,
  sectionWireId,
  spliceSectionLst,
} from "./sections";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const decoder = new TextDecoder();

function deck(n: number): PresentationModel {
  const model = createBlankPresentation();
  for (let i = 1; i < n; i++) {
    addSlide(model, i - 1);
  }
  delete model.structureDirty;
  return model;
}

function goldenModel(): PresentationModel {
  const model = deck(3);
  model.sections = structuredClone(GOLDEN_SECTIONS.sections);
  model.slides.forEach((slide, i) => (slide.sectionKey = GOLDEN_SECTIONS.sectionKeys[i]));
  return model;
}

function presentationXml(extLst = "", rootAttrs = ""): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:presentation xmlns:p="${P_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${rootAttrs}>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/>${extLst}</p:presentation>`
  );
}

async function presentationPart(bytes: Uint8Array): Promise<string> {
  return decoder.decode((await unzip(bytes)).get("ppt/presentation.xml")!);
}

describe("SECTIONS_GOLDEN", () => {
  it("is exactly what the writer emits for GOLDEN_SECTIONS", () => {
    expect(sectionsExtXml(goldenModel())).toBe(SECTIONS_GOLDEN);
  });

  it("reads back as GOLDEN_SECTIONS", () => {
    const doc = parseXml(presentationXml(`<p:extLst>${SECTIONS_GOLDEN}</p:extLst>`));
    const slides = deck(3).slides;
    const sections = readSections(doc.documentElement, slides, ["256", "257", "258"]);
    expect(sections).toEqual(GOLDEN_SECTIONS.sections);
    expect(slides.map((slide) => slide.sectionKey)).toEqual(GOLDEN_SECTIONS.sectionKeys);
  });
});

describe("readSections", () => {
  it("ignores unknown ids and joins an unassigned slide to the preceding run", () => {
    const list =
      `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="${P14_NAMESPACE}">` +
      `<p14:section name="A" id="{11111111-1111-4111-8111-111111111111}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="999"/></p14:sldIdLst></p14:section>` +
      `<p14:section name="B" id="{22222222-2222-4222-8222-222222222222}"><p14:sldIdLst><p14:sldId id="258"/></p14:sldIdLst></p14:section>` +
      `</p14:sectionLst></p:ext></p:extLst>`;
    const slides = deck(3).slides;
    const sections = readSections(parseXml(presentationXml(list)).documentElement, slides, ["256", "257", "258"]);
    expect(sections?.map((s) => s.name)).toEqual(["A", "B"]);
    // Slide 257 is in no section → joins A's run.
    expect(slides.map((slide) => slide.sectionKey)).toEqual([
      "{11111111-1111-4111-8111-111111111111}",
      "{11111111-1111-4111-8111-111111111111}",
      "{22222222-2222-4222-8222-222222222222}",
    ]);
  });

  it("is undefined when the deck declares no section list", () => {
    expect(readSections(parseXml(presentationXml()).documentElement, deck(3).slides, ["256", "257", "258"])).toBeUndefined();
  });
});

describe("section wire ids", () => {
  it("writes a source GUID verbatim and a minted key as a reserved GUID that reads back", () => {
    expect(sectionWireId("{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}")).toBe("{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}");
    const wire = sectionWireId("de-sec12");
    expect(wire).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
    expect(sectionKeyFromWireId(wire)).toBe("de-sec12");
    expect(sectionKeyFromWireId("not-a-guid")).toBeUndefined();
  });
});

describe("spliceSectionLst", () => {
  it("replaces an existing list in place, reusing its prefix and leaving every sibling p:ext byte-verbatim", () => {
    const before = `<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="urn:p15"/></p:ext>`;
    const after = `<p:ext uri="{2D200454-40CA-4A62-9FC3-DE9A4176ACB9}"><p15:notesGuideLst xmlns:p15="urn:p15"/></p:ext>`;
    const old =
      `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q14:sectionLst xmlns:q14="${P14_NAMESPACE}">` +
      `<q14:section name="Old" id="{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}"><q14:sldIdLst><q14:sldId id="258"/></q14:sldIdLst></q14:section>` +
      `</q14:sectionLst></p:ext>`;
    const xml = presentationXml(`<p:extLst>${before}${old}${after}</p:extLst>`);
    const out = spliceSectionLst(xml, goldenModel());
    expect(out.startsWith(xml.slice(0, xml.indexOf("<q14:sectionLst")))).toBe(true);
    expect(out).toContain(`${before}<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q14:sectionLst`);
    expect(out).toContain(`</q14:sectionLst></p:ext>${after}</p:extLst>`);
    expect(out).toContain(`<q14:section name="Intro" id="{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}">`);
    expect(out).not.toContain("Old");
    // The root is untouched (the prefix was already in use).
    expect(out.slice(0, out.indexOf("<p:sldIdLst>"))).toBe(xml.slice(0, xml.indexOf("<p:sldIdLst>")));
  });

  it("splices a p:ext before </p:extLst> when the list has none", () => {
    const sibling = `<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><x/></p:ext>`;
    const out = spliceSectionLst(presentationXml(`<p:extLst>${sibling}</p:extLst>`), goldenModel());
    expect(out).toContain(`<p:extLst>${sibling}${SECTIONS_GOLDEN}</p:extLst></p:presentation>`);
  });

  it("inserts a whole p:extLst last and declares p14 + mc:Ignorable on the root only when absent", () => {
    const out = spliceSectionLst(presentationXml(), goldenModel());
    expect(out).toContain(`<p:extLst>${SECTIONS_GOLDEN}</p:extLst></p:presentation>`);
    expect(out).toContain(`xmlns:p14="${P14_NAMESPACE}"`);
    expect(out).toContain(`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`);
    expect(out).toContain(`mc:Ignorable="p14"`);
    expect(() => parseXml(out)).not.toThrow();

    const declared = presentationXml(
      "",
      ` xmlns:p14="${P14_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p15"`,
    );
    const again = spliceSectionLst(declared, goldenModel());
    expect(again.match(/xmlns:p14=/g)).toHaveLength(2); // root + the list's own
    expect(again).toContain(`mc:Ignorable="p15 p14"`);
  });

  it("removes the section p:ext (and an emptied p:extLst) when the deck has no sections", () => {
    const xml = presentationXml(`<p:extLst>${SECTIONS_GOLDEN}</p:extLst>`);
    expect(spliceSectionLst(xml, deck(3))).toBe(presentationXml());
    expect(spliceSectionLst(presentationXml(), deck(3))).toBe(presentationXml());
  });
});

describe("writePptx with sections", () => {
  it("round-trips a from-scratch deck's sections and keys", async () => {
    const model = deck(4);
    addSection(model, 0, "Intro");
    addSection(model, 2, "Body");
    const reread = await readPresentationBytes((await writePptx(model)).bytes, "pptx");
    expect(reread.sections).toEqual(model.sections);
    expect(reread.slides.map((slide) => slide.sectionKey)).toEqual(model.slides.map((slide) => slide.sectionKey));
  });

  it("recomputes membership against the re-minted sldIds after a reorder", async () => {
    const source = deck(4);
    addSection(source, 0, "Intro");
    addSection(source, 2, "Body");
    const model = await readPresentationBytes((await writePptx(source)).bytes, "pptx");
    expect(model.opaque).toBeDefined();
    const order = model.slides.map((slide) => slide.id);
    moveSlide(model, 3, 0); // a Body slide to the front → joins Intro
    const reread = await readPresentationBytes((await writePptx(model)).bytes, "pptx");
    expect(reread.sections).toEqual(model.sections);
    expect(reread.slides.map((slide) => slide.sectionKey)).toEqual(model.slides.map((slide) => slide.sectionKey));
    expect(model.slides.map((slide) => slide.id)).toEqual([order[3], order[0], order[1], order[2]]);
  });

  it("a rename rewrites only the list; removing every section drops it", async () => {
    const source = deck(2);
    addSection(source, 0, "Only");
    const model = await readPresentationBytes((await writePptx(source)).bytes, "pptx");
    renameSection(model, model.sections![0].key, "Renamed");
    const renamed = await presentationPart((await writePptx(model)).bytes);
    expect(renamed).toContain(`name="Renamed"`);
    removeSection(model, model.sections![0].key, "sectionOnly");
    const removed = await presentationPart((await writePptx(model)).bytes);
    expect(removed).not.toContain("sectionLst");
    expect((await readPresentationBytes((await writePptx(model)).bytes, "pptx")).sections).toBeUndefined();
  });
});
