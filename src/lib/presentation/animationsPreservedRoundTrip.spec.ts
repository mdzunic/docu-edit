// SPEC-086 AC-4 on a REBUILT slide: out-of-subset timing is kept verbatim, but
// it names its targets by source id — so after an edit it must still point at
// the same shapes (pptx re-mints `p:cNvPr id`s 2, 3, …; odp drops ids a page
// rebuild does not re-emit), and a target the user removed must take the
// timing with it rather than leave it dangling. Mirrors
// animationsRoundTrip.spec.ts's read → mutate → write → re-read shape, on decks
// whose source ids are deliberately non-contiguous.
import { describe, expect, it } from "vitest";
import { buildOdpAnimationsFixtureBytes, buildPptxAnimationsFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { markSlideEdited } from "./text";
import { deleteFrame, duplicateFrame, groupFrames } from "./frames";
import { hasPreservedTiming } from "./animations";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import type { PresentationModel, Slide, TextFrame } from "./model";

const enc = (text: string): Uint8Array => new TextEncoder().encode(text);
const dec = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

/** Title text edit — the dirtying edit that leaves every timing target alone. */
function editTitle(slide: Slide, format: "pptx" | "odp"): void {
  const title = slide.frames.find((f): f is TextFrame => f.kind === "text" && f.role === "title")!;
  title.paragraphs[0].runs[0].text = "Edited";
  markSlideEdited(slide, format);
}

function frameByText(slide: Slide, text: string): Slide["frames"][number] {
  const frame = slide.frames.find(
    (f) => (f.kind === "text" || f.kind === "shape") && (f.paragraphs ?? []).some((p) => p.runs.some((r) => r.text === text)),
  );
  if (!frame) {
    throw new Error(`no frame with "${text}"`);
  }
  return frame;
}

// ---------------------------------------------------------------------------
// pptx
// ---------------------------------------------------------------------------

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SUB_SHAPE = "{4F0D2C6B-1A2B-4C3D-8E9F-0A1B2C3D4E5F}";

function pptxSp(id: number, name: string, ph: string, texts: string[], x: number): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/>${texts.map((t) => `<a:p><a:r><a:t>${t}</a:t></a:r></a:p>`).join("")}</p:txBody></p:sp>`
  );
}

/** A SmartArt-style graphicFrame: read as a foreign frame, re-emitted
 * verbatim WITH its source id 2 — which a naive re-mint would hand out again. */
const PPTX_DIAGRAM =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="6400800" y="3657600"/><a:ext cx="1828800" cy="1371600"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic></p:graphicFrame>`;

/** Out of subset on several counts (an emphasis preset, a by-paragraph text
 * build, an unmapped modern entrance, a diagram sub-shape build), targeting
 * the body (12), the box (20) and the diagram (2) — plus a `p:bldLst`. */
const PPTX_TIMING =
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
  `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="5" presetID="8" presetClass="emph" presetSubtype="0" grpId="0" nodeType="clickEffect">` +
  `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:animRot by="21600000"><p:cBhvr><p:cTn id="6" dur="2000" fill="hold"/>` +
  `<p:tgtEl><p:spTgt spid="12"><p:txEl><p:pRg st="1" end="1"/></p:txEl></p:spTgt></p:tgtEl>` +
  `<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>` +
  `</p:childTnLst></p:cTn></p:par>` +
  `<p:par><p:cTn id="7" presetID="53" presetClass="entr" presetSubtype="16" fill="hold" nodeType="withEffect">` +
  `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="8" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="20"/></p:tgtEl>` +
  `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par>` +
  `<p:par><p:cTn id="9" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">` +
  `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="10" dur="500"/>` +
  `<p:tgtEl><p:spTgt spid="2"><p:subSp spid="${SUB_SHAPE}"/></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect>` +
  `</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn>` +
  `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
  `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
  `</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>` +
  `<p:bldLst><p:bldP spid="12" grpId="0" build="p"/><p:bldGraphic spid="2" grpId="0"><p:bldAsOne/></p:bldGraphic></p:bldLst>` +
  `</p:timing>`;

/** Title (7), diagram (2), body (12), box (20), in z-order. */
const PPTX_SLIDE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  pptxSp(7, "Title 6", `<p:ph type="title"/>`, ["Exotic"], 457200) +
  PPTX_DIAGRAM +
  pptxSp(12, "Content 11", `<p:ph type="body" idx="1"/>`, ["First point", "Second point"], 3657600) +
  pptxSp(20, "Box 19", "", ["Box text"], 6858000) +
  `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${PPTX_TIMING}</p:sld>`;

/** The SPEC-086 animation fixture with slide 3 swapped for `slideXml`. */
async function pptxDeck(slideXml = PPTX_SLIDE): Promise<PresentationModel> {
  const parts = await unzip(await buildPptxAnimationsFixtureBytes());
  parts.set("ppt/slides/slide3.xml", enc(slideXml));
  const bytes = await zip([...parts].map(([name, data]) => ({ name, data })));
  return readPresentationBytes(bytes, "pptx", "Deck");
}

async function savedPptxSlide3(model: PresentationModel): Promise<string> {
  const { writePptx } = await import("./pptx/writer");
  const parts = await unzip((await writePptx(model)).bytes);
  return dec(parts.get("ppt/slides/slide3.xml")!);
}

/** The `p:cNvPr id` of the `p:sp` holding `text`. */
function spIdHolding(xml: string, text: string): string {
  const sp = Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (m) => m[0]).find((s) => s.includes(`>${text}<`));
  const id = sp ? /<p:cNvPr id="(\d+)"/.exec(sp)?.[1] : undefined;
  if (!id) {
    throw new Error(`no p:sp with "${text}"`);
  }
  return id;
}

const shapeIds = (xml: string): string[] => Array.from(xml.matchAll(/<p:cNvPr id="(\d+)"/g), (m) => m[1]);
/** The slide's `p:timing`, less the `xmlns:p` the reader's serialiser adds. */
const timingOf = (xml: string): string | undefined =>
  /<p:timing\b[\s\S]*<\/p:timing>/.exec(xml)?.[0].replace(/^<p:timing xmlns:p="[^"]*">/, "<p:timing>");
const spTgtIds = (timing: string): string[] => Array.from(timing.matchAll(/<p:spTgt spid="(\d+)"/g), (m) => m[1]);

describe("AC-4: pptx preserved timing on an edited slide", () => {
  it("records the source id of each frame the timing targets, and only those", async () => {
    const slide = (await pptxDeck()).slides[2];
    expect(slide.animations).toBeUndefined();
    expect(slide.frames.map((f) => f.sourceWireId)).toEqual([undefined, "2", "12", "20"]);
  });

  it("re-points every reference at the re-minted ids, stepping over the verbatim diagram's id", async () => {
    const model = await pptxDeck();
    editTitle(model.slides[2], "pptx");
    const xml = await savedPptxSlide3(model);
    // Ids stay unique: the diagram keeps 2, so the title re-mints to 3.
    const ids = shapeIds(xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(spIdHolding(xml, "Edited")).toBe("3");
    const body = spIdHolding(xml, "First point");
    const box = spIdHolding(xml, "Box text");
    expect([body, box]).toEqual(["4", "5"]);
    const timing = timingOf(xml)!;
    // Emphasis → body, entrance → box, diagram build → the diagram (2), in order.
    expect(spTgtIds(timing)).toEqual([body, box, "2"]);
    expect(timing).toContain(`<p:bldP spid="${body}" grpId="0" build="p"/>`);
    expect(timing).toContain(`<p:bldGraphic spid="2" grpId="0">`);
    // A diagram sub-shape is named by GUID, never re-pointed.
    expect(timing).toContain(`<p:subSp spid="${SUB_SHAPE}"/>`);
    // Nothing else in the tree moved.
    expect(timing.replace(/spid="\d+"/g, "")).toBe(timingOf(PPTX_SLIDE)!.replace(/spid="\d+"/g, ""));
  });

  it("reopens to the same targets and re-saves the rebuilt timing byte-stable", async () => {
    const model = await pptxDeck();
    editTitle(model.slides[2], "pptx");
    const first = await savedPptxSlide3(model);
    const { writePptx } = await import("./pptx/writer");
    const back = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const slide = back.slides[2];
    expect(slide.animations).toBeUndefined();
    expect(slide.frames.map((f) => f.sourceWireId)).toEqual([undefined, "2", "4", "5"]);
    markSlideEdited(slide, "pptx");
    expect(timingOf(first)).toBeDefined();
    expect(timingOf(await savedPptxSlide3(back))).toBe(timingOf(first));
  });

  it("deleting an untargeted frame keeps the timing, re-pointed past the gap", async () => {
    const model = await pptxDeck();
    const slide = model.slides[2];
    deleteFrame(model, 2, slide.frames[0].id); // the title
    expect(hasPreservedTiming(slide)).toBe(true);
    const xml = await savedPptxSlide3(model);
    expect(spTgtIds(timingOf(xml)!)).toEqual([spIdHolding(xml, "First point"), spIdHolding(xml, "Box text"), "2"]);
  });

  it("deleting a targeted frame drops the preserved timing whole — never a dangling spid", async () => {
    const model = await pptxDeck();
    const slide = model.slides[2];
    deleteFrame(model, 2, frameByText(slide, "Box text").id);
    expect(hasPreservedTiming(slide)).toBe(false);
    expect(slide.animations).toBeUndefined();
    expect(slide.frames.some((f) => f.sourceWireId !== undefined)).toBe(false);
    expect(timingOf(await savedPptxSlide3(model))).toBeUndefined();
  });

  it("grouping a targeted frame drops the preserved timing (grouped frames are unanimatable)", async () => {
    const model = await pptxDeck();
    const slide = model.slides[2];
    groupFrames(model, 2, [slide.frames[0].id, frameByText(slide, "First point").id]);
    expect(hasPreservedTiming(slide)).toBe(false);
  });

  it("a duplicated frame is no target: the timing keeps pointing at the original", async () => {
    const model = await pptxDeck();
    const slide = model.slides[2];
    const copyId = duplicateFrame(model, 2, frameByText(slide, "First point").id)!;
    expect(slide.frames.find((f) => f.id === copyId)?.sourceWireId).toBeUndefined();
    const xml = await savedPptxSlide3(model);
    const bodies = Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (m) => m[0]).filter((s) => s.includes(">First point<"));
    expect(bodies).toHaveLength(2);
    const [original, copy] = bodies.map((s) => /<p:cNvPr id="(\d+)"/.exec(s)![1]);
    expect(spTgtIds(timingOf(xml)!)[0]).toBe(original);
    expect(copy).not.toBe(original);
  });

  it("a rebuild that cannot resolve a reference drops the timing rather than guess", async () => {
    const model = await pptxDeck();
    const slide = model.slides[2];
    // A model stored before frames recorded their source ids.
    for (const frame of slide.frames) {
      delete frame.sourceWireId;
    }
    editTitle(slide, "pptx");
    expect(timingOf(await savedPptxSlide3(model))).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// pptx by-paragraph builds: `p:txEl/p:pRg` names text by `a:p` INDEX
// ---------------------------------------------------------------------------
// The model reads each `a:br` as a new paragraph (and drops empty ones), and a
// rebuild writes each model paragraph as its own `a:p` — so on a rebuilt
// slide every range must be re-pointed at the `a:p`s holding the same text,
// and a range whose paragraph is gone must take the timing with it.

/** A target narrowed to a text range of body 12. */
const bodyRangeTarget = (range: string): string =>
  `<p:tgtEl><p:spTgt spid="12"><p:txEl>${range}</p:txEl></p:spTgt></p:tgtEl>`;

/** A PowerPoint "Fade, by paragraph" build: one click group per range, each a
 * `p:set` + `p:animEffect` on body 12's `p:pRg` — a MODELLED preset, kept
 * preserved only because it targets text by range — plus the `p:bldP`. */
function byParagraphTiming(ranges: string[]): string {
  let id = 3;
  const group = (range: string): string =>
    `<p:par><p:cTn id="${id++}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id++}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id++}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:set><p:cBhvr><p:cTn id="${id++}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${bodyRangeTarget(range)}` +
    `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
    `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${id++}" dur="500"/>${bodyRangeTarget(range)}</p:cBhvr></p:animEffect>` +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>`;
  const groups = ranges.map(group).join("");
  return (
    `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groups}</p:childTnLst></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
    `</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>` +
    `<p:bldLst><p:bldP spid="12" grpId="0" build="p"/></p:bldLst></p:timing>`
  );
}

/** Body 12's source `a:p`s: 0 = "Line one" `a:br` "Line two", 1 = "Second
 * point", 2 = empty (PowerPoint's build skips it; the reader drops it), 3 =
 * "Third point". */
const PPTX_BR_BODY =
  `<p:sp><p:nvSpPr><p:cNvPr id="12" name="Content 11"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="457200" y="1828800"/><a:ext cx="8229600" cy="3657600"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/>` +
  `<a:p><a:r><a:t>Line one</a:t></a:r><a:br/><a:r><a:t>Line two</a:t></a:r></a:p>` +
  `<a:p><a:r><a:t>Second point</a:t></a:r></a:p>` +
  `<a:p/>` +
  `<a:p><a:r><a:t>Third point</a:t></a:r></a:p>` +
  `</p:txBody></p:sp>`;

/** Title (7) and body (12), under `timing`. */
function pptxBrSlide(timing: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    pptxSp(7, "Title 6", `<p:ph type="title"/>`, ["Exotic"], 457200) +
    PPTX_BR_BODY +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${timing}</p:sld>`
  );
}

const BY_PARAGRAPH_SLIDE = pptxBrSlide(
  byParagraphTiming([`<p:pRg st="0" end="0"/>`, `<p:pRg st="1" end="1"/>`, `<p:pRg st="3" end="3"/>`]),
);

/** The text of each `a:p` in the saved `p:sp` holding `text`. */
function spParagraphTexts(xml: string, text: string): string[] {
  const sp = Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (m) => m[0]).find((s) => s.includes(`>${text}<`));
  if (!sp) {
    throw new Error(`no p:sp with "${text}"`);
  }
  return Array.from(sp.matchAll(/<a:p>([\s\S]*?)<\/a:p>|<a:p\/>/g), (m) =>
    Array.from((m[1] ?? "").matchAll(/<a:t>([^<]*)<\/a:t>/g), (t) => t[1]).join(""),
  );
}

/** The `a:p` text each effect's paragraph range reveals in the saved body
 * (the `p:sp` holding `bodyText`) — one entry per effect (its `p:set` and
 * `p:animEffect` share the range). */
function revealed(xml: string, bodyText = "Third point"): string[][] {
  const timing = timingOf(xml);
  if (!timing) {
    throw new Error("no p:timing");
  }
  const body = spParagraphTexts(xml, bodyText);
  const ranges = Array.from(timing.matchAll(/<p:pRg st="(\d+)" end="(\d+)"\/>/g), (m) => [Number(m[1]), Number(m[2])]);
  return ranges.filter((_, i) => i % 2 === 0).map(([st, end]) => body.slice(st, end + 1));
}

const bodyOf = (slide: Slide): TextFrame => frameByText(slide, "Third point") as TextFrame;
/** A timing less its shape and paragraph references — what must not move. */
const withoutReferences = (timing: string): string => timing.replace(/spid="\d+"/g, "").replace(/<p:pRg [^>]*>/g, "");

describe("AC-4: pptx by-paragraph build (p:pRg) on an edited slide", () => {
  it("keeps the build preserved — a modelled preset on a paragraph range is out of subset — and records each line's source a:p", async () => {
    const slide = (await pptxDeck(BY_PARAGRAPH_SLIDE)).slides[2];
    expect(slide.animations).toBeUndefined();
    expect(hasPreservedTiming(slide)).toBe(true);
    const body = bodyOf(slide);
    expect(body.paragraphs.map((p) => p.runs.map((r) => r.text).join(""))).toEqual([
      "Line one",
      "Line two",
      "Second point",
      "Third point",
    ]);
    // Both a:br lines share a:p 0; the empty a:p 2 left no line.
    expect(body.paragraphs.map((p) => p.pptxSourceParagraph)).toEqual([0, 0, 1, 3]);
    // Only the frame the ranges address keeps indices.
    expect(slide.frames[0].kind === "text" && slide.frames[0].paragraphs[0].pptxSourceParagraph).toBeUndefined();
  });

  it("an untouched slide re-saves its source bytes, ranges and all", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    expect(await savedPptxSlide3(model)).toBe(BY_PARAGRAPH_SLIDE);
  });

  it("re-points each range at the a:p now holding its text when another frame's edit rebuilds the slide", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    editTitle(model.slides[2], "pptx");
    const xml = await savedPptxSlide3(model);
    expect(spParagraphTexts(xml, "Third point")).toEqual(["Line one", "Line two", "Second point", "Third point"]);
    // a:p 0 split at its a:br spans two a:p now; everything after shifts.
    expect(revealed(xml)).toEqual([["Line one", "Line two"], ["Second point"], ["Third point"]]);
    const timing = timingOf(xml)!;
    expect(Array.from(timing.matchAll(/<p:pRg [^>]*>/g), (m) => m[0])).toEqual([
      `<p:pRg st="0" end="1"/>`,
      `<p:pRg st="0" end="1"/>`,
      `<p:pRg st="2" end="2"/>`,
      `<p:pRg st="2" end="2"/>`,
      `<p:pRg st="3" end="3"/>`,
      `<p:pRg st="3" end="3"/>`,
    ]);
    // The shape reference moved with the rebuild; nothing else in the tree did.
    const body = spIdHolding(xml, "Third point");
    expect(new Set(spTgtIds(timing))).toEqual(new Set([body]));
    expect(timing).toContain(`<p:bldP spid="${body}" grpId="0" build="p"/>`);
    expect(withoutReferences(timing)).toBe(withoutReferences(timingOf(BY_PARAGRAPH_SLIDE)!));
  });

  it("reopens to the rebuilt numbering and re-saves the timing byte-stable", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    editTitle(model.slides[2], "pptx");
    const first = await savedPptxSlide3(model);
    const { writePptx } = await import("./pptx/writer");
    const back = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const slide = back.slides[2];
    expect(slide.animations).toBeUndefined();
    expect(bodyOf(slide).paragraphs.map((p) => p.pptxSourceParagraph)).toEqual([0, 1, 2, 3]);
    markSlideEdited(slide, "pptx");
    expect(timingOf(await savedPptxSlide3(back))).toBe(timingOf(first));
  });

  it("follows paragraphs the user adds: a new line before a target shifts it, a split line stays in its range", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    const slide = model.slides[2];
    const body = bodyOf(slide);
    // A new paragraph typed before "Second point", and "Third point" split in
    // two by Enter (the editor carries the source index onto the new line).
    body.paragraphs.splice(2, 0, { runs: [{ text: "Inserted" }] });
    body.paragraphs.splice(4, 1, { ...body.paragraphs[4], runs: [{ text: "Third" }] }, {
      ...body.paragraphs[4],
      runs: [{ text: "point" }],
    });
    markSlideEdited(slide, "pptx");
    const xml = await savedPptxSlide3(model);
    expect(spParagraphTexts(xml, "Inserted")).toEqual(["Line one", "Line two", "Inserted", "Second point", "Third", "point"]);
    expect(revealed(xml, "Inserted")).toEqual([["Line one", "Line two"], ["Second point"], ["Third", "point"]]);
  });

  it("a targeted paragraph the user deleted drops the timing whole — never a dangling or shifted range", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    const slide = model.slides[2];
    bodyOf(slide).paragraphs.splice(2, 1); // "Second point"
    markSlideEdited(slide, "pptx");
    const xml = await savedPptxSlide3(model);
    expect(spParagraphTexts(xml, "Third point")).toEqual(["Line one", "Line two", "Third point"]);
    expect(timingOf(xml)).toBeUndefined();
  });

  it("a rebuilt body with no record of its source paragraphs drops the timing rather than guess", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    const slide = model.slides[2];
    // A model stored before lines recorded their source a:p.
    for (const para of bodyOf(slide).paragraphs) {
      delete para.pptxSourceParagraph;
    }
    editTitle(slide, "pptx");
    expect(timingOf(await savedPptxSlide3(model))).toBeUndefined();
  });

  it("a character range on a rebuilt body drops the timing — its offsets have no record", async () => {
    const model = await pptxDeck(pptxBrSlide(byParagraphTiming([`<p:charRg st="0" end="7"/>`])));
    const slide = model.slides[2];
    expect(slide.animations).toBeUndefined();
    editTitle(slide, "pptx");
    expect(timingOf(await savedPptxSlide3(model))).toBeUndefined();
  });

  it("a duplicated body records no source paragraphs: the ranges keep pointing at the original", async () => {
    const model = await pptxDeck(BY_PARAGRAPH_SLIDE);
    const slide = model.slides[2];
    const copyId = duplicateFrame(model, 2, bodyOf(slide).id)!;
    const copy = slide.frames.find((f) => f.id === copyId) as TextFrame;
    expect(copy.paragraphs.every((p) => p.pptxSourceParagraph === undefined)).toBe(true);
    const xml = await savedPptxSlide3(model);
    const bodies = Array.from(xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g), (m) => m[0]).filter((s) => s.includes(">Third point<"));
    expect(bodies).toHaveLength(2);
    const original = /<p:cNvPr id="(\d+)"/.exec(bodies[0])![1];
    expect(new Set(spTgtIds(timingOf(xml)!))).toEqual(new Set([original]));
    expect(revealed(xml)).toEqual([["Line one", "Line two"], ["Second point"], ["Third point"]]);
  });
});

// ---------------------------------------------------------------------------
// odp
// ---------------------------------------------------------------------------

const ANIM_NS = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
const SMIL_NS = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";

/** Impress-shaped: a by-paragraph emphasis on the body's two paragraphs (id4,
 * id5 — `text:p` targets), and an interactive sequence triggered by clicking
 * the rect (id9 — referenced only through `smil:begin`/`smil:end`) that
 * reveals the body frame (id3). The prefixes are bound on the content.xml
 * root, as Impress writes them. */
const ODP_ANIM =
  `<anim:par presentation:node-type="timing-root">` +
  `<anim:seq presentation:node-type="main-sequence"><anim:par smil:begin="next"><anim:par smil:begin="0s">` +
  `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin">` +
  `<anim:animateTransform smil:dur="2s" smil:fill="hold" smil:targetElement="id4" svg:type="rotate" smil:by="360"/></anim:par>` +
  `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="with-previous" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin">` +
  `<anim:animateTransform smil:dur="2s" smil:fill="hold" smil:targetElement="id5" svg:type="rotate" smil:by="360"/></anim:par>` +
  `</anim:par></anim:par></anim:seq>` +
  `<anim:seq presentation:node-type="interactive-sequence" smil:begin="id9.click" smil:end="id9.click">` +
  `<anim:par smil:begin="0s"><anim:par smil:begin="0s">` +
  `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-appear">` +
  `<anim:set smil:begin="0s" smil:dur="0.001s" smil:fill="hold" smil:targetElement="id3" smil:attributeName="visibility" smil:to="visible"/>` +
  `</anim:par></anim:par></anim:par></anim:seq>` +
  `</anim:par>`;

const ODP_PAGE =
  `<draw:page draw:name="page3">` +
  `<draw:frame presentation:class="title" svg:x="2cm" svg:y="1cm" svg:width="29cm" svg:height="3cm"><draw:text-box><text:p>Exotic</text:p></draw:text-box></draw:frame>` +
  `<draw:frame presentation:class="body" xml:id="id3" draw:id="id3" svg:x="2cm" svg:y="5cm" svg:width="20cm" svg:height="8cm">` +
  `<draw:text-box><text:p xml:id="id4">First point</text:p><text:p xml:id="id5">Second point</text:p></draw:text-box></draw:frame>` +
  `<draw:rect xml:id="id9" draw:id="id9" svg:x="24cm" svg:y="5cm" svg:width="6cm" svg:height="4cm"><draw:text-box><text:p>Box text</text:p></draw:text-box></draw:rect>` +
  ODP_ANIM +
  `</draw:page>`;

/** The SPEC-086 odp animation fixture with page 3 swapped for ODP_PAGE (and
 * the anim/smil prefixes declared on the root). */
async function odpDeck(): Promise<PresentationModel> {
  const parts = await unzip(await buildOdpAnimationsFixtureBytes());
  const content = dec(parts.get("content.xml")!)
    .replace(/<office:document-content /, `<office:document-content xmlns:anim="${ANIM_NS}" xmlns:smil="${SMIL_NS}" `)
    .replace(/<draw:page draw:name="page3">[\s\S]*?<\/draw:page>/, ODP_PAGE);
  parts.set("content.xml", enc(content));
  // mimetype must stay first + stored.
  const ordered = [...parts].toSorted(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
  const bytes = await zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
  return readPresentationBytes(bytes, "odp", "Deck");
}

async function savedOdpPage3(model: PresentationModel): Promise<string> {
  const { writeOdp } = await import("./odp/writer");
  const parts = await unzip((await writeOdp(model)).bytes);
  const pages = Array.from(dec(parts.get("content.xml")!).matchAll(/<draw:page\b[\s\S]*?<\/draw:page>/g), (m) => m[0]);
  return pages[2];
}

describe("AC-4: odp preserved anim on an edited page", () => {
  it("keeps the anim byte-verbatim (no generated ns1: prefixes) and records its targets", async () => {
    const slide = (await odpDeck()).slides[2];
    expect(slide.animations).toBeUndefined();
    expect(slide.preserved?.["anim"]).toBe(ODP_ANIM);
    expect(slide.frames.map((f) => f.sourceWireId)).toEqual([undefined, "id3", "id9"]);
    const body = frameByText(slide, "First point") as TextFrame;
    expect(body.paragraphs.map((p) => p.sourceWireId)).toEqual(["id4", "id5"]);
  });

  it("re-emits each target's source id — frames and by-paragraph text:p — with the anim untouched", async () => {
    const model = await odpDeck();
    editTitle(model.slides[2], "odp");
    const page = await savedOdpPage3(model);
    expect(page).toContain("Edited");
    expect(page).toMatch(/<draw:frame presentation:class="body" xml:id="id3" draw:id="id3"/);
    expect(page).toMatch(/<text:p\b[^>]* xml:id="id4">First point</);
    expect(page).toMatch(/<text:p\b[^>]* xml:id="id5">Second point</);
    expect(page).toMatch(/<draw:rect xml:id="id9" draw:id="id9"/);
    expect(page).toContain(ODP_ANIM);
  });

  it("reopens to the same targets", async () => {
    const model = await odpDeck();
    editTitle(model.slides[2], "odp");
    const { writeOdp } = await import("./odp/writer");
    const back = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    const slide = back.slides[2];
    expect(slide.preserved?.["anim"]).toBe(ODP_ANIM);
    expect(slide.frames.map((f) => f.sourceWireId)).toEqual([undefined, "id3", "id9"]);
    expect((frameByText(slide, "First point") as TextFrame).paragraphs.map((p) => p.sourceWireId)).toEqual(["id4", "id5"]);
  });

  it("deleting the click-trigger shape drops the anim whole; deleting an untargeted one keeps it", async () => {
    const kept = await odpDeck();
    deleteFrame(kept, 2, kept.slides[2].frames[0].id); // the title
    expect(await savedOdpPage3(kept)).toContain(ODP_ANIM);

    const dropped = await odpDeck();
    const slide = dropped.slides[2];
    deleteFrame(dropped, 2, frameByText(slide, "Box text").id);
    expect(hasPreservedTiming(slide)).toBe(false);
    expect(await savedOdpPage3(dropped)).not.toContain("<anim:");
  });

  it("a targeted paragraph the edit removed drops the anim at save — never a dangling targetElement", async () => {
    const model = await odpDeck();
    const slide = model.slides[2];
    (frameByText(slide, "Second point") as TextFrame).paragraphs.splice(1, 1);
    markSlideEdited(slide, "odp");
    expect(await savedOdpPage3(model)).not.toContain("<anim:");
  });

  it("a duplicated frame re-emits no id: the targets stay unique on the page", async () => {
    const model = await odpDeck();
    const slide = model.slides[2];
    duplicateFrame(model, 2, frameByText(slide, "First point").id);
    const page = await savedOdpPage3(model);
    expect(page.match(/xml:id="id3"/g)).toHaveLength(1);
    expect(page.match(/xml:id="id4"/g)).toHaveLength(1);
    expect(page).toContain(ODP_ANIM);
  });
});
