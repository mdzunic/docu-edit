// SPEC-097 round trips over the spec097-text-depth fixtures, naming the
// relations: a deck given numbered and picture bullets with a start-at, 1.5×
// line spacing, 6pt space-before, 0.5pt letter spacing, shrink autofit, middle
// anchor, custom insets, two columns and both link kinds → save → reopen
// reproduces every modelled field, compared field-level with `toEqual` plus
// `equalStructural` (src/lib/equality.ts — its XML_STRING_KEYS does not list
// pptxPPr/pptxBodyPr/odpParaProps/odpBodyProps, so the four preservation slots
// compare RAW; `equalSemantic` is DocumentModel-typed and unusable here).
// Every part the edit did not touch is `equalBytes` after unzip, a no-edit
// download is byte-identical, writing the reopened model twice yields
// identical slide-part bytes, and pptx's fontScale/lnSpcReduction
// quantisation to 1000ths of a percent is the ONE sanctioned delta.
//
// odp: the three ODF details still going through the golden-fixture-first
// process (odp/textDepthGolden.ts — the shrink attribute, the numbering-restart
// attribute and the slide-jump link form) are omitted by the writer until a
// golden exists, so on odp an in-deck link and the shrink mode are pinned as
// NOT written (the fields round-trip on the pptx side); everything else is
// held to the same bar as pptx.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { equalStructural, firstStructuralDifference } from "../equality";
import { unzip } from "../core/zip";
import { buildOdpTextDepthFixtureBytes, buildPptxTextDepthFixtureBytes, EXPECTED_TEXT_DEPTH } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import type { PresentationModel, SlideFrame, SlideParagraph, TextFrame } from "./model";
import {
  applyParagraphFormat,
  markSlideEdited,
  setLetterSpacing,
  setLineSpacing,
  setListStyle,
  setParagraphSpacing,
  wholeFrameRange,
} from "./text";
import { applyTextBody } from "./textBody";
import { setFrameLink, setRunLink } from "./slideLinks";
import { paragraphText } from "./slideModel";

const FIXED_INSTANT = new Date(Date.UTC(2026, 9, 8, 12, 0, 0));

/** A 2×2 PNG for the picture bullet. */
const BULLET_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGMw7jhj3HGGAUIBACiuBh3+iAU4AAAAAElFTkSuQmCC";

async function openFixture(format: "pptx" | "odp"): Promise<{ model: PresentationModel; source: Map<string, Uint8Array> }> {
  const bytes = format === "pptx" ? await buildPptxTextDepthFixtureBytes() : await buildOdpTextDepthFixtureBytes();
  const model = await readPresentationBytes(bytes, format, "Text depth fixture");
  return { model, source: await unzip(bytes) };
}

async function writeBytes(model: PresentationModel): Promise<Uint8Array> {
  const options = { now: FIXED_INSTANT, modified: FIXED_INSTANT };
  return (model.format === "pptx" ? await writePptx(model, options) : await writeOdp(model, options)).bytes;
}

async function write(model: PresentationModel): Promise<Map<string, Uint8Array>> {
  return unzip(await writeBytes(model));
}

function textFrames(model: PresentationModel, slideIndex = 0): TextFrame[] {
  return model.slides[slideIndex].frames.filter((f): f is TextFrame => f.kind === "text");
}

/** The slide-1 text frame whose first paragraph reads `text`. */
function frameStartingWith(model: PresentationModel, text: string): TextFrame {
  const frame = textFrames(model).find((f) => f.paragraphs[0] && paragraphText(f.paragraphs[0]).startsWith(text));
  if (!frame) throw new Error(`no frame starting "${text}"`);
  return frame;
}

const r3 = (v: number): number => Math.round(v * 1000) / 1000;

/** Every paragraph of a frame, as a paragraph range. */
const all = (frame: TextFrame) => ({ startPara: 0, endPara: frame.paragraphs.length - 1 });

/** The modelled subset this spec promises to reproduce, per text frame: the
 * paragraphs (every list/spacing/preservation field and every run mark),
 * the body, the frame link and the box — never ids. */
function projection(frame: SlideFrame): Record<string, unknown> {
  const out: Record<string, unknown> = {
    kind: frame.kind,
    box: [r3(frame.xCm), r3(frame.yCm), r3(frame.widthCm), r3(frame.heightCm)],
  };
  for (const key of ["linkHref", "linkSlide", "linkTooltip"] as const) {
    if (frame[key] !== undefined) out[key] = frame[key];
  }
  if (frame.kind === "text" || frame.kind === "shape") {
    out.paragraphs = (frame.paragraphs ?? []).map((para: SlideParagraph) => ({ ...para }));
    if (frame.body) out.body = frame.body;
    if (frame.pptxBodyPr !== undefined) out.pptxBodyPr = frame.pptxBodyPr;
    if (frame.odpBodyProps !== undefined) out.odpBodyProps = frame.odpBodyProps;
  }
  return out;
}

/** Gives slide 1 every SPEC-097 property the round-trip AC names. */
function editSlide1(model: PresentationModel): void {
  const format = model.format;
  const body = frameStartingWith(model, EXPECTED_TEXT_DEPTH.bodyLines[0]);
  const continued = frameStartingWith(model, EXPECTED_TEXT_DEPTH.continuedLines[0]);
  const sentence = frameStartingWith(model, EXPECTED_TEXT_DEPTH.linkSentence);
  const title = frameStartingWith(model, EXPECTED_TEXT_DEPTH.title);

  // Numbered body with a start-at, 1.5× line spacing and 6pt space-before.
  applyParagraphFormat(body, all(body), { startAt: 3 });
  setLineSpacing(body, all(body), { mode: "multiple", value: 1.5 });
  setParagraphSpacing(body, all(body), { spaceBeforePt: 6 });
  // 0.5pt letter spacing on the body's first line.
  setLetterSpacing(body, { start: { para: 0, offset: 0 }, end: { para: 0, offset: paragraphText(body.paragraphs[0]).length } }, 0.5);
  // Two columns, middle anchor and custom insets on the body.
  applyTextBody(body, { columns: { count: 2, gapCm: 1 }, anchor: "middle", insetsCm: { left: 0.5, top: 0.3, right: 0.5, bottom: 0.3 } });

  // Picture bullets on the continued list (a media part minted by SPEC-082's
  // naming rule's shape — the next free image{n}).
  const mediaId = format === "pptx" ? "ppt/media/image97.png" : "Pictures/image97.png";
  model.media = [...(model.media ?? []), { mediaId, contentType: "image/png", data: BULLET_PNG_BASE64 }];
  setListStyle(continued, all(continued), { bulletImageId: mediaId });

  // Shrink autofit on the overflowing title (a ladder step: 62.5 / 90).
  applyTextBody(title, { autofit: "shrink", fontScalePct: 62.5, lineScalePct: 90 });

  // Both link kinds: "results" → slide 3, "detail" → an external URL with a
  // tooltip, and a whole-frame external link on the sentence's box.
  const sentenceText = paragraphText(sentence.paragraphs[0]);
  const at = sentenceText.indexOf(EXPECTED_TEXT_DEPTH.linkWord);
  setRunLink(sentence, { start: { para: 0, offset: at }, end: { para: 0, offset: at + EXPECTED_TEXT_DEPTH.linkWord.length } }, { slide: 3 });
  const detail = sentenceText.indexOf("detail");
  setRunLink(sentence, { start: { para: 0, offset: detail }, end: { para: 0, offset: detail + 6 } }, { href: "https://example.com/detail", tooltip: "More" });
  setFrameLink(sentence, { href: "https://example.com/" });

  markSlideEdited(model.slides[0], format);
}

/** What the odp writer deliberately omits until its golden exists: the
 * slide-jump link form and the shrink attribute (odp/textDepthGolden.ts). */
function withoutOdpGoldenPending(frame: Record<string, unknown>): Record<string, unknown> {
  const out = structuredClone(frame);
  const body = out.body as Record<string, unknown> | undefined;
  if (body && body.autofit === "shrink") {
    delete body.autofit;
    delete body.fontScalePct;
    delete body.lineScalePct;
    if (Object.keys(body).length === 0) delete out.body;
  }
  for (const para of (out.paragraphs as SlideParagraph[] | undefined) ?? []) {
    const runs: SlideParagraph["runs"] = [];
    for (const run of para.runs) {
      const next = { ...run };
      if (next.linkSlide !== undefined) {
        delete next.linkSlide;
        delete next.linkTooltip;
      }
      const prev = runs[runs.length - 1];
      const { text: _a, ...prevMarks } = prev ?? { text: "" };
      const { text: _b, ...marks } = next;
      if (prev && JSON.stringify(prevMarks) === JSON.stringify(marks)) prev.text += next.text;
      else runs.push(next);
    }
    para.runs = runs;
  }
  if (out.linkSlide !== undefined) {
    delete out.linkSlide;
    delete out.linkTooltip;
  }
  return out;
}

for (const format of ["pptx", "odp"] as const) {
  describe(`spec097-text-depth ${format}: the fixture reads`, () => {
    it("derives the numbered body, the start-at list and keeps the vertical frame's source", async () => {
      const { model } = await openFixture(format);
      expect(model.slides).toHaveLength(EXPECTED_TEXT_DEPTH.slideCount);
      const body = frameStartingWith(model, EXPECTED_TEXT_DEPTH.bodyLines[0]);
      expect(body.paragraphs.map((p) => p.bullet)).toEqual(["number", "number", "number"]);
      const continued = frameStartingWith(model, EXPECTED_TEXT_DEPTH.continuedLines[0]);
      expect(continued.paragraphs.map((p) => p.startAt)).toEqual([5, 5]);
      const vertical = frameStartingWith(model, EXPECTED_TEXT_DEPTH.verticalText);
      if (format === "pptx") {
        expect(vertical.pptxBodyPr).toContain('vert="vert270"');
        expect(vertical.body?.anchor).toBe("middle");
        expect(vertical.paragraphs[0].pptxPPr).toContain("a:tabLst");
        expect(body.paragraphs[0].pptxPPr).toContain("a:buFont");
        expect(body.paragraphs[0].hangingCm).toBe(EXPECTED_TEXT_DEPTH.bodyHangingCm);
      } else {
        expect(vertical.odpBodyProps).toContain('style:writing-mode="tb-rl"');
        expect(vertical.paragraphs[0].odpParaProps).toContain('style:punctuation-wrap="hanging"');
      }
    });
  });

  describe(`spec097-text-depth ${format}: no-edit round trip`, () => {
    it("a no-edit download stays byte-identical part-for-part under equalBytes", async () => {
      const { model, source } = await openFixture(format);
      const written = await write(model);
      expect([...written.keys()]).toEqual([...source.keys()]);
      for (const [name, data] of source) {
        expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
      }
    });
  });

  describe(`spec097-text-depth ${format}: lists + spacing + autofit + links → save → reopen`, () => {
    async function editedAndReopened(): Promise<{ edited: PresentationModel; reopened: PresentationModel; source: Map<string, Uint8Array>; written: Map<string, Uint8Array> }> {
      const { model, source } = await openFixture(format);
      editSlide1(model);
      const bytes = await writeBytes(model);
      const reopened = await readPresentationBytes(bytes, format, "Text depth fixture");
      return { edited: model, reopened, source, written: await unzip(bytes) };
    }

    it("reproduces every modelled field (toEqual + equalStructural, preservation slots raw)", async () => {
      const { edited, reopened } = await editedAndReopened();
      let expected = edited.slides[0].frames.map(projection);
      const actual = reopened.slides[0].frames.map(projection);
      if (format === "odp") expected = expected.map(withoutOdpGoldenPending);
      expect(actual).toEqual(expected);
      expect(firstStructuralDifference(expected, actual)).toBeNull();
      expect(equalStructural(expected, actual)).toBe(true);
      // The untouched slides 2 and 3 are unchanged too.
      for (const index of [1, 2]) {
        expect(reopened.slides[index].frames.map(projection)).toEqual(edited.slides[index].frames.map(projection));
      }
    });

    it("compares the preservation slots raw, not canonically", async () => {
      const { reopened } = await editedAndReopened();
      const vertical = frameStartingWith(reopened, EXPECTED_TEXT_DEPTH.verticalText);
      const clone = structuredClone(vertical);
      const slot = format === "pptx" ? "pptxBodyPr" : "odpBodyProps";
      const value = clone[slot];
      expect(value).toBeDefined();
      // Re-quote one attribute: canonically equal XML, a different raw string.
      clone[slot] = value!.replace(/"/, "'").replace(/"(\s|\/|>|$)/, "'$1");
      expect(clone[slot]).not.toBe(value);
      expect(equalStructural(projection(vertical), projection(clone))).toBe(false);
    });

    it("keeps every part the edit did not touch equalBytes", async () => {
      const { source, written } = await editedAndReopened();
      const touched =
        format === "pptx"
          ? new Set(["ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels", "[Content_Types].xml", "docProps/core.xml", "docProps/app.xml"])
          : new Set(["content.xml", "META-INF/manifest.xml", "meta.xml"]);
      for (const [name, data] of source) {
        if (touched.has(name)) continue;
        expect(written.has(name), `part ${name} must survive`).toBe(true);
        expect(equalBytes(written.get(name)!, data), `untouched part ${name} must stay byte-verbatim`).toBe(true);
      }
      // The picture bullet's media part was written alongside.
      expect(written.has(format === "pptx" ? "ppt/media/image97.png" : "Pictures/image97.png")).toBe(true);
    });

    it("writes the reopened model twice with identical slide-part bytes", async () => {
      const { reopened } = await editedAndReopened();
      markSlideEdited(reopened.slides[0], format);
      const first = await write(reopened);
      const second = await write(structuredClone(reopened));
      const slideParts = [...first.keys()].filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name) || name === "content.xml");
      expect(slideParts.length).toBeGreaterThan(0);
      for (const name of slideParts) {
        expect(equalBytes(first.get(name)!, second.get(name)!), `part ${name} must be write-stable`).toBe(true);
      }
    });
  });
}

describe("spec097-text-depth pptx: the one sanctioned delta", () => {
  it("quantises fontScale/lnSpcReduction to 1000ths of a percent and nothing else moves", async () => {
    const { model } = await openFixture("pptx");
    const title = frameStartingWith(model, EXPECTED_TEXT_DEPTH.title);
    applyTextBody(title, { autofit: "shrink" });
    // Bypass the command's own quantisation: a measured scale with more
    // precision than the wire carries.
    title.body = { ...title.body, autofit: "shrink", fontScalePct: 62.45678, lineScalePct: 91.23456 };
    markSlideEdited(model.slides[0], "pptx");
    const written = await write(model);
    const slideXml = new TextDecoder().decode(written.get("ppt/slides/slide1.xml")!);
    expect(slideXml).toContain('fontScale="62457"');
    expect(slideXml).toContain('lnSpcReduction="8765"');
    const reopened = await readPresentationBytes(await writeBytes(model), "pptx", "Text depth fixture");
    const back = frameStartingWith(reopened, EXPECTED_TEXT_DEPTH.title);
    expect(back.body).toEqual({ autofit: "shrink", fontScalePct: 62.457, lineScalePct: 91.235 });
    expect(back.paragraphs).toEqual(title.paragraphs);
  });
});

describe("spec097-text-depth: whole-frame range helper", () => {
  it("covers every paragraph of the numbered body", async () => {
    const { model } = await openFixture("pptx");
    const body = frameStartingWith(model, EXPECTED_TEXT_DEPTH.bodyLines[0]);
    expect(wholeFrameRange(body).end.para).toBe(EXPECTED_TEXT_DEPTH.bodyLines.length - 1);
  });
});
