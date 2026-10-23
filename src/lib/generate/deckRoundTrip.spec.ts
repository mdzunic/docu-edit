// SPEC-106 — generate → write → read: a deck built from every construct the
// API can emit is written by SPEC-080's writers (from scratch — a generated
// deck retains no package) and read back by its byte readers (readPptx/
// readOdp) to the same constructs under `equalStructural`, in both formats,
// geometry within ±0.01 cm. Generation is deterministic: the same deck
// generated twice with the same `modified` yields `equalBytes`-identical
// packages part for part (after `unzip`, never container bytes), and so
// does a second write of one deck. `equalSemantic` is typed against the text
// model and is not a relation claimed here.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { equalStructural, firstStructuralDifference } from "../equality";
import { buildJpegFixtureBytes } from "../media/fixtures/build";
import type { PresentationModel, SlideFrame, SlideParagraph, SlideRun } from "../presentation/model";
import { writeOdp } from "../presentation/odp/writer";
import { writePptx } from "../presentation/pptx/writer";
import { PresentationBuilder } from "./presentation";
import { readOdp, readPptx } from "./readDeck";
import { parseSlideMarkdown } from "./slideMarkdown";

const PNG_4X2 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";
// A 1×1 GIF, and the media product's real baseline JPEG fixture (passed as
// raw bytes, which the builder accepts beside base64) — every image type.
const GIF_1X1 = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
const JPEG = buildJpegFixtureBytes();
const MODIFIED = new Date(Date.UTC(2026, 8, 30, 9, 30, 0));

/** Every construct the generation API can emit. */
function everyConstruct(format: "pptx" | "odp", size: "16:9" | "4:3" = "16:9"): PresentationBuilder {
  return new PresentationBuilder({ title: "Every construct", format, size })
    .addTitleSlide(
      [{ text: "Q3 ", bold: true }, { text: "review", italic: true, underline: true }],
      [{ text: "Prepared 30 September", fontFamily: "Georgia", fontSizePt: 28, color: "#336699" }],
    )
    .setNotes(["Open with the headline.", "", "Pause."])
    .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%", ["Germany", "France"]], { text: "Deepest", level: 8 }, "Costs flat"])
    .addBullets(["Then numbered", ["and nested"]], { style: "number" })
    .addImageSlide("Pipeline", { data: PNG_4X2, contentType: "image/png" })
    .setBackground("#f0f0f0")
    .addBlankSlide()
    .addTextBox("A free text box", { x: { cm: 2 }, y: { cm: 2 }, width: { cm: 12 }, height: { cm: 3 } })
    .addTextBox([{ text: "Sized in points", fontSizePt: 14 }], { x: { pt: 400 }, y: { in: 3 }, width: { emu: 2700000 }, height: { pt: 60 } })
    .addImage({ data: GIF_1X1, contentType: "image/gif", x: { cm: 20 }, y: { cm: 2 }, width: { cm: 4 } })
    .addImage({ data: JPEG, contentType: "image/jpeg", x: { cm: 26 }, y: { cm: 10 }, width: { cm: 3 }, height: { cm: 2 } })
    .setBackground("#1f2937")
    .setNotes(["Blank slide notes."])
    .addTitleSlide("Thanks");
}

function emittedRun(run: SlideRun): Record<string, unknown> {
  const out: Record<string, unknown> = { text: run.text };
  for (const key of ["bold", "italic", "underline", "sizePt", "color", "linkHref"] as const) {
    if (run[key] !== undefined) out[key] = run[key];
  }
  if (typeof run.fontFamily === "string") out.fontFamily = run.fontFamily;
  return out;
}

/** The constructs the API emits, as plain data: a run's theme-reference
 * font (inherited from the written theme on read, never emitted) and the
 * pptx-only `phType`/`phIdx` wire attributes are left out, and frame
 * geometry is compared separately within ±0.01 cm. Image frames resolve to
 * their media part's type and bytes. */
function emitted(deck: PresentationModel): unknown {
  const media = new Map((deck.media ?? []).map((part) => [part.mediaId, part]));
  const paragraph = (p: SlideParagraph) => ({ runs: p.runs.map(emittedRun), bullet: p.bullet, level: p.level });
  const frame = (f: SlideFrame) => {
    if (f.kind === "text") return { kind: f.kind, role: f.role, paragraphs: f.paragraphs.map(paragraph) };
    if (f.kind === "image") {
      const part = media.get(f.mediaId);
      return { kind: f.kind, media: { contentType: part?.contentType, data: part?.data } };
    }
    return { kind: f.kind };
  };
  return {
    title: deck.title,
    format: deck.format,
    slides: deck.slides.map((slide) => ({ background: slide.background, notes: slide.notes, frames: slide.frames.map(frame) })),
  };
}

function expectGeometryWithin(read: PresentationModel, generated: PresentationModel): void {
  expect(Math.abs(read.slideWidthCm - generated.slideWidthCm)).toBeLessThanOrEqual(0.01);
  expect(Math.abs(read.slideHeightCm - generated.slideHeightCm)).toBeLessThanOrEqual(0.01);
  generated.slides.forEach((slide, s) => {
    expect(read.slides[s].frames).toHaveLength(slide.frames.length);
    slide.frames.forEach((frame, f) => {
      const back = read.slides[s].frames[f];
      for (const key of ["xCm", "yCm", "widthCm", "heightCm"] as const) {
        expect(Math.abs(back[key] - frame[key]), `slides[${s}].frames[${f}].${key}`).toBeLessThanOrEqual(0.01);
      }
    });
  });
}

async function readBack(bytes: Uint8Array, format: "pptx" | "odp", title: string): Promise<PresentationModel> {
  return format === "pptx" ? readPptx(bytes, title) : readOdp(bytes, title);
}

describe("a generated deck retains no package (AC-10, AC-11)", () => {
  it("has no opaque parts and no per-slide retention, and is plain JSON", () => {
    const deck = everyConstruct("pptx").toDeck();
    expect(deck.opaque).toBeUndefined();
    expect(deck.odpContentEnvelope).toBeUndefined();
    for (const slide of deck.slides) {
      expect([slide.partName, slide.sourceXml, slide.preserved]).toEqual([undefined, undefined, undefined]);
    }
    expect(JSON.parse(JSON.stringify(deck))).toEqual(deck);
  });
});

for (const format of ["pptx", "odp"] as const) {
  describe(`${format}: generate → write → read (AC-19)`, () => {
    it("reproduces every emitted construct under equalStructural, geometry within ±0.01 cm", async () => {
      const builder = everyConstruct(format);
      const generated = builder.toDeck();
      const read = await readBack(await builder.toBytes(format, { modified: MODIFIED }), format, generated.title);
      expect(firstStructuralDifference(emitted(read), emitted(generated))).toBeNull();
      expect(equalStructural(emitted(read), emitted(generated))).toBe(true);
      expectGeometryWithin(read, generated);
      expect(read.slides[0].notes).toEqual(["Open with the headline.", "", "Pause."]);
      expect(read.slides[2].background).toEqual({ type: "solid", color: "#f0f0f0" });
    });

    it("holds at 4:3 too", async () => {
      const builder = everyConstruct(format, "4:3");
      const generated = builder.toDeck();
      const read = await readBack(await builder.toBytes(), format, generated.title);
      expect(firstStructuralDifference(emitted(read), emitted(generated))).toBeNull();
      expectGeometryWithin(read, generated);
    });

    it("a deck parsed from slide Markdown round-trips the same way", async () => {
      const deck = parseSlideMarkdown(
        `---\ntitle: From Markdown\nformat: ${format}\n---\n# From **Markdown**\nA subtitle\n\n---\n\n## Rows\n- one\n  - _two_\n1. three\n\n![](data:image/png;base64,${PNG_4X2})\n\nNotes: Say it.`,
      );
      const bytes = format === "pptx" ? (await writePptx(deck)).bytes : (await writeOdp(deck)).bytes;
      const read = await readBack(bytes, format, deck.title);
      expect(firstStructuralDifference(emitted(read), emitted(deck))).toBeNull();
      expectGeometryWithin(read, deck);
    });

    it("is byte-deterministic: the same deck generated twice yields equalBytes parts", async () => {
      const first = await unzip(await everyConstruct(format).toBytes(format, { modified: MODIFIED }));
      const second = await unzip(await everyConstruct(format).toBytes(format, { modified: MODIFIED }));
      expect([...second.keys()]).toEqual([...first.keys()]);
      for (const [name, data] of first) {
        expect(equalBytes(second.get(name)!, data), `part ${name}`).toBe(true);
      }
    });

    it("a second write of one deck yields equalBytes parts", async () => {
      const builder = everyConstruct(format);
      const first = await unzip(await builder.toBytes(format, { modified: MODIFIED }));
      const second = await unzip(await builder.toBytes(format, { modified: MODIFIED }));
      for (const [name, data] of first) {
        expect(equalBytes(second.get(name)!, data), `part ${name}`).toBe(true);
      }
    });
  });
}

describe("cross-format writes take the same from-scratch path", () => {
  it("a pptx-born deck written as odp reads back the same constructs", async () => {
    const builder = everyConstruct("pptx");
    const generated = builder.toDeck();
    const read = await readOdp(await builder.toBytes("odp"), generated.title);
    expect(firstStructuralDifference(emitted(read), emitted({ ...generated, format: "odp" }))).toBeNull();
    expectGeometryWithin(read, generated);
  });
});
