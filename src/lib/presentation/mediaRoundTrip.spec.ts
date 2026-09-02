// SPEC-095 round trip: a deck carrying media (an embedded playable clip with a
// poster, an embedded unsupported clip, and a linked clip) opens, saves and
// reopens with every part byte-identical; a slide dirtied elsewhere re-emits
// each untouched media frame's verbatim sourceXml; insert → save → reopen
// matches the modelled media fields; and the model survives a JSON round trip.
import { describe, it, expect } from "vitest";
import { createBlankPresentation } from "./slides";
import { insertMediaFrame, insertTextBoxFrame, setMediaPoster } from "./frames";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { unzip } from "../core/zip";
import { equalBytes } from "../core/equality";
import { equalStructural } from "../equality";
import type { MediaFrame, PresentationFormat, PresentationModel, SlideFrame } from "./model";

const PNG_B64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGMw7jhj3HGGAUIBACiuBh3+iAU4AAAAAElFTkSuQmCC";

/** A slide carrying: an embedded audio clip (with a poster), an embedded video
 * clip, and a linked video clip. From-scratch, so the writers rebuild it. */
function buildMediaModel(format: PresentationFormat): PresentationModel {
  const model = createBlankPresentation();
  model.format = format;
  model.slides[0].frames = [];
  const audioId = insertMediaFrame(model, 0, { mediaId: "", contentType: "audio/wav", data: "QUJDRA==" }, "audio")!;
  insertMediaFrame(model, 0, { mediaId: "", contentType: "video/mp4", data: "RUZHSElK" }, "video");
  setMediaPoster(model, 0, audioId, { mediaId: "", contentType: "image/png", data: PNG_B64 });
  model.slides[0].frames.push({
    id: crypto.randomUUID(),
    kind: "media",
    media: "video",
    xCm: 20,
    yCm: 2,
    widthCm: 6,
    heightCm: 4,
    linkTarget: "../media/remote.mp4",
    contentType: "video/mp4",
  });
  return model;
}

function write(model: PresentationModel): Promise<Uint8Array> {
  return (model.format === "odp" ? writeOdp(model) : writePptx(model)).then((r) => r.bytes);
}

function r3(v: number): number {
  return Math.round(v * 1000) / 1000;
}

/** The modelled media fields a round trip must reproduce — never id or the
 * verbatim sourceXml (byte-exact, not part of the model comparison). */
function projectMedia(frame: SlideFrame): Record<string, unknown> {
  const out: Record<string, unknown> = {
    kind: frame.kind,
    box: [r3(frame.xCm), r3(frame.yCm), r3(frame.widthCm), r3(frame.heightCm)],
  };
  if (frame.kind === "media") {
    out.media = frame.media;
    if (frame.mediaId) out.mediaId = frame.mediaId;
    if (frame.posterMediaId) out.posterMediaId = frame.posterMediaId;
    if (frame.linkTarget) out.linkTarget = frame.linkTarget;
    if (frame.contentType) out.contentType = frame.contentType;
    if (frame.autoStart) out.autoStart = frame.autoStart;
    if (frame.loop) out.loop = frame.loop;
    if (frame.muted) out.muted = frame.muted;
    if (frame.volume !== undefined) out.volume = frame.volume;
    if (frame.fullScreen) out.fullScreen = frame.fullScreen;
    if (frame.oversize) out.oversize = frame.oversize;
  }
  return out;
}

for (const format of ["pptx", "odp"] as const) {
  describe(`${format} media round trip (SPEC-095)`, () => {
    it("re-emits every part byte-identical on a no-edit reopen→save (AC-43)", async () => {
      const bytesA = await write(buildMediaModel(format));
      const reopened = await readPresentationBytes(bytesA, format);
      const bytesB = await write(reopened);
      const a = await unzip(bytesA);
      const b = await unzip(bytesB);
      expect([...b.keys()].sort()).toEqual([...a.keys()].sort());
      for (const [name, data] of a) {
        expect(equalBytes(b.get(name)!, data), `part ${name} must round-trip byte-identical`).toBe(true);
      }
    });

    it("re-emits each untouched media frame's verbatim sourceXml when the slide is dirtied elsewhere (AC-44)", async () => {
      const bytesA = await write(buildMediaModel(format));
      const reopened = await readPresentationBytes(bytesA, format);
      const sources = reopened.slides[0].frames
        .filter((f): f is MediaFrame => f.kind === "media")
        .map((f) => f.sourceXml!);
      expect(sources.length).toBe(3);
      // Dirty the slide by inserting an unrelated frame; the media frames stay
      // untouched and must re-emit verbatim.
      insertTextBoxFrame(reopened, 0, { xCm: 1, yCm: 1, widthCm: 5, heightCm: 2 });
      const bytesB = await write(reopened);
      const slidePart = format === "odp" ? "content.xml" : "ppt/slides/slide1.xml";
      const xml = new TextDecoder().decode((await unzip(bytesB)).get(slidePart)!);
      for (const src of sources) {
        expect(xml).toContain(src);
      }
    });

    it("matches the modelled media fields under equalStructural on insert→save→reopen (AC-45)", async () => {
      const model = buildMediaModel(format);
      const inserted = model.slides[0].frames.map(projectMedia);
      const reopened = await readPresentationBytes(await write(model), format);
      const read = reopened.slides[0].frames.map(projectMedia);
      expect(equalStructural(inserted, read)).toBe(true);
    });

    it("survives a JSON round trip (AC-46)", () => {
      const model = buildMediaModel(format);
      expect(JSON.parse(JSON.stringify(model))).toEqual(model);
    });
  });
}
