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
import { zip } from "../core/zipWriter";
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
      expect([...b.keys()].toSorted()).toEqual([...a.keys()].toSorted());
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

// PowerPoint's "Start: Automatically" lives in p:timing, not on the p:pic: a
// playFrom call in the main sequence's first group, which begins on the slide's
// onBegin. The reader sets `autoStart` from it (the timing stays preserved).
function powerPointAutoStartTiming(spid: string, firstGroupStart: string): string {
  return (
    `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
    `<p:par><p:cTn id="3" fill="hold"><p:stCondLst>${firstGroupStart}</p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">` +
    `<p:stCondLst><p:cond delay="500"/></p:stCondLst><p:childTnLst>` +
    `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="6" dur="26880" fill="hold"/>` +
    `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:cmd>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
    `<p:video><p:cMediaNode vol="80000"><p:cTn id="7" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>` +
    `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:video>` +
    `</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  );
}

/** Saves `model`, splices PowerPoint's timing for its video into slide 1, and reopens. */
async function reopenWithTiming(model: PresentationModel, firstGroupStart: string): Promise<PresentationModel> {
  const parts = await unzip((await writePptx(model)).bytes);
  const slideXml = new TextDecoder().decode(parts.get("ppt/slides/slide1.xml"));
  const videoPic = slideXml.split("<p:pic>").find((pic) => pic.includes("a:videoFile")) ?? "";
  const spid = videoPic.match(/<p:cNvPr id="(\d+)"/)?.[1] ?? "";
  expect(spid).not.toBe("");
  const timed = slideXml.replace("</p:sld>", `${powerPointAutoStartTiming(spid, firstGroupStart)}</p:sld>`);
  parts.set("ppt/slides/slide1.xml", new TextEncoder().encode(timed));
  const bytes = await zip([...parts].map(([name, data]) => ({ name, data })));
  return readPresentationBytes(bytes, "pptx", "Deck");
}

function videoFrame(model: PresentationModel): MediaFrame {
  return model.slides[0].frames.find((f): f is MediaFrame => f.kind === "media" && f.media === "video" && !!f.mediaId)!;
}

describe("pptx media auto-start from PowerPoint's timing (SPEC-095)", () => {
  const WITH_SLIDE = `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>`;
  const ON_CLICK = `<p:cond delay="indefinite"/>`;

  it("a clip PowerPoint starts with the slide reads as autoStart, and its timing stays preserved", async () => {
    const back = await reopenWithTiming(buildMediaModel("pptx"), WITH_SLIDE);
    expect(videoFrame(back).autoStart).toBe(true);
    expect(back.slides[0].preserved?.["p:timing"]).toContain("playFrom(0.0)");
  });

  it("a clip PowerPoint starts on a click stays on click", async () => {
    const back = await reopenWithTiming(buildMediaModel("pptx"), ON_CLICK);
    expect(videoFrame(back).autoStart).toBeUndefined();
  });

  it("docu-edit's own media flags win over the timing cue", async () => {
    const model = buildMediaModel("pptx");
    videoFrame(model).loop = true; // written as docu-edit's flags extension, autoStart absent
    const back = await reopenWithTiming(model, WITH_SLIDE);
    expect(videoFrame(back).loop).toBe(true);
    expect(videoFrame(back).autoStart).toBeUndefined();
  });
});
