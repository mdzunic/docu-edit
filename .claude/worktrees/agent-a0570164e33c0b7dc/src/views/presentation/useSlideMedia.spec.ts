// SPEC-095: the Media panel controller — flag edits (record → mutate → commit,
// each dropping the clip's verbatim sourceXml and dirtying the slide), Replace
// poster…, and the editor's play/stop activation resetting on slide change /
// deselect.
import { computed, ref, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { insertMediaFrame } from "../../lib/presentation/frames";
import type { MediaFrame, PresentationModel } from "../../lib/presentation/model";
import { useSlideMedia } from "./useSlideMedia";

function fileStub(bytes: Uint8Array, type: string): File {
  return { type, arrayBuffer: async () => bytes.buffer } as unknown as File;
}

function setup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const id = insertMediaFrame(model.value!, 0, { mediaId: "", contentType: "video/mp4", data: "AAAA" }, "video")!;
  // A read clip carries verbatim sourceXml; edits must drop it.
  const clip = model.value!.slides[0].frames.find((f) => f.id === id) as MediaFrame;
  clip.sourceXml = "<p:pic/>";
  const currentIndex = ref(0);
  const currentSlide = computed(() => model.value?.slides[currentIndex.value] ?? null);
  const selectedFrameId = ref<string | null>(id);
  const record = vi.fn();
  const commit = vi.fn();
  const bumpMedia = vi.fn();
  const reportError = vi.fn();
  const media = useSlideMedia({
    model,
    currentSlide,
    currentIndex,
    record,
    commit,
    bumpMedia,
    selectedFrameId,
    selectFrame: (v) => (selectedFrameId.value = v),
    reportError,
  });
  return { model, media, currentIndex, selectedFrameId, record, commit, bumpMedia, reportError, id, clip };
}

describe("useSlideMedia flags (SPEC-095 AC-22..31)", () => {
  it("exposes the selected clip and its default flag values", () => {
    const { media } = setup();
    expect(media.isMedia.value).toBe(true);
    expect(media.startAutomatically.value).toBe(false);
    expect(media.loop.value).toBe(false);
    expect(media.muted.value).toBe(false);
    expect(media.volume.value).toBe(100);
    expect(media.fullScreen.value).toBe(false);
  });

  it("sets start mode, recording one undo entry, dropping sourceXml, and committing", () => {
    const { media, model, record, commit, clip } = setup();
    media.setStartMode(true);
    expect(record).toHaveBeenCalledWith("Media start mode");
    expect(commit).toHaveBeenCalledTimes(1);
    const f = model.value!.slides[0].frames.find((fr) => fr.kind === "media") as MediaFrame;
    expect(f.autoStart).toBe(true);
    expect(clip.sourceXml).toBeUndefined();
    expect(model.value!.slides[0].dirty).toBe(true);
  });

  it("toggles loop, mute and full screen with the SPEC-080 voice", () => {
    const { media, model, record } = setup();
    media.setLoop(true);
    media.setMuted(true);
    media.setFullScreen(true);
    expect(record).toHaveBeenNthCalledWith(1, "Media loop");
    expect(record).toHaveBeenNthCalledWith(2, "Media mute");
    expect(record).toHaveBeenNthCalledWith(3, "Media full screen");
    const f = model.value!.slides[0].frames.find((fr) => fr.kind === "media") as MediaFrame;
    expect(f.loop).toBe(true);
    expect(f.muted).toBe(true);
    expect(f.fullScreen).toBe(true);
  });

  it("stores a non-default volume but drops it at 100 (absent = full)", () => {
    const { media, model, record } = setup();
    media.setVolume(40);
    let f = model.value!.slides[0].frames.find((fr) => fr.kind === "media") as MediaFrame;
    expect(f.volume).toBe(40);
    expect(record).toHaveBeenCalledWith("Media volume");
    media.setVolume(100);
    f = model.value!.slides[0].frames.find((fr) => fr.kind === "media") as MediaFrame;
    expect(f.volume).toBeUndefined();
  });

  it("replaces the poster, records it, and bumps media (AC-28/31)", async () => {
    const { media, model, record, bumpMedia } = setup();
    await media.replacePoster(fileStub(new Uint8Array([1, 2, 3]), "image/png"));
    expect(record).toHaveBeenCalledWith("Replace poster");
    expect(bumpMedia).toHaveBeenCalled();
    const f = model.value!.slides[0].frames.find((fr) => fr.kind === "media") as MediaFrame;
    expect(f.posterMediaId).toBeTruthy();
    expect(model.value!.media!.some((m) => m.mediaId === f.posterMediaId)).toBe(true);
  });

  it("rejects a non-image poster and changes nothing", async () => {
    const { media, model, reportError } = setup();
    await media.replacePoster(fileStub(new Uint8Array([1]), "video/mp4"));
    expect(reportError).toHaveBeenCalled();
    expect((model.value!.slides[0].frames.find((fr) => fr.kind === "media") as MediaFrame).posterMediaId).toBeUndefined();
  });

  it("open state rides the presentationMediaOpen preference", () => {
    const { media } = setup();
    const initial = media.open.value;
    media.toggle();
    expect(media.open.value).toBe(!initial);
    media.toggle();
    expect(media.open.value).toBe(initial);
  });
});
