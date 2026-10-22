// SPEC-095: the media fixtures open to the modelled clips (an embedded audio
// with a poster, an embedded video, a linked clip) plus one preserved
// out-of-subset "Media" ForeignFrame.
import { describe, expect, it } from "vitest";
import { readPresentationBytes } from "../open";
import { buildOdpMediaFixtureBytes, buildPptxMediaFixtureBytes, EXPECTED_MEDIA_FRAMES } from "./build";
import type { MediaFrame, PresentationModel } from "../model";

function assertFixture(model: PresentationModel): void {
  const media = model.slides[0].frames.filter((f): f is MediaFrame => f.kind === "media");
  expect(media).toHaveLength(EXPECTED_MEDIA_FRAMES.count);

  const audio = media.find((f) => f.media === "audio");
  expect(audio).toBeDefined();
  expect(audio?.mediaId).toBeTruthy();
  expect(audio?.contentType).toBe(EXPECTED_MEDIA_FRAMES.audioContentType);
  expect(audio?.posterMediaId).toBeTruthy();

  const embeddedVideo = media.find((f) => f.media === "video" && f.mediaId);
  expect(embeddedVideo).toBeDefined();
  expect(embeddedVideo?.contentType).toBe(EXPECTED_MEDIA_FRAMES.videoContentType);

  const linked = media.find((f) => f.linkTarget);
  expect(linked).toBeDefined();
  expect(linked?.mediaId).toBeUndefined();
  expect(linked?.linkTarget?.endsWith(EXPECTED_MEDIA_FRAMES.linkedBasename)).toBe(true);

  // The out-of-subset element stays a preserved "Media" ForeignFrame.
  const foreign = model.slides[0].frames.find((f) => f.kind === "foreign");
  expect(foreign?.kind === "foreign" && foreign.label).toBe("Media");
}

describe("SPEC-095 media fixtures", () => {
  it("pptx: opens to the modelled clips + a preserved out-of-subset media element", async () => {
    const model = await readPresentationBytes(await buildPptxMediaFixtureBytes(), "pptx", "Media");
    assertFixture(model);
  });

  it("odp: opens to the modelled clips + a preserved out-of-subset media element", async () => {
    const model = await readPresentationBytes(await buildOdpMediaFixtureBytes(), "odp", "Media");
    assertFixture(model);
  });
});
