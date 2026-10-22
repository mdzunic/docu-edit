// Format-neutral helpers for the slide media frame (SPEC-095), shared by the
// pptx and odp readers/writers and the view. Framework-free like every other
// presentation/ module: pure functions over the model, no Vue, no DOM. The
// per-format wire encoding (a:videoFile / draw:plugin, the playback-flag
// spellings) lives in pptx/odp; what both formats agree on lives here.
import type { MediaFrame, SlideMediaKind } from "./model";

/** The video/audio discriminator from a clip's content type — the reader's
 * fallback when the wire says "media" without naming which (an odp
 * `draw:plugin`, a pptx `p14:media` with no explicit `a:videoFile`). An
 * `audio/*` type is audio; anything else (including an absent type) is video,
 * the safer default for an on-stage box. */
export function mediaKindForContentType(contentType: string | undefined): SlideMediaKind {
  return contentType !== undefined && contentType.startsWith("audio/") ? "audio" : "video";
}

/** A clip's display basename — the last path segment of its embedded part id
 * or its linked target — for the poster caption and the link placeholder.
 * Handles both `/` and `\` separators (a linked Windows path). "" when the
 * frame carries neither an id nor a target. */
export function mediaBasename(frame: MediaFrame): string {
  const source = frame.linkTarget ?? frame.mediaId ?? "";
  const slash = Math.max(source.lastIndexOf("/"), source.lastIndexOf("\\"));
  return slash >= 0 ? source.slice(slash + 1) : source;
}

/** An `HTMLMediaElement.volume` (0.0–1.0) from the model's 0–100 (absent = 100
 * = full), clamped so a preserved out-of-range value never throws on the
 * element. */
export function mediaElementVolume(volume: number | undefined): number {
  const pct = volume ?? 100;
  return Math.min(1, Math.max(0, pct / 100));
}

/** Whether a clip has an embedded, playable payload (an in-model `mediaId`) —
 * false for a linked (`linkTarget`) or oversize clip we did not embed. The
 * stage renders those two as a link placeholder / oversize chip rather than a
 * player. */
export function mediaIsEmbedded(frame: MediaFrame): boolean {
  return frame.mediaId !== undefined && !frame.oversize;
}
