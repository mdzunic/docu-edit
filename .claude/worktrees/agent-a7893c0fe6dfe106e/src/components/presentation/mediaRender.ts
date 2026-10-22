// SPEC-095: the typed transforms behind the stage's media frame — a pure
// classifier that maps a MediaFrame + its resolved poster + a "can the browser
// play this?" probe into the render mode and caption the SlideFrameView shows.
// Kept DOM-free (the probe is injected) so it unit-tests without mounting.
import type { MediaFrame, SlideMediaKind } from "../../lib/presentation/model";
import { mediaBasename } from "../../lib/presentation/mediaFrames";
import { canPlayMedia } from "../../lib/core/mediaPlayback";

/** How the stage renders a media frame:
 *  - `link`        — a clip we did not embed (`media-link-placeholder`)
 *  - `oversize`    — in the package but over 25 MB (`media-oversize`)
 *  - `unsupported` — the browser cannot play this codec (`media-unsupported`)
 *  - `poster`      — a playable clip showing its poster still
 *  - `plate`       — a playable clip with no poster (neutral plate + glyph) */
export type MediaRenderMode = "link" | "oversize" | "unsupported" | "poster" | "plate";

export interface MediaRenderDescriptor {
  mode: MediaRenderMode;
  /** "Video" | "Audio", for the accessible name and the plate glyph. */
  noun: string;
  /** The clip's display basename (poster caption, link/unsupported caption). */
  basename: string;
  /** The poster still as a `data:` URI, or "" when there is none. */
  posterSrc: string;
  /** The caption a degraded frame shows; "" when the frame is playable. */
  caption: string;
  /** Whether the badge can swap in native `<video>`/`<audio controls>`. */
  playable: boolean;
}

/** A `canPlayMedia` probe bound to a real `<video>`/`<audio>` element's own
 * `canPlayType` (SPEC-095) — the browser's verdict for a clip of this kind.
 * DOM-touching, so it lives here (a `.ts` helper) rather than in the espree-lint
 * `.vue`; returns "cannot play" when there is no DOM (SSR/tests without jsdom). */
export function createCanPlayProbe(mediaKind: SlideMediaKind): (contentType: string | undefined) => boolean {
  return (contentType) => {
    if (typeof document === "undefined") {
      return false;
    }
    const el = document.createElement(mediaKind === "audio" ? "audio" : "video");
    return canPlayMedia(contentType, (type) => el.canPlayType(type));
  };
}

export const MEDIA_LINK_CAPTION = "Linked file — not embedded";
export const MEDIA_OVERSIZE_CAPTION = "Not loaded — over 25 MB";
export const MEDIA_UNSUPPORTED_CAPTION = "This browser can't play that format";
export const MEDIA_BLOCKED_CAPTION = "Click to play";

/** Classifies how a media frame renders on the stage (SPEC-095). `resolvePoster`
 * maps a `posterMediaId` to a `data:` URI (""-tolerant); `canPlay` is the
 * browser's own verdict for the clip's content type (injected — see
 * core/mediaPlayback.ts). A `runtimeError` (a fired `error` event) forces the
 * unsupported degrade even when `canPlay` was optimistic (AC-33). */
export function mediaRenderDescriptor(
  frame: MediaFrame,
  resolvePoster: (mediaId: string) => string,
  canPlay: (contentType: string | undefined) => boolean,
  runtimeError = false,
): MediaRenderDescriptor {
  const noun = frame.media === "audio" ? "Audio" : "Video";
  const basename = mediaBasename(frame);
  const posterSrc = frame.posterMediaId ? resolvePoster(frame.posterMediaId) : "";

  if (frame.linkTarget && !frame.mediaId) {
    return { mode: "link", noun, basename, posterSrc, caption: MEDIA_LINK_CAPTION, playable: false };
  }
  if (frame.oversize || !frame.mediaId) {
    return { mode: "oversize", noun, basename, posterSrc, caption: MEDIA_OVERSIZE_CAPTION, playable: false };
  }
  if (runtimeError || !canPlay(frame.contentType)) {
    return { mode: "unsupported", noun, basename, posterSrc, caption: MEDIA_UNSUPPORTED_CAPTION, playable: false };
  }
  return {
    mode: posterSrc ? "poster" : "plate",
    noun,
    basename,
    posterSrc,
    caption: "",
    playable: true,
  };
}
