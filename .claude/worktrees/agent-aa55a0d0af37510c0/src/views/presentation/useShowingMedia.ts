// Per-showing URLs for animated images in the show (SPEC-084 playback). Every
// surface renders a slide's images from the same `data:` URI, and Chrome keeps
// ONE playback timeline per image URL — so a non-looping GIF that already
// played (on the editor stage, in a filmstrip thumbnail, or on an earlier visit
// to its slide) comes up frozen on its last frame, and one still playing comes
// up part-way through. PowerPoint restarts a GIF each time its slide is shown;
// the show matches that by giving each showing of a slide fresh object URLs for
// its GIFs. The previous showing's URLs stay live for the transition's outgoing
// stage, so a GIF leaving under a fade carries on rather than restarting.
import { shallowRef } from "vue";
import { ensureMediaEncoded, type MediaPart } from "../../lib/core/media";
import { base64ToBytes } from "../../lib/core/packageParts";

/** Media id → this showing's URL, or undefined to fall back to the `data:` URI. */
export type MediaUrlOverride = (mediaId: string) => string | undefined;

interface Showing {
  url: MediaUrlOverride;
  release: () => void;
}

function createShowing(media: readonly MediaPart[]): Showing {
  const parts = new Map(media.map((part) => [part.mediaId, part]));
  const urls = new Map<string, string>();
  let released = false;
  return {
    // Minted on first render of each GIF, so only the slide's own GIFs are copied.
    url(mediaId) {
      const known = urls.get(mediaId);
      const part = parts.get(mediaId);
      if (known || released || part?.contentType !== "image/gif" || typeof URL.createObjectURL !== "function") {
        return known;
      }
      ensureMediaEncoded(part);
      const url = URL.createObjectURL(new Blob([base64ToBytes(part.data) as BlobPart], { type: part.contentType }));
      urls.set(mediaId, url);
      return url;
    },
    release() {
      released = true;
      for (const url of urls.values()) {
        URL.revokeObjectURL(url);
      }
      urls.clear();
    },
  };
}

export function useShowingMedia() {
  const current = shallowRef<Showing | null>(null);
  const previous = shallowRef<Showing | null>(null);

  /** A slide was entered: its GIFs start over on fresh URLs. The slide being
   * left keeps its URLs until the next entry, then they are revoked. */
  function enter(media: readonly MediaPart[]): void {
    previous.value?.release();
    previous.value = current.value;
    current.value = createShowing(media);
  }

  function releaseAll(): void {
    current.value?.release();
    previous.value?.release();
    current.value = null;
    previous.value = null;
  }

  return {
    /** The override for the slide on screen (and a transition's incoming stage). */
    currentUrl: (): MediaUrlOverride | undefined => current.value?.url,
    /** The override for a transition's outgoing stage. */
    previousUrl: (): MediaUrlOverride | undefined => previous.value?.url,
    enter,
    releaseAll,
  };
}
